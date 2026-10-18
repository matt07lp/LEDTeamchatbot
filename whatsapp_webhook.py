"""
whatsapp_webhook.py - WhatsApp Cloud API webhook for the LED sales/rental agent.

Flow:
  WhatsApp -> POST /webhook (signature checked) -> 200 immediately
           -> background task: agent (LangChain + MCP tools) -> reply via Graph API

Run locally:   uvicorn whatsapp_webhook:app --port 8000
Test locally:  python simular_whatsapp.py        (needs WHATSAPP_DRY_RUN=1)
Unit tests:    pytest test_webhook.py -v

Environment variables (.env):
  WHATSAPP_VERIFY_TOKEN     string you choose; paste the same one in Meta's dashboard
  WHATSAPP_APP_SECRET       Meta app secret (App settings > Basic); signs every webhook
  WHATSAPP_ACCESS_TOKEN     token used to SEND messages
  WHATSAPP_PHONE_NUMBER_ID  the sender's phone number ID
  GRAPH_API_VERSION         default v23.0 (check Meta's dashboard for the current one)
  WHATSAPP_DRY_RUN          "1" = do not call Meta; keep replies in memory (for testing)
  WHATSAPP_ALLOWED_NUMBERS  optional comma-separated allowlist (pilot phase)
  MAX_HISTORY               max messages kept per customer (default 60)
  CONVERSATIONS_FILE        optional path to persist histories (experimental; off by default)
"""
import asyncio
import hashlib
import hmac
import importlib
import json
import logging
import os
from collections import OrderedDict
from contextlib import asynccontextmanager
from datetime import date
from pathlib import Path

import httpx
from fastapi import BackgroundTasks, FastAPI, HTTPException, Query, Request
from fastapi.responses import PlainTextResponse

try:  # load .env when running with uvicorn
    from dotenv import load_dotenv
    load_dotenv(Path(__file__).with_name(".env"))
except ImportError:
    pass

from trazas import _texto  # extracts plain text from a model message

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
logger = logging.getLogger("whatsapp")

MSG_NO_TEXTO = "Por ahora solo puedo leer mensajes de texto. ¿Me escribes tu consulta?"
MSG_ERROR = "Disculpa, tuve un problema técnico. ¿Puedes intentarlo de nuevo en un momento?"
MSG_VACIO = "Disculpa, no pude procesar tu mensaje. ¿Puedes repetirlo?"


# ---------------------------------------------------------------------------
# AGENT (built once; rebuilt each new day so the system prompt has today's date)
# ---------------------------------------------------------------------------
async def _crear_agente():
    from langchain.agents import create_agent
    from langchain_mcp_adapters.client import MultiServerMCPClient
    from config import PYTHON, llm_langchain

    # Reuse the prompt and server list from 07_mi_agente.py (reload = fresh date)
    mod = importlib.reload(importlib.import_module("07_mi_agente"))
    cliente = MultiServerMCPClient({
        nombre: {"command": PYTHON, "args": [ruta], "transport": "stdio"}
        for nombre, ruta in mod.SERVIDORES.items()
    })
    return create_agent(model=llm_langchain(), tools=await cliente.get_tools(),
                        system_prompt=mod.SYSTEM_PROMPT)


async def _obtener_agente():
    st = app.state
    if st.agente_fijo:  # injected agent (used by tests)
        return st.agente
    async with _build_lock:
        if st.agente is None or st.agente_fecha != date.today():
            st.agente = await _crear_agente()
            st.agente_fecha = date.today()
    return st.agente


@asynccontextmanager
async def lifespan(app: FastAPI):
    if app.state.agente is None:
        try:
            await _obtener_agente()  # warm up: fail early and visibly
            logger.info("Agent ready")
        except Exception:
            logger.exception("Could not build the agent at startup (will retry on first message)")
    yield


app = FastAPI(title="LED WhatsApp webhook", lifespan=lifespan)
app.state.agente = None
app.state.agente_fecha = None
app.state.agente_fijo = False
app.state.outbox = []          # replies kept here when WHATSAPP_DRY_RUN=1
_build_lock = asyncio.Lock()


# ---------------------------------------------------------------------------
# PER-CUSTOMER MEMORY (keyed by the customer's WhatsApp number)
# ---------------------------------------------------------------------------
def _tipo(m) -> str:
    t = getattr(m, "type", None)
    if t:
        return t
    return "human" if isinstance(m, dict) and m.get("role") == "user" else "other"


def _recortar(mensajes: list, maximo: int | None = None) -> list:
    """Trim old history, always starting at a customer message so tool-call
    pairs are never cut in half."""
    maximo = maximo or int(os.getenv("MAX_HISTORY", "60"))
    if len(mensajes) <= maximo:
        return mensajes
    for i in range(len(mensajes) - maximo, len(mensajes)):
        if _tipo(mensajes[i]) == "human":
            return mensajes[i:]
    return mensajes


class Conversaciones:
    def __init__(self):
        self.datos: dict[str, list] = {}
        self._cargado = False

    def _ruta(self):
        ruta = os.getenv("CONVERSATIONS_FILE", "")
        return Path(ruta) if ruta else None

    def _cargar(self):
        if self._cargado:
            return
        self._cargado = True
        ruta = self._ruta()
        if ruta and ruta.exists():
            try:
                from langchain_core.messages import messages_from_dict
                crudo = json.loads(ruta.read_text())
                self.datos = {u: messages_from_dict(ms) for u, ms in crudo.items()}
            except Exception:
                logger.exception("Could not load saved conversations; starting empty")

    def get(self, usuario: str) -> list:
        self._cargar()
        return list(self.datos.get(usuario, []))

    def set(self, usuario: str, mensajes: list) -> None:
        self.datos[usuario] = mensajes
        ruta = self._ruta()
        if ruta:
            try:
                from langchain_core.messages import messages_to_dict
                ruta.write_text(json.dumps(
                    {u: messages_to_dict(ms) for u, ms in self.datos.items()}, default=str))
            except Exception:
                logger.exception("Could not save conversations")


store = Conversaciones()
_locks: dict[str, asyncio.Lock] = {}
_vistos: OrderedDict[str, bool] = OrderedDict()  # message ids already handled (Meta retries)


def _lock_de(usuario: str) -> asyncio.Lock:
    return _locks.setdefault(usuario, asyncio.Lock())


def _ya_visto(msg_id: str) -> bool:
    if msg_id in _vistos:
        return True
    _vistos[msg_id] = True
    while len(_vistos) > 2000:
        _vistos.popitem(last=False)
    return False


def _permitido(usuario: str) -> bool:
    permitidos = {n.strip() for n in os.getenv("WHATSAPP_ALLOWED_NUMBERS", "").split(",") if n.strip()}
    return not permitidos or usuario in permitidos


# ---------------------------------------------------------------------------
# SECURITY + PARSING
# ---------------------------------------------------------------------------
def _firma_valida(cuerpo: bytes, cabecera: str) -> bool:
    """Meta signs the raw body with the app secret: X-Hub-Signature-256: sha256=<hex>."""
    secreto = os.getenv("WHATSAPP_APP_SECRET", "")
    if not secreto or not cabecera.startswith("sha256="):
        return False
    esperada = hmac.new(secreto.encode(), cuerpo, hashlib.sha256).hexdigest()
    return hmac.compare_digest(esperada.encode(), cabecera[len("sha256="):].encode())


def _extraer_mensajes(datos: dict) -> list[dict]:
    """Return [{id, from, type, text}] from a webhook payload. Delivery/read
    receipts ('statuses') and anything without messages are ignored."""
    salida = []
    for entrada in datos.get("entry", []):
        for cambio in entrada.get("changes", []):
            for m in cambio.get("value", {}).get("messages", []):
                salida.append({
                    "id": m.get("id", ""),
                    "from": m.get("from", ""),
                    "type": m.get("type", ""),
                    "text": (m.get("text") or {}).get("body", ""),
                })
    return salida


# ---------------------------------------------------------------------------
# SENDING
# ---------------------------------------------------------------------------
async def enviar_texto(destino: str, texto: str) -> None:
    texto = texto if len(texto) <= 4000 else texto[:3990] + "…"  # WhatsApp limit is 4096
    if os.getenv("WHATSAPP_DRY_RUN") == "1":
        app.state.outbox.append({"to": destino, "text": texto})
        logger.info("[DRY RUN] -> %s: %s", destino, texto)
        return
    url = (f"https://graph.facebook.com/{os.getenv('GRAPH_API_VERSION', 'v23.0')}/"
           f"{os.getenv('WHATSAPP_PHONE_NUMBER_ID', '')}/messages")
    cuerpo = {"messaging_product": "whatsapp", "to": destino, "type": "text",
              "text": {"body": texto}}
    try:
        async with httpx.AsyncClient(timeout=20) as http:
            r = await http.post(url, json=cuerpo, headers={
                "Authorization": f"Bearer {os.getenv('WHATSAPP_ACCESS_TOKEN', '')}"})
            if r.status_code >= 400:
                logger.error("Graph API error %s: %s", r.status_code, r.text[:300])
    except Exception:
        logger.exception("Could not send WhatsApp message to %s", destino)


# ---------------------------------------------------------------------------
# MESSAGE PROCESSING (runs in the background after we answered 200 to Meta)
# ---------------------------------------------------------------------------
async def procesar_mensaje(msg: dict) -> None:
    if not msg["id"] or _ya_visto(msg["id"]):
        return
    usuario = msg["from"]
    if not _permitido(usuario):
        logger.info("Ignoring message from non-allowed number %s", usuario)
        return
    if msg["type"] != "text" or not msg["text"].strip():
        await enviar_texto(usuario, MSG_NO_TEXTO)
        return

    async with _lock_de(usuario):  # one message at a time per customer
        try:
            agente = await _obtener_agente()
            historial = store.get(usuario)
            r = await agente.ainvoke(
                {"messages": historial + [{"role": "user", "content": msg["text"]}]})
            mensajes = r["messages"]
            store.set(usuario, _recortar(mensajes))
            respuesta = _texto(mensajes[-1].content).strip() or MSG_VACIO
        except Exception:
            logger.exception("Agent failed for %s", usuario)
            respuesta = MSG_ERROR
    await enviar_texto(usuario, respuesta)


# ---------------------------------------------------------------------------
# ENDPOINTS
# ---------------------------------------------------------------------------
@app.get("/webhook")
async def verificar(modo: str = Query("", alias="hub.mode"),
                    token: str = Query("", alias="hub.verify_token"),
                    challenge: str = Query("", alias="hub.challenge")):
    """Meta calls this once when you register the webhook."""
    esperado = os.getenv("WHATSAPP_VERIFY_TOKEN", "")
    if modo == "subscribe" and token and esperado and \
            hmac.compare_digest(token.encode(), esperado.encode()):
        return PlainTextResponse(challenge)
    raise HTTPException(status_code=403, detail="Verification failed")


@app.post("/webhook")
async def recibir(request: Request, background: BackgroundTasks):
    cuerpo = await request.body()
    if not _firma_valida(cuerpo, request.headers.get("X-Hub-Signature-256", "")):
        raise HTTPException(status_code=403, detail="Invalid signature")
    try:
        datos = json.loads(cuerpo)
    except ValueError:
        raise HTTPException(status_code=400, detail="Invalid JSON")
    for msg in _extraer_mensajes(datos):
        background.add_task(procesar_mensaje, msg)
    return {"status": "ok"}  # answer fast; Meta expects a 200 within seconds


@app.get("/health")
async def health():
    return {"status": "ok"}


@app.get("/debug/outbox")
async def debug_outbox(to: str = ""):
    """Only available in dry-run mode: lets the simulator read the bot's replies."""
    if os.getenv("WHATSAPP_DRY_RUN") != "1":
        raise HTTPException(status_code=404)
    return [m for m in app.state.outbox if not to or m["to"] == to]


if __name__ == "__main__":
    import uvicorn
    uvicorn.run("whatsapp_webhook:app", host="0.0.0.0", port=int(os.getenv("PORT", "8000")))
