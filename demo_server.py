"""
demo_server.py - Web demo of the LED sales/rental assistant.

It EXTENDS the app in whatsapp_webhook.py, so the demo uses exactly the same agent,
MCP tools and Lead Analyst. WhatsApp can be connected later without changes here.

Run:   uvicorn demo_server:app --port 8000
Open:  http://localhost:8000

Endpoints:
  GET  /                      the demo page (demo.html)
  POST /api/chat              {session_id, message} -> reply + tool activity
  GET  /api/score/{session}   lead analysis state: none | pending | ready | failed
  POST /api/reset             {session_id} -> forget the conversation

Keep WHATSAPP_DRY_RUN=1 in .env for the demo so HOT-lead alerts are not sent to Meta.
This is a DEMO server: no login and no rate limiting. Do not expose it to the internet.
"""
import asyncio
import logging
import re
from pathlib import Path

from fastapi import BackgroundTasks, HTTPException
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

import whatsapp_webhook as wh
from whatsapp_webhook import app
from lead_analyst import leer_resultado
from trazas import traza_para_web

logger = logging.getLogger("demo")
SESSION_RE = re.compile(r"^[A-Za-z0-9_-]{8,64}$")
HTML = Path(__file__).with_name("demo.html")

# session_id -> {"lead_id": str | None, "pendiente": bool, "intentado": bool}
app.state.web_sesiones = {}


class ChatIn(BaseModel):
    session_id: str
    message: str = Field(min_length=1, max_length=2000)


class ResetIn(BaseModel):
    session_id: str


def _validar(session_id: str) -> str:
    if not SESSION_RE.match(session_id or ""):
        raise HTTPException(status_code=400, detail="Invalid session_id")
    return session_id


def _usuario(sid: str) -> str:
    return f"web:{sid}"


@app.get("/")
async def pagina():
    if not HTML.exists():
        raise HTTPException(status_code=404, detail="demo.html not found next to demo_server.py")
    return FileResponse(HTML, media_type="text/html", headers={"Cache-Control": "no-store"})


async def _analizar_sesion(sid: str, lead_id: str, mensajes: list) -> None:
    """Agent 2 in the background; always clears the 'pending' flag when done."""
    try:
        await wh._analizar_y_alertar(lead_id, mensajes)
    finally:
        sesion = app.state.web_sesiones.get(sid)
        if sesion:
            sesion["pendiente"] = False


@app.post("/api/chat")
async def chat(body: ChatIn, background: BackgroundTasks):
    sid = _validar(body.session_id)
    texto = body.message.strip()
    if not texto:
        raise HTTPException(status_code=400, detail="Empty message")

    turno = await wh.ejecutar_turno(_usuario(sid), texto)

    sesion = app.state.web_sesiones.setdefault(
        sid, {"lead_id": None, "pendiente": False, "intentado": False})
    if turno["lead_id"]:
        sesion["lead_id"] = turno["lead_id"]

    pendiente = bool(turno["evento"] and turno["lead_id"])
    if pendiente:
        sesion["pendiente"] = True
        sesion["intentado"] = True
        background.add_task(_analizar_sesion, sid, turno["lead_id"], turno["mensajes"])

    return {"reply": turno["respuesta"], "lead_id": sesion["lead_id"],
            "analysis_pending": pendiente, "trace": traza_para_web(turno["nuevos"])}


@app.get("/api/score/{session_id}")
async def puntaje(session_id: str):
    sid = _validar(session_id)
    sesion = app.state.web_sesiones.get(sid)
    if not sesion or not sesion["lead_id"]:
        return {"status": "none"}
    if sesion["pendiente"]:
        return {"status": "pending"}
    resultado = await asyncio.to_thread(leer_resultado, sesion["lead_id"])
    if resultado:
        return {"status": "ready", "result": resultado}
    return {"status": "failed" if sesion["intentado"] else "none"}


@app.post("/api/reset")
async def reiniciar(body: ResetIn):
    sid = _validar(body.session_id)
    wh.store.datos.pop(_usuario(sid), None)
    app.state.web_sesiones.pop(sid, None)
    return {"status": "ok"}


if __name__ == "__main__":
    import os
    import uvicorn
    uvicorn.run("demo_server:app", host="127.0.0.1", port=int(os.getenv("PORT", "8000")))
