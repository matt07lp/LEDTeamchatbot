"""
Tests for whatsapp_webhook.py. They use a FAKE agent, so no LLM, API key,
MCP server or Excel file is needed.

Run:  pytest test_webhook.py -v
"""
import hashlib
import hmac
import json

import pytest
from fastapi.testclient import TestClient

import whatsapp_webhook as wh

SECRET = "test-secret"
VERIFY = "test-verify-token"
NUMERO = "5215511112222"


class Msg:
    def __init__(self, type, content):
        self.type = type
        self.content = content


class FakeAgent:
    """Echoes the customer's message and records how many messages it received."""
    def __init__(self):
        self.llamadas = []
        self.falla = False

    async def ainvoke(self, entrada):
        msgs = entrada["messages"]
        self.llamadas.append(len(msgs))
        if self.falla:
            raise RuntimeError("boom")
        previos = [m if hasattr(m, "type") else Msg("human", m["content"]) for m in msgs]
        return {"messages": previos + [Msg("ai", f"Echo: {msgs[-1]['content']}")]}


@pytest.fixture
def ctx(monkeypatch):
    monkeypatch.setenv("WHATSAPP_VERIFY_TOKEN", VERIFY)
    monkeypatch.setenv("WHATSAPP_APP_SECRET", SECRET)
    monkeypatch.setenv("WHATSAPP_DRY_RUN", "1")
    monkeypatch.setenv("WHATSAPP_ALLOWED_NUMBERS", "")
    monkeypatch.setenv("CONVERSATIONS_FILE", "")
    monkeypatch.delenv("SALES_ALERT_NUMBER", raising=False)
    agente = FakeAgent()
    wh.app.state.agente = agente
    wh.app.state.agente_fijo = True
    wh.app.state.outbox.clear()
    wh.store.datos.clear()
    wh._vistos.clear()
    wh._locks.clear()
    yield TestClient(wh.app), agente
    wh.app.state.agente = None
    wh.app.state.agente_fijo = False


def payload(texto="Hola", wa_id=NUMERO, msg_id="wamid.1", tipo="text"):
    msg = {"from": wa_id, "id": msg_id, "type": tipo}
    if tipo == "text":
        msg["text"] = {"body": texto}
    return {"object": "whatsapp_business_account", "entry": [{"changes": [{
        "field": "messages", "value": {"messages": [msg]}}]}]}


def post(client, datos, firmar=True, secreto=SECRET):
    cuerpo = json.dumps(datos).encode()
    headers = {"Content-Type": "application/json"}
    if firmar:
        headers["X-Hub-Signature-256"] = "sha256=" + hmac.new(
            secreto.encode(), cuerpo, hashlib.sha256).hexdigest()
    return client.post("/webhook", content=cuerpo, headers=headers)


# --- verification handshake -------------------------------------------------
def test_verification_ok(ctx):
    client, _ = ctx
    r = client.get("/webhook", params={
        "hub.mode": "subscribe", "hub.verify_token": VERIFY, "hub.challenge": "12345"})
    assert r.status_code == 200 and r.text == "12345"


def test_verification_wrong_token(ctx):
    client, _ = ctx
    r = client.get("/webhook", params={
        "hub.mode": "subscribe", "hub.verify_token": "nope", "hub.challenge": "1"})
    assert r.status_code == 403


# --- signature ----------------------------------------------------------------
def test_missing_signature_rejected(ctx):
    client, agente = ctx
    assert post(client, payload(), firmar=False).status_code == 403
    assert agente.llamadas == []


def test_wrong_signature_rejected(ctx):
    client, agente = ctx
    assert post(client, payload(), secreto="other-secret").status_code == 403
    assert agente.llamadas == []


# --- normal flow --------------------------------------------------------------
def test_text_message_gets_reply(ctx):
    client, agente = ctx
    assert post(client, payload("Quiero rentar una pantalla")).status_code == 200
    assert wh.app.state.outbox == [{"to": NUMERO, "text": "Echo: Quiero rentar una pantalla"}]
    assert agente.llamadas == [1]


def test_duplicate_message_processed_once(ctx):
    client, agente = ctx
    post(client, payload("Hola", msg_id="wamid.same"))
    post(client, payload("Hola", msg_id="wamid.same"))  # Meta retry
    assert len(wh.app.state.outbox) == 1
    assert agente.llamadas == [1]


def test_memory_per_customer(ctx):
    client, agente = ctx
    post(client, payload("Hola", msg_id="wamid.1"))
    post(client, payload("Quiero comprar", msg_id="wamid.2"))
    post(client, payload("Hola soy otro", wa_id="5215599998888", msg_id="wamid.3"))
    # 2nd call sees: first human + first ai + new human = 3 ; other customer starts fresh
    assert agente.llamadas == [1, 3, 1]


# --- edge cases ----------------------------------------------------------------
def test_non_text_message_gets_polite_reply(ctx):
    client, agente = ctx
    post(client, payload(tipo="image"))
    assert agente.llamadas == []
    assert len(wh.app.state.outbox) == 1
    assert "texto" in wh.app.state.outbox[0]["text"]


def test_status_updates_are_ignored(ctx):
    client, agente = ctx
    datos = {"entry": [{"changes": [{"value": {"statuses": [
        {"id": "x", "status": "delivered"}]}}]}]}
    assert post(client, datos).status_code == 200
    assert agente.llamadas == [] and wh.app.state.outbox == []


def test_allowlist_blocks_other_numbers(ctx, monkeypatch):
    client, agente = ctx
    monkeypatch.setenv("WHATSAPP_ALLOWED_NUMBERS", "5215500000000")
    post(client, payload())
    assert agente.llamadas == [] and wh.app.state.outbox == []


def test_agent_error_sends_apology_and_keeps_going(ctx):
    client, agente = ctx
    agente.falla = True
    assert post(client, payload("Hola", msg_id="wamid.1")).status_code == 200
    assert "problema" in wh.app.state.outbox[-1]["text"]
    agente.falla = False
    post(client, payload("Hola de nuevo", msg_id="wamid.2"))
    assert wh.app.state.outbox[-1]["text"] == "Echo: Hola de nuevo"


def test_history_trim_starts_at_customer_message():
    msgs = [Msg("human", "a"), Msg("ai", "b"), Msg("tool", "c"),
            Msg("human", "d"), Msg("ai", "e")]
    recortado = wh._recortar(msgs, maximo=3)
    assert recortado[0].content == "d"


def test_debug_outbox_hidden_without_dry_run(ctx, monkeypatch):
    client, _ = ctx
    monkeypatch.setenv("WHATSAPP_DRY_RUN", "0")
    assert client.get("/debug/outbox").status_code == 404


# --- agent 2 (lead analyst) coordination ------------------------------------------
REP = "5215500009999"


class FakeAgentConEvento(FakeAgent):
    """Like the real agent after it created a quotation (tool messages in the turn)."""
    async def ainvoke(self, entrada):
        r = await super().ainvoke(entrada)
        r["messages"] = r["messages"][:-1] + [
            Msg("tool", '{"status": "supported", "lead_id": "L-ABCDEF12"}'),
            Msg("tool", '{"status": "quotation_created", "quote_id": "Q-1"}'),
            Msg("ai", "Listo, te enviamos la cotizacion")]
        return r


def resultado(prioridad="HOT", notificar=True):
    return {"lead_id": "L-ABCDEF12", "score": 90, "priority": prioridad,
            "should_notify": notificar, "name": "Ana", "company": "Acme",
            "summary": "4 pantallas", "best_action": "showroom_visit",
            "deal_total": 1000.0, "currency": "MXN", "reasons": ["a"],
            "next_action": "Llamar hoy"}


def test_event_triggers_analyst_and_alert(ctx, monkeypatch):
    client, _ = ctx
    wh.app.state.agente = FakeAgentConEvento()
    monkeypatch.setenv("SALES_ALERT_NUMBER", REP)
    llamadas, marcados = [], []

    async def fake_analizar(lead_id, transcripcion):
        llamadas.append((lead_id, transcripcion))
        return resultado()

    monkeypatch.setattr(wh, "analizar_lead", fake_analizar)
    monkeypatch.setattr(wh, "marcar_notificado", lambda lead_id: marcados.append(lead_id))
    post(client, payload("Quiero cotizar"))
    assert llamadas[0][0] == "L-ABCDEF12" and "Quiero cotizar" in llamadas[0][1]
    assert wh.app.state.outbox[0]["to"] == NUMERO           # customer is answered first
    assert wh.app.state.outbox[1]["to"] == REP
    assert "HOT LEAD" in wh.app.state.outbox[1]["text"]
    assert marcados == ["L-ABCDEF12"]


def test_no_event_means_no_analysis(ctx, monkeypatch):
    client, _ = ctx
    llamadas = []

    async def fake_analizar(*a, **k):
        llamadas.append(1)

    monkeypatch.setattr(wh, "analizar_lead", fake_analizar)
    post(client, payload("Hola"))
    assert llamadas == []


def test_no_alert_when_not_hot_or_already_notified(ctx, monkeypatch):
    client, _ = ctx
    wh.app.state.agente = FakeAgentConEvento()
    monkeypatch.setenv("SALES_ALERT_NUMBER", REP)

    async def fake_analizar(lead_id, transcripcion):
        return resultado("WARM", notificar=False)

    monkeypatch.setattr(wh, "analizar_lead", fake_analizar)
    post(client, payload("Quiero cotizar"))
    assert len(wh.app.state.outbox) == 1                    # only the customer reply


def test_hot_without_alert_number_is_not_marked_notified(ctx, monkeypatch):
    client, _ = ctx
    wh.app.state.agente = FakeAgentConEvento()
    marcados = []

    async def fake_analizar(lead_id, transcripcion):
        return resultado()

    monkeypatch.setattr(wh, "analizar_lead", fake_analizar)
    monkeypatch.setattr(wh, "marcar_notificado", lambda lead_id: marcados.append(lead_id))
    post(client, payload("Quiero cotizar"))
    assert marcados == [] and len(wh.app.state.outbox) == 1


def test_analyst_failure_never_breaks_customer_reply(ctx, monkeypatch):
    client, _ = ctx
    wh.app.state.agente = FakeAgentConEvento()

    async def fake_analizar(lead_id, transcripcion):
        raise RuntimeError("excel locked")

    monkeypatch.setattr(wh, "analizar_lead", fake_analizar)
    assert post(client, payload("Quiero cotizar")).status_code == 200
    assert wh.app.state.outbox[0]["text"] == "Listo, te enviamos la cotizacion"
