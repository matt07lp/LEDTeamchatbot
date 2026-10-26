"""
Tests for demo_server.py. They reuse the fake agent fixture from test_webhook.py,
so no LLM, API key, MCP server or Excel file is needed.

Run:  pytest test_demo.py -v
"""
import pytest

import demo_server
import whatsapp_webhook as wh
from test_webhook import ctx, FakeAgent, FakeAgentConEvento, Msg, resultado  # noqa: F401
from trazas import traza_para_web

SID = "session-abc12345"


class AIMsg(Msg):
    def __init__(self, content, tool_calls=None):
        super().__init__("ai", content)
        self.tool_calls = tool_calls or []


class ToolMsg(Msg):
    def __init__(self, content, tool_call_id):
        super().__init__("tool", content)
        self.tool_call_id = tool_call_id


class FakeAgentConTrazas(FakeAgent):
    """Replies like the real agent after calling evaluate_request."""
    async def ainvoke(self, entrada):
        r = await super().ainvoke(entrada)
        r["messages"] = r["messages"][:-1] + [
            AIMsg("", [{"name": "evaluate_request", "args": {"intent": "rent"}, "id": "c1"}]),
            ToolMsg('{"status": "supported", "lead_id": "L-ABCDEF12"}', "c1"),
            AIMsg("¿Para qué fechas?")]
        return r


@pytest.fixture(autouse=True)
def limpiar_sesiones():
    demo_server.app.state.web_sesiones.clear()
    yield
    demo_server.app.state.web_sesiones.clear()


def chat(client, mensaje="Hola", sid=SID):
    return client.post("/api/chat", json={"session_id": sid, "message": mensaje})


# --- page ---------------------------------------------------------------------
def test_index_serves_the_demo_page(ctx):
    client, _ = ctx
    r = client.get("/")
    assert r.status_code == 200 and "text/html" in r.headers["content-type"]
    assert "LED Screens" in r.text


# --- chat ---------------------------------------------------------------------
def test_chat_returns_reply(ctx):
    client, agente = ctx
    r = chat(client, "Quiero rentar")
    assert r.status_code == 200
    d = r.json()
    assert d["reply"] == "Echo: Quiero rentar" and d["analysis_pending"] is False
    assert agente.llamadas == [1]


def test_chat_returns_tool_activity_and_lead_id(ctx):
    client, _ = ctx
    wh.app.state.agente = FakeAgentConTrazas()
    d = chat(client, "Quiero rentar").json()
    assert d["lead_id"] == "L-ABCDEF12"
    assert d["trace"][0]["tool"] == "evaluate_request"
    assert d["trace"][0]["args"] == {"intent": "rent"}
    assert d["trace"][0]["result"]["status"] == "supported"


def test_sessions_are_isolated_and_remember(ctx):
    client, agente = ctx
    chat(client, "Hola", sid="session-one-1111")
    chat(client, "Quiero comprar", sid="session-one-1111")
    chat(client, "Hola soy otro", sid="session-two-2222")
    assert agente.llamadas == [1, 3, 1]


def test_reset_forgets_the_conversation(ctx):
    client, agente = ctx
    chat(client, "Hola")
    assert client.post("/api/reset", json={"session_id": SID}).status_code == 200
    chat(client, "Hola otra vez")
    assert agente.llamadas == [1, 1]


def test_invalid_session_and_empty_message_rejected(ctx):
    client, agente = ctx
    assert client.post("/api/chat", json={"session_id": "x", "message": "hola"}).status_code == 400
    assert client.post("/api/chat", json={"session_id": SID, "message": "   "}).status_code == 400
    assert client.post("/api/chat", json={"session_id": SID, "message": ""}).status_code == 422
    assert client.get("/api/score/x").status_code == 400
    assert agente.llamadas == []


# --- score flow -----------------------------------------------------------------
def test_score_none_before_any_lead(ctx):
    client, _ = ctx
    assert client.get(f"/api/score/{SID}").json() == {"status": "none"}


def test_score_ready_after_a_quotation(ctx, monkeypatch):
    client, _ = ctx
    wh.app.state.agente = FakeAgentConEvento()

    async def fake_analizar(lead_id, transcripcion):
        return resultado("WARM", notificar=False)

    saved = {"lead_id": "L-ABCDEF12", "score": 72, "priority": "WARM", "reasons": ["a"],
             "next_action": "Llamar", "breakdown": {"urgency": 25}, "llm_ok": True}
    monkeypatch.setattr(wh, "analizar_lead", fake_analizar)
    monkeypatch.setattr(demo_server, "leer_resultado", lambda lead_id: saved)

    d = chat(client, "Quiero cotizar").json()
    assert d["analysis_pending"] is True and d["lead_id"] == "L-ABCDEF12"
    s = client.get(f"/api/score/{SID}").json()      # background task already finished
    assert s["status"] == "ready" and s["result"]["score"] == 72


def test_score_failed_when_analysis_saved_nothing(ctx, monkeypatch):
    client, _ = ctx
    wh.app.state.agente = FakeAgentConEvento()

    async def fake_analizar(lead_id, transcripcion):
        raise RuntimeError("excel locked")

    monkeypatch.setattr(wh, "analizar_lead", fake_analizar)
    monkeypatch.setattr(demo_server, "leer_resultado", lambda lead_id: None)
    chat(client, "Quiero cotizar")
    assert client.get(f"/api/score/{SID}").json() == {"status": "failed"}


def test_score_pending_while_analysis_runs(ctx):
    client, _ = ctx
    demo_server.app.state.web_sesiones[SID] = {
        "lead_id": "L-ABCDEF12", "pendiente": True, "intentado": True}
    assert client.get(f"/api/score/{SID}").json() == {"status": "pending"}


# --- activity summary helper ------------------------------------------------------
def test_trace_pairs_results_by_id_and_keeps_text_results():
    msgs = [AIMsg("", [{"name": "a", "args": {"x": 1}, "id": "1"},
                       {"name": "b", "args": {}, "id": "2"}]),
            ToolMsg("plain text result", "2"),
            ToolMsg('{"ok": true}', "1")]
    pasos = traza_para_web(msgs)
    assert pasos[0]["tool"] == "a" and pasos[0]["result"] == {"ok": True}
    assert pasos[1]["tool"] == "b" and pasos[1]["result"] == "plain text result"
