"""
trazas.py - Imprime lo que hace el agente paso a paso:
mensaje del cliente -> llamadas a herramientas -> resultados -> respuesta final.
"""
import json


def _texto(contenido) -> str:
    """El contenido puede ser un str o una lista de bloques (ej. Gemini)."""
    if isinstance(contenido, str):
        return contenido
    if isinstance(contenido, list):
        partes = []
        for bloque in contenido:
            if isinstance(bloque, str):
                partes.append(bloque)
            elif isinstance(bloque, dict) and bloque.get("type") == "text":
                partes.append(bloque.get("text", ""))
        return "".join(partes)
    return str(contenido)


def imprimir_traza_langchain(mensajes) -> None:
    for m in mensajes:
        if m.type == "human":
            print(f"Cliente: {_texto(m.content)}")
        elif m.type == "ai":
            for llamada in getattr(m, "tool_calls", None) or []:
                print(f"  [tool] {llamada['name']}({llamada['args']})")
            texto = _texto(m.content).strip()
            if texto:
                print(f"Agente: {texto}")
        elif m.type == "tool":
            resultado = _texto(m.content)
            print(f"  [resultado {m.name}] {resultado[:400]}")


async def ejecutar_con_traza_llamaindex(agente, pregunta: str) -> None:
    from llama_index.core.agent.workflow import ToolCall, ToolCallResult

    print(f"Cliente: {pregunta}")
    handler = agente.run(user_msg=pregunta)
    async for evento in handler.stream_events():
        if isinstance(evento, ToolCall):
            print(f"  [tool] {evento.tool_name}({evento.tool_kwargs})")
        elif isinstance(evento, ToolCallResult):
            print(f"  [resultado {evento.tool_name}] {str(evento.tool_output)[:400]}")
    respuesta = await handler
    print(f"Agente: {respuesta}\n")


def traza_para_web(mensajes, max_chars: int = 400) -> list:
    """Tool calls of one turn as [{tool, args, result}] for the web demo.
    Results are parsed as JSON when possible; long non-JSON text is truncated."""
    pasos, por_id = [], {}
    for m in mensajes:
        tipo = getattr(m, "type", None)
        if tipo == "ai":
            for llamada in getattr(m, "tool_calls", None) or []:
                paso = {"tool": llamada.get("name"), "args": llamada.get("args", {}),
                        "result": None}
                pasos.append(paso)
                if llamada.get("id"):
                    por_id[llamada["id"]] = paso
        elif tipo == "tool":
            crudo = _texto(m.content)
            try:
                resultado = json.loads(crudo)
            except ValueError:
                resultado = crudo[:max_chars]
            paso = por_id.get(getattr(m, "tool_call_id", None))
            if paso is None:
                paso = next((p for p in pasos if p["result"] is None), None)
            if paso is not None:
                paso["result"] = resultado
    return pasos
