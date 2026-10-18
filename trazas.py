"""Utilidades para imprimir el rastro de razonamiento de cada framework,
en el mismo formato Thought/Action/Observation de la Clase 2."""


def _texto(contenido):
    """Las tools MCP devuelven una lista de bloques de contenido; las locales, texto."""
    if isinstance(contenido, list):
        return " ".join(b.get("text", "") if isinstance(b, dict) else str(b) for b in contenido)
    return contenido


def imprimir_traza_langchain(mensajes):
    """Recorre la lista de mensajes que devuelve un agente de LangChain."""
    vuelta = 0
    for m in mensajes:
        tipo = m.type  # human, ai, tool
        if tipo == "human":
            print(f"USUARIO: {m.content}")
        elif tipo == "ai" and getattr(m, "tool_calls", None):
            vuelta += 1
            for tc in m.tool_calls:
                print(f"[Vuelta {vuelta}] Action: {tc['name']}({tc['args']})")
        elif tipo == "tool":
            print(f"           Observation: {_texto(m.content)}")
        elif tipo == "ai":
            texto = m.content if isinstance(m.content, str) else " ".join(
                b.get("text", "") for b in m.content if isinstance(b, dict))
            print(f"RESPUESTA FINAL: {texto}")


def _texto_llamaindex(salida):
    """Las tools MCP devuelven un CallToolResult; las locales, el valor directo."""
    crudo = getattr(salida, "raw_output", None)
    bloques = getattr(crudo, "content", None)
    if isinstance(bloques, list):
        return " ".join(getattr(b, "text", str(b)) for b in bloques)
    return salida.content


async def ejecutar_con_traza_llamaindex(agente, pregunta, ctx=None):
    """Corre un agente de LlamaIndex mostrando cada tool call en vivo
    (streaming de eventos) y devuelve la respuesta final."""
    print(f"USUARIO: {pregunta}")
    from llama_index.core.agent.workflow import ToolCallResult

    handler = agente.run(user_msg=pregunta, ctx=ctx) if ctx else agente.run(user_msg=pregunta)
    vuelta = 0
    async for evento in handler.stream_events():
        if isinstance(evento, ToolCallResult):
            vuelta += 1
            print(f"[Vuelta {vuelta}] Action: {evento.tool_name}({evento.tool_kwargs})")
            print(f"           Observation: {_texto_llamaindex(evento.tool_output)}")
    respuesta = await handler
    print(f"RESPUESTA FINAL: {respuesta}\n")
    return respuesta
