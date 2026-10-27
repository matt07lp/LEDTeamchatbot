"""
Paso 7 - TU agente: asistente de ventas y renta de pantallas LED.

Conecta el servidor MCP (servidor_led.py) a un LLM gratuito
(Gemini o Ollama, segun LLM_PROVIDER en .env).

Correr con:
  python 07_mi_agente.py          -> corre las PREGUNTAS de prueba
  python 07_mi_agente.py chat     -> conversacion interactiva (con memoria)
"""
import asyncio
import sys
from datetime import date
from pathlib import Path

from config import PYTHON, SERVIDOR_MCP, llm_langchain, llm_llamaindex
from trazas import ejecutar_con_traza_llamaindex, imprimir_traza_langchain

FRAMEWORK = "langchain"

# nombre -> ruta del .py (ruta absoluta, para que funcione desde cualquier carpeta)
SERVIDORES = {
    "mi_servidor": str(Path(__file__).with_name("servidor_led.py")),
}

SYSTEM_PROMPT = f"""Eres el asistente virtual de una empresa que vende y renta pantallas LED.
Hoy es {date.today().isoformat()}. Responde en el idioma del cliente, de forma amable y breve,
haciendo solo una o dos preguntas a la vez.

FLUJO:
1. Primero identifica si el cliente quiere COMPRAR o RENTAR y llama a evaluate_request.
   Nunca asumas la intencion: si no esta clara, pregunta.
   Cualquier otro tema (soporte, reparaciones, etc.) no esta soportado en este chat: explicalo con
   amabilidad y comparte el contacto del equipo de soporte que devuelva la herramienta
   (support_contact: telefono, correo y horario), tal como esta escrito, invitando al cliente a
   llamar o escribir. No inventes precios, tiempos de reparacion ni datos de contacto.
2. COMPRA: ofrece tres opciones: una llamada, una visita al showroom o una cotizacion.
   Luego usa handle_purchase_request.
3. RENTA: pide fechas de inicio y fin, medidas de la pantalla en metros (ancho y alto),
   interior o exterior, lugar del evento y datos de contacto. Luego usa handle_rental_request.
4. Siempre pasa el lead_id que devolvio evaluate_request a las demas herramientas.

REGLAS:
- No inventes precios, tiempos de instalacion, inventario ni disponibilidad.
  Solo informa lo que devuelvan las herramientas. Si no tienes el dato, dilo y ofrece
  una llamada con el equipo.
- No llames una herramienta hasta tener TODOS sus datos obligatorios (nombre, correo y telefono
  incluidos). Pidelos al cliente.
- Si una herramienta devuelve un error, sigue sus instrucciones y vuelve a preguntar al cliente.
- Convierte las fechas a formato ISO (AAAA-MM-DD, o AAAA-MM-DDTHH:MM para citas).
"""

# Preguntas representativas de tu proceso
PREGUNTAS = [
    # Preguntas vagas: el agente debe preguntar, no asumir ni inventar precios
    "Cuanto cuesta una pantalla LED?",
    "Cuanto cuesta una pantalla LED pitch 2.5?",
    "Cuanto cuesta rentar una pantalla?",
    # Fuera de alcance: no debe inventar la respuesta
    "Cuanto tiempo tardan en instalar?",
    # Solicitudes completas: el agente debe llamar a las herramientas
    "Quiero comprar 2 pantallas interiores pitch 2.5 de 3x2 metros. Mandenme una cotizacion. "
    "Soy Ana Ruiz, ana@acme.com, 5512345678.",
    "Necesito rentar una pantalla exterior de 6x3 metros en Ciudad de Mexico, "
    "del 2026-11-20 al 2026-11-22. Soy Luis Mora, luis@events.mx, 3398765432.",
]


async def con_langchain():
    from langchain.agents import create_agent
    from langchain_mcp_adapters.client import MultiServerMCPClient
    cliente = MultiServerMCPClient({
        nombre: {"command": PYTHON, "args": [ruta], "transport": "stdio"}
        for nombre, ruta in SERVIDORES.items()
    })
    agente = create_agent(model=llm_langchain(), tools=await cliente.get_tools(),
                          system_prompt=SYSTEM_PROMPT)
    for p in PREGUNTAS:
        r = await agente.ainvoke({"messages": [{"role": "user", "content": p}]})
        imprimir_traza_langchain(r["messages"])
        print()


async def chat_langchain():
    """Conversacion de varios turnos: el agente recuerda lo que ya se dijo."""
    from langchain.agents import create_agent
    from langchain_mcp_adapters.client import MultiServerMCPClient
    cliente = MultiServerMCPClient({
        nombre: {"command": PYTHON, "args": [ruta], "transport": "stdio"}
        for nombre, ruta in SERVIDORES.items()
    })
    agente = create_agent(model=llm_langchain(), tools=await cliente.get_tools(),
                          system_prompt=SYSTEM_PROMPT)
    historial = []
    print("Chat con el asistente LED (escribe 'salir' para terminar)\n")
    while True:
        texto = input("Tu: ").strip()
        if texto.lower() in {"salir", "exit"}:
            break
        historial.append({"role": "user", "content": texto})
        r = await agente.ainvoke({"messages": historial})
        historial = r["messages"]
        ultimo = historial[-1]
        from trazas import _texto
        print(f"Agente: {_texto(ultimo.content)}\n")


async def con_llamaindex():
    from llama_index.core.agent.workflow import FunctionAgent
    from llama_index.tools.mcp import BasicMCPClient, McpToolSpec
    tools = []
    for ruta in SERVIDORES.values():
        tools += await McpToolSpec(client=BasicMCPClient(PYTHON, args=[ruta])).to_tool_list_async()
    agente = FunctionAgent(tools=tools, llm=llm_llamaindex(), system_prompt=SYSTEM_PROMPT)
    for p in PREGUNTAS:
        await ejecutar_con_traza_llamaindex(agente, p)


if __name__ == "__main__":
    if len(sys.argv) > 1 and sys.argv[1] == "chat":
        asyncio.run(chat_langchain())
    else:
        asyncio.run(con_langchain() if FRAMEWORK == "langchain" else con_llamaindex())
