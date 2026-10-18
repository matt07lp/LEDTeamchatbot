import asyncio
import json
import sys
from datetime import datetime, timedelta

from mcp import ClientSession, StdioServerParameters
from mcp.client.stdio import stdio_client


def proxima_fecha_habil(hora: int = 10, dias_extra: int = 0) -> str:
    """Devuelve una fecha futura (lunes a viernes) en formato ISO 8601.
    El showroom solo abre de lunes a viernes, asi que evitamos fines de semana."""
    fecha = datetime.now() + timedelta(days=2 + dias_extra)
    while fecha.weekday() > 4:
        fecha += timedelta(days=1)
    return fecha.replace(hour=hora, minute=0, second=0, microsecond=0).isoformat()


async def main():
    # Parametros para lanzar servidor_led.py como subproceso.
    # command=sys.executable asegura que se use el MISMO interprete de Python
    # con el que corre este cliente.
    parametros_servidor = StdioServerParameters(
        command=sys.executable,
        args=["servidor_led.py"],
    )

    print("Conectando al servidor MCP ...")
    async with stdio_client(parametros_servidor) as (read, write):
        async with ClientSession(read, write) as session:
            # Handshake MCP: negociacion de version y capacidades
            await session.initialize()
            print("Conexion establecida.\n")

            # Descubrimiento: el cliente NO conoce las tools de antemano
            respuesta_tools = await session.list_tools()
            print("Herramientas descubiertas en el servidor:")
            for tool in respuesta_tools.tools:
                descripcion = (tool.description or "").strip().splitlines()[0]
                print(f"  - {tool.name}: {descripcion}")
            print()

            # El resultado viene como una lista de bloques de contenido.
            # Nuestras tools devuelven diccionarios, que llegan como JSON en .text
            def texto(resultado):
                return resultado.content[0].text

            async def llamar(nombre, argumentos):
                """Invoca una tool, imprime el resultado y lo devuelve como dict."""
                print(f"Invocando {nombre}({json.dumps(argumentos, ensure_ascii=False)})...")
                resultado = await session.call_tool(nombre, argumentos)
                crudo = texto(resultado)
                try:
                    datos = json.loads(crudo)
                    print(f"  Resultado: {json.dumps(datos, indent=2, ensure_ascii=False)}\n")
                    return datos
                except json.JSONDecodeError:
                    print(f"  Resultado: {crudo}\n")
                    return {"error": crudo}

            # Datos de contacto de prueba
            contacto = {"name": "Ana Ruiz", "email": "ana@acme.com", "phone": "5512345678"}

            # ---------------------------------------------------------------
            # ESCENARIO 1: solicitud NO soportada
            # ---------------------------------------------------------------
            print("=== Escenario 1: solicitud no soportada ===")
            await llamar("evaluate_request", {
                "intent": "other", "confidence": 0.95,
                "summary": "Pregunta por reparacion de una pantalla",
            })

            # ---------------------------------------------------------------
            # ESCENARIO 2: COMPRA -> visita al showroom
            # ---------------------------------------------------------------
            print("=== Escenario 2: compra + visita al showroom ===")
            r = await llamar("evaluate_request", {
                "intent": "buy", "confidence": 0.9,
                "summary": "Quiere comprar 2 pantallas interiores para un lobby",
            })
            lead_compra = r["lead_id"]  # el lead_id se reutiliza en las demas tools

            fecha_visita = proxima_fecha_habil(hora=10)
            await llamar("handle_purchase_request", {
                "lead_id": lead_compra, "action": "showroom_visit",
                "preferred_datetime": fecha_visita, **contacto,
            })

            # ---------------------------------------------------------------
            # ESCENARIO 3: conflicto de horario (mismo horario, otro lead)
            # ---------------------------------------------------------------
            print("=== Escenario 3: horario ocupado (debe rechazar) ===")
            r = await llamar("evaluate_request", {
                "intent": "buy", "confidence": 0.9,
                "summary": "Otro cliente quiere visitar el showroom",
            })
            await llamar("handle_purchase_request", {
                "lead_id": r["lead_id"], "action": "showroom_visit",
                "preferred_datetime": fecha_visita,
                "name": "Pedro Gil", "email": "pedro@mail.com", "phone": "5500000000",
            })

            # ---------------------------------------------------------------
            # ESCENARIO 4: compra -> cotizacion (primero sin pixel pitch)
            # ---------------------------------------------------------------
            print("=== Escenario 4: cotizacion de compra ===")
            await llamar("handle_purchase_request", {
                "lead_id": lead_compra, "action": "send_quotation",
                "screen_type": "indoor", "width_m": 3, "height_m": 2, "quantity": 2,
                **contacto,
            })  # falta pixel_pitch_mm: el servidor devuelve las opciones
            await llamar("handle_purchase_request", {
                "lead_id": lead_compra, "action": "send_quotation",
                "screen_type": "indoor", "pixel_pitch_mm": 2.5,
                "width_m": 3, "height_m": 2, "quantity": 2, **contacto,
            })

            # ---------------------------------------------------------------
            # ESCENARIO 5: RENTA -> cotizacion
            # ---------------------------------------------------------------
            print("=== Escenario 5: renta + cotizacion ===")
            r = await llamar("evaluate_request", {
                "intent": "rent", "confidence": 0.92,
                "summary": "Renta de pantalla para un concierto",
            })
            inicio = (datetime.now() + timedelta(days=15)).date()
            fin = inicio + timedelta(days=2)
            await llamar("handle_rental_request", {
                "lead_id": r["lead_id"],
                "start_date": inicio.isoformat(), "end_date": fin.isoformat(),
                "location": "Mexico City, Foro Sol",
                "screen_type": "outdoor", "width_m": 6, "height_m": 3,
                **contacto, "company": "Events MX",
            })

            # ---------------------------------------------------------------
            # ESCENARIO 6: validacion (tool de renta con un lead de compra)
            # ---------------------------------------------------------------
            print("=== Escenario 6: validacion de intencion (debe rechazar) ===")
            await llamar("handle_rental_request", {
                "lead_id": lead_compra,
                "start_date": inicio.isoformat(), "end_date": fin.isoformat(),
                "location": "Mexico City", "screen_type": "indoor",
                "width_m": 2, "height_m": 2, **contacto,
            })

            print("Todo el intercambio anterior uso el protocolo MCP real:")
            print("el cliente nunca importo ni conocio el codigo del servidor,")
            print("solo hablo list_tools() y call_tool() por stdio.")
            print("Revisa data/leads.xlsx para ver los leads, citas y cotizaciones guardados.")


if __name__ == "__main__":
    asyncio.run(main())
