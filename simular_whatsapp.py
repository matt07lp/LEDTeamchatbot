"""
simular_whatsapp.py - Chat with your webhook from the terminal, no WhatsApp needed.

1) In .env set WHATSAPP_DRY_RUN=1, WHATSAPP_APP_SECRET=<any test string>
2) Terminal A:  uvicorn whatsapp_webhook:app --port 8000
3) Terminal B:  python simular_whatsapp.py

It sends correctly signed payloads (exactly like Meta) and reads the bot's
replies from /debug/outbox.
"""
import hashlib
import hmac
import json
import os
import sys
import time
import uuid
from pathlib import Path

import httpx
from dotenv import load_dotenv

load_dotenv(Path(__file__).with_name(".env"))

URL = os.getenv("WEBHOOK_URL", "http://127.0.0.1:8000")
NUMERO = os.getenv("SIM_NUMBER", "5215512345678")   # fake customer number
SECRETO = os.getenv("WHATSAPP_APP_SECRET", "")


def payload(texto: str) -> dict:
    return {"object": "whatsapp_business_account", "entry": [{"id": "SIM", "changes": [{
        "field": "messages", "value": {
            "messaging_product": "whatsapp",
            "metadata": {"phone_number_id": "SIM"},
            "messages": [{"from": NUMERO, "id": f"wamid.{uuid.uuid4().hex}",
                          "type": "text", "text": {"body": texto}}]}}]}]}


def enviar(texto: str) -> None:
    cuerpo = json.dumps(payload(texto)).encode()
    firma = "sha256=" + hmac.new(SECRETO.encode(), cuerpo, hashlib.sha256).hexdigest()
    r = httpx.post(f"{URL}/webhook", content=cuerpo, timeout=10, headers={
        "Content-Type": "application/json", "X-Hub-Signature-256": firma})
    r.raise_for_status()


def respuestas() -> list:
    r = httpx.get(f"{URL}/debug/outbox", params={"to": NUMERO}, timeout=10)
    r.raise_for_status()
    return r.json()


def main():
    if not SECRETO:
        sys.exit("Set WHATSAPP_APP_SECRET in .env (any test string works).")
    try:
        vistas = len(respuestas())
    except httpx.HTTPError as e:
        sys.exit(f"Cannot reach {URL}/debug/outbox. Is the server running with "
                 f"WHATSAPP_DRY_RUN=1? ({e})")
    print(f"Simulated WhatsApp chat as {NUMERO}. Type 'salir' to quit.\n")
    while True:
        texto = input("Tu: ").strip()
        if texto.lower() in {"salir", "exit"}:
            break
        if not texto:
            continue
        enviar(texto)
        inicio = time.time()
        while time.time() - inicio < 120:       # LLM + tools can take a while
            nuevas = respuestas()
            if len(nuevas) > vistas:
                for m in nuevas[vistas:]:
                    print(f"Bot: {m['text']}\n")
                vistas = len(nuevas)
                break
            time.sleep(1)
        else:
            print("(no reply within 120 s - check the server logs)\n")


if __name__ == "__main__":
    main()
