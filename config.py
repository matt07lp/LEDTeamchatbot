"""
config.py - Configuracion central del agente.

Lee el archivo .env (misma carpeta) y expone:
  - PYTHON / SERVIDOR_MCP : para lanzar el servidor MCP como subproceso
  - llm_langchain()       : devuelve el modelo para LangChain
  - llm_llamaindex()      : devuelve el modelo para LlamaIndex

Proveedor gratuito elegido en .env con LLM_PROVIDER:
  "google" -> Gemini via Google AI Studio (API key gratuita)
  "ollama" -> modelo local con Ollama (sin internet, sin API key)
"""
import os
import sys
from pathlib import Path

from dotenv import load_dotenv

# Carga .env desde la carpeta de este archivo, sin importar desde donde se ejecute
load_dotenv(Path(__file__).parent / ".env")

# Mismo interprete de Python que corre el agente (evita "python" vs "python3")
PYTHON = sys.executable
SERVIDOR_MCP = "servidor_led.py"

PROVEEDOR = os.getenv("LLM_PROVIDER", "google").strip().lower()
GOOGLE_MODEL = os.getenv("GOOGLE_MODEL", "gemini-flash-latest")
OLLAMA_MODEL = os.getenv("OLLAMA_MODEL", "qwen3:8b")
OLLAMA_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")

# Temperatura opcional. Si la dejas vacia se usa el valor por defecto del modelo
# (recomendado para Gemini 3; para Ollama conviene algo bajo como 0).
_temp = os.getenv("LLM_TEMPERATURE", "").strip()
TEMPERATURA = float(_temp) if _temp else None


def _verificar_google() -> str:
    key = os.getenv("GOOGLE_API_KEY", "").strip()
    if not key or key == "tu_api_key_aqui":
        raise RuntimeError(
            "Falta GOOGLE_API_KEY en el archivo .env. "
            "Creala gratis en https://aistudio.google.com/apikey"
        )
    return key


def _proveedor_invalido():
    raise ValueError(f"LLM_PROVIDER='{PROVEEDOR}' no valido. Usa 'google' u 'ollama'.")


def llm_langchain(modelo: str | None = None):
    """Modelo de chat para LangChain. `modelo` permite usar otro modelo (ej. el analista)."""
    kwargs = {"temperature": TEMPERATURA} if TEMPERATURA is not None else {}
    if PROVEEDOR == "google":
        from langchain_google_genai import ChatGoogleGenerativeAI
        return ChatGoogleGenerativeAI(
            model=modelo or GOOGLE_MODEL, google_api_key=_verificar_google(), **kwargs)
    if PROVEEDOR == "ollama":
        from langchain_ollama import ChatOllama
        return ChatOllama(model=modelo or OLLAMA_MODEL, base_url=OLLAMA_URL,
                          **({"temperature": TEMPERATURA if TEMPERATURA is not None else 0}))
    _proveedor_invalido()


def llm_llamaindex():
    """Modelo LLM para LlamaIndex."""
    if PROVEEDOR == "google":
        from llama_index.llms.google_genai import GoogleGenAI
        return GoogleGenAI(model=GOOGLE_MODEL, api_key=_verificar_google())
    if PROVEEDOR == "ollama":
        from llama_index.llms.ollama import Ollama
        return Ollama(model=OLLAMA_MODEL, base_url=OLLAMA_URL,
                      request_timeout=180.0,
                      temperature=TEMPERATURA if TEMPERATURA is not None else 0)
    _proveedor_invalido()
