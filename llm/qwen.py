"""Shared optional Ollama transport. Empty output triggers deterministic fallback."""
import requests
from config import OLLAMA_URL, MODEL


def is_available(timeout=3):
    try:
        response = requests.get(f"{OLLAMA_URL}/api/tags", timeout=timeout)
        response.raise_for_status()
        names = [item.get("name", "") for item in response.json().get("models", [])]
        return any(name == MODEL or name == MODEL + ":latest" for name in names)
    except (requests.RequestException, ValueError, TypeError, AttributeError):
        return False


def generate(prompt, num_predict=500, temperature=0.2, timeout=120, json_mode=False):
    payload = {
        "model": MODEL, "prompt": prompt, "stream": False,
        "options": {"num_predict": num_predict, "temperature": temperature, "num_ctx": 8192},
    }
    if json_mode:
        payload["format"] = "json"
    try:
        response = requests.post(f"{OLLAMA_URL}/api/generate", json=payload, timeout=timeout)
        response.raise_for_status()
        text = response.json().get("response", "")
        return text if isinstance(text, str) else ""
    except (requests.RequestException, ValueError, TypeError, AttributeError):
        return ""
