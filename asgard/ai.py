from __future__ import annotations

import json
from typing import Optional

import requests

from .config import get_env


def ollama_generate(
    prompt: str,
    model: str,
    temperature: float = 0.2,
    max_tokens: int = 800,
) -> Optional[str]:
    base_url = get_env("OLLAMA_URL", "http://localhost:11434")
    url = f"{base_url.rstrip('/')}/api/generate"
    payload = {
        "model": model,
        "prompt": prompt,
        "stream": False,
        "options": {"temperature": temperature, "num_predict": max_tokens},
    }
    try:
        response = requests.post(url, json=payload, timeout=120)
        response.raise_for_status()
    except requests.RequestException:
        return None
    data = response.json()
    return data.get("response")


def build_prompt(template: str, content: str) -> str:
    return f"{template.strip()}\n\nТекст:\n{content.strip()}\n"


def generate_post(
    template: str,
    content: str,
    model: str,
    temperature: float,
    max_tokens: int,
) -> Optional[str]:
    prompt = build_prompt(template, content)
    return ollama_generate(prompt, model=model, temperature=temperature, max_tokens=max_tokens)
