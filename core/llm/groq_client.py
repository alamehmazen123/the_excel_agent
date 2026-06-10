"""Minimal Groq client (OpenAI-compatible chat completions).

Designed to fail softly: any network/auth/rate-limit problem returns None so
the caller falls back to the deterministic template. Never raises to the engine.
"""
from __future__ import annotations

from typing import Any, Optional

import requests

from .prompts import SYSTEM_PROMPT, build_user_prompt

GROQ_BASE_URL = "https://api.groq.com/openai/v1/chat/completions"
DEFAULT_MODEL = "llama-3.3-70b-versatile"


class GroqNarrator:
    """Callable that turns a metrics dict into narrative text, or None."""

    def __init__(self, api_key: str, model: str = DEFAULT_MODEL,
                 timeout: float = 30.0, retries: int = 1) -> None:
        self.api_key = (api_key or "").strip()
        self.model = model or DEFAULT_MODEL
        self.timeout = timeout
        self.retries = max(0, retries)
        self.last_error: Optional[str] = None

    def available(self) -> bool:
        return bool(self.api_key)

    def __call__(self, metrics: dict[str, Any]) -> Optional[str]:
        if not self.available():
            self.last_error = "No API key configured."
            return None

        payload = {
            "model": self.model,
            "messages": [
                {"role": "system", "content": SYSTEM_PROMPT},
                {"role": "user", "content": build_user_prompt(metrics)},
            ],
            "temperature": 0.4,
            "max_tokens": 900,
        }
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

        last_exc: Optional[str] = None
        for attempt in range(self.retries + 1):
            try:
                resp = requests.post(GROQ_BASE_URL, json=payload, headers=headers,
                                     timeout=self.timeout)
                if resp.status_code == 200:
                    data = resp.json()
                    text = data["choices"][0]["message"]["content"].strip()
                    self.last_error = None
                    return text or None
                if resp.status_code in (401, 403):
                    self.last_error = "Invalid or unauthorized API key."
                    return None      # no point retrying a bad key
                if resp.status_code == 429:
                    last_exc = "Rate limited by Groq (free-tier limit reached)."
                else:
                    last_exc = f"Groq returned HTTP {resp.status_code}."
            except requests.exceptions.RequestException as exc:
                last_exc = f"Network error contacting Groq: {exc}"
        self.last_error = last_exc
        return None

    def test_connection(self) -> tuple[bool, str]:
        """Lightweight check for the Settings dialog 'Test' button."""
        result = self.__call__({"record_count": 1, "measures": [],
                                "top_breakdowns": [], "source_sheet": "test",
                                "column_count": 1})
        if result:
            return True, "Connection successful."
        return False, self.last_error or "Connection failed."
