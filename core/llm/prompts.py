"""Prompt construction for the executive summary narrative."""
from __future__ import annotations

import json
from typing import Any

SYSTEM_PROMPT = (
    "You are a senior business analyst writing an executive summary for "
    "non-technical leadership. Be concise, concrete, and decision-oriented. "
    "Use plain language. Do not invent numbers that are not in the data."
)


def build_user_prompt(metrics: dict[str, Any]) -> str:
    """Render the computed metrics into an instruction for the model."""
    return (
        "Below are pre-computed metrics from a business dataset (JSON). Write a "
        "clear executive summary of 3 to 5 short paragraphs. Cover: what the data "
        "represents and its size; the most important figures and what they imply; "
        "any notable trend (e.g. period-over-period growth); and end with 2-3 "
        "concrete recommended actions. Do not use markdown headings or bullet "
        "symbols; separate paragraphs with blank lines.\n\n"
        f"METRICS:\n{json.dumps(metrics, indent=2, default=str)}"
    )
