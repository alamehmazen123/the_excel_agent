"""Prompt construction for the executive summary narrative.

We ask the model to act as a senior management consultant and return a STRICT
JSON object, which the analyzer renders into styled sections. JSON keeps the
output reliable to parse and consistently structured.
"""
from __future__ import annotations

import json
from typing import Any

SYSTEM_PROMPT = (
    "You are a top-tier management consultant and data analyst (the caliber of "
    "McKinsey/BCG) writing an executive briefing for a busy decision-maker. You "
    "turn raw numbers into sharp, professional, decision-ready guidance. You are "
    "specific and quantified: you cite the actual figures, shares, leaders and "
    "laggards, and trends from the data. You think about concentration risk, "
    "declines, negative contributors, and opportunities. You never invent numbers "
    "that are not present in the provided metrics. Your tone is confident, concise, "
    "and genuinely useful as a guide for the reader's job."
)

# The exact JSON shape we want back.
_SCHEMA = {
    "headline": "one punchy sentence: the single most important takeaway",
    "overview": "2-3 sentences framing what the data is and the big picture",
    "findings": ["3-5 quantified, insight-rich bullet strings using the real numbers"],
    "risks": ["2-4 concrete risks/watch-outs (concentration, declines, negatives, anomalies)"],
    "opportunities": ["2-3 specific upside opportunities grounded in the data"],
    "actions": ["3-5 prioritized, specific, actionable steps the reader should take"],
}


def build_user_prompt(metrics: dict[str, Any]) -> str:
    return (
        "Analyze the pre-computed business metrics below (JSON) and produce an "
        "executive briefing that GUIDES the reader's decisions.\n\n"
        "Return ONLY a valid JSON object with exactly these keys: "
        "headline (string), overview (string), findings (array of strings), "
        "risks (array of strings), opportunities (array of strings), and actions "
        "(array of strings). Each bullet must be a complete, self-contained sentence "
        "in plain professional English, quoting the real figures (use the *_display "
        "values for currency/percent so units are correct). Be substantive and "
        "specific; avoid generic filler. Do not include any text outside the JSON.\n\n"
        f"DESIRED JSON SHAPE (for structure only):\n{json.dumps(_SCHEMA, indent=2)}\n\n"
        f"METRICS:\n{json.dumps(metrics, indent=2, default=str)}"
    )
