"""Reusable, UI-agnostic analytics engine.

This package MUST NOT import any UI framework (PySide6). Every front-end --
the current desktop app, and future Ribbon add-in / web / API front-ends --
talks to the engine through the single entry point in ``core.pipeline.Engine``.
"""
