#!/usr/bin/env python3
"""Thin CLI wrapper that preserves the original entrypoint."""

from __future__ import annotations

import sys


def _configure_stdio() -> None:
    """Keep Vietnamese CLI output printable on Windows code pages."""
    for stream in (sys.stdout, sys.stderr):
        if stream is None:
            continue
        try:
            stream.reconfigure(encoding="utf-8")
        except (AttributeError, ValueError):
            pass


_configure_stdio()

from app.kv_runner import main


if __name__ == "__main__":
    sys.exit(main())
