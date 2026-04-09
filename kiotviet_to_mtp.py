#!/usr/bin/env python3
"""Thin CLI wrapper that preserves the original entrypoint."""

from __future__ import annotations

import sys

from app.kv_runner import main


if __name__ == "__main__":
    sys.exit(main())
