"""
qif_colors.py
-------------
Shared ANSI colour utilities for QImpExp.
Mirrors the t3_colors.py API so both tools can coexist.
No external dependencies.
"""

import os
import re
import sys

RESET  = "\033[0m"
BOLD   = "\033[1m"
DIM    = "\033[2m"
RED    = "\033[91m"
CYAN   = "\033[96m"
GREEN  = "\033[92m"
YELLOW = "\033[93m"
WHITE  = "\033[97m"
GREY   = "\033[90m"


def use_color() -> bool:
    """Return True when the terminal is likely to render ANSI codes."""
    if os.environ.get("NO_COLOR"):
        return False
    if sys.stdout.isatty():
        return True
    return bool(os.environ.get("WT_SESSION") or os.environ.get("TERM_PROGRAM"))


def strip_ansi(text: str) -> str:
    return re.sub(r"\033\[[0-9;]*m", "", text)


def col(code: str, text: str) -> str:
    """Wrap *text* in an ANSI code + RESET, only when colour is supported."""
    return f"{code}{text}{RESET}" if use_color() else text
