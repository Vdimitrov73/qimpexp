"""
ca_calendar.py
--------------
Canadian TSX market calendar helpers.
Pure stdlib — no external dependencies.

Provides:
  is_trading_day(d)         -> bool
  next_trading_day(d)       -> date   (T+1 settlement from trade date d)
  prev_trading_day(d)       -> date   (trade date from settlement date d)
  canadian_holidays(year)   -> frozenset[date]

Holiday rules (TSX-observed):
  New Year's Day       Jan 1 (observed Mon if weekend)
  Family Day           3rd Monday of February  (Ontario/TSX)
  Good Friday          Friday before Easter Sunday
  Victoria Day         Last Monday before May 25
  Canada Day           Jul 1 (observed Mon/Tue if weekend)
  Civic Holiday        1st Monday of August
  Labour Day           1st Monday of September
  Thanksgiving         2nd Monday of October
  Christmas Day        Dec 25 (observed Mon if weekend)
  Boxing Day           Dec 26 (observed, adjusted if Christmas shifts)

Note: Remembrance Day (Nov 11) is NOT a TSX holiday.
"""

from datetime import date, timedelta
from functools import lru_cache


# ---------------------------------------------------------------------------
# Easter — Spencer Jones algorithm (pure stdlib)
# ---------------------------------------------------------------------------

def _easter_sunday(year: int) -> date:
    a = year % 19
    b = year // 100
    c = year % 100
    d = b // 4
    e = b % 4
    f = (b + 8) // 25
    g = (b - f + 1) // 3
    h = (19 * a + b - d - g + 15) % 30
    i = c // 4
    k = c % 4
    L = (32 + 2 * e + 2 * i - h - k) % 7
    m = (a + 11 * h + 22 * L) // 451
    month = (h + L - 7 * m + 114) // 31
    day   = ((h + L - 7 * m + 114) % 31) + 1
    return date(year, month, day)


def _observed(d: date) -> date:
    """Move a Sat holiday to Mon, Sun holiday to Mon."""
    if d.weekday() == 5:   # Saturday
        return d + timedelta(2)
    if d.weekday() == 6:   # Sunday
        return d + timedelta(1)
    return d


def _nth_monday(year: int, month: int, n: int) -> date:
    d = date(year, month, 1)
    while d.weekday() != 0:
        d += timedelta(1)
    return d + timedelta(weeks=n - 1)


def _last_monday_before(year: int, month: int, day: int) -> date:
    d = date(year, month, day) - timedelta(1)
    while d.weekday() != 0:
        d -= timedelta(1)
    return d


# ---------------------------------------------------------------------------
# Holiday set
# ---------------------------------------------------------------------------

@lru_cache(maxsize=20)
def canadian_holidays(year: int) -> frozenset:
    """Return frozenset of TSX-closed dates for *year*."""
    h = set()

    # New Year's Day
    h.add(_observed(date(year, 1, 1)))

    # Family Day (3rd Monday of February — Ontario/TSX)
    h.add(_nth_monday(year, 2, 3))

    # Good Friday
    h.add(_easter_sunday(year) - timedelta(2))

    # Victoria Day (last Monday before May 25)
    h.add(_last_monday_before(year, 5, 25))

    # Canada Day (Jul 1, observed)
    h.add(_observed(date(year, 7, 1)))

    # Civic Holiday (1st Monday of August)
    h.add(_nth_monday(year, 8, 1))

    # Labour Day (1st Monday of September)
    h.add(_nth_monday(year, 9, 1))

    # Thanksgiving (2nd Monday of October)
    h.add(_nth_monday(year, 10, 2))

    # Christmas Day (observed)
    xmas = _observed(date(year, 12, 25))
    h.add(xmas)

    # Boxing Day (Dec 26, adjusted so it doesn't clash with observed Christmas)
    boxing = date(year, 12, 26)
    boxing = _observed(boxing)
    if boxing == xmas:           # e.g. Xmas on Sun->Mon, Boxing would also be Mon
        boxing += timedelta(1)
    h.add(boxing)

    return frozenset(h)


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def is_trading_day(d: date) -> bool:
    """Return True if *d* is a TSX trading day."""
    if d.weekday() >= 5:  # weekend
        return False
    return d not in canadian_holidays(d.year)


def next_trading_day(d: date) -> date:
    """
    Return the next TSX trading day after *d*.
    Used to convert trade date -> settlement date (T+1).
    """
    candidate = d + timedelta(1)
    # Cache holiday sets for both the starting year and the next
    # (Dec 31 → Jan 2 crosses a year boundary)
    h = canadian_holidays(candidate.year)
    while candidate.weekday() >= 5 or candidate in h:
        candidate += timedelta(1)
        h = canadian_holidays(candidate.year)
    return candidate


def prev_trading_day(d: date) -> date:
    """
    Return the TSX trading day immediately before *d*.
    Used to convert settlement date -> trade date (T-1).
    """
    candidate = d - timedelta(1)
    h = canadian_holidays(candidate.year)
    while candidate.weekday() >= 5 or candidate in h:
        candidate -= timedelta(1)
        h = canadian_holidays(candidate.year)
    return candidate
