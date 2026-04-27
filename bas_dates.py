"""
Dynamic BAS (Business Activity Statement) quarterly date calculator.

Calculates standard and tax-agent lodgement due dates for any Australian
financial year. Handles weekend adjustments (rolls to next business day).

Date pattern (same every year):
  Q1 (Jul-Sep): Standard due 28 Oct, Tax agent due 25 Nov
  Q2 (Oct-Dec): Standard due 28 Feb, Tax agent due 28 Feb (no extension)
  Q3 (Jan-Mar): Standard due 28 Apr, Tax agent due 26 May
  Q4 (Apr-Jun): Standard due 28 Jul, Tax agent due 25 Aug

Source: ATO BAS agent lodgement program.
"""
import logging
from datetime import date, timedelta
from typing import Dict, List, Optional

logger = logging.getLogger(__name__)


def _next_business_day(d: date) -> date:
    """If the date falls on a weekend, roll forward to Monday."""
    while d.weekday() >= 5:  # 5=Saturday, 6=Sunday
        d += timedelta(days=1)
    return d


def _subtract_business_days(d: date, n: int) -> date:
    """Subtract n business days from a date."""
    current = d
    count = 0
    while count < n:
        current -= timedelta(days=1)
        if current.weekday() < 5:  # Monday-Friday
            count += 1
    return current


def get_financial_year(reference_date: Optional[date] = None) -> int:
    """
    Return the financial year ending year for a given date.
    Australian FY runs Jul-Jun. FY2025-26 returns 2026.
    """
    if reference_date is None:
        reference_date = date.today()
    if reference_date.month >= 7:
        return reference_date.year + 1
    return reference_date.year


def get_bas_dates(financial_year: Optional[int] = None) -> List[Dict]:
    """
    Calculate BAS quarterly dates for a given financial year.

    Args:
        financial_year: The FY ending year (e.g., 2026 for FY2025-26).
                        If None, uses the current financial year.

    Returns:
        List of 4 dicts, one per quarter:
        {
            "quarter": "Q1",
            "period": "Jul-Sep 2025",
            "period_start": date(2025, 7, 1),
            "period_end": date(2025, 9, 30),
            "standard_due": date(2025, 10, 28),
            "agent_due": date(2025, 11, 25),
            "has_extension": True,
            "data_request_by": date(2025, 11, 11),  # 10 business days before agent due
            "description": "Q1 (Jul-Sep 2025): Standard due 28 Oct, Agent due 25 Nov"
        }
    """
    if financial_year is None:
        financial_year = get_financial_year()

    # FY start year (e.g., FY2025-26 starts in 2025)
    fy_start = financial_year - 1

    quarters = [
        {
            "quarter": "Q1",
            "period_start": date(fy_start, 7, 1),
            "period_end": date(fy_start, 9, 30),
            "standard_due_raw": date(fy_start, 10, 28),
            "agent_due_raw": date(fy_start, 11, 25),
            "has_extension": True,
        },
        {
            "quarter": "Q2",
            "period_start": date(fy_start, 10, 1),
            "period_end": date(fy_start, 12, 31),
            "standard_due_raw": date(financial_year, 2, 28),
            "agent_due_raw": date(financial_year, 2, 28),  # No extension for Q2
            "has_extension": False,
        },
        {
            "quarter": "Q3",
            "period_start": date(financial_year, 1, 1),
            "period_end": date(financial_year, 3, 31),
            "standard_due_raw": date(financial_year, 4, 28),
            "agent_due_raw": date(financial_year, 5, 26),
            "has_extension": True,
        },
        {
            "quarter": "Q4",
            "period_start": date(financial_year, 4, 1),
            "period_end": date(financial_year, 6, 30),
            "standard_due_raw": date(financial_year, 7, 28),
            "agent_due_raw": date(financial_year, 8, 25),
            "has_extension": True,
        },
    ]

    results = []
    for q in quarters:
        standard_due = _next_business_day(q["standard_due_raw"])
        agent_due = _next_business_day(q["agent_due_raw"])
        data_request_by = _subtract_business_days(agent_due, 10)

        period_label = (
            f"{q['period_start'].strftime('%b')}-{q['period_end'].strftime('%b %Y')}"
        )

        ext_note = "" if q["has_extension"] else " (no extension)"
        description = (
            f"{q['quarter']} ({period_label}): "
            f"Standard due {standard_due.strftime('%d %b %Y')}, "
            f"Agent due {agent_due.strftime('%d %b %Y')}{ext_note}"
        )

        results.append({
            "quarter": q["quarter"],
            "period": period_label,
            "period_start": q["period_start"],
            "period_end": q["period_end"],
            "standard_due": standard_due,
            "agent_due": agent_due,
            "has_extension": q["has_extension"],
            "data_request_by": data_request_by,
            "description": description,
        })

    return results


def get_current_quarter(reference_date: Optional[date] = None) -> Optional[Dict]:
    """Get the BAS quarter that the reference date falls within."""
    if reference_date is None:
        reference_date = date.today()
    fy = get_financial_year(reference_date)
    for q in get_bas_dates(fy):
        if q["period_start"] <= reference_date <= q["period_end"]:
            return q
    return None


def get_next_due_quarter(reference_date: Optional[date] = None) -> Optional[Dict]:
    """Get the next BAS quarter whose agent_due date is in the future."""
    if reference_date is None:
        reference_date = date.today()

    # Check current FY and next FY
    fy = get_financial_year(reference_date)
    all_quarters = get_bas_dates(fy) + get_bas_dates(fy + 1)

    for q in all_quarters:
        if q["agent_due"] >= reference_date:
            return q
    return None


def get_upcoming_deadlines(days_ahead: int = 30, reference_date: Optional[date] = None) -> List[Dict]:
    """Get all BAS deadlines within the next N days."""
    if reference_date is None:
        reference_date = date.today()
    cutoff = reference_date + timedelta(days=days_ahead)

    fy = get_financial_year(reference_date)
    all_quarters = get_bas_dates(fy) + get_bas_dates(fy + 1)

    upcoming = []
    for q in all_quarters:
        if reference_date <= q["agent_due"] <= cutoff:
            days_until = (q["agent_due"] - reference_date).days
            q_copy = dict(q)
            q_copy["days_until_due"] = days_until
            upcoming.append(q_copy)
        if reference_date <= q["data_request_by"] <= cutoff:
            days_until = (q["data_request_by"] - reference_date).days
            q_copy = dict(q)
            q_copy["days_until_data_request"] = days_until
            upcoming.append(q_copy)
    return upcoming


def format_bas_dates_for_prompt(financial_year: Optional[int] = None) -> str:
    """Format BAS dates as a text block for injection into Claude prompts."""
    dates = get_bas_dates(financial_year)
    today = date.today()
    next_due = get_next_due_quarter(today)

    lines = ["BAS Quarterly Lodgement Dates (Tax Agent Program):"]
    for q in dates:
        marker = " ← NEXT DUE" if next_due and q["quarter"] == next_due["quarter"] and q["period_start"] == next_due["period_start"] else ""
        lines.append(f"  {q['description']}{marker}")
        if marker:
            days_left = (q["agent_due"] - today).days
            lines.append(f"    Data needed from clients by: {q['data_request_by'].strftime('%d %b %Y')} ({days_left} days until agent due)")

    lines.append("")
    lines.append("MC&S lodges under the ATO tax agent lodgement program.")
    lines.append("Clients must provide data at least 10 business days before the agent due date.")
    return "\n".join(lines)
