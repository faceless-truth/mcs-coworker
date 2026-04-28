"""
Client name normalisation utility.

Standardises all client references to 'Surname, First Name' format
to match MC&S SharePoint folder structure:
  /Clients/Korkie, Gordon/
  /Clients/Korkie, Gordon/Korkie Family Trust/
  /Clients/Korkie, Gordon/Korkie Holdings Pty Ltd/

Memory storage uses:
  client_name = "Korkie, Gordon"          (primary key — the individual)
  entity_name = "Korkie Family Trust"     (optional — the entity within)
"""
import re
import logging
from typing import Optional, Tuple

logger = logging.getLogger(__name__)

# Common entity suffixes that indicate this is an entity, not a person
ENTITY_SUFFIXES = [
    "pty ltd", "pty. ltd.", "pty limited",
    "trust", "family trust", "unit trust", "discretionary trust",
    "smsf", "super fund", "superannuation fund",
    "holdings", "investments", "enterprises", "group",
    "partnership", "association", "incorporated", "inc",
    "foundation", "estate",
]


def is_entity_name(name: str) -> bool:
    """Check if a name is an entity (company/trust/SMSF) rather than a person."""
    lower = name.lower().strip()
    return any(lower.endswith(suffix) or suffix in lower for suffix in ENTITY_SUFFIXES)


def normalise_client_name(name: str) -> str:
    """
    Normalise a client name to 'Surname, First Name' format.

    Handles:
      'Gordon Korkie'        → 'Korkie, Gordon'
      'korkie, gordon'       → 'Korkie, Gordon'
      'KORKIE, GORDON'       → 'Korkie, Gordon'
      'gordon.korkie@...'    → 'Korkie, Gordon'  (email → name extraction)
      'Korkie Family Trust'  → 'Korkie Family Trust'  (entities kept as-is)
      'Gordon J Korkie'      → 'Korkie, Gordon J'

    Returns the normalised name string.
    """
    if not name or not name.strip():
        return "Unknown"

    name = name.strip()

    # If it's an entity, title-case it and return as-is
    if is_entity_name(name):
        return _title_case_entity(name)

    # If it looks like an email address, extract the name part
    if "@" in name:
        name = _name_from_email(name)

    # If already in 'Surname, First' format
    if "," in name:
        parts = [p.strip() for p in name.split(",", 1)]
        if len(parts) == 2 and parts[0] and parts[1]:
            return f"{parts[0].title()}, {parts[1].title()}"

    # Split into words — assume last word is surname
    words = name.split()
    if len(words) == 1:
        return words[0].title()

    # Last word is surname, everything else is first/middle names
    surname = words[-1]
    first_names = " ".join(words[:-1])

    return f"{surname.title()}, {first_names.title()}"


def _name_from_email(email: str) -> str:
    """Extract a plausible name from an email address."""
    local = email.split("@")[0]
    # Replace dots, underscores, hyphens with spaces
    local = re.sub(r"[._\-]", " ", local)
    # Remove numbers
    local = re.sub(r"\d+", "", local)
    return local.strip()


def _title_case_entity(name: str) -> str:
    """Title-case an entity name, preserving 'Pty Ltd' etc."""
    # Simple title case, then fix common abbreviations
    result = name.title()
    replacements = {
        "Pty Ltd": "Pty Ltd",
        "Pty. Ltd.": "Pty. Ltd.",
        "Pty Limited": "Pty Limited",
        "Smsf": "SMSF",
    }
    for find, replace in replacements.items():
        result = result.replace(find, replace)
    return result


def parse_client_and_entity(text: str) -> Tuple[Optional[str], Optional[str]]:
    """
    Parse a user input that might contain both a client name and entity.

    Examples:
      'Gordon Korkie'                    → ('Korkie, Gordon', None)
      'Korkie, Gordon'                   → ('Korkie, Gordon', None)
      'Korkie Family Trust'              → (None, 'Korkie Family Trust')
      'Gordon Korkie - Korkie Family Trust' → ('Korkie, Gordon', 'Korkie Family Trust')
      'Korkie, Gordon / Korkie Holdings' → ('Korkie, Gordon', 'Korkie Holdings')

    Returns (client_name, entity_name) tuple.
    """
    if not text or not text.strip():
        return ("Unknown", None)

    text = text.strip()

    # Check for separators indicating both person and entity
    for sep in [" - ", " / ", " | ", " — ", " – "]:
        if sep in text:
            parts = text.split(sep, 1)
            person_part = parts[0].strip()
            entity_part = parts[1].strip()

            if is_entity_name(entity_part):
                return (normalise_client_name(person_part), _title_case_entity(entity_part))
            elif is_entity_name(person_part):
                return (normalise_client_name(entity_part), _title_case_entity(person_part))

    # Single value — determine if it's a person or entity
    if is_entity_name(text):
        # Try to extract the surname for the parent folder
        # 'Korkie Family Trust' → parent might be 'Korkie, ...'
        # But we don't know the first name, so just return the entity
        return (None, _title_case_entity(text))

    return (normalise_client_name(text), None)


def format_memory_metadata(
    client_name: Optional[str] = None,
    entity_name: Optional[str] = None,
    **extra
) -> dict:
    """
    Build a standardised metadata dict for memory storage.
    Always includes client_name in normalised format.
    """
    meta = {}
    if client_name:
        meta["client_name"] = normalise_client_name(client_name)
    if entity_name:
        meta["entity_name"] = _title_case_entity(entity_name)
        # If we have an entity but no client name, use the entity
        if not client_name:
            meta["client_name"] = meta["entity_name"]
    meta.update(extra)
    return meta
