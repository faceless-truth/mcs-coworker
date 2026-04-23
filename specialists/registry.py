"""
Specialist Agent Registry.

Each specialist has:
  - id: unique slug
  - name: display name for the dropdown
  - description: one-line description
  - icon: emoji for the UI
  - category: grouping (tax, compliance, documents, general)
  - system_prompt: the full system prompt text
  - supports_files: whether this agent can process uploaded files
  - file_types: list of accepted file extensions
  - model_preference: which Claude model to use
    ("sonnet" for complex analysis, "haiku" for simple tasks)
"""
from pathlib import Path
from typing import Dict, List, Optional
from dataclasses import dataclass, field


@dataclass
class SpecialistAgent:
    id: str
    name: str
    description: str
    icon: str
    category: str
    system_prompt: str
    supports_files: bool = False
    file_types: List[str] = field(default_factory=list)
    model_preference: str = "sonnet"


def _load_prompt(filename: str) -> str:
    """Load a system prompt from the specialists/prompts/ directory."""
    prompt_path = Path(__file__).parent / "prompts" / filename
    if prompt_path.exists():
        return prompt_path.read_text(encoding="utf-8")
    raise FileNotFoundError(f"Specialist prompt not found: {prompt_path}")


def get_all_agents() -> Dict[str, SpecialistAgent]:
    """Return all registered specialist agents keyed by id."""
    return {a.id: a for a in [
        SpecialistAgent(
            id="plugin_builder",
            name="Plugin Builder",
            description="Build automation plugins from templates or custom Python",
            icon="🔧",
            category="general",
            system_prompt=_load_prompt("plugin_builder.md"),
            supports_files=False,
            model_preference="sonnet",
        ),
        SpecialistAgent(
            id="gst",
            name="GST Specialist",
            description="Australian GST law analysis with section citations from the GST Act",
            icon="📘",
            category="tax",
            system_prompt=_load_prompt("gst.md"),
            supports_files=True,
            file_types=[".pdf", ".docx", ".xlsx", ".csv", ".txt"],
            model_preference="sonnet",
        ),
        SpecialistAgent(
            id="smsf",
            name="SMSF Specialist",
            description="Self-managed super fund compliance, taxation, and audit advice",
            icon="🏦",
            category="tax",
            system_prompt=_load_prompt("smsf.md"),
            supports_files=True,
            file_types=[".pdf", ".docx", ".xlsx", ".csv", ".txt"],
            model_preference="sonnet",
        ),
        SpecialistAgent(
            id="div7a",
            name="Division 7A Specialist",
            description="Div 7A analysis, Distributable Surplus calculations, loan compliance",
            icon="🔢",
            category="tax",
            system_prompt=_load_prompt("div7a.md"),
            supports_files=True,
            file_types=[".pdf", ".docx", ".xlsx", ".csv", ".txt"],
            model_preference="sonnet",
        ),
        SpecialistAgent(
            id="trusts",
            name="Trust Tax Specialist",
            description="Australian trust taxation — distributions, streaming, s100A, Div 7A/UPE",
            icon="📋",
            category="tax",
            system_prompt=_load_prompt("trusts.md"),
            supports_files=True,
            file_types=[".pdf", ".docx", ".xlsx", ".csv", ".txt"],
            model_preference="sonnet",
        ),
        SpecialistAgent(
            id="individual_tax",
            name="Individual Tax Deductions",
            description="Occupation-specific deductions, rental property, audit risk assessment",
            icon="🧾",
            category="tax",
            system_prompt=_load_prompt("individual_tax.md"),
            supports_files=True,
            file_types=[".pdf", ".docx", ".xlsx", ".csv", ".txt", ".jpg", ".png"],
            model_preference="sonnet",
        ),
        SpecialistAgent(
            id="ato_portal",
            name="ATO Portal Assistant",
            description="Generate remission letters, payment arrangement requests, and portal notes",
            icon="📝",
            category="documents",
            system_prompt=_load_prompt("ato_portal.md"),
            supports_files=True,
            file_types=[".pdf", ".xlsx", ".csv", ".txt"],
            model_preference="sonnet",
        ),
        SpecialistAgent(
            id="net_wealth",
            name="Net Wealth / SMSF Audit",
            description="Extract financial data from statements, produce BGL import files and audit reports",
            icon="💰",
            category="documents",
            system_prompt=_load_prompt("net_wealth.md"),
            supports_files=True,
            file_types=[".pdf", ".xlsx", ".csv", ".txt"],
            model_preference="sonnet",
        ),
    ]}


def get_agent(agent_id: str) -> Optional[SpecialistAgent]:
    """Get a specific agent by id."""
    return get_all_agents().get(agent_id)


def get_agents_by_category(category: str) -> List[SpecialistAgent]:
    """Get all agents in a category."""
    return [a for a in get_all_agents().values() if a.category == category]
