"""Saved daily GEO/RECO agent rotation and consistent account assignments."""

from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import date
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Any

from automation import AutomationError
from export_config import GEO_RECO_AGENTS, GEO_RECO_CYCLE_START


AGENT_CYCLE_VERSION = 1
SETTINGS_PATH = Path(__file__).resolve().parent / "settings" / "geo_reco_agent_cycle.json"


@dataclass(frozen=True)
class AgentCycle:
    start_date: date
    agent_codes: tuple[str, ...]


def default_cycle() -> AgentCycle:
    return AgentCycle(date.fromisoformat(GEO_RECO_CYCLE_START), tuple(code for _, code in GEO_RECO_AGENTS))


def validate_cycle(cycle: AgentCycle) -> None:
    expected = {code for _, code in GEO_RECO_AGENTS}
    if len(cycle.agent_codes) != len(expected) or set(cycle.agent_codes) != expected:
        raise AutomationError("Choose each GEO/RECO agent exactly once in the cycle.")
    if type(cycle.start_date) is not date:
        raise AutomationError("Choose a valid cycle start date.")


def cycle_agents(cycle: AgentCycle, assignment_date: date) -> tuple[tuple[str, str], ...]:
    validate_cycle(cycle)
    offset = (assignment_date - cycle.start_date).days % len(cycle.agent_codes)
    codes = cycle.agent_codes[offset:] + cycle.agent_codes[:offset]
    names = {code: name for name, code in GEO_RECO_AGENTS}
    return tuple((names[code], code) for code in codes)


def load_cycle(path: Path | None = None) -> AgentCycle:
    path = SETTINGS_PATH if path is None else path
    try:
        if not path.exists():
            return default_cycle()
        data = json.loads(path.read_text(encoding="utf-8"))
        if data["version"] != AGENT_CYCLE_VERSION:
            raise ValueError("Unsupported cycle version")
        cycle = AgentCycle(date.fromisoformat(data["start_date"]), tuple(data["agent_codes"]))
        validate_cycle(cycle)
        return cycle
    except (OSError, ValueError, KeyError, TypeError) as exc:
        raise AutomationError("The saved GEO/RECO agent cycle cannot be read. Reset the cycle or restore its settings file.") from exc


def save_cycle(cycle: AgentCycle, path: Path | None = None) -> None:
    validate_cycle(cycle)
    path = SETTINGS_PATH if path is None else path
    temporary = None
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as output:
            temporary = Path(output.name)
            json.dump({"version": AGENT_CYCLE_VERSION, "start_date": cycle.start_date.isoformat(),
                       "agent_codes": cycle.agent_codes}, output, indent=2)
            output.write("\n")
        os.replace(temporary, path)
    except OSError as exc:
        raise AutomationError("Could not save the agent cycle. Check write access to the app's settings folder.") from exc
    finally:
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def assign_agents_to_accounts(
    records: list[dict[str, Any]], agents: tuple[tuple[str, str], ...],
) -> list[dict[str, Any]]:
    """Balance distinct accounts in their current order before channel separation."""
    if not agents or any(not name or not code for name, code in agents):
        raise AutomationError("Choose at least one agent with a name and code before preparing files.")
    assigned, result = {}, []
    for number, record in enumerate(records):
        identity = tuple(str(record.get(field) or "").strip().casefold()
                         for field in ("primary_key", "cardnumber", "indexoneuid"))
        key = identity if any(identity) else ("row", number)
        if key not in assigned:
            assigned[key] = agents[len(assigned) % len(agents)]
        name, code = assigned[key]
        result.append({**record, "Monthly": name, "Terms": code})
    return result
