"""Persist the agents available for each assignment day."""

from __future__ import annotations

import json
import os
from datetime import date
from pathlib import Path
from tempfile import NamedTemporaryFile
from typing import Sequence

from automation import AutomationError
from export_config import COLLECTION_AGENTS, GEO_RECO_AGENTS


AGENT_AVAILABILITY_VERSION = 1
SETTINGS_PATH = Path(__file__).resolve().parent / "settings" / "agent_availability.json"
AGENT_ROSTERS = {"geo_reco": GEO_RECO_AGENTS, "b2_b4": COLLECTION_AGENTS}


def agent_roster(group: str) -> tuple[tuple[str, str], ...]:
    try:
        return AGENT_ROSTERS[group]
    except (KeyError, TypeError) as exc:
        raise AutomationError("Choose the GEO/RECO or B2/B4 agent group.") from exc


def _day_key(assignment_date: date) -> str:
    if type(assignment_date) is not date:
        raise AutomationError("Choose a valid assignment date for agent availability.")
    return assignment_date.isoformat()


def _validated_codes(group: str, codes: Sequence[str]) -> list[str]:
    roster = agent_roster(group)
    known = {code for _, code in roster}
    if (not isinstance(codes, (list, tuple)) or any(not isinstance(code, str) for code in codes)
            or len(set(codes)) != len(codes) or not set(codes) <= known):
        raise AutomationError("Choose each available agent once from the displayed roster.")
    return [code for _, code in roster if code in codes]


def _read_settings(path: Path) -> dict:
    try:
        if not path.exists():
            return {"version": AGENT_AVAILABILITY_VERSION, "days": {}}
        data = json.loads(path.read_text(encoding="utf-8"))
        if (not isinstance(data, dict) or data.get("version") != AGENT_AVAILABILITY_VERSION
                or not isinstance(data.get("days"), dict)):
            raise ValueError("Invalid availability settings")
        for day, groups in data["days"].items():
            if not isinstance(day, str) or date.fromisoformat(day).isoformat() != day or not isinstance(groups, dict):
                raise ValueError("Invalid availability day")
            for group, codes in groups.items():
                if not isinstance(codes, list):
                    raise ValueError("Invalid availability roster")
                _validated_codes(group, codes)
        return data
    except (OSError, ValueError, TypeError, AutomationError) as exc:
        raise AutomationError(
            "The saved agent availability cannot be read. Restore its settings file or remove "
            "settings/agent_availability.json to start with all agents available."
        ) from exc


def available_agents(
    group: str, assignment_date: date, agents: tuple[tuple[str, str], ...], *, path: Path | None = None,
) -> tuple[tuple[str, str], ...]:
    """Filter an ordered roster while retaining its rotation or fixed order."""
    roster = agent_roster(group)
    day = _day_key(assignment_date)
    if (any(agent not in roster for agent in agents)
            or len({code for _, code in agents}) != len(agents)):
        raise AutomationError("Use the configured names and codes for this agent group.")
    data = _read_settings(SETTINGS_PATH if path is None else path)
    codes = data["days"].get(day, {}).get(group, [code for _, code in roster])
    return tuple(agent for agent in agents if agent[1] in codes)


def save_available_agents(
    group: str, assignment_date: date, codes: Sequence[str], *, path: Path | None = None,
) -> None:
    """Save one group's day without replacing other groups or dates."""
    day = _day_key(assignment_date)
    selected = _validated_codes(group, codes)
    path = SETTINGS_PATH if path is None else path
    data = _read_settings(path)
    data["days"].setdefault(day, {})[group] = selected
    temporary = None
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with NamedTemporaryFile(mode="w", encoding="utf-8", dir=path.parent, delete=False) as output:
            temporary = Path(output.name)
            json.dump(data, output, indent=2)
            output.write("\n")
        os.replace(temporary, path)
    except OSError as exc:
        raise AutomationError(
            "Could not save agent availability. Check write access to the app's settings folder and try again."
        ) from exc
    finally:
        if temporary is not None:
            try:
                temporary.unlink(missing_ok=True)
            except OSError:
                pass
