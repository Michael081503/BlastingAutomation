"""Daily availability controls shared by workflows with the same agents."""

from __future__ import annotations

from datetime import date
from functools import partial
from typing import Callable

import streamlit as st

from agent_availability import agent_roster, available_agents, save_available_agents
from automation import AutomationError


AGENT_AVAILABILITY_PAGE_VERSION = 1
AGENT_ABSENCE_CHECKBOX_VERSION = 1


def _error_key(group: str, assignment_date: date) -> str:
    return f"agent_availability_error_{group}_{assignment_date.isoformat()}"


def _save_selection(group: str, assignment_date: date, key: str) -> None:
    try:
        save_available_agents(group, assignment_date, st.session_state[key])
        roster = agent_roster(group)
        st.session_state[key + "_saved"] = tuple(
            code for _, code in available_agents(group, assignment_date, roster)
        )
        st.session_state.pop(_error_key(group, assignment_date), None)
    except AutomationError as exc:
        st.session_state[_error_key(group, assignment_date)] = {
            "message": str(exc), "codes": list(st.session_state[key]),
        }
    result_keys = ("result", "geo_autostat_result") if group == "geo_reco" else ("b2_result", "b4_result")
    for result_key in result_keys:
        st.session_state.pop(result_key, None)


def _save_absence(group: str, assignment_date: date, key: str, code: str, checkbox_key: str) -> None:
    try:
        pending = st.session_state.get(_error_key(group, assignment_date))
        codes = list(pending["codes"]) if pending else [
            agent_code for _, agent_code in available_agents(group, assignment_date, agent_roster(group))
        ]
        if st.session_state[checkbox_key]:
            codes = [value for value in codes if value != code]
        elif code not in codes:
            codes.append(code)
        st.session_state[key] = codes
        _save_selection(group, assignment_date, key)
    except AutomationError as exc:
        st.session_state[_error_key(group, assignment_date)] = {
            "message": str(exc), "codes": list(st.session_state.get(key, [])),
        }
        for result_key in ("result", "geo_autostat_result", "b2_result", "b4_result"):
            st.session_state.pop(result_key, None)


def render_agent_availability(
    group: str, assignment_date: date, agents: tuple[tuple[str, str], ...], *, key_prefix: str,
    on_move: Callable[[str, int], None] | None = None,
) -> tuple[tuple[str, str], ...] | None:
    """Return available agents in the caller's order, or block preparation."""
    try:
        roster = agent_roster(group)
        saved_codes = tuple(code for _, code in available_agents(group, assignment_date, roster))
    except AutomationError as exc:
        st.error(str(exc))
        return None
    key = f"agent_availability_{key_prefix}_{group}_{assignment_date.isoformat()}"
    save_error = st.session_state.get(_error_key(group, assignment_date))
    displayed_codes = save_error["codes"] if save_error else list(saved_codes)
    st.session_state[key] = displayed_codes
    st.caption(f"Mark absent agents for {assignment_date.isoformat()}. Accounts go to the remaining agents.")
    for index, (name, code) in enumerate(agents):
        columns = st.columns([6, 2, 1, 1] if on_move else [6, 2], vertical_alignment="center")
        columns[0].markdown(f"**{index + 1}. {name}**  \n{code}")
        checkbox_key = f"{key}_{code}_absent"
        absent = code not in displayed_codes
        if checkbox_key not in st.session_state or st.session_state.get(checkbox_key + "_saved") != absent:
            st.session_state[checkbox_key] = absent
            st.session_state[checkbox_key + "_saved"] = absent
        columns[1].checkbox(
            "Absent", key=checkbox_key, help=f"Exclude {name} from assignments for this date.",
            on_change=partial(_save_absence, group, assignment_date, key, code, checkbox_key),
        )
        if on_move:
            columns[2].button("↑", key=f"{key}_{code}_up", disabled=index == 0,
                              help=f"Move {name} up", on_click=on_move, args=(code, -1))
            columns[3].button("↓", key=f"{key}_{code}_down", disabled=index == len(agents) - 1,
                              help=f"Move {name} down", on_click=on_move, args=(code, 1))
    if save_error:
        st.error(save_error["message"])
        st.button("Retry saving availability", key=key + "_retry",
                  on_click=partial(_save_selection, group, assignment_date, key))
        return None
    try:
        selected = available_agents(group, assignment_date, agents)
    except AutomationError as exc:
        st.error(str(exc))
        return None
    if not selected:
        st.info("All agents are marked absent. Uncheck Absent for at least one agent before preparing files.")
        return None
    st.caption("Availability is saved for this date. Each new day starts with all agents available.")
    return selected
