"""Daily availability controls shared by workflows with the same agents."""

from __future__ import annotations

from datetime import date
from functools import partial

import streamlit as st

from agent_availability import agent_roster, available_agents, save_available_agents
from automation import AutomationError


AGENT_AVAILABILITY_PAGE_VERSION = 1


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


def render_agent_availability(
    group: str, assignment_date: date, agents: tuple[tuple[str, str], ...], *, key_prefix: str,
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
    if (key not in st.session_state or st.session_state.get(key + "_saved") != saved_codes
            or save_error and st.session_state[key] != displayed_codes):
        st.session_state[key] = displayed_codes
        st.session_state[key + "_saved"] = saved_codes
    labels = {code: f"{name} ({code})" for name, code in roster}
    st.multiselect(
        f"Available agents for {assignment_date.isoformat()}", [code for _, code in roster],
        key=key, format_func=labels.get,
        help="Remove unavailable agents. Accounts are divided evenly among the selected agents. "
             "This selection is saved for the date and shared by workflows using this roster.",
        on_change=partial(_save_selection, group, assignment_date, key),
    )
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
        st.info("Select at least one available agent before preparing files.")
        return None
    st.caption("Availability is saved for this date. Each new day starts with all agents available.")
    return selected
