"""Visible, editable daily agent cycle for RECO campaign exports."""

from __future__ import annotations

import hashlib
from datetime import datetime, timedelta
from functools import partial

import streamlit as st

from agent_cycle import AgentCycle, cycle_agents, default_cycle, load_cycle, reorder_cycle, save_cycle
from agent_order import render_agent_order
from automation import AutomationError
from export_config import GEO_RECO_AGENTS
from folder_output import MANILA_TIMEZONE


AGENT_CYCLE_PAGE_VERSION = 4
AGENT_DRAG_ORDER_VERSION = 1
AGENT_STABLE_ORDER_VERSION = 1
RECO_AGENT_AVAILABILITY_VERSION = 1


def manila_today():
    return datetime.now(MANILA_TIMEZONE).date()


def _cycle_revision(cycle):
    return hashlib.sha256(repr(cycle).encode()).hexdigest()[:12]


def _save_agent_order(key, assignment_date):
    """Save the widget event before rendering, without a second page rerun."""
    st.session_state.pop("agent_cycle_order_error", None)
    st.session_state.pop("result", None)
    try:
        cycle = load_cycle()
        change = st.session_state.get(key)
        if not isinstance(change, dict) or change.get("revision") != _cycle_revision(cycle):
            raise AutomationError("The RECO agent cycle changed. Please reorder the updated list.")
        save_cycle(reorder_cycle(cycle, assignment_date, change.get("order")))
    except AutomationError as exc:
        st.session_state["agent_cycle_order_error"] = {"key": key, "message": str(exc)}


def render_agent_cycle():
    try:
        cycle = load_cycle()
    except AutomationError as exc:
        st.error(str(exc))
        if st.button("Reset agent cycle to defaults", key="agent_cycle_reset"):
            try:
                save_cycle(default_cycle())
                st.session_state.pop("agent_cycle_order_error", None)
                st.rerun()
            except AutomationError as error:
                st.error(str(error))
        return None
    today = manila_today()
    with st.expander("RECO agent cycle", expanded=True):
        assignment_date = st.date_input(
            "Assignment date (Asia/Manila)", today, key=f"agent_cycle_date_{today.isoformat()}",
            help="Defaults to today's Manila date each day. Choose another date to preview or prepare that day's assignments.",
        )
        if assignment_date is None:
            st.info("Choose an assignment date to prepare RECO files.")
            return None
        agents = cycle_agents(cycle, assignment_date)
        revision = _cycle_revision(cycle)
        st.caption("Check Absent beside unavailable agents. Use the arrows to change the order. Changes save automatically.")
        component_key = f"agent_cycle_drag_{assignment_date.isoformat()}"
        error = st.session_state.get("agent_cycle_order_error", {})
        save_error = error.get("message") if error.get("key") == component_key else None
        available = render_agent_order(
            agents, key=component_key, revision=revision, save_error=save_error,
            assignment_date=assignment_date,
            on_change=partial(_save_agent_order, component_key, assignment_date),
        )
        if save_error:
            st.error(save_error)
            return None
        st.caption(
            "Accounts in each output group are split evenly in this order, starting with the highest OB. "
            "The same account keeps its agent in Email and Viber. The first agent moves to the end each calendar day."
        )
        st.caption("The top agent receives the first account. Names fill Monthly and codes fill Terms. Change the cycle start date in the sidebar.")
        with st.expander("See upcoming days"):
            st.table([
                {"Date": (assignment_date + timedelta(days=day)).isoformat(),
                 **{label: code for label, (_, code) in zip(("First", "Second", "Third"), cycle_agents(cycle, assignment_date + timedelta(days=day)))}}
                for day in range(3)
            ])
        with st.sidebar.expander("Change RECO agent cycle"):
            start = st.date_input("Cycle start date", cycle.start_date, key=f"agent_cycle_start_{revision}")
            codes = [code for _, code in GEO_RECO_AGENTS]
            labels = {code: f"{name} ({code})" for name, code in GEO_RECO_AGENTS}
            selected = []
            for number in range(len(codes)):
                selected.append(st.selectbox(
                    f"Agent {number + 1} on the start date", codes,
                    index=codes.index(cycle.agent_codes[number]), format_func=labels.get,
                    key=f"agent_cycle_order_{revision}_{number}",
                ))
            valid = len(set(selected)) == len(codes)
            if not valid:
                st.caption("Choose each agent once before saving. The saved cycle remains active until you save changes.")
            if st.button("Save agent cycle", key="agent_cycle_save", disabled=not valid):
                try:
                    save_cycle(AgentCycle(start, tuple(selected)))
                    st.session_state.pop("agent_cycle_order_error", None)
                    st.session_state["agent_cycle_saved"] = True
                    st.rerun()
                except AutomationError as exc:
                    st.error(str(exc))
            st.caption("This cycle applies only to RECO and is stored on the computer or server running this app.")
        if st.session_state.pop("agent_cycle_saved", False):
            st.success("Agent cycle saved.")
        if available is None:
            return None
    return assignment_date, available
