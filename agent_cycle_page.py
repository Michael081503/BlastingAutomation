"""Visible, editable daily agent cycle for GEO and RECO campaign exports."""

from __future__ import annotations

import hashlib
from datetime import datetime, timedelta

import streamlit as st

from agent_cycle import AgentCycle, cycle_agents, default_cycle, load_cycle, save_cycle
from automation import AutomationError
from export_config import GEO_RECO_AGENTS
from folder_output import MANILA_TIMEZONE


AGENT_CYCLE_PAGE_VERSION = 1


def manila_today():
    return datetime.now(MANILA_TIMEZONE).date()


def render_agent_cycle():
    try:
        cycle = load_cycle()
    except AutomationError as exc:
        st.error(str(exc))
        if st.button("Reset agent cycle to defaults", key="agent_cycle_reset"):
            try:
                save_cycle(default_cycle())
                st.rerun()
            except AutomationError as error:
                st.error(str(error))
        return None
    today = manila_today()
    with st.expander("GEO & RECO agent cycle", expanded=True):
        assignment_date = st.date_input(
            "Assignment date (Asia/Manila)", today, key=f"agent_cycle_date_{today.isoformat()}",
            help="Defaults to today's Manila date each day. Choose another date to preview or prepare that day's assignments.",
        )
        if assignment_date is None:
            st.info("Choose an assignment date to prepare GEO and RECO files.")
            return None
        agents = cycle_agents(cycle, assignment_date)
        st.table([{"Order": number, "Monthly": name, "Terms": code} for number, (name, code) in enumerate(agents, 1)])
        st.caption(
            "Accounts in each output group are split evenly in this order, starting with the highest OB. "
            "The same account keeps its agent in Email and Viber. The first agent moves to the end each calendar day."
        )
        st.caption("Change the saved order and start date under Change GEO & RECO agent cycle in the sidebar.")
        with st.expander("See upcoming days"):
            st.table([
                {"Date": (assignment_date + timedelta(days=day)).isoformat(),
                 **{label: code for label, (_, code) in zip(("First", "Second", "Third"), cycle_agents(cycle, assignment_date + timedelta(days=day)))}}
                for day in range(3)
            ])
        with st.sidebar.expander("Change GEO & RECO agent cycle"):
            revision = hashlib.sha256(repr(cycle).encode()).hexdigest()[:12]
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
                    st.session_state["agent_cycle_saved"] = True
                    st.rerun()
                except AutomationError as exc:
                    st.error(str(exc))
            st.caption("The saved cycle is shared by GEO and RECO and stored on the computer or server running this app.")
        if st.session_state.pop("agent_cycle_saved", False):
            st.success("Agent cycle saved.")
    return assignment_date, agents
