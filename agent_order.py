"""Native agent rows with absence checkboxes and order controls."""

from datetime import date
from functools import partial
from typing import Callable

import streamlit as st

from agent_availability_page import render_agent_availability


AGENT_ORDER_VERSION = 3
AGENT_ORDER_CALLBACK_VERSION = 1
AGENT_NATIVE_ROWS_VERSION = 1


def _move_agent(key, revision, codes, on_change, code, delta):
    order = list(codes)
    source = order.index(code)
    destination = source + delta
    if not 0 <= destination < len(order):
        return
    order[source], order[destination] = order[destination], order[source]
    st.session_state[key] = {"order": order, "revision": revision}
    on_change()


def render_agent_order(
    agents: tuple[tuple[str, str], ...], *, key: str, revision: str,
    assignment_date: date, on_change: Callable[[], None], save_error: str | None = None,
) -> tuple[tuple[str, str], ...] | None:
    """Render names and absence controls directly in Streamlit."""
    return render_agent_availability(
        "geo_reco", assignment_date, agents, key_prefix="reco",
        on_move=partial(_move_agent, key, revision, tuple(code for _, code in agents), on_change),
    )
