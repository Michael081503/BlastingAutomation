"""Self-contained drag-and-drop ordering for the RECO agents."""

from pathlib import Path

import streamlit.components.v1 as components


AGENT_ORDER_VERSION = 1
_ASSETS = Path(__file__).resolve().parent / "assets" / "agent-order"
_agent_order_component = components.declare_component(
    "reco_agent_order",
    path=str(_ASSETS),
)


def render_agent_order(agents: tuple[tuple[str, str], ...], *, key: str) -> list[str] | None:
    """Return a new order only after a drop or an arrow-button click."""
    return _agent_order_component(
        key=key,
        data=[{"name": name, "code": code} for name, code in agents],
        default=None,
    )
