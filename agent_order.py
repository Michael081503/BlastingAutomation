"""Self-contained drag-and-drop ordering for the RECO agents."""

from pathlib import Path
from typing import Callable

import streamlit.components.v1 as components


AGENT_ORDER_VERSION = 2
AGENT_ORDER_CALLBACK_VERSION = 1
_ASSETS = Path(__file__).resolve().parent / "assets" / "agent-order"
_agent_order_component = components.declare_component(
    "reco_agent_order",
    path=str(_ASSETS),
)


def render_agent_order(
    agents: tuple[tuple[str, str], ...], *, key: str, revision: str,
    on_change: Callable[[], None], save_error: str | None = None,
) -> None:
    """Keep the list mounted and save changes before the next page render."""
    _agent_order_component(
        key=key,
        data=[{"name": name, "code": code} for name, code in agents],
        revision=revision,
        save_error=save_error,
        on_change=on_change,
        default=None,
    )
