"""Regression coverage for page navigation during a Streamlit module update."""

from __future__ import annotations

import importlib.util
import subprocess
import sys
import unittest
from pathlib import Path
from types import ModuleType
from unittest.mock import MagicMock, patch


ROOT = Path(__file__).resolve().parents[1]


class InboundStartupTests(unittest.TestCase):
    def test_cached_recovery_lookup_policy_refreshes_page(self):
        result = subprocess.run(
            [sys.executable, "-c", "\n".join((
                "import app_startup",
                "app_startup.refresh_application_modules()",
                "import inbound_reply_page",
                "del inbound_reply_page.PL_RECOVERY_LOOKUP_VERSION",
                "assert app_startup.refresh_application_modules() == {'inbound_reply_page'}",
                "assert inbound_reply_page.PL_RECOVERY_LOOKUP_VERSION == 2",
                "assert app_startup.refresh_application_modules() == set()",
            ))], cwd=ROOT, capture_output=True, text=True, timeout=30,
        )
        self.assertEqual(result.returncode, 0, result.stdout + result.stderr)

    def test_app_navigation_accepts_cached_page_without_title_constant(self):
        import app_startup

        legacy_page = ModuleType("inbound_reply_page")
        legacy_page.render_inbound_reply = MagicMock()
        # An already-loaded startup helper can also predate the new page constant.
        with patch.dict(sys.modules, {"inbound_reply_page": legacy_page}), \
                patch.object(app_startup, "refresh_application_modules", return_value=set()):
            spec = importlib.util.spec_from_file_location("app_title_regression", ROOT / "app.py")
            app = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(app)
            with patch.object(app.st, "set_page_config"), patch.object(app.st, "Page") as pages, \
                    patch.object(app.st, "navigation") as navigation:
                app.main()
            inbound = next(call for call in pages.call_args_list
                           if call.args[0] == "pages/INBOUND_REPLY.py")
            self.assertEqual(inbound.kwargs["title"], "INBOUND REPLY")
            navigation.return_value.run.assert_called_once_with()
            self.assertFalse(hasattr(legacy_page, "INBOUND_REPLY_TITLE"))


if __name__ == "__main__":
    unittest.main()
