from __future__ import annotations

import json
import unittest
from datetime import date, datetime, timedelta
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from agent_availability import agent_roster, available_agents, save_available_agents
from automation import AutomationError
from export_config import COLLECTION_AGENTS, GEO_RECO_AGENTS
from upload_state import clear_all_files


TODAY = date(2026, 11, 4)
APP = """
from datetime import date
import streamlit as st
from agent_availability_page import render_agent_availability
from export_config import COLLECTION_AGENTS
day = st.date_input('Assignment day', date(2026, 11, 4), key='day')
for flow in ('b2', 'b4'):
    selected = render_agent_availability('b2_b4', day, COLLECTION_AGENTS, key_prefix=flow)
    st.text(flow + ':' + (','.join(code for _, code in selected) if selected else 'blocked'))
"""


def absence_controls(app, flow=None):
    prefix = f"agent_availability_{flow}_" if flow else "agent_availability_"
    return [widget for widget in app.checkbox if widget.key.startswith(prefix)]


def set_available(app, codes, flow="b2"):
    for widget in absence_controls(app, flow):
        code = next(code for _, code in COLLECTION_AGENTS if widget.key.endswith(f"_{code}_absent"))
        widget.set_value(code not in codes)
    return app.run(timeout=20)


class AgentAvailabilityTests(unittest.TestCase):
    def setUp(self):
        directory = TemporaryDirectory()
        self.addCleanup(directory.cleanup)
        self.path = Path(directory.name) / "settings" / "availability.json"
        change = patch("agent_availability.SETTINGS_PATH", self.path)
        change.start()
        self.addCleanup(change.stop)

    def test_defaults_and_each_new_day_include_every_agent_without_writing_settings(self):
        self.assertEqual(available_agents("geo_reco", TODAY, GEO_RECO_AGENTS), GEO_RECO_AGENTS)
        self.assertEqual(available_agents("b2_b4", TODAY, COLLECTION_AGENTS), COLLECTION_AGENTS)
        self.assertFalse(self.path.exists())
        save_available_agents("geo_reco", TODAY, ["ROSVINOYA"])
        self.assertEqual(available_agents("geo_reco", TODAY + timedelta(days=1), GEO_RECO_AGENTS), GEO_RECO_AGENTS)

    def test_saved_days_and_groups_are_independent_and_keep_the_supplied_assignment_order(self):
        save_available_agents("geo_reco", TODAY, ["JBATALLER", "APMAMARIL"])
        save_available_agents("b2_b4", TODAY, ["KPALIMA"])
        save_available_agents("geo_reco", TODAY + timedelta(days=1), ["ROSVINOYA"])
        rotated = GEO_RECO_AGENTS[2:] + GEO_RECO_AGENTS[:2]
        self.assertEqual(available_agents("geo_reco", TODAY, rotated),
                         (GEO_RECO_AGENTS[2], GEO_RECO_AGENTS[0]))
        self.assertEqual(available_agents("b2_b4", TODAY, COLLECTION_AGENTS), (COLLECTION_AGENTS[1],))
        self.assertEqual(available_agents("geo_reco", TODAY + timedelta(days=1), rotated), (GEO_RECO_AGENTS[1],))
        saved = json.loads(self.path.read_text(encoding="utf-8"))
        self.assertEqual(saved["days"][TODAY.isoformat()]["geo_reco"], ["APMAMARIL", "JBATALLER"])

    def test_explicitly_unavailable_roster_remains_empty_after_reload(self):
        save_available_agents("b2_b4", TODAY, [])
        self.assertEqual(available_agents("b2_b4", TODAY, COLLECTION_AGENTS), ())
        self.assertEqual(available_agents("b2_b4", TODAY + timedelta(days=1), COLLECTION_AGENTS), COLLECTION_AGENTS)

    def test_unknown_duplicate_or_malformed_selections_do_not_create_settings(self):
        for group, day, codes in (
            ("other", TODAY, []), ("geo_reco", TODAY, ["KPALIMA"]),
            ("geo_reco", TODAY, ["APMAMARIL", "APMAMARIL"]),
            ("b2_b4", TODAY, "KPALIMA"), ("b2_b4", datetime(2026, 11, 4), []),
        ):
            with self.subTest(group=group, day=day, codes=codes), self.assertRaises(AutomationError):
                save_available_agents(group, day, codes)
        self.assertFalse(self.path.exists())
        with self.assertRaises(AutomationError):
            available_agents("geo_reco", TODAY, (("Wrong name", "APMAMARIL"),))

    def test_corrupt_settings_block_reads_and_saves_without_overwriting_the_file(self):
        self.path.parent.mkdir()
        for content in ("{broken", '{"version":2,"days":{}}',
                        '{"version":1,"days":{"2026-11-04":{"geo_reco":["UNKNOWN"]}}}',
                        '{"version":1,"days":{"not-a-date":{}}}'):
            with self.subTest(content=content):
                self.path.write_text(content, encoding="utf-8")
                with self.assertRaisesRegex(AutomationError, "saved agent availability cannot be read"):
                    available_agents("geo_reco", TODAY, GEO_RECO_AGENTS)
                with self.assertRaises(AutomationError):
                    save_available_agents("geo_reco", TODAY, [])
                self.assertEqual(self.path.read_text(encoding="utf-8"), content)

    def test_failed_atomic_save_preserves_previous_availability(self):
        save_available_agents("b2_b4", TODAY, ["KPALIMA"])
        previous = self.path.read_bytes()
        with patch("agent_availability.os.replace", side_effect=PermissionError("Denied")):
            with self.assertRaisesRegex(AutomationError, "Could not save agent availability"):
                save_available_agents("b2_b4", TODAY, ["AMRANCELI"])
        self.assertEqual(self.path.read_bytes(), previous)
        self.assertEqual(available_agents("b2_b4", TODAY, COLLECTION_AGENTS), (COLLECTION_AGENTS[1],))

    def test_shared_widgets_synchronize_both_directions_and_block_an_empty_day(self):
        app = AppTest.from_string(APP).run(timeout=20)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(absence_controls(app)), 4)
        self.assertTrue(all(not widget.value for widget in absence_controls(app)))
        self.assertEqual(len(app.multiselect), 0)
        set_available(app, ["KPALIMA"])
        self.assertEqual(len(app.exception), 0)
        self.assertEqual([widget.value for widget in absence_controls(app)], [True, False, True, False])
        self.assertEqual([item.value for item in app.text], ["b2:KPALIMA", "b4:KPALIMA"])
        set_available(app, [], flow="b4")
        self.assertEqual(len(app.exception), 0)
        self.assertTrue(all(widget.value for widget in absence_controls(app)))
        self.assertEqual([item.value for item in app.text], ["b2:blocked", "b4:blocked"])
        set_available(app, ["AMRANCELI"])
        self.assertEqual([widget.value for widget in absence_controls(app)], [False, True, False, True])
        self.assertEqual([item.value for item in app.text], ["b2:AMRANCELI", "b4:AMRANCELI"])

    def test_widgets_restore_each_dates_saved_selection_and_new_sessions(self):
        app = AppTest.from_string(APP).run(timeout=20)
        set_available(app, ["KPALIMA"])
        app.date_input(key="day").set_value(TODAY + timedelta(days=1)).run(timeout=20)
        self.assertTrue(all(not widget.value for widget in absence_controls(app)))
        app.date_input(key="day").set_value(TODAY).run(timeout=20)
        self.assertEqual([widget.value for widget in absence_controls(app)], [True, False, True, False])
        reopened = AppTest.from_string(APP).run(timeout=20)
        self.assertEqual([widget.value for widget in absence_controls(reopened)], [True, False, True, False])

    def test_save_errors_block_results_until_retry_succeeds(self):
        app = AppTest.from_string(APP).run(timeout=20)
        app.session_state["b2_result"] = {"old": True}
        app.session_state["b4_result"] = {"old": True}
        with patch("agent_availability_page.save_available_agents", side_effect=AutomationError("Cannot save availability")):
            set_available(app, ["KPALIMA"])
        self.assertNotIn("b2_result", app.session_state)
        self.assertNotIn("b4_result", app.session_state)
        self.assertTrue(any(error.value == "Cannot save availability" for error in app.error))
        self.assertEqual([item.value for item in app.text], ["b2:blocked", "b4:blocked"])
        self.assertEqual([widget.value for widget in absence_controls(app)], [True, False, True, False])
        app.button[0].click().run(timeout=20)
        self.assertEqual(len(app.exception), 0)
        self.assertEqual(len(app.error), 0)
        self.assertEqual([widget.value for widget in absence_controls(app)], [True, False, True, False])

    def test_clear_files_keeps_availability_preferences_and_saved_day(self):
        save_available_agents("b2_b4", TODAY, ["KPALIMA"])
        state = {"agent_availability_b2_b2_b4_2026-11-04": ["KPALIMA"], "b2_result": {"old": True}}
        with patch("upload_state.st.session_state", state):
            clear_all_files()
        self.assertEqual(state["agent_availability_b2_b2_b4_2026-11-04"], ["KPALIMA"])
        self.assertNotIn("b2_result", state)
        self.assertEqual(available_agents("b2_b4", TODAY, COLLECTION_AGENTS), (COLLECTION_AGENTS[1],))


if __name__ == "__main__":
    unittest.main()
