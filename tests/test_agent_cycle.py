from __future__ import annotations

import csv
import json
import unittest
from collections import Counter
from datetime import date, timedelta
from io import BytesIO, StringIO
from itertools import permutations
from pathlib import Path
from tempfile import TemporaryDirectory
from unittest.mock import patch
from uuid import uuid4
from zipfile import ZipFile

from openpyxl import Workbook, load_workbook
from streamlit.testing.v1 import AppTest

from agent_cycle import AgentCycle, assign_agents_to_accounts, cycle_agents, default_cycle, load_cycle, reorder_cycle, save_cycle
from automation import AutomationError, DEFAULT_TEMPLATE_HEADERS, create_default_csv_template
from export_config import GEO_AGENTS, GEO_RECO_AGENTS
from exports import build_excel_zip, build_export_excel, build_geo_exports, build_negative_exports, build_sheet_exports


ROOT = Path(__file__).resolve().parents[1]
TODAY = date(2026, 10, 30)


def records():
    return [{"primary_key": f"{number:018d}", "cardnumber": f"{number + 100:018d}", "indexoneuid": f"UID-{number}",
             "email": f"account{number}@example.com" if number != 1 else None,
             "phone": "09171234567" if number != 2 else None, "OB": number, "level": 1,
             "product": "MC" if number < 8 else "PL", "AREA": "Mindanao", "ChristmasAmnesty": None}
            for number in range(1, 12)]


def rows(file):
    return list(csv.DictReader(StringIO(file.data.decode("cp1252"))))


def agent_order_args(app):
    return json.loads(app.get("component_instance")[0].proto.json_args)


def displayed_agent_codes(app):
    return [agent["code"] for agent in agent_order_args(app)["data"]]


def change_agent_order(app, codes, *, revision=None):
    args = agent_order_args(app)
    states = app._tree.get_widget_states()
    states.widgets.add(
        id=app.get("component_instance")[0].proto.id,
        json_value=json.dumps({"order": codes, "revision": revision or args["revision"], "event_id": str(uuid4())}),
    )
    return app._run(widget_state=states, timeout=20)


class AgentCycleTests(unittest.TestCase):
    def test_reordering_any_day_preserves_start_date_and_rotates_from_selected_order(self):
        cycle = default_cycle()
        for offset in range(-2, 5):
            selected_date = cycle.start_date + timedelta(days=offset)
            for codes in permutations(cycle.agent_codes):
                with self.subTest(offset=offset, codes=codes):
                    reordered = reorder_cycle(cycle, selected_date, list(codes))
                    self.assertEqual(reordered.start_date, cycle.start_date)
                    for day in range(3):
                        actual = cycle_agents(reordered, selected_date + timedelta(days=day))
                        self.assertEqual(tuple(code for _, code in actual), codes[day:] + codes[:day])

    def test_reordering_rejects_missing_duplicate_and_unknown_agents(self):
        for codes in ([], ["APMAMARIL"] * 3, ["APMAMARIL", "ROSVINOYA", "UNKNOWN"], "APMAMARIL", [{}]):
            with self.subTest(codes=codes), self.assertRaisesRegex(AutomationError, "exactly once"):
                reorder_cycle(default_cycle(), TODAY, codes)

    def test_daily_first_agent_moves_to_last_and_repeats_with_matching_names(self):
        cycle = default_cycle()
        for offset in range(-1, 8):
            agents = cycle_agents(cycle, TODAY + timedelta(days=offset))
            first = offset % 3
            self.assertEqual(agents, GEO_RECO_AGENTS[first:] + GEO_RECO_AGENTS[:first])

    def test_changed_cycle_is_saved_and_loaded_without_losing_the_start_date(self):
        with TemporaryDirectory() as directory:
            path = Path(directory) / "settings" / "cycle.json"
            self.assertEqual(load_cycle(path), default_cycle())
            cycle = AgentCycle(TODAY, ("ROSVINOYA", "JBATALLER", "APMAMARIL"))
            save_cycle(cycle, path)
            self.assertEqual(load_cycle(path), cycle)
            self.assertEqual(cycle_agents(load_cycle(path), TODAY + timedelta(days=1))[0], GEO_RECO_AGENTS[2])
            saved = path.read_bytes()
            with patch("agent_cycle.os.replace", side_effect=PermissionError):
                with self.assertRaisesRegex(AutomationError, "Could not save"):
                    save_cycle(default_cycle(), path)
            self.assertEqual(path.read_bytes(), saved)
            self.assertEqual(list(path.parent.iterdir()), [path])
            path.write_text("broken json", encoding="utf-8")
            with self.assertRaisesRegex(AutomationError, "cannot be read"):
                load_cycle(path)

    def test_duplicate_agents_are_rejected_and_duplicate_account_rows_keep_one_agent(self):
        with self.assertRaisesRegex(AutomationError, "exactly once"):
            cycle_agents(AgentCycle(TODAY, ("APMAMARIL",) * 3), TODAY)
        source = records()
        source.insert(1, dict(source[0]))
        assigned = assign_agents_to_accounts(source, GEO_RECO_AGENTS)
        self.assertEqual(assigned[0]["Terms"], assigned[1]["Terms"])
        self.assertEqual(assigned[2]["Terms"], "ROSVINOYA")
        self.assertNotIn("Terms", source[0])

    def test_reco_and_all_accounts_use_balanced_paired_assignments_in_all_formats(self):
        source = records()
        original = [dict(row) for row in source]
        headers = [header for header in DEFAULT_TEMPLATE_HEADERS if header not in ("Monthly", "Terms")]
        for mode in ("RECO", "All accounts"):
            for include_headers in (True, False):
                with self.subTest(mode=mode, include_headers=include_headers):
                    template = create_default_csv_template() if include_headers else (",".join(headers) + "\n").encode()
                    if mode == "All accounts":
                        bundle = build_sheet_exports(template, source, "", 1, base_name="Accounts", label="Accounts",
                                                     template_extension=".csv", agents=GEO_RECO_AGENTS)
                    else:
                        builder = build_geo_exports if mode == "GEO" else build_negative_exports
                        bundle = builder(template, source, "", 1, ".csv", agents=GEO_RECO_AGENTS)
                    assignments, group_accounts = {}, {}
                    for file in bundle.files:
                        exported = rows(file)
                        for row in exported:
                            pair = row["Monthly"], row["Terms"]
                            self.assertIn(pair, GEO_RECO_AGENTS)
                            self.assertEqual(assignments.setdefault(row["primary_key"], pair), pair)
                            group_accounts.setdefault(file.group_key, {})[row["primary_key"]] = row["Terms"]
                        self.assertEqual([row["Monthly"] for row in exported], [row["Monthly"] for row in file.preview])
                        workbook = load_workbook(BytesIO(build_export_excel(file)))
                        saved_headers = [cell.value for cell in workbook.active[1]]
                        saved_rows = list(workbook.active.values)[1:]
                        self.assertEqual([row[saved_headers.index("Terms")] for row in saved_rows], [row["Terms"] for row in exported])
                        self.assertEqual([row[saved_headers.index("Monthly")] for row in saved_rows], [row["Monthly"] for row in exported])
                        workbook.close()
                    self.assertEqual(len(assignments), len(source))
                    for accounts in group_accounts.values():
                        counts = Counter(accounts.values())
                        self.assertLessEqual(max(counts.get(code, 0) for _, code in GEO_RECO_AGENTS)
                                             - min(counts.get(code, 0) for _, code in GEO_RECO_AGENTS), 1)
                    self.assertEqual(assignments[f"{11:018d}"][1], "APMAMARIL")
                    self.assertEqual(source, original)

    def test_geo_keeps_fixed_order_when_rotated_reco_agents_are_supplied(self):
        source = records()
        self.assertEqual(GEO_AGENTS, (
            ("Andrea Mamaril", "APMAMARIL"),
            ("Rosalie Vinoya", "ROSVINOYA"),
            ("Jenayrine Bataller", "JBATALLER"),
        ))
        rotated_agents = GEO_RECO_AGENTS[1:] + GEO_RECO_AGENTS[:1]
        bundle = build_geo_exports(create_default_csv_template(), source, "", 1, ".csv", agents=rotated_agents)
        for file in bundle.files:
            self.assertTrue(all((row["Monthly"], row["Terms"]) in GEO_AGENTS for row in rows(file)))
            self.assertEqual((file.records[0]["Monthly"], file.records[0]["Terms"]), GEO_AGENTS[0])
            workbook = load_workbook(BytesIO(build_export_excel(file)))
            headers = [cell.value for cell in workbook.active[1]]
            for row in list(workbook.active.values)[1:]:
                self.assertIn((row[headers.index("Monthly")], row[headers.index("Terms")]), GEO_AGENTS)
            workbook.close()
        self.assertTrue(all("Monthly" not in row and "Terms" not in row for row in source))

    def test_geo_fixed_agents_balance_accounts_with_custom_headers_and_both_zip_formats(self):
        source = records()
        source[0]["AREA"] = "Region 3"
        source.append(dict(source[3]))
        source.append({**source[0], "primary_key": "NO-CONTACT", "email": None, "phone": None, "OB": 999})
        original = [dict(row) for row in source]
        expected_codes = {1: "APMAMARIL", 2: "JBATALLER", 3: "ROSVINOYA", 4: "APMAMARIL", 5: "JBATALLER",
                          6: "ROSVINOYA", 7: "APMAMARIL", 8: "APMAMARIL", 9: "JBATALLER", 10: "ROSVINOYA", 11: "APMAMARIL"}
        names = {code: name for name, code in GEO_AGENTS}
        headers = [header for header in reversed(DEFAULT_TEMPLATE_HEADERS) if header not in ("Monthly", "Terms")]
        template = (",".join(headers) + "\n").encode()
        bundle = build_geo_exports(template, source, "", 1, ".csv", agents=GEO_RECO_AGENTS)
        assignments = {}
        with ZipFile(BytesIO(bundle.zip_data)) as csv_zip, ZipFile(BytesIO(build_excel_zip(bundle.files))) as excel_zip:
            for file in bundle.files:
                self.assertEqual(file.headers, [*headers, "Monthly", "Terms"])
                exported = rows(file)
                for row in exported:
                    code = expected_codes[int(row["primary_key"])]
                    self.assertEqual((row["Monthly"], row["Terms"]), (names[code], code))
                    self.assertEqual(assignments.setdefault(row["primary_key"], code), code)
                self.assertEqual(csv_zip.read(file.filename), file.data)
                self.assertEqual([(row["Monthly"], row["Terms"]) for row in file.preview],
                                 [(row["Monthly"], row["Terms"]) for row in exported])
                workbook = load_workbook(BytesIO(excel_zip.read(Path(file.filename).with_suffix(".xlsx").name)))
                saved_headers = [cell.value for cell in workbook.active[1]]
                for row, expected in zip(list(workbook.active.values)[1:], exported):
                    self.assertEqual((row[saved_headers.index("Monthly")], row[saved_headers.index("Terms")]),
                                     (expected["Monthly"], expected["Terms"]))
                workbook.close()
        self.assertEqual(len(assignments), 11)
        self.assertEqual(sum(bundle.counts.values()), 12)
        self.assertEqual(source, original)


class Upload(BytesIO):
    def __init__(self, *, geo=False):
        workbook = Workbook()
        workbook.active.title = "NEGATIVE"
        headers = ["primary_key", "cardnumber", "indexoneuid", "email", "phone", "OB", "product", "level", "ChristmasAmnesty"]
        if geo:
            headers.append("AREA")
        workbook.active.append(headers)
        for row in records():
            workbook.active.append([row.get(header) for header in headers])
        output = BytesIO()
        workbook.save(output)
        workbook.close()
        super().__init__(output.getvalue())
        self.name = "GEO.xlsx" if geo else "RECO.xlsx"


class AgentCyclePageTests(unittest.TestCase):
    def setUp(self):
        self.directory = TemporaryDirectory()
        self.addCleanup(self.directory.cleanup)
        self.path = Path(self.directory.name) / "cycle.json"
        self.addCleanup(patch.stopall)
        patch("agent_cycle.SETTINGS_PATH", self.path).start()
        patch("agent_cycle_page.manila_today", return_value=TODAY).start()

    def test_cycle_visible_without_uploads_and_next_day_rotates_to_rosalie(self):
        with patch("streamlit.file_uploader", return_value=None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(displayed_agent_codes(app), ["APMAMARIL", "ROSVINOYA", "JBATALLER"])
            app.date_input(key=f"agent_cycle_date_{TODAY.isoformat()}").set_value(TODAY + timedelta(days=1)).run(timeout=20)
            self.assertEqual(displayed_agent_codes(app), ["ROSVINOYA", "JBATALLER", "APMAMARIL"])
            self.assertFalse(self.path.exists())

    def test_saving_a_new_cycle_invalidates_exports_and_survives_clear_and_new_sessions(self):
        upload = Upload()
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw.get("key") == "source_upload" else None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            app.run(timeout=20)
            self.assertIn("result", app.session_state)
            choices = [widget for widget in app.selectbox if widget.key.startswith("agent_cycle_order_")]
            choices[0].select("ROSVINOYA").run(timeout=20)
            self.assertTrue(app.button(key="agent_cycle_save").disabled)
            choices = [widget for widget in app.selectbox if widget.key.startswith("agent_cycle_order_")]
            choices[1].select("JBATALLER")
            choices[2].select("APMAMARIL")
            app.run(timeout=20)
            app.button(key="agent_cycle_save").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertEqual(app.session_state["result"]["files"][0].records[0]["Terms"], "ROSVINOYA")
            self.assertEqual(load_cycle().agent_codes, ("ROSVINOYA", "JBATALLER", "APMAMARIL"))
            app.run(timeout=20)
            for file in app.session_state["result"]["files"]:
                self.assertEqual(file.records[0]["Terms"], "ROSVINOYA")
            self.assertTrue(all(value in {name for name, _ in GEO_RECO_AGENTS}
                                for value in app.main.dataframe[-1].value["Monthly"]))
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(load_cycle().agent_codes[0], "ROSVINOYA")
            fresh = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(displayed_agent_codes(fresh), ["ROSVINOYA", "JBATALLER", "APMAMARIL"])

    def test_dragging_for_a_rotated_date_saves_order_and_updates_reco_exports(self):
        upload = Upload()
        selected_date = TODAY + timedelta(days=1)
        codes = ["JBATALLER", "ROSVINOYA", "APMAMARIL"]
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw.get("key") == "source_upload" else None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            app.date_input(key=f"agent_cycle_date_{TODAY.isoformat()}").set_value(selected_date).run(timeout=20)
            app.run(timeout=20)
            self.assertIn("result", app.session_state)
            component_id = app.get("component_instance")[0].proto.id
            with patch("agent_cycle_page.st.rerun", side_effect=AssertionError("Reordering should not force another rerun")):
                change_agent_order(app, codes)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertEqual(app.session_state["result"]["files"][0].records[0]["Terms"], "JBATALLER")
            self.assertEqual(load_cycle().start_date, default_cycle().start_date)
            self.assertEqual(displayed_agent_codes(app), codes)
            self.assertEqual(app.get("component_instance")[0].proto.id, component_id)
            self.assertEqual([code for _, code in cycle_agents(load_cycle(), selected_date)], codes)
            self.assertEqual(app.table[0].value.iloc[1]["First"], "ROSVINOYA")
            app.run(timeout=20)
            for file in app.session_state["result"]["files"]:
                self.assertEqual(file.records[0]["Terms"], "JBATALLER")
                self.assertEqual(file.records[0]["Monthly"], "Jenayrine Bataller")
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual([code for _, code in cycle_agents(load_cycle(), selected_date)], codes)
            fresh = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            fresh.date_input(key=f"agent_cycle_date_{TODAY.isoformat()}").set_value(selected_date).run(timeout=20)
            self.assertEqual(displayed_agent_codes(fresh), codes)

    def test_failed_drag_save_blocks_prepared_downloads_and_keeps_saved_order(self):
        upload = Upload()
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw.get("key") == "source_upload" else None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            app.run(timeout=20)
            with patch("agent_cycle_page.save_cycle", side_effect=AutomationError("Could not save the agent cycle.")):
                change_agent_order(app, ["JBATALLER", "ROSVINOYA", "APMAMARIL"])
            self.assertEqual(len(app.exception), 0)
            self.assertTrue(any("Could not save" in error.value for error in app.error))
            self.assertNotIn("result", app.session_state)
            self.assertEqual(load_cycle(), default_cycle())
            change_agent_order(app, ["JBATALLER", "ROSVINOYA", "APMAMARIL"])
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertEqual(displayed_agent_codes(app), ["JBATALLER", "ROSVINOYA", "APMAMARIL"])

    def test_saved_drag_event_does_not_reapply_on_reruns_or_sidebar_changes(self):
        with patch("streamlit.file_uploader", return_value=None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            with patch("agent_cycle_page.save_cycle", wraps=save_cycle) as save:
                change_agent_order(app, ["JBATALLER", "APMAMARIL", "ROSVINOYA"])
                app.run(timeout=20)
                self.assertEqual(save.call_count, 1)
            save_cycle(default_cycle())
            app.run(timeout=20)
            self.assertEqual(displayed_agent_codes(app), ["APMAMARIL", "ROSVINOYA", "JBATALLER"])
            change_agent_order(app, ["JBATALLER", "APMAMARIL", "ROSVINOYA"])
            self.assertEqual(displayed_agent_codes(app), ["JBATALLER", "APMAMARIL", "ROSVINOYA"])
            self.assertEqual(load_cycle().agent_codes[0], "JBATALLER")

    def test_stale_drag_event_keeps_the_newer_saved_cycle_and_can_be_retried(self):
        with patch("streamlit.file_uploader", return_value=None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            old_revision = agent_order_args(app)["revision"]
            newer_cycle = AgentCycle(TODAY, ("ROSVINOYA", "APMAMARIL", "JBATALLER"))
            save_cycle(newer_cycle)
            change_agent_order(app, ["JBATALLER", "APMAMARIL", "ROSVINOYA"], revision=old_revision)
            self.assertEqual(len(app.exception), 0)
            self.assertTrue(any("cycle changed" in error.value for error in app.error))
            self.assertEqual(load_cycle(), newer_cycle)
            self.assertEqual(displayed_agent_codes(app), list(newer_cycle.agent_codes))
            change_agent_order(app, ["JBATALLER", "APMAMARIL", "ROSVINOYA"])
            self.assertEqual(len(app.error), 0)
            self.assertEqual(load_cycle().agent_codes[0], "JBATALLER")

    def test_new_manila_day_rotates_automatically_and_rebuilds_exports(self):
        upload = Upload()
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw.get("key") == "source_upload" else None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            app.run(timeout=20)
            self.assertIn("result", app.session_state)
            with patch("agent_cycle_page.manila_today", return_value=TODAY + timedelta(days=1)):
                app.run(timeout=20)
                self.assertEqual(len(app.exception), 0)
                self.assertEqual(app.session_state["result"]["files"][0].records[0]["Terms"], "ROSVINOYA")
                self.assertEqual(displayed_agent_codes(app), ["ROSVINOYA", "JBATALLER", "APMAMARIL"])
                self.assertEqual(next(widget for widget in app.date_input if widget.label == "Assignment date (Asia/Manila)").value,
                                 TODAY + timedelta(days=1))

    def test_corrupt_saved_cycle_blocks_assignments_and_can_be_reset(self):
        self.path.write_text("bad json", encoding="utf-8")
        with patch("streamlit.file_uploader", return_value=None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertTrue(any("cannot be read" in item.value for item in app.error))
            self.assertNotIn("result", app.session_state)
            app.button(key="agent_cycle_reset").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertEqual(load_cycle(), default_cycle())

    def test_geo_grouped_and_all_accounts_exports_do_not_use_or_load_reco_settings(self):
        self.path.write_text("bad reco cycle json", encoding="utf-8")
        upload = Upload(geo=True)
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw.get("key") == "source_upload" else None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertEqual(len(app.error), 0)
            self.assertFalse(any(widget.key == "agent_cycle_save" for widget in app.button))
            self.assertFalse(any(widget.label == "Assignment date (Asia/Manila)" for widget in app.date_input))
            for mode in ("Separate Area and Christmas Amnesty groups", "All accounts in one group"):
                with self.subTest(mode=mode):
                    app.radio[0].set_value(mode).run(timeout=20)
                    app.run(timeout=20)
                    self.assertEqual(len(app.exception), 0)
                    self.assertEqual(len(app.error), 0)
                    result = app.session_state["result"]
                    self.assertTrue(result["geo"])
                    for file in result["files"]:
                        self.assertTrue(all((row["Monthly"], row["Terms"]) in GEO_AGENTS for row in rows(file)))
                        self.assertEqual((file.records[0]["Monthly"], file.records[0]["Terms"]), GEO_AGENTS[0])
                    signature = result["signature"]
                    zip_data = result["zip_data"]
                    save_cycle(AgentCycle(TODAY, ("ROSVINOYA", "JBATALLER", "APMAMARIL")))
                    with patch("agent_cycle_page.manila_today", return_value=TODAY + timedelta(days=1)):
                        app.run(timeout=20)
                        self.assertEqual(app.session_state["result"]["signature"], signature)
                        self.assertEqual(app.session_state["result"]["zip_data"], zip_data)

    def test_geo_fixed_agents_replace_stale_manual_settings_and_downloads(self):
        upload = Upload(geo=True)
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw.get("key") == "source_upload" else None):
            app = AppTest.from_file(str(ROOT / "app.py"))
            app.session_state["geo_agent_name"] = "Old Manual Agent"
            app.session_state["geo_agent_code"] = "OLDAGENT"
            app.session_state["result"] = {"signature": ("old-manual-assignment",)}
            app.run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertFalse(any(widget.key in ("geo_agent_name", "geo_agent_code") for widget in app.text_input))
            self.assertEqual(list(app.table[0].value["Monthly"]), [name for name, _ in GEO_AGENTS])
            self.assertEqual(list(app.table[0].value["Terms"]), [code for _, code in GEO_AGENTS])
            for mode in ("Separate Area and Christmas Amnesty groups", "All accounts in one group"):
                with self.subTest(mode=mode):
                    app.radio[0].set_value(mode).run(timeout=20)
                    self.assertEqual(len(app.exception), 0)
                    self.assertEqual(len(app.error), 0)
                    result = app.session_state["result"]
                    self.assertTrue(result["geo"])
                    accounts = {}
                    for file in result["files"]:
                        for row in rows(file):
                            pair = row["Monthly"], row["Terms"]
                            self.assertIn(pair, GEO_AGENTS)
                            key = file.group_key, row["primary_key"]
                            self.assertEqual(accounts.setdefault(key, pair), pair)
                    for group in {key[0] for key in accounts}:
                        group_codes = Counter(pair[1] for key, pair in accounts.items() if key[0] == group)
                        counts = [group_codes[code] for _, code in GEO_AGENTS]
                        self.assertEqual(counts, sorted(counts, reverse=True))
                        self.assertLessEqual(max(counts) - min(counts), 1)

    def test_geo_fixed_first_agent_survives_clear_all_files_and_new_sessions(self):
        upload = Upload(geo=True)
        with patch("streamlit.file_uploader", side_effect=lambda label, **kw: upload if kw.get("key", "").startswith("source_upload") else None):
            app = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            upload = None
            app.button(key="clear_all_files").click().run(timeout=20)
            self.assertEqual(len(app.exception), 0)
            self.assertNotIn("result", app.session_state)
            upload = Upload(geo=True)
            app.run(timeout=20)
            fresh = AppTest.from_file(str(ROOT / "app.py")).run(timeout=20)
            for session in (app, fresh):
                self.assertEqual(len(session.exception), 0)
                for file in session.session_state["result"]["files"]:
                    self.assertEqual((file.records[0]["Monthly"], file.records[0]["Terms"]), GEO_AGENTS[0])


if __name__ == "__main__":
    unittest.main()
