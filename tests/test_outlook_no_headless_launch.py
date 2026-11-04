"""GURI must attach to an open Outlook and never start a headless one (BUG-040)."""

from __future__ import annotations

import sys
import types
import unittest
from pathlib import Path
from unittest import mock

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from guri import outlook_scraper  # noqa: E402


def _fake_win32com(get_object, dispatch):
    client = types.SimpleNamespace(GetObject=get_object, Dispatch=dispatch)
    package = types.ModuleType("win32com")
    package.client = client
    module = types.ModuleType("win32com.client")
    module.GetObject = get_object
    module.Dispatch = dispatch
    return {"win32com": package, "win32com.client": module}


class OutlookAppTests(unittest.TestCase):
    def test_closed_outlook_is_not_started(self):
        dispatch = mock.Mock()
        modules = _fake_win32com(mock.Mock(side_effect=OSError("no active object")), dispatch)
        with mock.patch.dict(sys.modules, modules), mock.patch.object(
            outlook_scraper, "outlook_is_running", return_value=False
        ):
            with self.assertRaises(RuntimeError):
                outlook_scraper._outlook_app()
        dispatch.assert_not_called()

    def test_running_outlook_missed_by_getobject_still_attaches(self):
        app = object()
        dispatch = mock.Mock(return_value=app)
        modules = _fake_win32com(mock.Mock(side_effect=OSError("missed")), dispatch)
        with mock.patch.dict(sys.modules, modules), mock.patch.object(
            outlook_scraper, "outlook_is_running", return_value=True
        ):
            self.assertIs(outlook_scraper._outlook_app(), app)

    def test_getobject_is_preferred(self):
        app = object()
        dispatch = mock.Mock()
        modules = _fake_win32com(mock.Mock(return_value=app), dispatch)
        with mock.patch.dict(sys.modules, modules):
            self.assertIs(outlook_scraper._outlook_app(), app)
        dispatch.assert_not_called()

    def test_process_check_reads_tasklist(self):
        running = types.SimpleNamespace(stdout='"OUTLOOK.EXE","123","Console","1","1 K"\n')
        closed = types.SimpleNamespace(stdout="INFO: No tasks are running.\n")
        with mock.patch.object(outlook_scraper.subprocess, "run", return_value=running):
            self.assertTrue(outlook_scraper.outlook_is_running())
        with mock.patch.object(outlook_scraper.subprocess, "run", return_value=closed):
            self.assertFalse(outlook_scraper.outlook_is_running())


if __name__ == "__main__":
    unittest.main()
