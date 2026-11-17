"""The suite must attach to an open Outlook and never start a headless one (BUG-040)."""

from __future__ import annotations

import re
import sys
import types
import unittest
from pathlib import Path
from unittest import mock

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from geofooter import outlook_com  # noqa: E402


def _fake_win32com(get_object, dispatch):
    client = types.SimpleNamespace(GetObject=get_object, Dispatch=dispatch)
    package = types.ModuleType("win32com")
    package.client = client
    module = types.ModuleType("win32com.client")
    module.GetObject = get_object
    module.Dispatch = dispatch
    return {"win32com": package, "win32com.client": module}


class AttachOutlookTests(unittest.TestCase):
    def test_closed_outlook_is_not_started(self):
        dispatch = mock.Mock()
        modules = _fake_win32com(mock.Mock(side_effect=OSError("no active object")), dispatch)
        with mock.patch.dict(sys.modules, modules), mock.patch.object(
            outlook_com, "outlook_is_running", return_value=False
        ):
            with self.assertRaises(RuntimeError):
                outlook_com.attach_outlook()
        dispatch.assert_not_called()

    def test_running_outlook_missed_by_getobject_still_attaches(self):
        app = object()
        dispatch = mock.Mock(return_value=app)
        modules = _fake_win32com(mock.Mock(side_effect=OSError("missed")), dispatch)
        with mock.patch.dict(sys.modules, modules), mock.patch.object(
            outlook_com, "outlook_is_running", return_value=True
        ):
            self.assertIs(outlook_com.attach_outlook(), app)

    def test_getobject_is_preferred(self):
        app = object()
        dispatch = mock.Mock()
        modules = _fake_win32com(mock.Mock(return_value=app), dispatch)
        with mock.patch.dict(sys.modules, modules):
            self.assertIs(outlook_com.attach_outlook(), app)
        dispatch.assert_not_called()

    def test_window_check_never_raises(self):
        with mock.patch.object(outlook_com.ctypes, "windll", create=True, new=None):
            self.assertFalse(outlook_com.outlook_is_running())

    def test_scraper_uses_the_shared_helper(self):
        from guri import outlook_scraper

        with mock.patch.object(outlook_scraper, "attach_outlook", return_value="app") as attach:
            self.assertEqual(outlook_scraper._outlook_app(), "app")
        attach.assert_called_once_with()


class NoDirectDispatchTests(unittest.TestCase):
    """Only geofooter/outlook_com.py may create Outlook.Application."""

    def test_no_python_file_dispatches_outlook_itself(self):
        pattern = re.compile(r"""(Dispatch(Ex)?|EnsureDispatch)\(\s*['"]Outlook\.Application""")
        offenders = []
        for package in ("aes", "guri", "aura", "geofooter", "scripts"):
            for path in (ROOT / package).rglob("*.py"):
                if path.name == "outlook_com.py":
                    continue
                if pattern.search(path.read_text(encoding="utf-8", errors="replace")):
                    offenders.append(str(path.relative_to(ROOT)))
        self.assertEqual(offenders, [])


if __name__ == "__main__":
    unittest.main()
