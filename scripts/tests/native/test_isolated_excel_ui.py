from pathlib import Path
from tempfile import TemporaryDirectory
from types import SimpleNamespace
import unittest
from xml.etree import ElementTree
from zipfile import ZipFile

from isolated_excel_ui import (
    ExcelOwner, ProbeError, bounded_points, check_synthetic_compile_dialog, color_state,
    failure_record, write_startup_book,
)


class IsolatedExcelUiTests(unittest.TestCase):
    def test_startup_fixture_has_one_synthetic_sheet_and_no_macro_or_links(self):
        with TemporaryDirectory() as temp:
            source = Path(temp) / "startup.xlsx"
            write_startup_book(source)
            with ZipFile(source) as archive:
                self.assertEqual(
                    set(archive.namelist()),
                    {"[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                     "xl/_rels/workbook.xml.rels", "xl/worksheets/sheet1.xml"})
                sheet = ElementTree.fromstring(archive.read("xl/workbook.xml")).find(".//{*}sheet")
                self.assertEqual(sheet.attrib["name"], "Origin")
                for name in ("_rels/.rels", "xl/_rels/workbook.xml.rels"):
                    for relationship in ElementTree.fromstring(archive.read(name)):
                        self.assertNotIn("TargetMode", relationship.attrib)
            with self.assertRaises(FileExistsError):
                write_startup_book(source)

    def test_bounds_reject_empty_large_outside_and_non_integer_points(self):
        self.assertEqual(bounded_points(((0, 0), (31, 31)), 32, 32), ((0, 0), (31, 31)))
        for points in ((), ((0, 0),) * 17, ((32, 0),), ((0, -1),), ((True, 0),), ((1.5, 0),)):
            with self.subTest(points=points), self.assertRaises(ProbeError):
                bounded_points(points, 32, 32)

    def test_clipped_pixels_do_not_count_as_transient_color(self):
        allowed = {(10, 20), (30, 40)}
        self.assertEqual(color_state((10, 20), allowed), "expected")
        self.assertEqual(color_state((30, 40), allowed), "expected")
        self.assertEqual(color_state((30, 20), allowed), "other")
        self.assertEqual(color_state((30, 0xFFFFFFFF), allowed), "invalid")

    def test_window_must_match_process_and_private_desktop(self):
        owner = object.__new__(ExcelOwner)
        owner.pid, owner.desktop = 17, "Private"
        api = owner.windows = SimpleNamespace(
            gui=SimpleNamespace(IsWindow=lambda hwnd: hwnd != 0),
            process=SimpleNamespace(GetWindowThreadProcessId=lambda hwnd: (hwnd, 17 if hwnd != 3 else 18)),
            user=SimpleNamespace(GetThreadDesktop=lambda thread: thread),
            desktop_name=lambda desktop: "Private" if desktop == 1 else "Default")
        self.assertTrue(owner.matches(1))
        self.assertFalse(owner.matches(0))
        self.assertFalse(owner.matches(2))
        self.assertFalse(owner.matches(3))
        mutations = []
        owner.assert_process = lambda: None
        api.input_desktop_name = lambda: "Private"
        api.gui.PostMessage = lambda *arguments: mutations.append(arguments)
        with self.assertRaisesRegex(ProbeError, "private_desktop_became_input_desktop"):
            owner.click(1, (0, 0))
        self.assertEqual(mutations, [])

    def test_drawing_release_does_not_hide_primary_failure(self):
        owner = object.__new__(ExcelOwner)
        owner.assert_window = lambda hwnd: None
        calls = []
        owner.windows = SimpleNamespace(
            gui=SimpleNamespace(GetClientRect=lambda hwnd: (0, 0, 32, 32)),
            user=SimpleNamespace(GetDC=lambda hwnd: 7,
                                 ReleaseDC=lambda hwnd, dc: calls.append((hwnd, dc)) or 0))

        def fail(_):
            raise ProbeError("observation_failed")

        with self.assertRaises(ExceptionGroup) as caught:
            owner.get_pixels(1, ((0, 0),), fail)
        self.assertEqual(calls, [(1, 7)])
        data = failure_record("sample_hover", caught.exception)
        self.assertEqual([item["code"] for item in data["errors"]],
                         ["observation_failed", "window_device_context_not_released"])

    def test_receipt_does_not_copy_untrusted_error_text(self):
        data = failure_record("invoke", RuntimeError("a private workbook/path"))
        self.assertEqual(data, {"phase": "invoke", "type": "RuntimeError"})

    def test_compile_dialog_requires_exact_owner_label_and_unique_button(self):
        labels = {10: "Microsoft Visual Basic for Applications",
                  11: "Compile error:\n\nSub or Function not defined", 12: "OK", 13: "OK"}
        children = [11, 12]
        owner = SimpleNamespace(assert_window=lambda hwnd: None, matches=lambda hwnd: True)
        owner.windows = SimpleNamespace(
            con=SimpleNamespace(GW_OWNER=4),
            gui=SimpleNamespace(
                GetClassName=lambda hwnd: {10: "#32770", 11: "Static", 12: "Button", 13: "Button"}[hwnd],
                GetWindow=lambda hwnd, command: 9,
                GetWindowText=labels.__getitem__,
                IsWindowVisible=lambda hwnd: True,
                IsWindowEnabled=lambda hwnd: True,
                GetParent=lambda hwnd: 10,
                GetDlgCtrlID=lambda hwnd: 2,
                EnumChildWindows=lambda hwnd, callback, arg: [callback(child, arg) for child in children]))
        self.assertEqual(check_synthetic_compile_dialog(owner, 10, 9), 12)
        self.assertIsNone(check_synthetic_compile_dialog(owner, 10, 8))
        original = labels[11]
        labels[11] = "Compile error: Syntax error"
        self.assertIsNone(check_synthetic_compile_dialog(owner, 10, 9))
        labels[11] = original
        children.append(13)
        self.assertIsNone(check_synthetic_compile_dialog(owner, 10, 9))


if __name__ == "__main__":
    unittest.main()
