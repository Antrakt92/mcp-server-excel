"""Opt-in Excel UI test runner. The controller never switches the input desktop."""

import argparse
import ctypes
from ctypes import wintypes
import hashlib
import json
from pathlib import Path
import struct
import subprocess
import sys
import threading
import time
import traceback
import uuid
from zipfile import ZipFile


class ProbeError(RuntimeError):
    """A stable, nonsecret failure code for the receipt."""


def require(condition, code):
    if not condition:
        raise ProbeError(code)


def failure_record(phase, error):
    result = {"phase": phase, "type": type(error).__name__}
    if isinstance(error, ProbeError):
        result["code"] = str(error)
    for name in ("winerror", "hresult"):
        code = getattr(error, name, None)
        if type(code) is int:
            result[name] = code
    if isinstance(error, ExceptionGroup):
        result["errors"] = [failure_record(phase, item) for item in error.exceptions]
    own_frames = [frame for frame in traceback.extract_tb(error.__traceback__)
                  if Path(frame.filename).name == "isolated_excel_ui.py" and frame.name != "require"]
    if own_frames:
        frame = own_frames[-1]
        result["operation"] = frame.name
        result["line"] = frame.lineno
    return result


def attempt_cleanup(failures, phase, action):
    try:
        action()
    except Exception as error:
        failures.append(failure_record(phase, error))


def bounded_wait(read, seconds, code):
    deadline = time.monotonic() + seconds
    while True:
        result = read()
        if result:
            return result
        require(time.monotonic() < deadline, code)
        time.sleep(0.025)


def write_json(path, result):
    with path.open("x", encoding="utf-8") as stream:
        json.dump(result, stream, indent=2)
        stream.write("\n")


def write_startup_book(path):
    """Supply a known blank workbook so Excel has an EXCEL7 client at startup."""
    parts = {
        "[Content_Types].xml":
            '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            '<Override PartName="/xl/workbook.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
            '<Override PartName="/xl/worksheets/sheet1.xml" '
            'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
            '</Types>',
        "_rels/.rels":
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" '
            'Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml":
            '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
            'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
            '<sheets><sheet name="Origin" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels":
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" '
            'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
            'Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml":
            '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
            '<sheetData/></worksheet>',
    }
    with ZipFile(path, "x") as archive:
        for name, value in parts.items():
            archive.writestr(name, value)


def bounded_points(points, width, height):
    require(1 <= len(points) <= 16, "point_count_out_of_bounds")
    require(all(type(x) is int and type(y) is int and 0 <= x < width and 0 <= y < height
                for x, y in points), "point_outside_window")
    return tuple(points)


def color_state(colors, allowed):
    if 0xFFFFFFFF in colors:
        return "invalid"
    return "expected" if tuple(colors) in allowed else "other"


MENU_CAPTION = "ExcelMcp Synthetic Navigation"
COMPILE_MODULE = "ExcelMcpProbe"
MENU_FORM = "ExcelMcpProbeMenu"
OPEN_SHAPE = "OpenSyntheticMenu"
TARGET_CONTROL = "MenuDestination"
SOURCE_SHEET = "Origin"
TARGET_SHEET = "Destination"
FORM_PACKAGES = ("form-baseline.xlsm", "form-code.xlsm", "form-designer.xlsm")


class Windows:
    def __init__(self):
        # Test tooling only. No new product dependency or package installation.
        import pythoncom
        import win32com.client
        import win32con
        import win32event
        import win32gui
        import win32process

        self.com = pythoncom
        self.dispatch = win32com.client.Dispatch
        self.con = win32con
        self.event = win32event
        self.gui = win32gui
        self.process = win32process
        self.user = ctypes.WinDLL("user32", use_last_error=True)
        self.ole = ctypes.WinDLL("oleacc", use_last_error=True)
        self.gdi = ctypes.WinDLL("gdi32", use_last_error=True)
        h, ptr = wintypes.HANDLE, ctypes.c_void_p
        signatures = (
            (self.user, "CreateDesktopW", h, [wintypes.LPCWSTR, ptr, ptr, wintypes.DWORD, wintypes.DWORD, ptr]),
            (self.user, "OpenInputDesktop", h, [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]),
            (self.user, "GetThreadDesktop", h, [wintypes.DWORD]),
            (self.user, "CloseDesktop", wintypes.BOOL, [h]),
            (self.user, "GetUserObjectInformationW", wintypes.BOOL,
             [h, ctypes.c_int, ptr, wintypes.DWORD, ctypes.POINTER(wintypes.DWORD)]),
            (self.user, "GetDC", wintypes.HDC, [wintypes.HWND]),
            (self.user, "ReleaseDC", ctypes.c_int, [wintypes.HWND, wintypes.HDC]),
            (self.gdi, "GetPixel", wintypes.DWORD, [wintypes.HDC, ctypes.c_int, ctypes.c_int]),
            (self.gdi, "CreateCompatibleDC", wintypes.HDC, [wintypes.HDC]),
            (self.gdi, "CreateCompatibleBitmap", h, [wintypes.HDC, ctypes.c_int, ctypes.c_int]),
            (self.gdi, "SelectObject", h, [wintypes.HDC, h]),
            (self.gdi, "BitBlt", wintypes.BOOL, [wintypes.HDC, ctypes.c_int, ctypes.c_int,
             ctypes.c_int, ctypes.c_int, wintypes.HDC, ctypes.c_int, ctypes.c_int, wintypes.DWORD]),
            (self.gdi, "GetDIBits", ctypes.c_int, [wintypes.HDC, h, wintypes.UINT, wintypes.UINT,
             ptr, ptr, wintypes.UINT]),
            (self.gdi, "DeleteObject", wintypes.BOOL, [h]),
            (self.gdi, "DeleteDC", wintypes.BOOL, [wintypes.HDC]),
            (self.ole, "AccessibleObjectFromWindow", ctypes.c_long,
             [wintypes.HWND, wintypes.DWORD, ptr, ctypes.POINTER(ptr)]),
        )
        for dll, name, result, arguments in signatures:
            method = getattr(dll, name)
            method.restype, method.argtypes = result, arguments

    def desktop_name(self, handle):
        require(handle, "desktop_handle_unavailable")
        name = ctypes.create_unicode_buffer(256)
        needed = wintypes.DWORD()
        if not self.user.GetUserObjectInformationW(handle, 2, name, ctypes.sizeof(name), ctypes.byref(needed)):
            raise ctypes.WinError(ctypes.get_last_error())
        return name.value

    def input_desktop_name(self):
        handle = self.user.OpenInputDesktop(0, False, 1)
        try:
            return self.desktop_name(handle)
        finally:
            if handle:
                require(self.user.CloseDesktop(handle), "input_desktop_handle_not_closed")

    def own_thread_desktop(self):
        kernel = ctypes.WinDLL("kernel32")
        return self.desktop_name(self.user.GetThreadDesktop(kernel.GetCurrentThreadId()))

    def launch(self, arguments, desktop, cwd):
        info = self.process.STARTUPINFO()
        info.lpDesktop = "WinSta0\\" + desktop
        process, thread, pid, _ = self.process.CreateProcess(
            str(arguments[0]), subprocess.list2cmdline(list(map(str, arguments))), None, None, False,
            self.con.CREATE_NO_WINDOW, None, str(cwd), info)
        thread.Close()
        return process, pid

    def filetime(self, process):
        kernel = ctypes.WinDLL("kernel32", use_last_error=True)
        call = kernel.GetProcessTimes
        filetime = ctypes.POINTER(wintypes.FILETIME)
        call.argtypes = [wintypes.HANDLE, filetime, filetime, filetime, filetime]
        call.restype = wintypes.BOOL
        created, exited, used_kernel, used_user = (wintypes.FILETIME() for _ in range(4))
        if not call(int(process), ctypes.byref(created), ctypes.byref(exited),
                    ctypes.byref(used_kernel), ctypes.byref(used_user)):
            raise ctypes.WinError(ctypes.get_last_error())
        return (created.dwHighDateTime << 32) | created.dwLowDateTime

    def native_window(self, window):
        iid = (ctypes.c_ubyte * 16).from_buffer_copy(uuid.UUID("00020400-0000-0000-C000-000000000046").bytes_le)
        pointer = ctypes.c_void_p()
        status = self.ole.AccessibleObjectFromWindow(window, 0xFFFFFFF0, iid, ctypes.byref(pointer))
        require(status == 0 and pointer.value, "excel_native_object_unavailable")
        try:
            return self.dispatch(self.com.ObjectFromAddress(pointer.value, self.com.IID_IDispatch))
        finally:
            vtable = ctypes.cast(pointer, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
            ctypes.WINFUNCTYPE(ctypes.c_ulong, ctypes.c_void_p)(vtable[2])(pointer)


class ExcelOwner:
    def __init__(self, windows, desktop, identity):
        self.windows = windows
        self.desktop = desktop
        self.pid = identity["processId"]
        self.created = int(identity["startedAtUtcFileTime"])
        require(windows.own_thread_desktop() == desktop and windows.input_desktop_name() != desktop,
                "worker_not_on_private_desktop")
        import win32api
        self.handle = win32api.OpenProcess(0x101000, False, self.pid)
        try:
            self.assert_process()
        except Exception:
            self.handle.Close()
            raise

    def assert_process(self):
        api = self.windows
        require(api.event.WaitForSingleObject(self.handle, 0) == api.event.WAIT_TIMEOUT
                and api.filetime(self.handle) == self.created, "excel_process_identity_changed")

    def matches(self, hwnd):
        api = self.windows
        if not api.gui.IsWindow(hwnd):
            return False
        thread, owner = api.process.GetWindowThreadProcessId(hwnd)
        if owner != self.pid:
            return False
        return api.desktop_name(api.user.GetThreadDesktop(thread)) == self.desktop

    def assert_window(self, hwnd):
        self.assert_process()
        require(self.windows.input_desktop_name() != self.desktop, "private_desktop_became_input_desktop")
        require(self.matches(hwnd), "target_window_identity_changed")

    def top_windows(self, class_name=None):
        api = self.windows
        self.assert_process()
        windows = []

        def visit(hwnd, _):
            if self.matches(hwnd) and api.gui.IsWindowVisible(hwnd):
                if class_name is None or api.gui.GetClassName(hwnd) == class_name:
                    windows.append(hwnd)
            return True

        api.gui.EnumWindows(visit, None)
        return windows

    def client(self, window):
        self.assert_window(window)
        api = self.windows
        clients = []

        def visit(hwnd, _):
            if self.matches(hwnd) and api.gui.GetClassName(hwnd) == "EXCEL7" and api.gui.IsWindowVisible(hwnd):
                clients.append(hwnd)
            return True

        api.gui.EnumChildWindows(window, visit, None)
        require(len(clients) == 1, "excel_document_client_ambiguous")
        self.assert_window(clients[0])
        return clients[0]

    def click(self, hwnd, point):
        api = self.windows
        self.assert_window(hwnd)
        _, _, width, height = api.gui.GetClientRect(hwnd)
        x, y = bounded_points((point,), width, height)[0]
        require(x <= 32767 and y <= 32767, "click_message_coordinate_out_of_bounds")
        packed = x | (y << 16)
        for message, flags in ((api.con.WM_MOUSEMOVE, 0), (api.con.WM_LBUTTONDOWN, 1), (api.con.WM_LBUTTONUP, 0)):
            self.assert_window(hwnd)
            api.gui.PostMessage(hwnd, message, flags, packed)

    def move(self, hwnd, x, y):
        self.assert_window(hwnd)
        api = self.windows
        _, _, width, height = api.gui.GetClientRect(hwnd)
        bounded_points(((x, y),), width, height)
        require(x <= 32767 and y <= 32767, "hover_message_coordinate_out_of_bounds")
        api.gui.PostMessage(hwnd, api.con.WM_MOUSEMOVE, 0, x | (y << 16))

    def escape(self, hwnd):
        self.assert_window(hwnd)
        api = self.windows
        api.gui.PostMessage(hwnd, api.con.WM_KEYDOWN, api.con.VK_ESCAPE, 1)

    def check_hit(self, hwnd, point):
        self.assert_window(hwnd)
        api = self.windows
        screen_point = api.gui.ClientToScreen(hwnd, point)
        hit = api.gui.WindowFromPoint(screen_point)
        require(hit and self.matches(hit) and (hit == hwnd or api.gui.IsChild(hwnd, hit)),
                "native_point_not_in_target_window")
        return hit, api.gui.ScreenToClient(hit, screen_point)

    def get_pixels(self, hwnd, points, read):
        self.assert_window(hwnd)
        api = self.windows
        _, _, width, height = api.gui.GetClientRect(hwnd)
        selected = bounded_points(points, width, height)
        device = api.user.GetDC(hwnd)
        require(device, "window_device_context_unavailable")
        try:
            value = read(lambda: tuple(api.gdi.GetPixel(device, x, y) for x, y in selected))
        except Exception as primary:
            try:
                require(api.user.ReleaseDC(hwnd, device) == 1, "window_device_context_not_released")
            except Exception as cleanup:
                raise ExceptionGroup("pixel probe and cleanup failed", [primary, cleanup]) from None
            raise
        require(api.user.ReleaseDC(hwnd, device) == 1, "window_device_context_not_released")
        return value

    def read_pixels(self, hwnd, points):
        return self.get_pixels(hwnd, points, lambda read: read())

    def sample(self, hwnd, points, allowed, seconds=0.35, period=0.01, on_tick=None, calibration=None):
        require(0 < seconds <= 3 and 0.005 <= period <= 0.25, "sampling_budget_out_of_bounds")
        require(1 <= len(allowed) <= 8 and all(len(x) == len(points) for x in allowed),
                "sample_color_states_invalid")
        require(calibration is None or len(calibration) == len(points) and calibration not in allowed,
                "calibration_color_state_invalid")

        def collect(read):
            count = {"expected": 0, "other": 0, "invalid": 0}
            expected_states = {value: 0 for value in allowed}
            start = time.monotonic()
            first, last = None, None
            calibration_samples = 0
            while time.monotonic() - start < seconds:
                self.assert_window(hwnd)
                if on_tick is not None:
                    on_tick(sum(count.values()))
                colors = read()
                first = colors if first is None else first
                last = colors
                state = color_state(colors, allowed)
                count[state] += 1
                if state == "expected":
                    expected_states[colors] += 1
                elif colors == calibration:
                    calibration_samples += 1
                time.sleep(period)
            require(sum(count.values()) > 0, "no_samples_observed")
            return {"counts": count, "first": first, "last": last,
                    "calibrationSamples": calibration_samples,
                    "expectedStates": [{"colors": colors, "samples": samples}
                                       for colors, samples in sorted(expected_states.items())],
                    "durationMs": round((time.monotonic() - start) * 1000)}

        return self.get_pixels(hwnd, points, collect)

    def capture_control(self, hwnd, path):
        self.assert_window(hwnd)
        api = self.windows
        _, _, width, height = api.gui.GetClientRect(hwnd)
        require(1 <= width <= 800 and 1 <= height <= 600, "control_frame_dimensions_out_of_bounds")
        device = api.user.GetDC(hwnd)
        require(device, "control_frame_device_unavailable")
        memory = bitmap = previous = None
        errors = []

        def restore_selection():
            require(api.gdi.SelectObject(memory, previous), "control_frame_selection_not_restored")

        try:
            memory = api.gdi.CreateCompatibleDC(device)
            require(memory, "control_frame_memory_device_unavailable")
            bitmap = api.gdi.CreateCompatibleBitmap(device, width, height)
            require(bitmap, "control_frame_bitmap_unavailable")
            previous = api.gdi.SelectObject(memory, bitmap)
            require(previous, "control_frame_bitmap_not_selected")
            self.assert_window(hwnd)
            require(api.gdi.BitBlt(memory, 0, 0, width, height, device, 0, 0, 0x00CC0020),
                    "control_frame_copy_failed")
            restore_selection()
            previous = None
            pixels = ctypes.create_string_buffer(width * height * 4)
            header = ctypes.create_string_buffer(struct.pack(
                "<IiiHHIIiiII", 40, width, -height, 1, 32, 0, len(pixels), 0, 0, 0, 0))
            require(api.gdi.GetDIBits(device, bitmap, 0, height, pixels, header, 0) == height,
                    "control_frame_bitmap_read_incomplete")
            self.assert_window(hwnd)
            with path.open("xb") as stream:
                stream.write(struct.pack("<2sIHHI", b"BM", 54 + len(pixels), 0, 0, 54))
                stream.write(header.raw[:40])
                stream.write(pixels.raw)
        except Exception as error:
            errors.append(error)
        finally:
            released = []
            if previous:
                attempt_cleanup(released, "restore_control_frame_selection", restore_selection)
            if memory and not api.gdi.DeleteDC(memory):
                released.append({"phase": "delete_control_frame_device"})
            if bitmap and not api.gdi.DeleteObject(bitmap):
                released.append({"phase": "delete_control_frame_bitmap"})
            if api.user.ReleaseDC(hwnd, device) != 1:
                released.append({"phase": "release_control_frame_device"})
            errors.extend(ProbeError(item["phase"]) for item in released)
        if len(errors) > 1:
            raise ExceptionGroup("control frame and cleanup failed", errors)
        if errors:
            raise errors[0]
        return {"width": width, "height": height, "sha256": hashlib.sha256(path.read_bytes()).hexdigest()}

    def assert_active(self, app, expected_hwnd, expected_book, expected_sheet=None):
        self.assert_window(expected_hwnd)
        require(app.ActiveWindow.Hwnd == expected_hwnd, "active_excel_window_changed")
        require(Path(app.ActiveWorkbook.FullName).resolve() == expected_book.resolve(), "active_workbook_changed")
        if expected_sheet is not None:
            require(app.ActiveSheet.Name == expected_sheet, "active_sheet_changed")


def prepare_macro_fixture(owner, app, directory, add_menu, compile_error=False):
    """Author code only in the freshly created workbook on the private desktop."""
    app.DisplayAlerts = False
    app.EnableEvents = False
    workbook = app.ActiveWorkbook
    workbook.EnableAutoRecover = False
    path = directory / "fixture.xlsm"
    workbook.SaveAs(str(path), 52)
    owner.assert_active(app, app.ActiveWindow.Hwnd, path, SOURCE_SHEET)
    project = workbook.VBProject
    project.Name = "ExcelMcpProbeProject"
    component = project.VBComponents.Add(1)
    component.Name = COMPILE_MODULE
    if not add_menu:
        source = ("Option Explicit\r\nPublic Sub SyntheticCompileError()\r\n"
                  "    ExcelMcpSyntheticUndefinedProcedure\r\nEnd Sub") if compile_error else (
                      "Option Explicit\r\nPublic Function SyntheticValue() As Long\r\n"
                      "    SyntheticValue = 73\r\nEnd Function")
        component.CodeModule.AddFromString(source)
        return workbook, path, project, component

    worksheet = workbook.Worksheets(SOURCE_SHEET)
    destination = workbook.Worksheets.Add(After=worksheet)
    destination.Name = TARGET_SHEET
    worksheet.Activate()
    component.CodeModule.AddFromString(
        "Option Explicit\r\nPublic Sub ShowSyntheticMenu()\r\n"
        f"    {MENU_FORM}.Show vbModeless\r\nEnd Sub")
    form = project.VBComponents.Add(3)
    form.Name = MENU_FORM
    properties = form.Properties
    properties.Item("Caption").Value = MENU_CAPTION
    properties.Item("Width").Value = 220
    properties.Item("Height").Value = 115
    button = form.Designer.Controls.Add("Forms.CommandButton.1", TARGET_CONTROL, True)
    button.Caption = "Go to Destination"
    button.Left, button.Top, button.Width, button.Height = 20, 20, 170, 32
    cancel = form.Designer.Controls.Add("Forms.CommandButton.1", "CancelMenu", True)
    cancel.Caption = "Cancel"
    cancel.Cancel = True
    cancel.Left, cancel.Top, cancel.Width, cancel.Height = 20, 62, 170, 20
    form.CodeModule.AddFromString(
        f"Option Explicit\r\nPrivate Sub {TARGET_CONTROL}_Click()\r\n"
        f'    ThisWorkbook.Worksheets("{SOURCE_SHEET}").Range("A1").Value = "{TARGET_CONTROL}"\r\n'
        f'    ThisWorkbook.Worksheets("{TARGET_SHEET}").Activate\r\n'
        "    Unload Me\r\nEnd Sub\r\n"
        "Private Sub CancelMenu_Click()\r\n"
        f'    ThisWorkbook.Worksheets("{SOURCE_SHEET}").Range("A1").Value = "CancelMenu"\r\n'
        "    Unload Me\r\nEnd Sub\r\n"
        f"Private Sub {TARGET_CONTROL}_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, "
        "ByVal X As Single, ByVal Y As Single)\r\n"
        "    Me.BackColor = RGB(140, 185, 230)\r\nEnd Sub\r\n"
        "Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, "
        "ByVal X As Single, ByVal Y As Single)\r\n"
        "    If X < 11 Then\r\n"
        "        Me.BackColor = RGB(91, 91, 91)\r\n"
        "    Else\r\n"
        "        Me.BackColor = RGB(245, 230, 190)\r\n"
        "    End If\r\nEnd Sub\r\n")
    anchor = worksheet.Range("C4")
    shape = worksheet.Shapes.AddShape(1, anchor.Left, anchor.Top, 72, 24)
    shape.Name = OPEN_SHAPE
    shape.TextFrame.Characters().Text = "Open menu"
    shape.OnAction = f"'{workbook.Name}'!{COMPILE_MODULE}.ShowSyntheticMenu"
    workbook.Save()
    return workbook, path, project, component


def check_synthetic_compile_dialog(owner, hwnd, vbe_window):
    api = owner.windows
    owner.assert_window(hwnd)
    if (api.gui.GetClassName(hwnd) != "#32770" or api.gui.GetWindow(hwnd, api.con.GW_OWNER) != vbe_window
            or api.gui.GetWindowText(hwnd) != "Microsoft Visual Basic for Applications"):
        return None
    labels, buttons = [], []

    def visit(child, _):
        if owner.matches(child) and api.gui.IsWindowVisible(child):
            kind = api.gui.GetClassName(child)
            if kind == "Static":
                labels.append(api.gui.GetWindowText(child))
            if kind == "Button" and api.gui.GetParent(child) == hwnd:
                buttons.append(child)
        return True

    api.gui.EnumChildWindows(hwnd, visit, None)
    text = " ".join(" ".join(labels).split())
    if text != "Compile error: Sub or Function not defined":
        return None
    okay = [button for button in buttons if api.gui.GetWindowText(button) == "OK"
            and api.gui.IsWindowEnabled(button) and api.gui.GetDlgCtrlID(button) == 2]
    return okay[0] if len(okay) == 1 else None


def compile_project(owner, app, path, project, component, directory, expect_error=False):
    api = owner.windows
    app_hwnd = app.ActiveWindow.Hwnd
    owner.assert_active(app, app_hwnd, path, SOURCE_SHEET)
    vbe = project.VBE
    vbe.ActiveVBProject = project
    component.CodeModule.CodePane.Show()
    vbe_window = vbe.MainWindow.HWnd
    owner.assert_window(vbe_window)

    def assert_selected():
        owner.assert_active(app, app_hwnd, path, SOURCE_SHEET)
        owner.assert_window(vbe_window)
        active = vbe.ActiveVBProject
        require(active.Name == project.Name and Path(active.FileName).resolve() == path.resolve(),
                "compile_project_selection_changed")
        require(vbe.ActiveCodePane.CodeModule.Parent.Name == component.Name,
                "compile_code_pane_selection_changed")

    def read_control():
        assert_selected()
        command = vbe.CommandBars.FindControl(1, 578)
        require(command is not None and command.Id == 578, "vbe_compile_command_unavailable")
        return command

    command = read_control()
    require(command.Enabled, "compile_initial_availability_unconfirmed")
    dialogs_before = set(owner.top_windows("#32770"))
    observed = []
    stop = threading.Event()
    ready = threading.Event()
    observed_done = threading.Event()

    def observe():
        try:
            require(api.own_thread_desktop() == owner.desktop, "compile_observer_desktop_changed")
            ready.set()
            while not stop.is_set():
                new = set(owner.top_windows("#32770")) - dialogs_before
                if new:
                    outcome = "unconfirmed"
                    if expect_error and len(new) == 1:
                        hwnd = next(iter(new))
                        button = check_synthetic_compile_dialog(owner, hwnd, vbe_window)
                        if button is not None:
                            owner.assert_window(vbe_window)
                            owner.assert_window(hwnd)
                            owner.assert_window(button)
                            require(check_synthetic_compile_dialog(owner, hwnd, vbe_window) == button,
                                    "compile_dialog_identity_changed")
                            outcome = "compileError"
                            api.gui.PostMessage(button, api.con.BM_CLICK, 0, 0)
                    observed.append(outcome)
                    write_json(directory / "compile-dialog.json",
                               {"outcome": outcome, "newDialogCount": len(new), "windows": sorted(new)})
                    return
                stop.wait(0.025)
        except Exception as error:
            observed.append("unconfirmed")
            write_json(directory / "compile-dialog.json",
                       {"outcome": "unconfirmed", "failure": failure_record("observe_compile_dialog", error)})
        finally:
            observed_done.set()

    observer = threading.Thread(target=observe, daemon=True)
    observer.start()
    try:
        require(ready.wait(2), "compile_observer_not_ready")
        command.Execute()
        if expect_error:
            require(observed_done.wait(5), "compile_dialog_observation_timed_out")
    finally:
        stop.set()
        observer.join(2)
    require(not observer.is_alive(), "compile_observer_unconfirmed")
    if expect_error:
        require(observed == ["compileError"], "compile_error_unconfirmed")
        bounded_wait(lambda: set(owner.top_windows("#32770")) == dialogs_before, 5, "compile_dialog_still_open")
        assert_selected()
        return {"outcome": "compileError", "projectBound": True, "dialogClosed": True, "hwnd": vbe_window}
    require(not observed and set(owner.top_windows("#32770")) == dialogs_before, "compile_result_unconfirmed")
    require(not read_control().Enabled, "compile_result_unconfirmed")
    require(not read_control().Enabled, "compiled_state_unconfirmed")
    return {"outcome": "compiled", "repeatOutcome": "alreadyCompiled",
            "projectBound": True, "hwnd": vbe_window}


def run_navigation(owner, app, directory, workbook, path):
    api = owner.windows
    first_window = app.ActiveWindow
    second_window = workbook.NewWindow()
    require(first_window.Hwnd != second_window.Hwnd, "distinct_excel_windows_unconfirmed")
    actions = []

    def active_menu():
        matches = [hwnd for hwnd in owner.top_windows() if api.gui.GetWindowText(hwnd) == MENU_CAPTION]
        require(len(matches) <= 1, "synthetic_menu_window_ambiguous")
        return matches[0] if matches else None

    def open_menu(window):
        window.Activate()
        workbook.Worksheets(SOURCE_SHEET).Activate()
        marker = workbook.Worksheets(SOURCE_SHEET).Range("A1")
        marker.ClearContents()
        frame = window.Hwnd
        owner.assert_active(app, frame, path, SOURCE_SHEET)
        require(active_menu() is None, "synthetic_menu_already_visible")
        client = owner.client(frame)
        shape = app.ActiveSheet.Shapes.Item(OPEN_SHAPE)
        x = window.PointsToScreenPixelsX(round(shape.Left + shape.Width / 2))
        y = window.PointsToScreenPixelsY(round(shape.Top + shape.Height / 2))
        point = api.gui.ScreenToClient(client, (x, y))
        owner.check_hit(client, point)
        owner.click(client, point)
        menu = bounded_wait(active_menu, 5, "synthetic_menu_did_not_open")
        require(api.gui.GetWindow(menu, api.con.GW_OWNER) == frame, "menu_not_owned_by_active_workbook_window")
        return frame, menu, marker

    def select_destination(window):
        frame, menu, marker = open_menu(window)
        _, _, width, height = api.gui.GetClientRect(menu)
        # The test form declares the destination in the upper button row.
        target = (width // 2, height // 3)
        native_target, target_point = owner.check_hit(menu, target)
        require(native_target == menu or not api.gui.GetWindowText(native_target)
                or api.gui.GetWindowText(native_target) == "Go to Destination", "native_control_caption_changed")
        trace = None
        if not actions:
            sample_point = (7, 10)
            gold_target, gold_point = owner.check_hit(menu, (27, 10))
            gray_target, gray_point = owner.check_hit(menu, sample_point)
            gold = (245 | 230 << 8 | 190 << 16,)
            blue = (140 | 185 << 8 | 230 << 16,)
            gray = (91 | 91 << 8 | 91 << 16,)
            owner.move(gold_target, *gold_point)
            bounded_wait(lambda: owner.read_pixels(menu, (sample_point,)) == gold,
                         2, "menu_hover_reference_color_not_observed")
            calibration_times = []

            def change_hover(index):
                if index == 6:
                    owner.move(gray_target, *gray_point)
                    calibration_times.append(time.monotonic())
                elif index == 10:
                    owner.move(gold_target, *gold_point)
                    calibration_times.append(time.monotonic())
                elif index == 15:
                    owner.move(native_target, *target_point)

            trace = owner.sample(menu, (sample_point,), {gold, blue}, on_tick=change_hover, calibration=gray)
            if len(calibration_times) == 2:
                trace["calibrationDispatchIntervalMs"] = round((calibration_times[1] - calibration_times[0]) * 1000)
            write_json(directory / "hover-control.json", trace)
            require(all(state["samples"] > 0 for state in trace["expectedStates"]), "menu_hover_color_not_observed")
            require(trace["first"] == gold and trace["last"] == blue and trace["calibrationSamples"] > 0
                    and 0 < trace.get("calibrationDispatchIntervalMs", 0) <= 150,
                    "menu_transient_calibration_not_observed")
            trace["controlFrame"] = owner.capture_control(menu, directory / "hover-control.bmp")
        owner.click(native_target, target_point)
        bounded_wait(lambda: active_menu() is None, 5, "synthetic_menu_did_not_close")
        owner.assert_active(app, frame, path, TARGET_SHEET)
        require(marker.Value == TARGET_CONTROL, "synthetic_named_control_did_not_run")
        actions.append({"window": frame, "shape": OPEN_SHAPE, "menuWindow": menu,
                        "control": TARGET_CONTROL, "sheet": TARGET_SHEET, "hoverTrace": trace})

    for window in (first_window, second_window, first_window):
        select_destination(window)
    frame, menu, marker = open_menu(second_window)
    owner.escape(menu)
    bounded_wait(lambda: active_menu() is None, 5, "synthetic_escape_did_not_close_menu")
    owner.assert_active(app, frame, path, SOURCE_SHEET)
    require(marker.Value == "CancelMenu", "synthetic_escape_handler_did_not_run")
    old_path = path
    path = path.with_name("renamed.xlsm")
    workbook.SaveAs(str(path), 52)
    try:
        owner.assert_active(app, frame, old_path)
    except ProbeError as error:
        require(str(error) == "active_workbook_changed", "save_as_identity_unconfirmed")
    else:
        raise ProbeError("save_as_accepted_stale_workbook_identity")
    client = owner.client(frame)
    rebound_app = api.native_window(client).Application
    owner.assert_active(rebound_app, frame, path, SOURCE_SHEET)
    shape = workbook.Worksheets(SOURCE_SHEET).Shapes.Item(OPEN_SHAPE)
    shape.OnAction = f"'{workbook.Name}'!{COMPILE_MODULE}.ShowSyntheticMenu"
    select_destination(second_window)
    return {"actions": actions, "escapeConfirmed": True, "stalePathRejectedAfterSaveAs": True}


def read_form_code(component):
    module = component.CodeModule
    return module.Lines(1, module.CountOfLines).replace("\r\n", "\n").rstrip("\n")


def read_form_state(component):
    return read_form_code(component), component.Designer.Caption


def create_form_packages(owner, app, directory, workbook, path, project):
    app.AutomationSecurity = 3
    component = project.VBComponents.Item(MENU_FORM)
    baseline, caption = read_form_state(component)
    module = component.CodeModule
    baseline_lines = module.CountOfLines
    baseline_path, code_path, designer_path = (directory / name for name in FORM_PACKAGES)
    owner.assert_active(app, app.ActiveWindow.Hwnd, path, SOURCE_SHEET)
    workbook.SaveCopyAs(str(baseline_path))

    procedure = "Private Sub SyntheticPackageWitness()\r\nEnd Sub"
    module.InsertLines(baseline_lines + 1, procedure)
    edited, unchanged_caption = read_form_state(component)
    require(edited.startswith(baseline) and edited[len(baseline):].strip() == procedure.replace("\r\n", "\n")
            and unchanged_caption == caption, "form_code_change_readback_failed")
    workbook.SaveCopyAs(str(code_path))
    module.DeleteLines(baseline_lines + 1, module.CountOfLines - baseline_lines)
    require(read_form_state(component) == (baseline, caption), "form_code_restore_readback_failed")

    designer_caption = caption + " (designer witness)"
    component.Designer.Caption = designer_caption
    require(read_form_state(component) == (baseline, designer_caption), "form_designer_change_readback_failed")
    workbook.SaveCopyAs(str(designer_path))
    workbook.Close(False)

    expected = ((baseline_path, baseline, caption), (code_path, edited, caption),
                (designer_path, baseline, designer_caption))
    snapshots = []
    for snapshot_path, expected_source, expected_caption in expected:
        owner.assert_process()
        require(snapshot_path.is_file() and not snapshot_path.is_symlink(), "form_package_missing")
        selected = app.Workbooks.Open(str(snapshot_path), UpdateLinks=0, ReadOnly=True)
        owner.assert_active(app, app.ActiveWindow.Hwnd, snapshot_path)
        require(selected.ReadOnly, "reopened_form_package_not_read_only")
        selected_component = selected.VBProject.VBComponents.Item(MENU_FORM)
        require(read_form_state(selected_component) == (expected_source, expected_caption),
                "reopened_form_package_readback_failed")
        selected.Close(False)
        snapshots.append({"file": snapshot_path.name, "sha256": hashlib.sha256(snapshot_path.read_bytes()).hexdigest(),
                          "reopenedReadOnly": True, "codeVerified": True, "designerCaptionVerified": True})
    return {"snapshots": snapshots}


def connect_startup(owner, path):
    def find_frame():
        frames = owner.top_windows("XLMAIN")
        require(len(frames) <= 1, "startup_excel_window_ambiguous")
        if frames:
            try:
                return frames[0], owner.client(frames[0])
            except ProbeError as error:
                if str(error) != "excel_document_client_ambiguous":
                    raise
        return None

    frame, client = bounded_wait(find_frame, 20, "startup_excel_window_unavailable")
    owner.assert_window(client)
    app = owner.windows.native_window(client).Application
    owner.assert_active(app, frame, path, "Origin")
    require(app.Workbooks.Count == 1, "unexpected_workbooks_in_private_excel")
    return app


def worker(directory, desktop, scenario):
    api = Windows()
    report = {"nativeObservationPassed": False, "primaryFailure": None, "cleanupFailures": []}
    owner = None
    app = None
    phase = "initialize"
    api.com.CoInitialize()
    try:
        require(api.own_thread_desktop() == desktop and api.input_desktop_name() != desktop,
                "worker_not_on_private_desktop")
        write_json(directory / "worker-desktop.json", {"desktop": desktop, "isolated": True})
        def read_identity():
            try:
                return json.loads((directory / "identity.json").read_text(encoding="utf-8"))
            except (FileNotFoundError, json.JSONDecodeError):
                return None

        identity = bounded_wait(read_identity, 15, "excel_launch_identity_unavailable")
        owner = ExcelOwner(api, desktop, identity)
        phase = "bind_startup"
        app = connect_startup(owner, directory / "startup.xlsx")
        write_json(directory / "worker-ready.json", {"excelIdentityVerified": True, "privateDesktopVerified": True})
        phase = scenario
        report["observation"] = run_scenario(owner, app, directory, scenario)
        report["nativeObservationPassed"] = True
    except Exception as error:
        report["primaryFailure"] = failure_record(phase, error)
        write_json(directory / "worker-primary.json", report["primaryFailure"])
    finally:
        if app is not None:
            def close_books():
                owner.assert_process()
                owned_names = {"startup.xlsx", "fixture.xlsm", "renamed.xlsm", *FORM_PACKAGES}
                books = [app.Workbooks.Item(index + 1) for index in range(app.Workbooks.Count)]
                require(all(Path(book.FullName).resolve().parent == directory
                            and book.Name in owned_names for book in books), "unknown_workbook_blocks_cleanup")
                for book in books:
                    book.Close(False)

            def quit_app():
                owner.assert_process()
                require(app.Workbooks.Count == 0, "workbooks_remain_open")
                app.Quit()

            attempt_cleanup(report["cleanupFailures"], "close_workbooks", close_books)
            attempt_cleanup(report["cleanupFailures"], "quit_excel", quit_app)
            app = None
        if owner is not None:
            attempt_cleanup(report["cleanupFailures"], "close_worker_process_handle", owner.handle.Close)
        attempt_cleanup(report["cleanupFailures"], "release_com", api.com.CoUninitialize)
        write_json(directory / "worker-result.json", report)
    return 0 if report["nativeObservationPassed"] and not report["cleanupFailures"] else 1


def run_scenario(owner, app, directory, scenario):
    # The startup screen itself provides the first isolation and pixel witness.
    path = directory / "startup.xlsx"
    hwnd = app.ActiveWindow.Hwnd
    owner.assert_active(app, hwnd, path, "Origin")
    client = owner.client(hwnd)
    result = {"window": hwnd, "clientWindow": client, "startSheet": "Origin"}

    if scenario == "sample":
        points = ((20, 20), (24, 24))
        reference = owner.read_pixels(client, points)
        require(0xFFFFFFFF not in reference, "reference_sample_invalid")
        result["sample"] = owner.sample(client, points, {reference})
        require(result["sample"]["counts"]["expected"] > 0, "reference_colors_not_observed")
    if scenario in ("compile", "compile-error", "navigation", "form-packages"):
        compile_case = scenario in ("compile", "compile-error")
        workbook, path, project, component = prepare_macro_fixture(
            owner, app, directory, not compile_case, compile_error=scenario == "compile-error")
        if compile_case:
            result["compile"] = compile_project(
                owner, app, path, project, component, directory, expect_error=scenario == "compile-error")
        elif scenario == "navigation":
            result["navigation"] = run_navigation(owner, app, directory, workbook, path)
        else:
            result["formPackages"] = create_form_packages(owner, app, directory, workbook, path, project)
    return result


def controller(excel, directory, scenario):
    api = Windows()
    directory.mkdir()
    report = {"schemaVersion": 1, "passed": False, "scenario": scenario, "primaryFailure": None,
              "cleanupFailures": [], "workerExitCode": None, "excelExitCode": None,
              "inputDesktopUnchanged": False}
    input_desktop = api.input_desktop_name()
    desktop = "ExcelMcpTest_" + uuid.uuid4().hex
    handle = None
    excel_process = None
    worker_process = None
    phase = "create_desktop"
    try:
        handle = api.user.CreateDesktopW(desktop, None, None, 0, 0x01FF, None)
        require(api.desktop_name(handle) == desktop and desktop != input_desktop, "private_desktop_unconfirmed")
        phase = "write_startup_fixture"
        startup = directory / "startup.xlsx"
        write_startup_book(startup)
        report["fixtureSha256"] = hashlib.sha256(startup.read_bytes()).hexdigest()
        phase = "launch_worker"
        worker_process, _ = api.launch(
            [sys.executable, Path(__file__).resolve(), "--worker", "--directory", directory,
             "--desktop", desktop, "--scenario", scenario], desktop, directory)
        phase = "verify_worker_desktop"

        def read_desktop():
            try:
                return json.loads((directory / "worker-desktop.json").read_text(encoding="utf-8"))
            except (FileNotFoundError, json.JSONDecodeError):
                require(api.event.WaitForSingleObject(worker_process, 0) == api.event.WAIT_TIMEOUT,
                        "private_desktop_worker_exited")
                return None

        reported = bounded_wait(read_desktop, 8, "private_desktop_worker_unavailable")
        require(reported == {"desktop": desktop, "isolated": True} and api.input_desktop_name() != desktop,
                "private_desktop_worker_mismatch")
        phase = "launch_excel"
        excel_process, pid = api.launch([excel, "/x", startup], desktop, directory)
        identity = {"processId": pid, "startedAtUtcFileTime": str(api.filetime(excel_process))}
        report["excel"] = identity
        write_json(directory / "identity.json", identity)
        phase = "wait_worker"
        require(api.event.WaitForSingleObject(worker_process, 90_000) == api.event.WAIT_OBJECT_0, "worker_deadline")
        report["workerExitCode"] = api.process.GetExitCodeProcess(worker_process)
        report["worker"] = json.loads((directory / "worker-result.json").read_text(encoding="utf-8"))
    except Exception as error:
        report["primaryFailure"] = failure_record(phase, error)
    finally:
        if excel_process is not None:
            def confirm_excel_exit():
                require(api.event.WaitForSingleObject(excel_process, 10_000) == api.event.WAIT_OBJECT_0,
                        "excel_exit_unconfirmed")
                report["excelExitCode"] = api.process.GetExitCodeProcess(excel_process)
                require(report["excelExitCode"] == 0, "excel_exit_nonzero")

            attempt_cleanup(report["cleanupFailures"], "confirm_excel_exit", confirm_excel_exit)
            attempt_cleanup(report["cleanupFailures"], "close_excel_process_handle", excel_process.Close)
        if worker_process is not None:
            attempt_cleanup(report["cleanupFailures"], "close_worker_process_handle", worker_process.Close)
        if handle:
            def close_desktop():
                require(api.user.CloseDesktop(handle), "private_desktop_handle_not_closed")

            attempt_cleanup(report["cleanupFailures"], "close_private_desktop_handle", close_desktop)
        try:
            report["inputDesktopUnchanged"] = api.input_desktop_name() == input_desktop
            require(report["inputDesktopUnchanged"], "input_desktop_changed")
        except Exception as error:
            report["cleanupFailures"].append(failure_record("verify_input_desktop", error))
        for filename, key in (("worker-ready.json", "workerReadiness"),
                              ("worker-primary.json", "workerPrimaryFailure"),
                              ("compile-dialog.json", "ownedCompileDialogObserved")):
            checkpoint = directory / filename
            if checkpoint.is_file():
                try:
                    report[key] = json.loads(checkpoint.read_text(encoding="utf-8"))
                except (OSError, ValueError) as error:
                    report["cleanupFailures"].append(failure_record("read_worker_checkpoint", error))
        report["passed"] = (
            report["primaryFailure"] is None and not report["cleanupFailures"] and report["workerExitCode"] == 0
            and report["excelExitCode"] == 0 and report["worker"]["nativeObservationPassed"]
            and not report["worker"]["cleanupFailures"])
        write_json(directory / "result.json", report)
    return 0 if report["passed"] else 1


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--excel", type=Path, help="Absolute path to installed EXCEL.EXE")
    parser.add_argument("--directory", type=Path, required=True, help="New, private working directory")
    parser.add_argument("--scenario",
                        choices=("isolation", "sample", "compile", "compile-error", "navigation", "form-packages"),
                        default="isolation")
    parser.add_argument("--worker", action="store_true", help=argparse.SUPPRESS)
    parser.add_argument("--desktop", help=argparse.SUPPRESS)
    args = parser.parse_args()
    if args.worker:
        require(args.desktop and args.desktop.startswith("ExcelMcpTest_"), "private_desktop_required")
        return worker(args.directory.resolve(), args.desktop, args.scenario)
    require(sys.platform == "win32", "windows_required")
    require(args.excel is not None and args.excel.is_absolute() and args.excel.name.casefold() == "excel.exe",
            "absolute_excel_executable_required")
    excel = args.excel.resolve(strict=True)
    require(excel.is_file(), "excel_executable_required")
    directory = args.directory.resolve()
    require(not directory.exists() and directory.parent.is_dir(), "new_working_directory_required")
    return controller(excel, directory, args.scenario)


if __name__ == "__main__":
    sys.exit(main())
