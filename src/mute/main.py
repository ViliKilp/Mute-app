from __future__ import annotations

import ctypes
import math
import struct
import sys
import wave
import winsound
import winreg
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

import comtypes
import psutil
import wx
import wx.adv
from pycaw.pycaw import AudioUtilities
from pynput import keyboard, mouse


APP_NAME = "Mute"
RUN_KEY_PATH = r"Software\Microsoft\Windows\CurrentVersion\Run"
ASSETS_DIR = Path(__file__).resolve().parent / "assets"


def _enable_dpi_awareness() -> None:
    try:
        ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4))
        return
    except Exception:
        pass
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass


@dataclass
class HotkeyState:
    listener: Optional[keyboard.GlobalHotKeys] = None
    mouse_listener: Optional[mouse.Listener] = None
    hotkey_text: str = "<pause>"


def _get_foreground_pid() -> Optional[int]:
    user32 = ctypes.windll.user32
    hwnd = user32.GetForegroundWindow()
    if not hwnd:
        return None
    pid = ctypes.c_ulong(0)
    user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return pid.value or None


def _toggle_mute_for_target(pid: int, proc_name: Optional[str]) -> Optional[bool]:
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process:
            if session.Process.pid == pid:
                volume = session.SimpleAudioVolume
                is_muted = bool(volume.GetMute())
                volume.SetMute(0 if is_muted else 1, None)
                return not is_muted
            if proc_name and session.Process.name().lower() == proc_name:
                volume = session.SimpleAudioVolume
                is_muted = bool(volume.GetMute())
                volume.SetMute(0 if is_muted else 1, None)
                return not is_muted
        if proc_name:
            display = (session.DisplayName or "").lower()
            if display and proc_name in display:
                volume = session.SimpleAudioVolume
                is_muted = bool(volume.GetMute())
                volume.SetMute(0 if is_muted else 1, None)
                return not is_muted
    return None


def toggle_active_app_mute() -> Optional[bool]:
    comtypes.CoInitialize()
    try:
        pid = _get_foreground_pid()
        if not pid:
            return None
        try:
            process = psutil.Process(pid)
        except psutil.Error:
            return None
        proc_name = process.name().lower() if process.name() else None
        return _toggle_mute_for_target(pid, proc_name)
    finally:
        comtypes.CoUninitialize()


def _get_autostart_command() -> str:
    python_exe = Path(sys.executable)
    pythonw = python_exe.with_name("pythonw.exe")
    exe = pythonw if pythonw.exists() else python_exe
    script_path = Path(__file__).resolve()
    return f'"{exe}" "{script_path}"'


def _set_autostart(enabled: bool) -> None:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_SET_VALUE) as key:
        if enabled:
            winreg.SetValueEx(key, APP_NAME, 0, winreg.REG_SZ, _get_autostart_command())
        else:
            try:
                winreg.DeleteValue(key, APP_NAME)
            except FileNotFoundError:
                pass


def _get_autostart_enabled() -> bool:
    with winreg.OpenKey(winreg.HKEY_CURRENT_USER, RUN_KEY_PATH, 0, winreg.KEY_READ) as key:
        try:
            winreg.QueryValueEx(key, APP_NAME)
            return True
        except FileNotFoundError:
            return False


def _normalize_hotkey(text: str) -> str:
    raw = text.strip().lower()
    if not raw:
        return "<pause>"
    if raw == "pause":
        return "<pause>"
    if raw in {"mouse4", "mouse5", "mouse1", "mouse2", "mouse3"}:
        return f"<{raw}>"
    if raw.startswith("<") and raw.endswith(">"):
        return raw
    return raw


def _key_to_hotkey(key: keyboard.Key | keyboard.KeyCode) -> Optional[str]:
    if isinstance(key, keyboard.KeyCode) and key.char:
        return key.char.lower()
    if isinstance(key, keyboard.Key):
        name = key.name
        if name:
            return f"<{name}>"
    return None


def _mouse_button_to_hotkey(button: mouse.Button) -> Optional[str]:
    mapping = {
        mouse.Button.left: "<mouse1>",
        mouse.Button.right: "<mouse2>",
        mouse.Button.middle: "<mouse3>",
        mouse.Button.x1: "<mouse4>",
        mouse.Button.x2: "<mouse5>",
    }
    return mapping.get(button)


def _build_tray_icon(size: int = 16) -> wx.Icon:
    bmp = wx.Bitmap(size, size)
    dc = wx.MemoryDC(bmp)
    dc.SetBackground(wx.Brush(wx.Colour("#111214")))
    dc.Clear()
    dc.SetBrush(wx.Brush(wx.Colour("#3c9cff")))
    dc.SetPen(wx.Pen(wx.Colour("#3c9cff")))
    dc.DrawRectangle(2, 4, 6, 8)
    dc.DrawPolygon([(8, 5), (14, 2), (14, 14), (8, 11)])
    dc.SelectObject(wx.NullBitmap)
    icon = wx.Icon()
    icon.CopyFromBitmap(bmp)
    return icon


def _ensure_sound_assets(volume: float) -> tuple[Path, Path]:
    ASSETS_DIR.mkdir(parents=True, exist_ok=True)
    mute_path = ASSETS_DIR / "mute.wav"
    unmute_path = ASSETS_DIR / "unmute.wav"
    _write_tone_wav(mute_path, [320, 260], duration_ms=180, volume=volume)
    _write_tone_wav(unmute_path, [440, 520], duration_ms=160, volume=volume)
    return mute_path, unmute_path


def _write_tone_wav(path: Path, freqs: list[int], duration_ms: int, volume: float) -> None:
    sample_rate = 44100
    samples = int(sample_rate * (duration_ms / 1000.0))
    max_amp = int(32767 * max(0.0, min(volume, 1.0)))
    with wave.open(str(path), "wb") as wav:
        wav.setnchannels(1)
        wav.setsampwidth(2)
        wav.setframerate(sample_rate)
        fade_len = int(samples * 0.12)
        for i in range(samples):
            t = i / sample_rate
            value = 0.0
            for freq in freqs:
                value += math.sin(2 * math.pi * freq * t)
            value /= max(1, len(freqs))
            if i < fade_len:
                env = i / fade_len
            elif i > samples - fade_len:
                env = (samples - i) / fade_len
            else:
                env = 1.0
            packed = struct.pack("<h", int(max_amp * value * env))
            wav.writeframesraw(packed)


class TrayIcon(wx.adv.TaskBarIcon):
    def __init__(self, frame: "MuteFrame") -> None:
        super().__init__()
        self.frame = frame
        self.SetIcon(_build_tray_icon(), APP_NAME)
        self.Bind(wx.adv.EVT_TASKBAR_LEFT_DCLICK, self.on_show)

    def CreatePopupMenu(self) -> wx.Menu:
        menu = wx.Menu()
        show_item = menu.Append(wx.ID_ANY, "Show")
        quit_item = menu.Append(wx.ID_EXIT, "Quit")
        self.Bind(wx.EVT_MENU, self.on_show, show_item)
        self.Bind(wx.EVT_MENU, self.on_quit, quit_item)
        return menu

    def on_show(self, _event: wx.Event) -> None:
        self.frame.show_from_tray()

    def on_quit(self, _event: wx.Event) -> None:
        self.frame.shutdown()


class MuteFrame(wx.Frame):
    def __init__(self) -> None:
        super().__init__(None, title=APP_NAME, size=(400, 290))
        self.SetMinSize((400, 290))

        self.hotkey_state = HotkeyState()
        self.autostart_enabled = _get_autostart_enabled()
        self.tray_on_close = True
        self._recording_hotkey = False
        self._record_stop: Optional[callable] = None
        self.tray_icon: Optional[TrayIcon] = None
        self._status_reset: Optional[wx.CallLater] = None
        self.sound_volume = 0.05
        self.mute_sound, self.unmute_sound = _ensure_sound_assets(self.sound_volume)

        self._build_ui()
        self._configure_hotkey()
        self.Bind(wx.EVT_CLOSE, self._on_close)

    def _build_ui(self) -> None:
        bg = wx.Colour("#111214")
        fg = wx.Colour("#e6e6e6")
        subtle = wx.Colour("#9aa0a6")
        entry_bg = wx.Colour("#1c1f24")
        accent = wx.Colour("#3c9cff")

        panel = wx.Panel(self)
        panel.SetBackgroundColour(bg)

        title = wx.StaticText(panel, label="Mute active app")
        title_font = title.GetFont()
        title_font.SetPointSize(13)
        title_font.SetWeight(wx.FONTWEIGHT_BOLD)
        title.SetFont(title_font)
        title.SetForegroundColour(fg)

        desc = wx.StaticText(panel, label="Toggle the active app's sound via a hotkey.")
        desc.SetForegroundColour(subtle)

        hotkey_label = wx.StaticText(panel, label="Hotkey")
        hotkey_label.SetForegroundColour(fg)
        self.hotkey_entry = wx.TextCtrl(panel, style=wx.TE_READONLY)
        self.hotkey_entry.SetValue("§")
        self.hotkey_entry.SetBackgroundColour(entry_bg)
        self.hotkey_entry.SetForegroundColour(fg)

        record_button = wx.Button(panel, label="Record")
        record_button.SetBackgroundColour(wx.Colour("#2a2f36"))
        record_button.SetForegroundColour(fg)
        record_button.Bind(wx.EVT_BUTTON, self._record_hotkey)

        self.autostart_checkbox = wx.CheckBox(panel, label="Run on Windows startup")
        self.autostart_checkbox.SetValue(self.autostart_enabled)
        self.autostart_checkbox.SetForegroundColour(fg)
        self.autostart_checkbox.Bind(wx.EVT_CHECKBOX, self._on_toggle_autostart)

        self.tray_checkbox = wx.CheckBox(panel, label="Minimize to tray on close")
        self.tray_checkbox.SetValue(True)
        self.tray_checkbox.SetForegroundColour(fg)
        self.tray_checkbox.Bind(wx.EVT_CHECKBOX, self._on_toggle_tray)

        volume_label = wx.StaticText(panel, label="Sound volume")
        volume_label.SetForegroundColour(fg)
        self.volume_value = wx.StaticText(panel, label="5%")
        self.volume_value.SetForegroundColour(subtle)
        self.volume_slider = wx.Slider(panel, minValue=0, maxValue=100, value=5)
        self.volume_slider.Bind(wx.EVT_SLIDER, self._on_volume_change)

        toggle_button = wx.Button(panel, label="Toggle active app now")
        toggle_button.SetBackgroundColour(accent)
        toggle_button.SetForegroundColour(wx.Colour("#ffffff"))
        toggle_button.Bind(wx.EVT_BUTTON, self._toggle_now)

        self.status_text = wx.StaticText(panel, label="Ready")
        self.status_text.SetForegroundColour(subtle)

        top_sizer = wx.BoxSizer(wx.VERTICAL)
        top_sizer.Add(title, 0, wx.BOTTOM, 4)
        top_sizer.Add(desc, 0, wx.BOTTOM, 12)

        hotkey_sizer = wx.BoxSizer(wx.HORIZONTAL)
        hotkey_sizer.Add(hotkey_label, 0, wx.ALIGN_CENTER_VERTICAL)
        hotkey_sizer.Add(self.hotkey_entry, 1, wx.LEFT | wx.RIGHT, 8)
        hotkey_sizer.Add(record_button, 0)

        top_sizer.Add(hotkey_sizer, 0, wx.EXPAND | wx.BOTTOM, 12)
        top_sizer.Add(self.autostart_checkbox, 0, wx.BOTTOM, 6)
        top_sizer.Add(self.tray_checkbox, 0, wx.BOTTOM, 14)
        volume_sizer = wx.BoxSizer(wx.HORIZONTAL)
        volume_sizer.Add(volume_label, 0, wx.ALIGN_CENTER_VERTICAL)
        volume_sizer.Add(self.volume_slider, 1, wx.LEFT | wx.RIGHT, 8)
        volume_sizer.Add(self.volume_value, 0, wx.ALIGN_CENTER_VERTICAL)
        top_sizer.Add(volume_sizer, 0, wx.EXPAND | wx.BOTTOM, 14)
        top_sizer.Add(toggle_button, 0, wx.EXPAND | wx.BOTTOM, 8)
        top_sizer.Add(self.status_text, 0)

        panel.SetSizer(top_sizer)
        panel.Layout()
        self.Centre()

    def _configure_hotkey(self) -> None:
        text = _normalize_hotkey(self.hotkey_entry.GetValue())
        self.hotkey_state.hotkey_text = text
        if self.hotkey_state.listener:
            self.hotkey_state.listener.stop()
            self.hotkey_state.listener = None
        if self.hotkey_state.mouse_listener:
            self.hotkey_state.mouse_listener.stop()
            self.hotkey_state.mouse_listener = None
        try:
            if text.startswith("<mouse") and text.endswith(">"):
                self._setup_mouse_hotkey(text)
            else:
                self.hotkey_state.listener = keyboard.GlobalHotKeys({text: self._toggle_now_from_hook})
                self.hotkey_state.listener.start()
            self._set_status(f"Hotkey set to {text}")
        except ValueError:
            self._set_status("Invalid hotkey format")

    def _setup_mouse_hotkey(self, text: str) -> None:
        mapping = {
            "<mouse1>": mouse.Button.left,
            "<mouse2>": mouse.Button.right,
            "<mouse3>": mouse.Button.middle,
            "<mouse4>": mouse.Button.x1,
            "<mouse5>": mouse.Button.x2,
        }
        button = mapping.get(text)
        if not button:
            raise ValueError("Unknown mouse button")

        def on_click(_x: int, _y: int, clicked: mouse.Button, pressed: bool) -> None:
            if pressed and clicked == button:
                self._toggle_now_from_hook()

        self.hotkey_state.mouse_listener = mouse.Listener(on_click=on_click)
        self.hotkey_state.mouse_listener.start()

    def _record_hotkey(self, _event: wx.CommandEvent) -> None:
        if self._recording_hotkey:
            return
        self._recording_hotkey = True
        self._set_status("Press a key or mouse button...")
        self.hotkey_entry.SetValue("press any button")
        self.hotkey_entry.Enable(False)

        def finish(text: str) -> None:
            wx.CallAfter(self._finish_record_hotkey, text)

        def on_key_press(key: keyboard.Key | keyboard.KeyCode) -> Optional[bool]:
            text = _key_to_hotkey(key)
            if text:
                finish(text)
                return False
            return None

        def on_click(_x: int, _y: int, button: mouse.Button, pressed: bool) -> Optional[bool]:
            if pressed:
                text = _mouse_button_to_hotkey(button)
                if text:
                    finish(text)
                    return False
            return None

        keyboard_listener = keyboard.Listener(on_press=on_key_press)
        mouse_listener = mouse.Listener(on_click=on_click)
        keyboard_listener.start()
        mouse_listener.start()

        def stop_listeners() -> None:
            keyboard_listener.stop()
            mouse_listener.stop()

        self._record_stop = stop_listeners

    def _finish_record_hotkey(self, text: str) -> None:
        if not self._recording_hotkey:
            return
        self._recording_hotkey = False
        if self._record_stop:
            self._record_stop()
            self._record_stop = None
        self.hotkey_entry.Enable(True)
        self.hotkey_entry.SetValue(text)
        self._configure_hotkey()

    def _toggle_now_from_hook(self) -> None:
        wx.CallAfter(self._toggle_now, None)

    def _toggle_now(self, _event: Optional[wx.CommandEvent]) -> None:
        result = toggle_active_app_mute()
        if result is None:
            self._set_status("No active app audio session found", reset_ms=5000)
            return
        if result:
            winsound.PlaySound(str(self.mute_sound), winsound.SND_FILENAME | winsound.SND_ASYNC)
            self._set_status("Muted", reset_ms=2000)
        else:
            winsound.PlaySound(str(self.unmute_sound), winsound.SND_FILENAME | winsound.SND_ASYNC)
            self._set_status("Unmuted", reset_ms=2000)

    def _set_status(self, text: str, reset_ms: Optional[int] = None) -> None:
        self.status_text.SetLabel(text)
        if self._status_reset:
            self._status_reset.Stop()
            self._status_reset = None
        if reset_ms:
            self._status_reset = wx.CallLater(reset_ms, self._clear_status)

    def _clear_status(self) -> None:
        self.status_text.SetLabel("Ready")

    def _on_toggle_autostart(self, _event: wx.CommandEvent) -> None:
        self.autostart_enabled = self.autostart_checkbox.GetValue()
        _set_autostart(self.autostart_enabled)

    def _on_toggle_tray(self, _event: wx.CommandEvent) -> None:
        self.tray_on_close = self.tray_checkbox.GetValue()

    def _on_volume_change(self, _event: wx.CommandEvent) -> None:
        value = self.volume_slider.GetValue()
        self.sound_volume = max(0.0, min(1.0, value / 100.0))
        self.volume_value.SetLabel(f"{value}%")
        self.mute_sound, self.unmute_sound = _ensure_sound_assets(self.sound_volume)

    def _on_close(self, event: wx.CloseEvent) -> None:
        if self.tray_on_close:
            self.hide_to_tray()
            event.Veto()
        else:
            self.shutdown()

    def hide_to_tray(self) -> None:
        if not self.tray_icon:
            self.tray_icon = TrayIcon(self)
        self.Hide()

    def show_from_tray(self) -> None:
        self.Show()
        self.Raise()

    def shutdown(self) -> None:
        if self.hotkey_state.listener:
            self.hotkey_state.listener.stop()
        if self.hotkey_state.mouse_listener:
            self.hotkey_state.mouse_listener.stop()
        if self.tray_icon:
            self.tray_icon.RemoveIcon()
            self.tray_icon.Destroy()
            self.tray_icon = None
        self.Destroy()


class MuteApp(wx.App):
    def OnInit(self) -> bool:
        frame = MuteFrame()
        frame.Show()
        self.SetTopWindow(frame)
        return True


def main() -> None:
    _enable_dpi_awareness()
    app = MuteApp(False)
    app.MainLoop()


if __name__ == "__main__":
    main()
