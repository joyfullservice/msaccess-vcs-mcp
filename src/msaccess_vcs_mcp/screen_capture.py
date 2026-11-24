"""Win32 window capture for Access screenshots, without touching COM.

Everything here talks to Windows, never to Access. That is what lets a
screenshot work while Access sits on a MsgBox, a compile error, break mode,
or a call that never returns: none of these stop the window manager, DWM,
or the Restart Manager from answering.

The pure helpers (``plan_canvas``, ``scale_for``, ``blit``, ``crop_buffer``,
``bgra_to_rgb``, ``encode_png``, and the window classifiers) take plain data
so they can be tested without a desktop.
"""

from __future__ import annotations

import ctypes
import hashlib
import struct
import sys
import threading
import zlib
from contextlib import contextmanager
from ctypes import wintypes
from dataclasses import dataclass, field
from typing import Callable, Iterable, Iterator

MAIN_WINDOW_CLASS = "OMain"
VBE_WINDOW_CLASS = "wndclass_desked_gsk"
MDI_CLIENT_CLASS = "MDIClient"
DIALOG_CLASSES = frozenset({"#32770", "NUIDialog"})
DIALOG_CLASS_PREFIXES = ("bosa_sdm_",)
# Office paints its window glow as separate layered top-level windows, and the
# shell adds shadows and tooltips; composited over a capture they paint black.
CHROME_CLASSES = frozenset({
    "MSO_BORDEREFFECT_WINDOW_CLASS",
    "SysShadow",
    "tooltips_class32",
    "OfficeTooltip",
    "IME",
    "MSCTFIME UI",
    "Default IME",
})
TEXT_CONTROL_CLASSES = ("Static", "Button", "Edit", "RichEdit")

NEUTRAL_GRAY = 0x80
MAX_DIALOG_TEXT = 4000
MAX_DIALOG_CHILDREN = 200
PRINT_WINDOW_TIMEOUT_SEC = 2.0
RESPONSIVE_PROBE_MS = 1000
CHILD_TEXT_TIMEOUT_MS = 250

GWL_STYLE = -16
GWL_EXSTYLE = -20
GW_OWNER = 4
GA_ROOT = 2
WS_CHILD = 0x40000000
WS_EX_LAYERED = 0x00080000
WS_EX_TRANSPARENT = 0x00000020
ES_PASSWORD = 0x0020
WM_NULL = 0x0000
WM_GETTEXT = 0x000D
WM_MDIGETACTIVE = 0x0229
BM_CLICK = 0x00F5
SMTO_NORMAL = 0x0000
SMTO_ABORTIFHUNG = 0x0002
PW_RENDERFULLCONTENT = 0x00000002
SRCCOPY = 0x00CC0020
CAPTUREBLT = 0x40000000
HALFTONE = 4
DIB_RGB_COLORS = 0
BI_RGB = 0
DWMWA_EXTENDED_FRAME_BOUNDS = 9
DWMWA_CLOAKED = 14
SM_XVIRTUALSCREEN = 76
SM_YVIRTUALSCREEN = 77
SM_CXVIRTUALSCREEN = 78
SM_CYVIRTUALSCREEN = 79
DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = -4
ERROR_MORE_DATA = 234
CCH_RM_SESSION_KEY = 32
CCH_RM_MAX_APP_NAME = 255
CCH_RM_MAX_SVC_NAME = 63


class CaptureError(RuntimeError):
    """A capture that could not produce pixels, with an MCP error pattern."""

    def __init__(self, message: str, error_pattern: str = "capture_failed") -> None:
        super().__init__(message)
        self.error_pattern = error_pattern


# ---------------------------------------------------------------------------
# Plain data
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Rect:
    left: int
    top: int
    right: int
    bottom: int

    @property
    def width(self) -> int:
        return self.right - self.left

    @property
    def height(self) -> int:
        return self.bottom - self.top

    @property
    def empty(self) -> bool:
        return self.width <= 0 or self.height <= 0

    def intersect(self, other: "Rect") -> "Rect | None":
        result = Rect(
            max(self.left, other.left),
            max(self.top, other.top),
            min(self.right, other.right),
            min(self.bottom, other.bottom),
        )
        return None if result.empty else result

    def union(self, other: "Rect") -> "Rect":
        return Rect(
            min(self.left, other.left),
            min(self.top, other.top),
            max(self.right, other.right),
            max(self.bottom, other.bottom),
        )

    def offset(self, dx: int, dy: int) -> "Rect":
        return Rect(self.left + dx, self.top + dy, self.right + dx, self.bottom + dy)

    def as_dict(self) -> dict[str, int]:
        return {"left": self.left, "top": self.top, "width": self.width, "height": self.height}


@dataclass(frozen=True)
class WindowInfo:
    """One top-level window. ``z`` is its EnumWindows index; 0 is topmost."""

    hwnd: int
    pid: int
    cls: str
    title: str
    rect: Rect
    frame: Rect
    owner: int = 0
    style: int = 0
    ex_style: int = 0
    visible: bool = True
    enabled: bool = True
    minimized: bool = False
    cloaked: bool = False
    hung: bool = False
    z: int = 0


@dataclass
class ChildControl:
    cls: str
    text: str
    style: int = 0


@dataclass
class Capture:
    png: bytes
    width: int
    height: int
    original_width: int
    original_height: int
    scale: float
    method: str
    responsive: bool = True
    occluded: bool = False
    warnings: list[str] = field(default_factory=list)

    def image_meta(self) -> dict[str, object]:
        return {
            "width": self.width,
            "height": self.height,
            "original_width": self.original_width,
            "original_height": self.original_height,
            "scale": round(self.scale, 4),
            "capture_method": self.method,
            "occluded": self.occluded,
        }


# ---------------------------------------------------------------------------
# Pure helpers
# ---------------------------------------------------------------------------


def is_dialog_class(cls: str) -> bool:
    return cls in DIALOG_CLASSES or cls.startswith(DIALOG_CLASS_PREFIXES)


def is_chrome(info: WindowInfo) -> bool:
    if info.cls in CHROME_CLASSES:
        return True
    return bool(info.ex_style & WS_EX_LAYERED and info.ex_style & WS_EX_TRANSPARENT)


def owner_chain_reaches(info: WindowInfo, target: int, by_hwnd: dict[int, WindowInfo]) -> bool:
    seen: set[int] = set()
    owner = info.owner
    while owner and owner not in seen:
        if owner == target:
            return True
        seen.add(owner)
        parent = by_hwnd.get(owner)
        owner = parent.owner if parent else 0
    return False


def owner_label(info: WindowInfo, main: int, vbe: int, by_hwnd: dict[int, WindowInfo]) -> str:
    if vbe and owner_chain_reaches(info, vbe, by_hwnd):
        return "vbe"
    if main and owner_chain_reaches(info, main, by_hwnd):
        return "application"
    return "other"


def find_window(windows: Iterable[WindowInfo], cls: str) -> WindowInfo | None:
    """Topmost window of ``cls``, preferring a visible one."""
    matches = sorted((w for w in windows if w.cls == cls), key=lambda w: (not w.visible, w.z))
    return matches[0] if matches else None


def vbe_state(windows: Iterable[WindowInfo]) -> str | None:
    """``break``, ``running``, or ``design`` from the VBE title; None with no VBE.

    The VBE keeps its title current while hidden, so a hidden editor still
    reports break mode.
    """
    vbe = find_window(windows, VBE_WINDOW_CLASS)
    if vbe is None:
        return None
    title = vbe.title.lower()
    if "[break]" in title:
        return "break"
    if "[running]" in title:
        return "running"
    return "design"


def select_overlays(windows: list[WindowInfo], base: WindowInfo) -> list[WindowInfo]:
    """Same-process windows to paint over ``base``, bottom-to-top.

    An overlay is owned (directly or through its owner chain) by the base.
    Dialogs above the base whose owner chain is broken still count, unless
    they belong to the other primary window: a VBE compile error is not
    painted over the Access window, only listed in ``dialogs``. A window
    that does not touch the base (a dialog on another monitor) is not
    painted either; the canvas would stretch across the gap between them.
    """
    by_hwnd = {w.hwnd: w for w in windows}
    sibling_cls = VBE_WINDOW_CLASS if base.cls == MAIN_WINDOW_CLASS else MAIN_WINDOW_CLASS
    siblings = {w.hwnd for w in windows if w.cls == sibling_cls}
    picked: list[WindowInfo] = []
    for w in windows:
        if w.hwnd == base.hwnd or w.hwnd in siblings:
            continue
        if not w.visible or w.minimized or w.cloaked or w.frame.empty or is_chrome(w):
            continue
        if w.frame.intersect(base.frame) is None:
            continue
        if owner_chain_reaches(w, base.hwnd, by_hwnd):
            picked.append(w)
        elif is_dialog_class(w.cls) and w.z < base.z and not any(
            owner_chain_reaches(w, s, by_hwnd) for s in siblings
        ):
            picked.append(w)
    return sorted(picked, key=lambda w: w.z, reverse=True)


def is_occluded(base_frame: Rect, base_z: int, pid: int, windows: Iterable[WindowInfo]) -> bool:
    """True when another process's visible window sits above and overlaps."""
    for w in windows:
        if w.z >= base_z or w.pid == pid:
            continue
        if not w.visible or w.minimized or w.cloaked or is_chrome(w):
            continue
        if w.frame.intersect(base_frame) is not None:
            return True
    return False


def strip_accelerators(text: str) -> str:
    return text.replace("&&", "\0").replace("&", "").replace("\0", "&")


def dialog_id(hwnd: int, cls: str, title: str, text: str, buttons: list[str]) -> str:
    """Short fingerprint of what a dialog shows, so a click lands on the dialog
    the agent read and not on whatever replaced it."""
    raw = "\x1f".join([str(hwnd), cls, title, text, "\x1e".join(buttons)])
    return hashlib.sha1(raw.encode("utf-8")).hexdigest()[:12]


def _clip_text(lines: list[str]) -> str:
    body = "\n".join(lines)
    return body[:MAX_DIALOG_TEXT] + "..." if len(body) > MAX_DIALOG_TEXT else body


def dialog_entry(
    info: WindowInfo,
    children: list[ChildControl],
    owner: str,
    overlaps_target: bool,
) -> dict[str, object]:
    lines: list[str] = []
    buttons: list[str] = []
    for child in children:
        text = child.text.strip()
        if not text:
            continue
        if child.cls == "Button":
            buttons.append(strip_accelerators(text))
        elif child.cls == "Static" or child.cls.startswith(("Edit", "RichEdit")):
            if text not in lines:
                lines.append(text)
    body = _clip_text(lines)
    return {
        "dialog_id": dialog_id(info.hwnd, info.cls, info.title, body, buttons),
        "title": info.title,
        "class": info.cls,
        "owner": owner,
        "text": body,
        "buttons": buttons,
        "text_source": "window",
        "overlaps_target": overlaps_target,
    }


def dialog_windows(windows: Iterable[WindowInfo]) -> list[WindowInfo]:
    """Visible dialogs, topmost first: the order ``read_dialogs`` reports them in."""
    return [
        w for w in sorted(windows, key=lambda w: w.z)
        if w.visible and not w.cloaked and is_dialog_class(w.cls)
    ]


def read_accessible_dialog(hwnd: int, title: str) -> tuple[list[str], list[str]] | None:
    """MSAA text and buttons, only for a window that answers messages."""
    if not is_responsive(hwnd):
        return None
    from . import accessibility

    return accessibility.read_dialog(hwnd, title)


def read_dialogs(
    windows: list[WindowInfo],
    main: int,
    vbe: int,
    target_frame: Rect | None = None,
    read_children: Callable[[int], list[ChildControl]] | None = None,
    read_accessible: Callable[[int, str], tuple[list[str], list[str]] | None] | None = None,
) -> tuple[list[dict[str, object]], list[str]]:
    """Every visible dialog in the process, whatever was captured.

    Window calls read standard dialogs. A dialog that shows nothing to them
    (Office's ``NUIDialog``) is read through its accessibility tree instead.
    """
    reader = read_children or read_child_controls
    accessible = read_accessible or read_accessible_dialog
    by_hwnd = {w.hwnd: w for w in windows}
    dialogs: list[dict[str, object]] = []
    warnings: list[str] = []
    for w in dialog_windows(windows):
        overlaps = bool(target_frame and w.frame.intersect(target_frame))
        entry = dialog_entry(w, reader(w.hwnd), owner_label(w, main, vbe, by_hwnd), overlaps)
        if not entry["text"] and not entry["buttons"]:
            content = accessible(w.hwnd, w.title)
            if content and (content[0] or content[1]):
                body = _clip_text(content[0])
                entry.update(
                    dialog_id=dialog_id(w.hwnd, w.cls, w.title, body, content[1]),
                    text=body,
                    buttons=content[1],
                    text_source="accessibility",
                )
        if not entry["text"] and not entry["buttons"]:
            warnings.append(
                f"Dialog '{w.title or w.cls}' exposes no readable text; read it from the image."
            )
        dialogs.append(entry)
    return dialogs, warnings


def plan_canvas(
    base_frame: Rect,
    overlay_frames: list[Rect],
    crop: Rect | None = None,
    virtual_screen: Rect | None = None,
) -> tuple[Rect, list[tuple[int, int]]]:
    """Canvas rectangle and each window's offset in it (base first).

    The canvas is the union of every frame, so a dialog hanging past the
    Access window is not cut off, clipped to the virtual screen and then to
    ``crop``.
    """
    canvas = base_frame
    for frame in overlay_frames:
        canvas = canvas.union(frame)
    if virtual_screen is not None:
        clipped = canvas.intersect(virtual_screen)
        if clipped is None:
            raise CaptureError("The window is entirely off screen.", "capture_failed")
        canvas = clipped
    if crop is not None:
        clipped = canvas.intersect(crop)
        if clipped is None:
            raise CaptureError("The object is outside the captured window.", "capture_failed")
        canvas = clipped
    placements = [(f.left - canvas.left, f.top - canvas.top) for f in [base_frame, *overlay_frames]]
    return canvas, placements


def scale_for(width: int, height: int, max_dimension: int) -> tuple[int, int, float]:
    longest = max(width, height)
    if longest <= max_dimension:
        return width, height, 1.0
    scale = max_dimension / longest
    return max(1, round(width * scale)), max(1, round(height * scale)), scale


def new_canvas(width: int, height: int) -> bytearray:
    return bytearray([NEUTRAL_GRAY, NEUTRAL_GRAY, NEUTRAL_GRAY, 0xFF]) * (width * height)


def blit(
    dst: bytearray, dst_w: int, dst_h: int, src: bytes, src_w: int, src_h: int, x: int, y: int
) -> None:
    """Copy a BGRA buffer into another at (x, y), clipped to the destination."""
    sx = max(0, -x)
    sy = max(0, -y)
    dx = max(0, x)
    dy = max(0, y)
    w = min(src_w - sx, dst_w - dx)
    h = min(src_h - sy, dst_h - dy)
    if w <= 0 or h <= 0:
        return
    span = w * 4
    for row in range(h):
        s = ((sy + row) * src_w + sx) * 4
        d = ((dy + row) * dst_w + dx) * 4
        dst[d:d + span] = src[s:s + span]


def fill_gray(dst: bytearray, dst_w: int, dst_h: int, region: Rect) -> None:
    region = region.intersect(Rect(0, 0, dst_w, dst_h))
    if region is None:
        return
    row = bytes([NEUTRAL_GRAY, NEUTRAL_GRAY, NEUTRAL_GRAY, 0xFF]) * region.width
    for y in range(region.top, region.bottom):
        start = (y * dst_w + region.left) * 4
        dst[start:start + len(row)] = row


def mask_foreign(
    pixels: bytes, canvas: Rect, base_z: int, pid: int, windows: Iterable[WindowInfo]
) -> tuple[bytearray, int]:
    """Gray out other processes' windows above Access in a screen copy.

    Windows are applied bottom to top: a foreign window is grayed, and an
    Access window above it gets its screen pixels back, since on screen it
    is the one showing. Returns the buffer and how many foreign windows
    were masked.
    """
    original = bytes(pixels)
    result = bytearray(original)
    masked = 0
    above = sorted((w for w in windows if w.z < base_z), key=lambda w: w.z, reverse=True)
    for w in above:
        if not w.visible or w.minimized or w.cloaked or is_chrome(w):
            continue
        overlap = w.frame.intersect(canvas)
        if overlap is None:
            continue
        local = overlap.offset(-canvas.left, -canvas.top)
        if w.pid == pid:
            piece, pw, ph = crop_buffer(original, canvas.width, canvas.height, local)
            blit(result, canvas.width, canvas.height, piece, pw, ph, local.left, local.top)
        else:
            fill_gray(result, canvas.width, canvas.height, local)
            masked += 1
    return result, masked


def crop_buffer(src: bytes, src_w: int, src_h: int, region: Rect) -> tuple[bytes, int, int]:
    """Cut ``region`` (in buffer coordinates) out of a BGRA buffer."""
    region = region.intersect(Rect(0, 0, src_w, src_h))
    if region is None:
        raise CaptureError("Crop region is outside the captured bitmap.")
    span = region.width * 4
    rows = [
        src[(y * src_w + region.left) * 4:(y * src_w + region.left) * 4 + span]
        for y in range(region.top, region.bottom)
    ]
    return b"".join(rows), region.width, region.height


def bgra_to_rgb(bgra: bytes) -> bytearray:
    pixels = len(bgra) // 4
    rgb = bytearray(pixels * 3)
    rgb[0::3] = bgra[2::4]
    rgb[1::3] = bgra[1::4]
    rgb[2::3] = bgra[0::4]
    return rgb


def is_blank(bgra: bytes) -> bool:
    return not bgra_to_rgb(bgra).strip(b"\x00")


def _png_chunk(kind: bytes, data: bytes) -> bytes:
    return (
        struct.pack(">I", len(data))
        + kind
        + data
        + struct.pack(">I", zlib.crc32(kind + data) & 0xFFFFFFFF)
    )


def encode_png(rgb: bytes, width: int, height: int) -> bytes:
    """24-bit RGB PNG with filter 0 on every row."""
    stride = width * 3
    if len(rgb) != stride * height:
        raise ValueError(f"RGB buffer is {len(rgb)} bytes; expected {stride * height}.")
    raw = b"".join(b"\x00" + bytes(rgb[y * stride:(y + 1) * stride]) for y in range(height))
    return (
        b"\x89PNG\r\n\x1a\n"
        + _png_chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + _png_chunk(b"IDAT", zlib.compress(raw, 6))
        + _png_chunk(b"IEND", b"")
    )


# ---------------------------------------------------------------------------
# Win32 bindings
# ---------------------------------------------------------------------------


class _RECT(ctypes.Structure):
    _fields_ = [("left", wintypes.LONG), ("top", wintypes.LONG),
                ("right", wintypes.LONG), ("bottom", wintypes.LONG)]

    def to_rect(self) -> Rect:
        return Rect(self.left, self.top, self.right, self.bottom)


class _BITMAPINFOHEADER(ctypes.Structure):
    _fields_ = [
        ("biSize", wintypes.DWORD),
        ("biWidth", wintypes.LONG),
        ("biHeight", wintypes.LONG),
        ("biPlanes", wintypes.WORD),
        ("biBitCount", wintypes.WORD),
        ("biCompression", wintypes.DWORD),
        ("biSizeImage", wintypes.DWORD),
        ("biXPelsPerMeter", wintypes.LONG),
        ("biYPelsPerMeter", wintypes.LONG),
        ("biClrUsed", wintypes.DWORD),
        ("biClrImportant", wintypes.DWORD),
    ]


class _BITMAPINFO(ctypes.Structure):
    _fields_ = [("bmiHeader", _BITMAPINFOHEADER), ("bmiColors", wintypes.DWORD * 3)]


class _RM_UNIQUE_PROCESS(ctypes.Structure):
    _fields_ = [("dwProcessId", wintypes.DWORD), ("ProcessStartTime", wintypes.FILETIME)]


class _RM_PROCESS_INFO(ctypes.Structure):
    _fields_ = [
        ("Process", _RM_UNIQUE_PROCESS),
        ("strAppName", wintypes.WCHAR * (CCH_RM_MAX_APP_NAME + 1)),
        ("strServiceShortName", wintypes.WCHAR * (CCH_RM_MAX_SVC_NAME + 1)),
        ("ApplicationType", ctypes.c_int),
        ("AppStatus", wintypes.ULONG),
        ("TSSessionId", wintypes.DWORD),
        ("bRestartable", wintypes.BOOL),
    ]


class _Api:
    """Lazily bound Win32 functions with explicit signatures.

    Explicit ``argtypes``/``restype`` keep 64-bit handles from being
    truncated to ``int``.
    """

    def __init__(self) -> None:
        if sys.platform != "win32":
            raise CaptureError("Screenshots need Windows.", "capture_failed")
        u = ctypes.WinDLL("user32", use_last_error=True)
        g = ctypes.WinDLL("gdi32", use_last_error=True)
        d = ctypes.WinDLL("dwmapi")
        rm = ctypes.WinDLL("rstrtmgr")
        H, HDC, HBMP, HGDI = wintypes.HWND, wintypes.HDC, wintypes.HBITMAP, wintypes.HGDIOBJ
        self.WNDENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, H, wintypes.LPARAM)

        def bind(dll, name, restype, *argtypes):
            fn = getattr(dll, name)
            fn.restype = restype
            fn.argtypes = argtypes
            setattr(self, name, fn)

        bind(u, "EnumWindows", wintypes.BOOL, self.WNDENUMPROC, wintypes.LPARAM)
        bind(u, "EnumChildWindows", wintypes.BOOL, H, self.WNDENUMPROC, wintypes.LPARAM)
        bind(u, "GetWindowThreadProcessId", wintypes.DWORD, H, ctypes.POINTER(wintypes.DWORD))
        bind(u, "GetClassNameW", ctypes.c_int, H, wintypes.LPWSTR, ctypes.c_int)
        bind(u, "InternalGetWindowText", ctypes.c_int, H, wintypes.LPWSTR, ctypes.c_int)
        bind(u, "GetWindowRect", wintypes.BOOL, H, ctypes.POINTER(_RECT))
        bind(u, "GetWindow", H, H, wintypes.UINT)
        bind(u, "GetAncestor", H, H, wintypes.UINT)
        bind(u, "GetWindowLongW", wintypes.LONG, H, ctypes.c_int)
        bind(u, "IsWindowVisible", wintypes.BOOL, H)
        bind(u, "IsWindowEnabled", wintypes.BOOL, H)
        bind(u, "IsIconic", wintypes.BOOL, H)
        bind(u, "IsWindow", wintypes.BOOL, H)
        bind(u, "IsHungAppWindow", wintypes.BOOL, H)
        bind(u, "SendMessageTimeoutW", wintypes.LPARAM, H, wintypes.UINT, wintypes.WPARAM,
             wintypes.LPARAM, wintypes.UINT, wintypes.UINT, ctypes.POINTER(ctypes.c_size_t))
        bind(u, "PostMessageW", wintypes.BOOL, H, wintypes.UINT, wintypes.WPARAM,
             wintypes.LPARAM)
        bind(u, "PrintWindow", wintypes.BOOL, H, HDC, wintypes.UINT)
        bind(u, "GetDC", HDC, H)
        bind(u, "ReleaseDC", ctypes.c_int, H, HDC)
        bind(u, "GetSystemMetrics", ctypes.c_int, ctypes.c_int)
        self.SetThreadDpiAwarenessContext = getattr(u, "SetThreadDpiAwarenessContext", None)
        if self.SetThreadDpiAwarenessContext is not None:
            self.SetThreadDpiAwarenessContext.restype = ctypes.c_void_p
            self.SetThreadDpiAwarenessContext.argtypes = (ctypes.c_void_p,)

        bind(g, "CreateCompatibleDC", HDC, HDC)
        bind(g, "DeleteDC", wintypes.BOOL, HDC)
        bind(g, "CreateDIBSection", HBMP, HDC, ctypes.POINTER(_BITMAPINFO), wintypes.UINT,
             ctypes.POINTER(ctypes.c_void_p), wintypes.HANDLE, wintypes.DWORD)
        bind(g, "SelectObject", HGDI, HDC, HGDI)
        bind(g, "DeleteObject", wintypes.BOOL, HGDI)
        bind(g, "BitBlt", wintypes.BOOL, HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int,
             ctypes.c_int, HDC, ctypes.c_int, ctypes.c_int, wintypes.DWORD)
        bind(g, "StretchDIBits", ctypes.c_int, HDC, ctypes.c_int, ctypes.c_int, ctypes.c_int,
             ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
             ctypes.c_void_p, ctypes.POINTER(_BITMAPINFO), wintypes.UINT, wintypes.DWORD)
        bind(g, "SetStretchBltMode", ctypes.c_int, HDC, ctypes.c_int)
        bind(g, "SetBrushOrgEx", wintypes.BOOL, HDC, ctypes.c_int, ctypes.c_int,
             ctypes.c_void_p)
        bind(g, "GdiFlush", wintypes.BOOL)

        bind(d, "DwmGetWindowAttribute", ctypes.c_long, H, wintypes.DWORD, ctypes.c_void_p,
             wintypes.DWORD)

        bind(rm, "RmStartSession", wintypes.DWORD, ctypes.POINTER(wintypes.DWORD),
             wintypes.DWORD, wintypes.LPWSTR)
        bind(rm, "RmRegisterResources", wintypes.DWORD, wintypes.DWORD, wintypes.UINT,
             ctypes.POINTER(wintypes.LPCWSTR), wintypes.UINT, ctypes.c_void_p,
             wintypes.UINT, ctypes.c_void_p)
        bind(rm, "RmGetList", wintypes.DWORD, wintypes.DWORD, ctypes.POINTER(wintypes.UINT),
             ctypes.POINTER(wintypes.UINT), ctypes.POINTER(_RM_PROCESS_INFO),
             ctypes.POINTER(wintypes.DWORD))
        bind(rm, "RmEndSession", wintypes.DWORD, wintypes.DWORD)


_api_instance: _Api | None = None
_api_lock = threading.Lock()


def api() -> _Api:
    global _api_instance
    with _api_lock:
        if _api_instance is None:
            _api_instance = _Api()
        return _api_instance


@contextmanager
def dpi_aware() -> Iterator[None]:
    """Per-monitor-v2 DPI awareness for this thread, restored on exit.

    Worker and gate threads are reused, so leaving the context changed would
    leak into unrelated work. Inside, every rectangle is in physical pixels.
    """
    setter = api().SetThreadDpiAwarenessContext
    previous = None
    if setter is not None:
        previous = setter(ctypes.c_void_p(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2))
    try:
        yield
    finally:
        if setter is not None and previous:
            setter(ctypes.c_void_p(previous))


def _class_name(hwnd: int) -> str:
    buf = ctypes.create_unicode_buffer(256)
    api().GetClassNameW(hwnd, buf, 256)
    return buf.value


def _stored_text(hwnd: int, limit: int = 512) -> str:
    """Text Windows stores for the window; never sends a message, never blocks."""
    buf = ctypes.create_unicode_buffer(limit)
    api().InternalGetWindowText(hwnd, buf, limit)
    return buf.value


def _message_text(hwnd: int, limit: int = 2048) -> str:
    """WM_GETTEXT with a timeout, for controls that keep their own text."""
    buf = ctypes.create_unicode_buffer(limit)
    result = ctypes.c_size_t(0)
    ok = api().SendMessageTimeoutW(
        hwnd, WM_GETTEXT, limit, ctypes.addressof(buf),
        SMTO_NORMAL | SMTO_ABORTIFHUNG, CHILD_TEXT_TIMEOUT_MS, ctypes.byref(result),
    )
    return buf.value if ok else ""


def window_rect(hwnd: int) -> Rect:
    r = _RECT()
    if not api().GetWindowRect(hwnd, ctypes.byref(r)):
        raise CaptureError(f"Window {hwnd:#x} no longer exists.", "capture_failed")
    return r.to_rect()


def _frame_rect(hwnd: int, fallback: Rect) -> Rect:
    r = _RECT()
    hr = api().DwmGetWindowAttribute(
        hwnd, DWMWA_EXTENDED_FRAME_BOUNDS, ctypes.byref(r), ctypes.sizeof(r)
    )
    if hr != 0:
        return fallback
    frame = r.to_rect()
    return fallback if frame.empty else frame


def _cloaked(hwnd: int) -> bool:
    value = wintypes.DWORD(0)
    hr = api().DwmGetWindowAttribute(hwnd, DWMWA_CLOAKED, ctypes.byref(value), ctypes.sizeof(value))
    return hr == 0 and value.value != 0


def window_pid(hwnd: int) -> int:
    pid = wintypes.DWORD(0)
    api().GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
    return int(pid.value)


def describe_window(hwnd: int, pid: int, z: int) -> WindowInfo:
    a = api()
    rect = window_rect(hwnd)
    return WindowInfo(
        hwnd=hwnd,
        pid=pid,
        cls=_class_name(hwnd),
        title=_stored_text(hwnd),
        rect=rect,
        frame=_frame_rect(hwnd, rect),
        owner=int(a.GetWindow(hwnd, GW_OWNER) or 0),
        style=int(a.GetWindowLongW(hwnd, GWL_STYLE)) & 0xFFFFFFFF,
        ex_style=int(a.GetWindowLongW(hwnd, GWL_EXSTYLE)) & 0xFFFFFFFF,
        visible=bool(a.IsWindowVisible(hwnd)),
        enabled=bool(a.IsWindowEnabled(hwnd)),
        minimized=bool(a.IsIconic(hwnd)),
        cloaked=_cloaked(hwnd),
        hung=bool(a.IsHungAppWindow(hwnd)),
        z=z,
    )


def enum_windows(pids: set[int] | None = None, visible_only: bool = False) -> list[WindowInfo]:
    """Top-level windows in z-order (topmost first).

    ``pids`` limits the full description to those processes; windows of
    other processes are described only when ``pids`` is None.
    """
    a = api()
    found: list[tuple[int, int, int]] = []

    def callback(hwnd, _lparam):
        handle = int(hwnd or 0)
        if handle:
            pid = window_pid(handle)
            if pids is None or pid in pids:
                found.append((handle, pid, len(found)))
            else:
                found.append((0, pid, len(found)))
        return True

    a.EnumWindows(a.WNDENUMPROC(callback), 0)
    windows: list[WindowInfo] = []
    for hwnd, pid, z in found:
        if not hwnd:
            continue
        if visible_only and not a.IsWindowVisible(hwnd):
            continue
        try:
            windows.append(describe_window(hwnd, pid, z))
        except CaptureError:
            continue
    return windows


def access_window_pids() -> dict[int, str]:
    """PID to main-window title for every process owning an Access window."""
    result: dict[int, str] = {}
    for w in enum_windows():
        if w.cls == MAIN_WINDOW_CLASS:
            result[w.pid] = w.title
        elif w.cls == VBE_WINDOW_CLASS:
            result.setdefault(w.pid, "")
    return result


def virtual_screen() -> Rect:
    a = api()
    left = a.GetSystemMetrics(SM_XVIRTUALSCREEN)
    top = a.GetSystemMetrics(SM_YVIRTUALSCREEN)
    return Rect(
        left,
        top,
        left + a.GetSystemMetrics(SM_CXVIRTUALSCREEN),
        top + a.GetSystemMetrics(SM_CYVIRTUALSCREEN),
    )


def read_child_controls(hwnd: int) -> list[ChildControl]:
    """Text of a dialog's child controls, read without risking a hang.

    Stored text first: it covers ``Static`` and ``Button`` and cannot block.
    Controls that keep their own text (``Edit``) need ``WM_GETTEXT``, sent
    with a timeout. Password fields are never read.
    """
    a = api()
    handles: list[int] = []

    def callback(child, _lparam):
        if len(handles) >= MAX_DIALOG_CHILDREN:
            return False
        handles.append(int(child or 0))
        return True

    a.EnumChildWindows(hwnd, a.WNDENUMPROC(callback), 0)
    controls: list[ChildControl] = []
    for child in handles:
        if not child or not a.IsWindowVisible(child):
            continue
        cls = _class_name(child)
        if not cls.startswith(TEXT_CONTROL_CLASSES):
            continue
        style = int(a.GetWindowLongW(child, GWL_STYLE)) & 0xFFFFFFFF
        if cls.startswith(("Edit", "RichEdit")) and style & ES_PASSWORD:
            continue
        text = _stored_text(child, MAX_DIALOG_TEXT)
        if not text:
            text = _message_text(child)
        controls.append(ChildControl(cls=cls, text=text, style=style))
    return controls


def click_button(dialog_hwnd: int, label: str) -> bool:
    """Post BM_CLICK to the dialog's button labelled ``label``.

    Posted, not sent: the click runs when the dialog's thread next pumps
    messages, so a busy Access cannot block the caller.
    """
    a = api()
    wanted = strip_accelerators(label).strip().casefold()
    hits: list[int] = []

    def callback(child, _lparam):
        handle = int(child or 0)
        if (
            handle
            and _class_name(handle) == "Button"
            and a.IsWindowVisible(handle)
            and strip_accelerators(_stored_text(handle)).strip().casefold() == wanted
        ):
            hits.append(handle)
            return False
        return True

    a.EnumChildWindows(dialog_hwnd, a.WNDENUMPROC(callback), 0)
    return bool(hits) and bool(a.PostMessageW(hits[0], BM_CLICK, 0, 0))


def window_gone(hwnd: int) -> bool:
    a = api()
    return not a.IsWindow(hwnd) or not a.IsWindowVisible(hwnd)


def is_responsive(hwnd: int) -> bool:
    """A MsgBox, compile error, or break mode still pumps messages; a busy loop does not."""
    a = api()
    if a.IsHungAppWindow(hwnd):
        return False
    result = ctypes.c_size_t(0)
    return bool(
        a.SendMessageTimeoutW(
            hwnd, WM_NULL, 0, 0, SMTO_ABORTIFHUNG, RESPONSIVE_PROBE_MS, ctypes.byref(result)
        )
    )


def mdi_active_child(main_hwnd: int) -> int:
    """Active MDI child of the Access main window, or 0."""
    a = api()
    clients: list[int] = []

    def callback(child, _lparam):
        handle = int(child or 0)
        if handle and _class_name(handle) == MDI_CLIENT_CLASS:
            clients.append(handle)
            return False
        return True

    a.EnumChildWindows(main_hwnd, a.WNDENUMPROC(callback), 0)
    if not clients:
        return 0
    result = ctypes.c_size_t(0)
    ok = a.SendMessageTimeoutW(
        clients[0], WM_MDIGETACTIVE, 0, 0, SMTO_ABORTIFHUNG, 500, ctypes.byref(result)
    )
    return int(result.value) if ok else 0


def mdi_client_rect(main_hwnd: int) -> Rect | None:
    a = api()
    clients: list[int] = []

    def callback(child, _lparam):
        handle = int(child or 0)
        if handle and _class_name(handle) == MDI_CLIENT_CLASS:
            clients.append(handle)
            return False
        return True

    a.EnumChildWindows(main_hwnd, a.WNDENUMPROC(callback), 0)
    return window_rect(clients[0]) if clients else None


def is_top_level(hwnd: int) -> bool:
    return int(api().GetAncestor(hwnd, GA_ROOT) or 0) == hwnd


def is_visible(hwnd: int) -> bool:
    return bool(api().IsWindowVisible(hwnd))


def is_minimized(hwnd: int) -> bool:
    return bool(api().IsIconic(hwnd))


# ---------------------------------------------------------------------------
# Restart Manager
# ---------------------------------------------------------------------------


def pids_holding_file(path: str) -> set[int] | None:
    """PIDs with ``path`` open, from the Restart Manager; None if the query fails.

    The Restart Manager asks the kernel, not the process, so it answers while
    Access is hung. It can be blind to files on network shares, which shows
    up as an empty set rather than an error.
    """
    a = api()
    session = wintypes.DWORD(0)
    key = ctypes.create_unicode_buffer(CCH_RM_SESSION_KEY + 1)
    if a.RmStartSession(ctypes.byref(session), 0, key) != 0:
        return None
    try:
        files = (wintypes.LPCWSTR * 1)(path)
        if a.RmRegisterResources(session, 1, files, 0, None, 0, None) != 0:
            return None
        needed = wintypes.UINT(0)
        count = wintypes.UINT(0)
        reasons = wintypes.DWORD(0)
        for _ in range(4):
            infos = (_RM_PROCESS_INFO * max(1, needed.value))()
            count.value = len(infos)
            rc = a.RmGetList(session, ctypes.byref(needed), ctypes.byref(count), infos,
                             ctypes.byref(reasons))
            if rc == 0:
                return {int(infos[i].Process.dwProcessId) for i in range(count.value)}
            if rc != ERROR_MORE_DATA:
                return None
        return None
    finally:
        a.RmEndSession(session)


# ---------------------------------------------------------------------------
# Pixel capture
# ---------------------------------------------------------------------------


def _bitmap_info(width: int, height: int) -> _BITMAPINFO:
    bmi = _BITMAPINFO()
    bmi.bmiHeader.biSize = ctypes.sizeof(_BITMAPINFOHEADER)
    bmi.bmiHeader.biWidth = width
    bmi.bmiHeader.biHeight = -height
    bmi.bmiHeader.biPlanes = 1
    bmi.bmiHeader.biBitCount = 32
    bmi.bmiHeader.biCompression = BI_RGB
    return bmi


@contextmanager
def _dib_dc(width: int, height: int) -> Iterator[tuple[int, int]]:
    """A memory DC with a top-down 32 bpp DIB selected; yields (hdc, bits)."""
    a = api()
    screen = a.GetDC(None)
    mem = a.CreateCompatibleDC(screen)
    bits = ctypes.c_void_p()
    bmi = _bitmap_info(width, height)
    bitmap = a.CreateDIBSection(mem, ctypes.byref(bmi), DIB_RGB_COLORS, ctypes.byref(bits), None, 0)
    old = a.SelectObject(mem, bitmap) if bitmap else None
    try:
        if not bitmap or not bits.value:
            raise CaptureError(f"Could not allocate a {width}x{height} bitmap.")
        yield mem, bits.value
    finally:
        if old:
            a.SelectObject(mem, old)
        if bitmap:
            a.DeleteObject(bitmap)
        a.DeleteDC(mem)
        a.ReleaseDC(None, screen)


_stuck_lock = threading.Lock()
_stuck_print_workers: dict[int, threading.Thread] = {}


def _print_window_stuck(hwnd: int) -> bool:
    with _stuck_lock:
        worker = _stuck_print_workers.get(hwnd)
        if worker is not None and not worker.is_alive():
            del _stuck_print_workers[hwnd]
            worker = None
        return worker is not None


def print_window(hwnd: int, rect: Rect, timeout: float = PRINT_WINDOW_TIMEOUT_SEC) -> bytes | None:
    """Pixels of ``rect`` (the full window rect) via PrintWindow, or None.

    PrintWindow asks the target to render and can block on a stuck thread,
    so it runs on a daemon worker. A timed-out worker is remembered per
    window, and later calls skip straight to the fallback until it exits.
    """
    if _print_window_stuck(hwnd):
        return None
    box: dict[str, bytes | None] = {"bgra": None}

    def worker() -> None:
        with dpi_aware(), _dib_dc(rect.width, rect.height) as (mem, bits):
            if api().PrintWindow(hwnd, mem, PW_RENDERFULLCONTENT):
                api().GdiFlush()
                box["bgra"] = ctypes.string_at(bits, rect.width * rect.height * 4)

    thread = threading.Thread(target=worker, daemon=True, name="vcs-print-window")
    thread.start()
    thread.join(timeout)
    if thread.is_alive():
        with _stuck_lock:
            _stuck_print_workers[hwnd] = thread
        return None
    return box["bgra"]


def screen_copy(region: Rect) -> bytes:
    """Pixels of a screen region, including whatever covers it."""
    a = api()
    with _dib_dc(region.width, region.height) as (mem, bits):
        screen = a.GetDC(None)
        try:
            ok = a.BitBlt(mem, 0, 0, region.width, region.height, screen, region.left,
                          region.top, SRCCOPY | CAPTUREBLT)
        finally:
            a.ReleaseDC(None, screen)
        if not ok:
            raise CaptureError("Could not copy the screen region.")
        a.GdiFlush()
        return ctypes.string_at(bits, region.width * region.height * 4)


def capture_frame(info: WindowInfo) -> bytes | None:
    """The window's visible frame via PrintWindow, or None to fall back."""
    bgra = print_window(info.hwnd, info.rect)
    if bgra is None or is_blank(bgra):
        return None
    inner = info.frame.offset(-info.rect.left, -info.rect.top)
    pixels, _, _ = crop_buffer(bgra, info.rect.width, info.rect.height, inner)
    return pixels


def stretch(bgra: bytes, width: int, height: int, new_w: int, new_h: int) -> bytes:
    """Downscale with GDI HALFTONE, which averages instead of dropping pixels."""
    a = api()
    src = ctypes.create_string_buffer(bytes(bgra), len(bgra))
    bmi = _bitmap_info(width, height)
    with _dib_dc(new_w, new_h) as (mem, bits):
        a.SetStretchBltMode(mem, HALFTONE)
        a.SetBrushOrgEx(mem, 0, 0, None)
        a.StretchDIBits(mem, 0, 0, new_w, new_h, 0, 0, width, height, src, ctypes.byref(bmi),
                        DIB_RGB_COLORS, SRCCOPY)
        a.GdiFlush()
        return ctypes.string_at(bits, new_w * new_h * 4)


def capture_composite(
    base: WindowInfo,
    overlays: list[WindowInfo],
    all_windows: list[WindowInfo],
    *,
    crop: Rect | None = None,
    max_dimension: int = 1568,
) -> Capture:
    """Capture ``base`` with ``overlays`` painted over it, then scale and encode.

    A responsive window is rendered with PrintWindow, which ignores whatever
    covers it. A hung one is copied from the screen, with every other
    application's window above it grayed out so none of it leaks into the
    image.
    """
    warnings: list[str] = []
    canvas, placements = plan_canvas(
        base.frame, [o.frame for o in overlays], crop, virtual_screen()
    )
    pixels: bytes | bytearray | None = None
    method = "print_window"
    responsive = is_responsive(base.hwnd)
    if responsive:
        base_pixels = capture_frame(base)
        if base_pixels is not None:
            buffer = new_canvas(canvas.width, canvas.height)
            blit(buffer, canvas.width, canvas.height, base_pixels, base.frame.width,
                 base.frame.height, *placements[0])
            for index, overlay in enumerate(overlays, start=1):
                overlay_pixels = capture_frame(overlay)
                if overlay_pixels is None:
                    warnings.append(f"Could not render the window '{overlay.title or overlay.cls}'.")
                    continue
                blit(buffer, canvas.width, canvas.height, overlay_pixels, overlay.frame.width,
                     overlay.frame.height, *placements[index])
            pixels = buffer
    occluded = False
    if pixels is None:
        method = "screen"
        warnings.append(
            "Access is not responding to window messages, so this is a copy of the screen region."
            if not responsive
            else "PrintWindow returned no pixels, so this is a copy of the screen region."
        )
        occluded = is_occluded(canvas, base.z, base.pid, all_windows)
        pixels, masked = mask_foreign(screen_copy(canvas), canvas, base.z, base.pid, all_windows)
        if masked:
            warnings.append(
                f"{masked} window(s) of other applications covered Access and are grayed out."
            )
    new_w, new_h, scale = scale_for(canvas.width, canvas.height, max_dimension)
    if scale < 1.0:
        pixels = stretch(pixels, canvas.width, canvas.height, new_w, new_h)
    png = encode_png(bgra_to_rgb(pixels), new_w, new_h)
    return Capture(
        png=png,
        width=new_w,
        height=new_h,
        original_width=canvas.width,
        original_height=canvas.height,
        scale=scale,
        method=method,
        responsive=responsive,
        occluded=occluded,
        warnings=warnings,
    )
