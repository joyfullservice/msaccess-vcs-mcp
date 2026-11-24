"""Tests for the Win32 capture helpers that take plain data.

Window geometry, overlay choice, dialog text, canvas planning, and PNG
encoding are all pure; the Win32 calls they sit between are patched.
"""

from __future__ import annotations

import struct
import zlib

import pytest

import msaccess_vcs_mcp.screen_capture as sc
from msaccess_vcs_mcp.screen_capture import ChildControl, Rect, WindowInfo

MAIN = 0x100
VBE = 0x200


def window(hwnd, cls, *, pid=10, z=5, owner=0, frame=None, **kwargs) -> WindowInfo:
    frame = frame or Rect(0, 0, 100, 80)
    return WindowInfo(
        hwnd=hwnd, pid=pid, cls=cls, title=kwargs.pop("title", cls), rect=frame,
        frame=frame, owner=owner, z=z, **kwargs,
    )


def decode_png(data: bytes) -> tuple[int, int, bytes]:
    """Width, height, and RGB rows of a filter-0 24-bit PNG."""
    assert data[:8] == b"\x89PNG\r\n\x1a\n"
    pos, idat, width, height = 8, b"", 0, 0
    while pos < len(data):
        (length,) = struct.unpack(">I", data[pos:pos + 4])
        kind = data[pos + 4:pos + 8]
        body = data[pos + 8:pos + 8 + length]
        (crc,) = struct.unpack(">I", data[pos + 8 + length:pos + 12 + length])
        assert crc == zlib.crc32(kind + body) & 0xFFFFFFFF
        if kind == b"IHDR":
            width, height, depth, color = struct.unpack(">IIBB", body[:10])
            assert (depth, color) == (8, 2)
        elif kind == b"IDAT":
            idat += body
        pos += 12 + length
    raw = zlib.decompress(idat)
    stride = width * 3 + 1
    assert all(raw[y * stride] == 0 for y in range(height))
    rgb = b"".join(raw[y * stride + 1:(y + 1) * stride] for y in range(height))
    return width, height, rgb


def coordinate_buffer(width: int, height: int) -> bytes:
    """BGRA pixels whose blue is x and green is y, to check where pixels land."""
    return b"".join(bytes([x, y, 0, 255]) for y in range(height) for x in range(width))


# ---------------------------------------------------------------------------
# Pixels and PNG
# ---------------------------------------------------------------------------


def test_png_round_trip():
    rgb = bytes(range(2 * 3 * 3))
    width, height, decoded = decode_png(sc.encode_png(rgb, 3, 2))
    assert (width, height) == (3, 2)
    assert decoded == rgb


def test_png_rejects_wrong_buffer_size():
    with pytest.raises(ValueError):
        sc.encode_png(b"\x00" * 10, 3, 2)


def test_bgra_to_rgb_swaps_channels_and_drops_alpha():
    assert sc.bgra_to_rgb(bytes([1, 2, 3, 255, 4, 5, 6, 0])) == bytearray([3, 2, 1, 6, 5, 4])


def test_is_blank():
    assert sc.is_blank(bytes([0, 0, 0, 255]) * 4)
    assert not sc.is_blank(bytes([0, 0, 0, 255, 0, 9, 0, 255]))


@pytest.mark.parametrize(
    "size,limit,expected",
    [
        ((800, 600), 1568, (800, 600, 1.0)),
        ((1568, 900), 1568, (1568, 900, 1.0)),
        ((3136, 1000), 1568, (1568, 500, 0.5)),
        ((1000, 4000), 400, (100, 400, 0.1)),
    ],
)
def test_scale_for_only_scales_down(size, limit, expected):
    assert sc.scale_for(*size, limit) == expected


def test_blit_clips_to_destination():
    dst = sc.new_canvas(4, 3)
    src = bytes([9, 9, 9, 255]) * 4
    sc.blit(dst, 4, 3, src, 2, 2, 3, 2)
    painted = [i // 4 for i in range(0, len(dst), 4) if dst[i] == 9]
    assert painted == [2 * 4 + 3]


def test_blit_negative_offset():
    dst = sc.new_canvas(2, 2)
    sc.blit(dst, 2, 2, coordinate_buffer(3, 3), 3, 3, -1, -1)
    assert (dst[0], dst[1]) == (1, 1)


def test_crop_buffer():
    pixels, w, h = sc.crop_buffer(coordinate_buffer(5, 4), 5, 4, Rect(1, 2, 4, 4))
    assert (w, h) == (3, 2)
    assert (pixels[0], pixels[1]) == (1, 2)
    assert (pixels[-4], pixels[-3]) == (3, 3)


def test_crop_buffer_outside_raises():
    with pytest.raises(sc.CaptureError):
        sc.crop_buffer(coordinate_buffer(2, 2), 2, 2, Rect(5, 5, 8, 8))


def test_capture_frame_crops_invisible_borders(monkeypatch):
    """GetWindowRect includes the invisible resize border (7px left, right,
    bottom on Windows 10/11); the visible frame is the DWM extended bounds."""
    rect = Rect(93, 100, 117, 117)
    frame = Rect(100, 100, 110, 110)
    info = WindowInfo(hwnd=1, pid=1, cls=sc.MAIN_WINDOW_CLASS, title="", rect=rect, frame=frame)
    monkeypatch.setattr(sc, "print_window", lambda hwnd, r: coordinate_buffer(r.width, r.height))

    pixels = sc.capture_frame(info)

    assert len(pixels) == frame.width * frame.height * 4
    assert (pixels[0], pixels[1]) == (7, 0)
    assert (pixels[-4], pixels[-3]) == (16, 9)


def test_capture_frame_blank_print_window_falls_back(monkeypatch):
    info = window(1, sc.MAIN_WINDOW_CLASS, frame=Rect(0, 0, 4, 4))
    monkeypatch.setattr(sc, "print_window", lambda hwnd, r: bytes([0, 0, 0, 255]) * 16)
    assert sc.capture_frame(info) is None


# ---------------------------------------------------------------------------
# Canvas planning
# ---------------------------------------------------------------------------


def test_plan_canvas_includes_overhanging_dialog():
    canvas, placements = sc.plan_canvas(Rect(100, 100, 500, 400), [Rect(450, 150, 650, 250)])
    assert canvas == Rect(100, 100, 650, 400)
    assert placements == [(0, 0), (350, 50)]


def test_plan_canvas_clips_to_virtual_screen():
    canvas, placements = sc.plan_canvas(
        Rect(-50, 0, 300, 200), [], virtual_screen=Rect(0, 0, 1920, 1080)
    )
    assert canvas == Rect(0, 0, 300, 200)
    assert placements == [(-50, 0)]


def test_plan_canvas_crop():
    canvas, placements = sc.plan_canvas(Rect(0, 0, 800, 600), [], crop=Rect(100, 120, 400, 320))
    assert canvas == Rect(100, 120, 400, 320)
    assert placements == [(-100, -120)]


def test_plan_canvas_off_screen_raises():
    with pytest.raises(sc.CaptureError):
        sc.plan_canvas(Rect(5000, 0, 5100, 100), [], virtual_screen=Rect(0, 0, 1920, 1080))


# ---------------------------------------------------------------------------
# Window classification
# ---------------------------------------------------------------------------


def test_dialog_classes():
    assert sc.is_dialog_class("#32770")
    assert sc.is_dialog_class("NUIDialog")
    assert sc.is_dialog_class("bosa_sdm_Microsoft Office Access")
    assert not sc.is_dialog_class("OForm")


def test_chrome_detection():
    assert sc.is_chrome(window(1, "MSO_BORDEREFFECT_WINDOW_CLASS"))
    layered = window(2, "Something", ex_style=sc.WS_EX_LAYERED | sc.WS_EX_TRANSPARENT)
    assert sc.is_chrome(layered)
    assert not sc.is_chrome(window(3, "#32770", ex_style=sc.WS_EX_LAYERED))


def test_select_overlays_filters_and_orders_bottom_to_top():
    main = window(MAIN, sc.MAIN_WINDOW_CLASS, z=10)
    vbe = window(VBE, sc.VBE_WINDOW_CLASS, z=3)
    msgbox = window(1, "#32770", owner=MAIN, z=1)
    nested = window(2, "#32770", owner=1, z=0)
    popup_form = window(3, "OForm", owner=MAIN, z=4)
    glow = window(4, "MSO_BORDEREFFECT_WINDOW_CLASS", owner=MAIN, z=9)
    hidden = window(5, "#32770", owner=MAIN, z=2, visible=False)
    compile_error = window(6, "#32770", owner=VBE, z=2)
    orphan_dialog = window(7, "#32770", owner=0, z=6)
    windows = [nested, msgbox, compile_error, hidden, vbe, popup_form, orphan_dialog, glow, main]

    overlays = sc.select_overlays(windows, main)

    assert [w.hwnd for w in overlays] == [7, 3, 1, 2]


def test_select_overlays_for_vbe_excludes_main_window():
    main = window(MAIN, sc.MAIN_WINDOW_CLASS, z=1)
    vbe = window(VBE, sc.VBE_WINDOW_CLASS, z=5)
    compile_error = window(6, "#32770", owner=VBE, z=0)
    app_dialog = window(1, "#32770", owner=MAIN, z=0)

    overlays = sc.select_overlays([compile_error, app_dialog, main, vbe], vbe)

    assert [w.hwnd for w in overlays] == [6]


def test_owner_label():
    main = window(MAIN, sc.MAIN_WINDOW_CLASS)
    vbe = window(VBE, sc.VBE_WINDOW_CLASS)
    by_hwnd = {w.hwnd: w for w in (main, vbe)}
    assert sc.owner_label(window(1, "#32770", owner=VBE), MAIN, VBE, by_hwnd) == "vbe"
    assert sc.owner_label(window(2, "#32770", owner=MAIN), MAIN, VBE, by_hwnd) == "application"
    assert sc.owner_label(window(3, "#32770"), MAIN, VBE, by_hwnd) == "other"


def test_owner_chain_survives_cycles():
    a = window(1, "#32770", owner=2)
    b = window(2, "#32770", owner=1)
    assert not sc.owner_chain_reaches(a, MAIN, {1: a, 2: b})


def test_is_occluded():
    base = Rect(0, 0, 500, 400)
    above_other = window(1, "Notepad", pid=99, z=1, frame=Rect(400, 300, 700, 600))
    below_other = window(2, "Notepad", pid=99, z=9, frame=Rect(0, 0, 100, 100))
    own_dialog = window(3, "#32770", pid=10, z=0, frame=Rect(0, 0, 100, 100))
    far_away = window(4, "Notepad", pid=99, z=0, frame=Rect(900, 900, 1000, 1000))

    assert sc.is_occluded(base, 5, 10, [above_other])
    assert not sc.is_occluded(base, 5, 10, [below_other, own_dialog, far_away])


def test_find_window_prefers_visible_then_topmost():
    hidden = window(1, sc.MAIN_WINDOW_CLASS, z=0, visible=False)
    lower = window(2, sc.MAIN_WINDOW_CLASS, z=4)
    upper = window(3, sc.MAIN_WINDOW_CLASS, z=2)
    assert sc.find_window([hidden, lower, upper], sc.MAIN_WINDOW_CLASS).hwnd == 3
    assert sc.find_window([hidden], sc.VBE_WINDOW_CLASS) is None


# ---------------------------------------------------------------------------
# Dialog text
# ---------------------------------------------------------------------------


def test_strip_accelerators():
    assert sc.strip_accelerators("&Yes") == "Yes"
    assert sc.strip_accelerators("Save && &Close") == "Save & Close"


def test_read_dialogs_with_injected_reader():
    main = window(MAIN, sc.MAIN_WINDOW_CLASS, z=10, frame=Rect(0, 0, 800, 600))
    vbe = window(VBE, sc.VBE_WINDOW_CLASS, z=8, frame=Rect(900, 0, 1500, 600))
    compile_error = window(
        1, "#32770", owner=VBE, z=0, title="Microsoft Visual Basic for Applications",
        frame=Rect(1000, 100, 1300, 250),
    )
    msgbox = window(2, "#32770", owner=MAIN, z=1, title="Example", frame=Rect(200, 200, 500, 350))
    no_text = window(3, "NUIDialog", owner=MAIN, z=2, title="", frame=Rect(10, 10, 50, 50))
    children = {
        1: [
            ChildControl("Static", "Compile error:"),
            ChildControl("Static", "Variable not defined"),
            ChildControl("Button", "OK"),
            ChildControl("Button", "&Help"),
        ],
        2: [ChildControl("Static", "Delete 3 records?"), ChildControl("Button", "&Yes"),
            ChildControl("Button", "&No"), ChildControl("Static", "  ")],
        3: [],
    }

    dialogs, warnings = sc.read_dialogs(
        [main, vbe, compile_error, msgbox, no_text], MAIN, VBE, main.frame, children.__getitem__,
        lambda hwnd, title: None,
    )

    assert [d["owner"] for d in dialogs] == ["vbe", "application", "application"]
    assert dialogs[0]["text"] == "Compile error:\nVariable not defined"
    assert dialogs[0]["buttons"] == ["OK", "Help"]
    assert dialogs[0]["overlaps_target"] is False
    assert dialogs[1]["text"] == "Delete 3 records?"
    assert dialogs[1]["buttons"] == ["Yes", "No"]
    assert dialogs[1]["overlaps_target"] is True
    assert len(warnings) == 1 and "NUIDialog" in warnings[0]


def test_read_dialogs_skips_hidden_and_cloaked():
    hidden = window(1, "#32770", visible=False)
    cloaked = window(2, "#32770", cloaked=True)
    dialogs, warnings = sc.read_dialogs([hidden, cloaked], MAIN, 0, None, lambda h: [])
    assert dialogs == [] and warnings == []


def test_read_dialogs_reads_office_dialogs_through_accessibility():
    nui = window(3, "NUIDialog", owner=MAIN, z=0, title="Microsoft Access")
    asked: list[tuple[int, str]] = []

    def accessible(hwnd, title):
        asked.append((hwnd, title))
        return ["Cannot Compile Project."], ["OK"]

    dialogs, warnings = sc.read_dialogs([nui], MAIN, 0, None, lambda h: [], accessible)

    assert asked == [(3, "Microsoft Access")]
    assert dialogs[0]["text"] == "Cannot Compile Project."
    assert dialogs[0]["buttons"] == ["OK"]
    assert dialogs[0]["text_source"] == "accessibility"
    assert warnings == []


def test_accessibility_is_not_asked_when_window_text_is_readable():
    msgbox = window(2, "#32770", owner=MAIN, z=0)
    dialogs, _ = sc.read_dialogs(
        [msgbox], MAIN, 0, None, lambda h: [ChildControl("Button", "OK")],
        lambda hwnd, title: pytest.fail("accessibility read for a readable dialog"),
    )
    assert dialogs[0]["text_source"] == "window"


def test_dialog_id_is_stable_and_tracks_content():
    info = window(1, "#32770", title="Example")
    first = sc.dialog_entry(info, [ChildControl("Static", "Delete?"), ChildControl("Button", "OK")],
                            "application", False)
    again = sc.dialog_entry(info, [ChildControl("Static", "Delete?"), ChildControl("Button", "OK")],
                            "application", True)
    changed = sc.dialog_entry(info, [ChildControl("Static", "Save?"), ChildControl("Button", "OK")],
                              "application", False)
    other_window = sc.dialog_entry(window(2, "#32770", title="Example"),
                                   [ChildControl("Static", "Delete?"), ChildControl("Button", "OK")],
                                   "application", False)

    assert first["dialog_id"] == again["dialog_id"]
    assert len(first["dialog_id"]) == 12
    assert changed["dialog_id"] != first["dialog_id"]
    assert other_window["dialog_id"] != first["dialog_id"]


@pytest.mark.parametrize(
    "title, expected",
    [
        ("Microsoft Visual Basic for Applications - Db1 [break] - [modX (Code)]", "break"),
        ("Microsoft Visual Basic for Applications - Db1 [running]", "running"),
        ("Microsoft Visual Basic for Applications - Db1", "design"),
    ],
)
def test_vbe_state_from_title(title, expected):
    assert sc.vbe_state([window(VBE, sc.VBE_WINDOW_CLASS, title=title, visible=False)]) == expected


def test_vbe_state_without_editor():
    assert sc.vbe_state([window(MAIN, sc.MAIN_WINDOW_CLASS)]) is None


def test_select_overlays_skips_windows_that_do_not_touch_the_base():
    main = window(MAIN, sc.MAIN_WINDOW_CLASS, z=5, frame=Rect(0, 0, 800, 600))
    near = window(1, "#32770", owner=MAIN, z=1, frame=Rect(700, 500, 900, 700))
    other_monitor = window(2, "#32770", owner=MAIN, z=0, frame=Rect(2000, 100, 2300, 300))

    overlays = sc.select_overlays([other_monitor, near, main], main)

    assert [w.hwnd for w in overlays] == [1]


def test_dialog_text_is_bounded():
    info = window(1, "#32770")
    entry = sc.dialog_entry(info, [ChildControl("Static", "x" * 10000)], "application", False)
    assert len(entry["text"]) == sc.MAX_DIALOG_TEXT + 3


# ---------------------------------------------------------------------------
# Restart Manager
# ---------------------------------------------------------------------------


class _FailingRm:
    def __init__(self, start_rc=0, register_rc=0):
        self.start_rc = start_rc
        self.register_rc = register_rc
        self.ended = False

    def RmStartSession(self, session, flags, key):
        return self.start_rc

    def RmRegisterResources(self, *args):
        return self.register_rc

    def RmEndSession(self, session):
        self.ended = True
        return 0


def test_pids_holding_file_start_failure_returns_none(monkeypatch):
    fake = _FailingRm(start_rc=5)
    monkeypatch.setattr(sc, "api", lambda: fake)
    assert sc.pids_holding_file(r"C:\path\to\example.accdb") is None
    assert fake.ended is False


def test_pids_holding_file_register_failure_ends_session(monkeypatch):
    fake = _FailingRm(register_rc=87)
    monkeypatch.setattr(sc, "api", lambda: fake)
    assert sc.pids_holding_file(r"C:\path\to\example.accdb") is None
    assert fake.ended is True


# ---------------------------------------------------------------------------
# Composite capture
# ---------------------------------------------------------------------------


@pytest.fixture
def desktop(monkeypatch):
    monkeypatch.setattr(sc, "virtual_screen", lambda: Rect(0, 0, 1920, 1080))
    monkeypatch.setattr(sc, "dpi_aware", _null_context)


class _null_context:
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def _solid(info: WindowInfo, value: int) -> bytes:
    return bytes([value, value, value, 255]) * (info.frame.width * info.frame.height)


def test_composite_paints_dialog_over_base(desktop, monkeypatch):
    base = window(MAIN, sc.MAIN_WINDOW_CLASS, frame=Rect(0, 0, 4, 3))
    dialog = window(1, "#32770", owner=MAIN, z=0, frame=Rect(2, 1, 6, 2))
    monkeypatch.setattr(sc, "is_responsive", lambda hwnd: True)
    monkeypatch.setattr(sc, "capture_frame", lambda info: _solid(info, 200 if info is dialog else 50))

    capture = sc.capture_composite(base, [dialog], [base, dialog])

    width, height, rgb = decode_png(capture.png)
    assert (width, height) == (6, 3)
    assert capture.method == "print_window" and capture.responsive
    row1 = [rgb[(1 * 6 + x) * 3] for x in range(6)]
    assert row1 == [50, 50, 200, 200, 200, 200]
    assert rgb[(0 * 6 + 5) * 3] == sc.NEUTRAL_GRAY


def test_composite_falls_back_to_screen_when_hung(desktop, monkeypatch):
    base = window(MAIN, sc.MAIN_WINDOW_CLASS, frame=Rect(0, 0, 4, 4), z=5)
    cover = window(9, "Notepad", pid=99, z=1, frame=Rect(2, 2, 10, 10))
    monkeypatch.setattr(sc, "is_responsive", lambda hwnd: False)
    monkeypatch.setattr(sc, "capture_frame", lambda info: pytest.fail("PrintWindow on a hung window"))
    monkeypatch.setattr(sc, "screen_copy", lambda region: _solid(base, 77))

    capture = sc.capture_composite(base, [], [base, cover])

    assert capture.method == "screen"
    assert capture.responsive is False
    assert capture.occluded is True
    assert any("not responding" in w for w in capture.warnings)
    assert any("grayed out" in w for w in capture.warnings)
    _, _, rgb = decode_png(capture.png)
    assert rgb[(0 * 4 + 0) * 3] == 77
    assert rgb[(3 * 4 + 3) * 3] == sc.NEUTRAL_GRAY


def test_mask_foreign_restores_access_windows_above_a_foreign_one():
    canvas = Rect(0, 0, 6, 1)
    pixels = bytes(b for v in range(6) for b in (v, v, v, 255))
    foreign = window(9, "Notepad", pid=99, z=3, frame=Rect(0, 0, 4, 1))
    own_dialog = window(1, "#32770", pid=10, z=1, frame=Rect(2, 0, 3, 1))
    below = window(8, "Notepad", pid=99, z=9, frame=Rect(4, 0, 6, 1))
    glow = window(7, "MSO_BORDEREFFECT_WINDOW_CLASS", pid=99, z=0, frame=Rect(5, 0, 6, 1))

    result, masked = sc.mask_foreign(pixels, canvas, 5, 10, [own_dialog, foreign, below, glow])

    assert masked == 1
    assert [result[x * 4] for x in range(6)] == [
        sc.NEUTRAL_GRAY, sc.NEUTRAL_GRAY, 2, sc.NEUTRAL_GRAY, 4, 5
    ]


def test_composite_falls_back_when_print_window_returns_nothing(desktop, monkeypatch):
    base = window(MAIN, sc.MAIN_WINDOW_CLASS, frame=Rect(0, 0, 4, 4))
    monkeypatch.setattr(sc, "is_responsive", lambda hwnd: True)
    monkeypatch.setattr(sc, "capture_frame", lambda info: None)
    monkeypatch.setattr(sc, "screen_copy", lambda region: _solid(base, 77))

    capture = sc.capture_composite(base, [], [base])

    assert capture.method == "screen"
    assert capture.responsive is True
    assert any("PrintWindow returned no pixels" in w for w in capture.warnings)


def test_composite_scales_down(desktop, monkeypatch):
    base = window(MAIN, sc.MAIN_WINDOW_CLASS, frame=Rect(0, 0, 800, 400))
    monkeypatch.setattr(sc, "is_responsive", lambda hwnd: True)
    monkeypatch.setattr(sc, "capture_frame", lambda info: _solid(info, 10))
    monkeypatch.setattr(
        sc, "stretch", lambda px, w, h, nw, nh: bytes([10, 10, 10, 255]) * (nw * nh)
    )

    capture = sc.capture_composite(base, [], [base], max_dimension=400)

    assert (capture.width, capture.height) == (400, 200)
    assert (capture.original_width, capture.original_height) == (800, 400)
    assert capture.image_meta()["scale"] == 0.5
