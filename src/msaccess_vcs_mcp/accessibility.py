"""Read and press Office dialogs through MSAA, which window calls cannot see.

Office draws ``NUIDialog`` windows (for example "Cannot Compile Project.")
without child controls, so ``InternalGetWindowText`` finds no text or
buttons. Their accessibility tree has both. ``IAccessible`` is bound by
vtable through ctypes because pywin32 cannot pass a raw interface pointer
to ``AccessibleChildren``.

Every MSAA call sends ``WM_GETOBJECT`` to the dialog's thread, so callers
check the window is responsive first, and each operation runs on a daemon
worker with a deadline.
"""

from __future__ import annotations

import ctypes
import sys
import threading
from ctypes import wintypes
from dataclasses import dataclass
from typing import Callable, Iterable

VT_I4 = 3
VT_DISPATCH = 9
OBJID_CLIENT = 0xFFFFFFFC
ROLE_TITLEBAR = 0x01
ROLE_MENUBAR = 0x02
ROLE_SCROLLBAR = 0x03
ROLE_STATICTEXT = 0x29
ROLE_TEXT = 0x2A
ROLE_PUSHBUTTON = 0x2B
STATE_INVISIBLE = 0x8000
NON_CLIENT_ROLES = frozenset({ROLE_TITLEBAR, ROLE_MENUBAR, ROLE_SCROLLBAR})
TEXT_ROLES = frozenset({ROLE_STATICTEXT, ROLE_TEXT})
MAX_DEPTH = 12
MAX_NODES = 400
DEFAULT_TIMEOUT_SEC = 1.5
IID_IACCESSIBLE = "{618736E0-3C3D-11CF-810C-00AA00389B71}"

# IAccessible vtable slots after IUnknown (0-2) and IDispatch (3-6).
_RELEASE = 2
_GET_CHILD_COUNT = 8
_GET_NAME = 10
_GET_ROLE = 13
_GET_STATE = 14
_DO_DEFAULT_ACTION = 25


@dataclass(frozen=True)
class Node:
    """One element of the tree; ``ancestors`` are the roles above it."""

    ancestors: tuple[int, ...]
    role: int | None
    name: str
    invisible: bool = False
    handle: int = 0
    child_id: int = 0


def _label(name: str) -> str:
    return name.replace("&&", "\0").replace("&", "").replace("\0", "&").strip()


def _client_nodes(nodes: Iterable[Node]) -> list[Node]:
    return [
        n for n in nodes
        if not n.invisible and not NON_CLIENT_ROLES.intersection(n.ancestors)
        and n.role not in NON_CLIENT_ROLES
    ]


def dialog_content(nodes: Iterable[Node], title: str = "") -> tuple[list[str], list[str]]:
    """Message lines and button labels, without the title bar or scroll bars."""
    lines: list[str] = []
    buttons: list[str] = []
    for node in _client_nodes(nodes):
        text = node.name.strip()
        if not text:
            continue
        if node.role == ROLE_PUSHBUTTON:
            label = _label(text)
            if label and label not in buttons:
                buttons.append(label)
        elif node.role in TEXT_ROLES and text != title and text not in lines:
            lines.append(text)
    return lines, buttons


def find_button(nodes: Iterable[Node], label: str) -> Node | None:
    wanted = _label(label).casefold()
    for node in _client_nodes(nodes):
        if node.role == ROLE_PUSHBUTTON and _label(node.name).casefold() == wanted:
            return node
    return None


# ---------------------------------------------------------------------------
# Raw IAccessible
# ---------------------------------------------------------------------------


class _GUID(ctypes.Structure):
    _fields_ = [("d1", wintypes.DWORD), ("d2", wintypes.WORD), ("d3", wintypes.WORD),
                ("d4", ctypes.c_ubyte * 8)]


class _VARIANT(ctypes.Structure):
    _fields_ = [("vt", ctypes.c_ushort), ("r1", ctypes.c_ushort), ("r2", ctypes.c_ushort),
                ("r3", ctypes.c_ushort), ("val", ctypes.c_void_p), ("pad", ctypes.c_void_p)]


class _Msaa:
    def __init__(self) -> None:
        self.oleacc = ctypes.WinDLL("oleacc")
        self.oleaut = ctypes.WinDLL("oleaut32")
        self.oleaut.SysFreeString.argtypes = [ctypes.c_void_p]
        self.oleacc.AccessibleObjectFromWindow.argtypes = [
            wintypes.HWND, wintypes.DWORD, ctypes.POINTER(_GUID), ctypes.POINTER(ctypes.c_void_p)
        ]
        self.oleacc.AccessibleChildren.argtypes = [
            ctypes.c_void_p, ctypes.c_long, ctypes.c_long, ctypes.POINTER(_VARIANT),
            ctypes.POINTER(ctypes.c_long),
        ]
        self.iid = _GUID()
        ctypes.WinDLL("ole32").CLSIDFromString(ctypes.c_wchar_p(IID_IACCESSIBLE),
                                              ctypes.byref(self.iid))

    @staticmethod
    def _slot(ptr: int, index: int, restype, *argtypes):
        vtbl = ctypes.cast(ptr, ctypes.POINTER(ctypes.POINTER(ctypes.c_void_p))).contents
        return ctypes.WINFUNCTYPE(restype, ctypes.c_void_p, *argtypes)(vtbl[index])

    @staticmethod
    def _child(child_id: int) -> _VARIANT:
        v = _VARIANT()
        v.vt = VT_I4
        v.val = child_id
        return v

    def release(self, ptr: int) -> None:
        self._slot(ptr, _RELEASE, ctypes.c_ulong)(ptr)

    def from_window(self, hwnd: int) -> int:
        ptr = ctypes.c_void_p()
        hr = self.oleacc.AccessibleObjectFromWindow(hwnd, OBJID_CLIENT, ctypes.byref(self.iid),
                                                    ctypes.byref(ptr))
        if hr != 0 or not ptr.value:
            raise OSError(f"AccessibleObjectFromWindow failed: {hr & 0xFFFFFFFF:#x}")
        return int(ptr.value)

    def child_count(self, ptr: int) -> int:
        n = ctypes.c_long()
        self._slot(ptr, _GET_CHILD_COUNT, ctypes.HRESULT, ctypes.POINTER(ctypes.c_long))(
            ptr, ctypes.byref(n))
        return int(n.value)

    def name(self, ptr: int, child_id: int = 0) -> str:
        bstr = ctypes.c_void_p()
        try:
            self._slot(ptr, _GET_NAME, ctypes.HRESULT, _VARIANT, ctypes.POINTER(ctypes.c_void_p))(
                ptr, self._child(child_id), ctypes.byref(bstr))
        except OSError:
            return ""
        if not bstr.value:
            return ""
        try:
            return ctypes.wstring_at(bstr.value)
        finally:
            self.oleaut.SysFreeString(bstr)

    def _variant_int(self, ptr: int, slot: int, child_id: int) -> int | None:
        v = _VARIANT()
        try:
            self._slot(ptr, slot, ctypes.HRESULT, _VARIANT, ctypes.POINTER(_VARIANT))(
                ptr, self._child(child_id), ctypes.byref(v))
        except OSError:
            return None
        return ctypes.c_long(v.val or 0).value if v.vt == VT_I4 else None

    def role(self, ptr: int, child_id: int = 0) -> int | None:
        return self._variant_int(ptr, _GET_ROLE, child_id)

    def invisible(self, ptr: int, child_id: int = 0) -> bool:
        state = self._variant_int(ptr, _GET_STATE, child_id)
        return bool(state is not None and state & STATE_INVISIBLE)

    def do_default_action(self, ptr: int, child_id: int = 0) -> None:
        self._slot(ptr, _DO_DEFAULT_ACTION, ctypes.HRESULT, _VARIANT)(ptr, self._child(child_id))

    def children(self, ptr: int) -> list[tuple[str, int]]:
        count = self.child_count(ptr)
        if count <= 0:
            return []
        arr = (_VARIANT * count)()
        got = ctypes.c_long()
        if self.oleacc.AccessibleChildren(ptr, 0, count, arr, ctypes.byref(got)) < 0:
            return []
        out: list[tuple[str, int]] = []
        for v in arr[: got.value]:
            if v.vt == VT_DISPATCH and v.val:
                out.append(("obj", int(v.val)))
            elif v.vt == VT_I4:
                out.append(("id", ctypes.c_long(v.val or 0).value))
        return out

    def walk(self, root: int, owned: list[int]) -> list[Node]:
        """Every element under ``root``. Interface pointers go into ``owned``."""
        nodes: list[Node] = []

        def visit(ptr: int, ancestors: tuple[int, ...]) -> None:
            if len(ancestors) > MAX_DEPTH:
                return
            for kind, value in self.children(ptr):
                if len(nodes) >= MAX_NODES:
                    if kind == "obj":
                        self.release(value)
                    continue
                if kind == "obj":
                    owned.append(value)
                    role = self.role(value)
                    nodes.append(Node(ancestors, role, self.name(value), self.invisible(value),
                                      value, 0))
                    visit(value, (*ancestors, role or 0))
                else:
                    nodes.append(Node(ancestors, self.role(ptr, value), self.name(ptr, value),
                                      self.invisible(ptr, value), ptr, value))

        visit(root, ())
        return nodes


def _run_with_tree(hwnd: int, action: Callable[["_Msaa", list[Node]], object],
                   timeout: float) -> object | None:
    """Walk the dialog's tree on a COM worker and run ``action``; None on timeout."""
    if sys.platform != "win32":
        return None
    box: dict[str, object] = {}

    def worker() -> None:
        import pythoncom

        pythoncom.CoInitialize()
        msaa = _Msaa()
        owned: list[int] = []
        root = 0
        try:
            root = msaa.from_window(hwnd)
            box["result"] = action(msaa, msaa.walk(root, owned))
        except Exception as exc:
            box["error"] = exc
        finally:
            for ptr in owned:
                try:
                    msaa.release(ptr)
                except Exception:
                    pass
            if root:
                try:
                    msaa.release(root)
                except Exception:
                    pass
            pythoncom.CoUninitialize()

    thread = threading.Thread(target=worker, daemon=True, name="vcs-msaa")
    thread.start()
    thread.join(timeout)
    if thread.is_alive() or "error" in box:
        return None
    return box.get("result")


def read_dialog(hwnd: int, title: str = "", timeout: float = DEFAULT_TIMEOUT_SEC
                ) -> tuple[list[str], list[str]] | None:
    """(message lines, button labels), or None when the tree could not be read."""
    result = _run_with_tree(hwnd, lambda _m, nodes: dialog_content(nodes, title), timeout)
    return result if isinstance(result, tuple) else None


def press_button(hwnd: int, label: str, timeout: float = DEFAULT_TIMEOUT_SEC) -> bool:
    """Run the default action of the button labelled ``label``."""

    def action(msaa: _Msaa, nodes: list[Node]) -> bool:
        node = find_button(nodes, label)
        if node is None:
            return False
        msaa.do_default_action(node.handle, node.child_id)
        return True

    return _run_with_tree(hwnd, action, timeout) is True
