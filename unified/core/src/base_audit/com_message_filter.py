"""线程级 COM IMessageFilter：自动重试被 Excel/WPS 拒绝的呼入呼叫。

Win7 等慢速机器上，Excel 重算或保存工作簿时忙，COM 调用会被以
RPC_E_CALL_REJECTED（“被呼叫方拒绝接收呼叫”，0x80010001）拒绝。
注册 IMessageFilter 后，COM 会先询问本过滤器，我们回答“立即重试”，
让审核流程在 Excel 繁忙期间自动等待而不是直接失败。

pywin32 新版不再附带 win32com.messagefilter 示例模块，这里用纯
ctypes 实现（无新增依赖）。仅在 Windows 上可用，注册失败时静默降级
（保持旧行为，错误信息仍由调用链正常抛出）。

用法（每线程一次，随 ExcelSession 生命周期）：
    ok = register_com_message_filter()
    ...
    revoke_com_message_filter()
"""

from __future__ import annotations

import ctypes
from ctypes import WINFUNCTYPE, byref, c_void_p, c_ulong

# COM 约定常量
_S_OK = 0
_E_NOINTERFACE = -2147467262  # 0x80004002
_SERVERCALL_ISHANDLED = 0    # HandleInComingCall：随时可以接收呼叫
_PENDINGMSG_WAITDEFPROCESS = 2  # MessagePending：等待时继续默认消息处理

_filter = None    # 持有 COM 对象与回调，防止被垃圾回收
_previous = None  # 注册前线程里原有的过滤器，注销时恢复


def _make_vtable() -> ctypes.Structure:
    # IUnknown：QueryInterface / AddRef / Release
    # IMessageFilter：HandleInComingCall / RetryRejectedCall / MessagePending
    _QueryInterface = WINFUNCTYPE(c_ulong, c_void_p, c_void_p, c_void_p)
    _AddRef = WINFUNCTYPE(c_ulong, c_void_p)
    _Release = WINFUNCTYPE(c_ulong, c_void_p)
    _HandleInComingCall = WINFUNCTYPE(c_ulong, c_void_p, c_ulong, c_void_p, c_ulong, c_void_p)
    _RetryRejectedCall = WINFUNCTYPE(c_ulong, c_void_p, c_void_p, c_ulong, c_ulong)
    _MessagePending = WINFUNCTYPE(c_ulong, c_void_p, c_void_p, c_ulong, c_ulong)

    class _VTable(ctypes.Structure):
        _fields_ = [
            ("QueryInterface", _QueryInterface),
            ("AddRef", _AddRef),
            ("Release", _Release),
            ("HandleInComingCall", _HandleInComingCall),
            ("RetryRejectedCall", _RetryRejectedCall),
            ("MessagePending", _MessagePending),
        ]

    class _MessageFilter(ctypes.Structure):
        _fields_ = [("lpVtbl", ctypes.POINTER(_VTable))]

    vtable = _VTable(
        QueryInterface=_QueryInterface(lambda self, riid, ppv: _E_NOINTERFACE),
        AddRef=_AddRef(lambda self: 2),
        Release=_Release(lambda self: 1),
        # HandleInComingCall：随时可以接收呼叫。
        HandleInComingCall=_HandleInComingCall(
            lambda self, call_type, caller, tick, info: _SERVERCALL_ISHANDLED
        ),
        # RetryRejectedCall：返回 0..99 表示“立即重试”，<0 放弃，>=100 为等待毫秒数。
        RetryRejectedCall=_RetryRejectedCall(
            lambda self, callee, tick, reject_type: 1
        ),
        # MessagePending：等待时继续默认消息处理。
        MessagePending=_MessagePending(
            lambda self, callee, tick, pending_type: _PENDINGMSG_WAITDEFPROCESS
        ),
    )
    obj = _MessageFilter(lpVtbl=ctypes.pointer(vtable))
    # 回调与结构体必须以实例属性持有引用，防止被垃圾回收后指针悬空。
    obj._keepalive = vtable
    return obj


def register_com_message_filter() -> bool:
    """在当前线程注册重试过滤器；失败（非 Windows/COM 拒绝）返回 False。"""
    global _filter, _previous
    if _filter is not None:
        return True  # 本线程已注册
    try:
        ole32 = ctypes.oledll.ole32
        obj = _make_vtable()
        previous = c_void_p()
        ole32.CoRegisterMessageFilter(byref(obj), byref(previous))
        _filter = obj
        _previous = previous
        return True
    except Exception:
        return False


def revoke_com_message_filter() -> None:
    """注销过滤器并恢复线程原有过滤器；未注册时是空操作。"""
    global _filter, _previous
    if _filter is None:
        return
    try:
        ctypes.oledll.ole32.CoRegisterMessageFilter(
            byref(_previous) if _previous else None, None
        )
    except Exception:
        pass
    finally:
        _filter = None
        _previous = None
