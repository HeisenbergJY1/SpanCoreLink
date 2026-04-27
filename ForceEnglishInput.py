# -*- coding: utf-8 -*-
"""
ForceEnglishInput.py - Rhino/GH 自动锁定英文输入

功能：
- 切到 Rhino/GH 窗口时：自动切换到英文模式
- 切到其他程序时：自动恢复中文输入法

使用方式：Toggle 模式，点击一次开启，再点击一次关闭

通过临时标志文件实现跨脚本执行的状态通信，
解决 Rhino 8 每次 RunPythonScript 独立执行、全局变量不共享的问题。
"""
import ctypes
import ctypes.wintypes
import os
import tempfile
import threading
import time

try:
    import rhinoscriptsyntax as rs

    RHINO_AVAILABLE = True
except ImportError:
    RHINO_AVAILABLE = False

# ==================== Windows API ====================
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
imm32 = ctypes.windll.imm32

# 常量
WM_INPUTLANGCHANGEREQUEST = 0x0050
LANG_EN_US = 0x04090409  # 英文(美国)
LANG_ZH_CN = 0x08040804  # 中文(简体)

# SendInput 相关
INPUT_KEYBOARD = 1
KEYEVENTF_KEYUP = 0x0002
VK_SHIFT = 0x10


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", ctypes.wintypes.WORD),
        ("wScan", ctypes.wintypes.WORD),
        ("dwFlags", ctypes.wintypes.DWORD),
        ("time", ctypes.wintypes.DWORD),
        ("dwExtraInfo", ctypes.POINTER(ctypes.c_ulong)),
    ]


class INPUT(ctypes.Structure):
    class _INPUT(ctypes.Union):
        _fields_ = [("ki", KEYBDINPUT)]

    _anonymous_ = ("_input",)
    _fields_ = [
        ("type", ctypes.wintypes.DWORD),
        ("_input", _INPUT),
    ]


# 标志文件路径
_FLAG_FILE = os.path.join(tempfile.gettempdir(), "spancore_force_english.flag")


def _get_rhino_pid():
    """获取当前 Rhino 进程 ID"""
    try:
        import System.Diagnostics
        return System.Diagnostics.Process.GetCurrentProcess().Id
    except Exception:
        return 0


def _is_rhino_window(hwnd, rhino_pid):
    """判断窗口是否属于 Rhino/GH"""
    if not hwnd:
        return False
    if rhino_pid:
        pid = ctypes.wintypes.DWORD()
        user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
        if pid.value == rhino_pid:
            return True
    title = ctypes.create_unicode_buffer(512)
    user32.GetWindowTextW(hwnd, title, 512)
    title_str = title.value
    return "Rhino" in title_str or "Grasshopper" in title_str


def _get_layout_lang(hwnd):
    """获取窗口键盘布局的语言 ID"""
    tid = user32.GetWindowThreadProcessId(hwnd, None)
    layout = user32.GetKeyboardLayout(tid)
    return layout & 0xFFFF


def _is_chinese_mode(hwnd):
    """检测窗口是否处于中文输入模式（IME 或布局）"""
    # 先检查键盘布局
    if _get_layout_lang(hwnd) == 0x0804:
        # 中文布局，再检查 IME 转换状态
        tid = user32.GetWindowThreadProcessId(hwnd, None)
        current_tid = kernel32.GetCurrentThreadId()
        attached = False
        try:
            attached = user32.AttachThreadInput(current_tid, tid, True)
            focus = user32.GetFocus()
            target = focus if focus else hwnd
            himc = imm32.ImmGetContext(target)
            if himc:
                conversion = ctypes.wintypes.DWORD()
                sentence = ctypes.wintypes.DWORD()
                if imm32.ImmGetConversionStatus(
                    himc, ctypes.byref(conversion), ctypes.byref(sentence)
                ):
                    is_cn = bool(conversion.value & 0x0001)
                    imm32.ImmReleaseContext(target, himc)
                    return is_cn
                imm32.ImmReleaseContext(target, himc)
        except Exception:
            pass
        finally:
            if attached:
                user32.AttachThreadInput(current_tid, tid, False)
        return True  # 中文布局默认认为是中文模式
    return False


def _force_english(hwnd):
    """强制切换到英文输入（多重方案）"""
    tid = user32.GetWindowThreadProcessId(hwnd, None)
    current_tid = kernel32.GetCurrentThreadId()
    attached = False
    try:
        attached = user32.AttachThreadInput(current_tid, tid, True)

        # 方案1：ActivateKeyboardLayout 切换到英文布局
        hkl = user32.LoadKeyboardLayoutW("00000409", 0x00000001)
        if hkl:
            user32.ActivateKeyboardLayout(hkl, 0)

        # 方案2：IMM 设置转换模式为英文
        focus = user32.GetFocus()
        target = focus if focus else hwnd
        himc = imm32.ImmGetContext(target)
        if himc:
            conversion = ctypes.wintypes.DWORD()
            sentence = ctypes.wintypes.DWORD()
            if imm32.ImmGetConversionStatus(
                himc, ctypes.byref(conversion), ctypes.byref(sentence)
            ):
                if conversion.value & 0x0001:
                    new_mode = conversion.value & ~0x0001
                    imm32.ImmSetConversionStatus(himc, new_mode, sentence.value)
            imm32.ImmReleaseContext(target, himc)

        # 方案3：PostMessage 请求切换
        user32.PostMessageW(hwnd, WM_INPUTLANGCHANGEREQUEST, 0, LANG_EN_US)
    except Exception:
        pass
    finally:
        if attached:
            user32.AttachThreadInput(current_tid, tid, False)

    # 方案4：如果还是中文，模拟 Shift 键（微软拼音中英切换）
    time.sleep(0.05)
    if _is_chinese_mode(hwnd):
        _send_shift_key()


def _force_chinese(hwnd):
    """强制切换到中文输入（多重方案）"""
    tid = user32.GetWindowThreadProcessId(hwnd, None)
    current_tid = kernel32.GetCurrentThreadId()
    attached = False
    try:
        attached = user32.AttachThreadInput(current_tid, tid, True)

        # 方案1：ActivateKeyboardLayout 切换到中文布局
        hkl = user32.LoadKeyboardLayoutW("00000804", 0x00000001)
        if hkl:
            user32.ActivateKeyboardLayout(hkl, 0)

        # 方案2：IMM 设置转换模式为中文
        focus = user32.GetFocus()
        target = focus if focus else hwnd
        himc = imm32.ImmGetContext(target)
        if himc:
            conversion = ctypes.wintypes.DWORD()
            sentence = ctypes.wintypes.DWORD()
            if imm32.ImmGetConversionStatus(
                himc, ctypes.byref(conversion), ctypes.byref(sentence)
            ):
                if not (conversion.value & 0x0001):
                    new_mode = conversion.value | 0x0001
                    imm32.ImmSetConversionStatus(himc, new_mode, sentence.value)
            imm32.ImmReleaseContext(target, himc)

        # 方案3：PostMessage 请求切换
        user32.PostMessageW(hwnd, WM_INPUTLANGCHANGEREQUEST, 0, LANG_ZH_CN)
    except Exception:
        pass
    finally:
        if attached:
            user32.AttachThreadInput(current_tid, tid, False)

    # 方案4：如果还是英文，模拟 Shift 键切回中文
    time.sleep(0.05)
    if not _is_chinese_mode(hwnd):
        _send_shift_key()


def _send_shift_key():
    """模拟按下并释放 Shift 键（微软拼音中英切换）"""
    inputs = (INPUT * 2)()
    inputs[0].type = INPUT_KEYBOARD
    inputs[0].ki.wVk = VK_SHIFT
    inputs[0].ki.dwFlags = 0
    inputs[1].type = INPUT_KEYBOARD
    inputs[1].ki.wVk = VK_SHIFT
    inputs[1].ki.dwFlags = KEYEVENTF_KEYUP
    user32.SendInput(2, ctypes.byref(inputs), ctypes.sizeof(INPUT))


def _is_running():
    """通过标志文件检查是否正在运行"""
    return os.path.exists(_FLAG_FILE)


def _monitor_loop():
    """后台监控循环"""
    rhino_pid = _get_rhino_pid()
    print(f"[ForceEnglish] Started (Rhino PID: {rhino_pid})")

    prev_was_rhino = False
    cooldown = 0  # 切换冷却，避免频繁切换

    while _is_running():
        try:
            hwnd = user32.GetForegroundWindow()
            if not hwnd:
                time.sleep(0.3)
                continue

            is_rhino = _is_rhino_window(hwnd, rhino_pid)

            if is_rhino and not prev_was_rhino:
                # 从其他程序切到 Rhino → 切英文
                _force_english(hwnd)
                prev_was_rhino = True
                cooldown = 4

            elif is_rhino and prev_was_rhino:
                # 持续在 Rhino 中，检测是否又变成中文
                if cooldown > 0:
                    cooldown -= 1
                elif _is_chinese_mode(hwnd):
                    _force_english(hwnd)
                    cooldown = 4

            elif not is_rhino and prev_was_rhino:
                # 从 Rhino 切到其他程序 → 切中文
                _force_chinese(hwnd)
                prev_was_rhino = False
                cooldown = 4

            elif not is_rhino and not prev_was_rhino:
                # 在其他程序之间切换，不干预
                cooldown = 0

        except Exception:
            pass

        time.sleep(0.25)

    print("[ForceEnglish] Stopped")


def start_force_english():
    """开启强制英文输入"""
    if _is_running():
        print("[ForceEnglish] Already running")
        return
    with open(_FLAG_FILE, "w") as f:
        f.write(str(os.getpid()))
    t = threading.Thread(target=_monitor_loop, daemon=True)
    t.start()
    print("[ForceEnglish] ON")


def stop_force_english():
    """关闭强制英文输入"""
    if not _is_running():
        print("[ForceEnglish] Already stopped")
        return
    try:
        os.remove(_FLAG_FILE)
    except OSError:
        pass
    print("[ForceEnglish] OFF")


if __name__ == "__main__":
    if _is_running():
        stop_force_english()
    else:
        start_force_english()
