import os
import time
import tempfile
from pathlib import Path


def _ensure_windows():
    if os.name != 'nt':
        raise RuntimeError("当前系统不支持 COM 转换，请在 Windows 上运行")


def _safe_quit(app):
    """安全退出 COM 应用，兼容 Word 和 WPS"""
    if app is None:
        return
    try:
        app.Quit()
    except Exception:
        pass
    time.sleep(0.5)


def _safe_close(doc):
    """安全关闭文档"""
    if doc is None:
        return
    try:
        doc.Close(SaveChanges=False)
    except Exception:
        pass


def _create_app(prog_id):
    """创建 COM 应用实例"""
    import win32com.client

    try:
        app = win32com.client.DispatchEx(prog_id)
    except Exception:
        app = win32com.client.Dispatch(prog_id)

    try:
        app.Visible = False
    except Exception:
        pass
    try:
        app.DisplayAlerts = False
    except Exception:
        pass

    return app


def _detect_all_apps():
    """
    检测系统中所有可用的 Office 应用。
    返回列表：[(prog_id, name), ...]
    """
    _ensure_windows()
    import pythoncom
    import win32com.client

    candidates = [
        ('Kwps.Application', 'WPS'),          # WPS 2019+
        ('wps.Application', 'WPS'),            # WPS 旧版
        ('Word.Application', 'Microsoft Word'),
    ]

    available = []
    pythoncom.CoInitialize()
    try:
        for prog_id, name in candidates:
            try:
                app = win32com.client.Dispatch(prog_id)
                _safe_quit(app)
                available.append((prog_id, name))
            except Exception:
                continue
    finally:
        pythoncom.CoUninitialize()

    return available


def detect_office_app(prefer_wps=False):
    """
    检测安装的 Office 应用，返回 COM ProgID。

    Args:
        prefer_wps: 如果为 True，优先返回 WPS（即使 Word 也可用）

    Returns:
        (prog_id, app_name) 或 (None, None)
    """
    available = _detect_all_apps()
    if not available:
        return None, None

    if prefer_wps:
        # 优先找 WPS
        for prog_id, name in available:
            if name == 'WPS':
                return prog_id, name

    # 返回第一个可用的（默认顺序：WPS > Word）
    return available[0]


def convert_to_docx(input_path, output_path=None):
    """
    将 .doc/.wps 转换为 .docx
    """
    _ensure_windows()
    try:
        import pythoncom
    except ModuleNotFoundError as e:
        raise RuntimeError("缺少 pywin32（pythoncom）。请重新下载最新版 EXE 或安装 pywin32 后再试。") from e

    input_path = Path(input_path).resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"文件不存在: {input_path}")

    if output_path is None:
        fd, temp_path = tempfile.mkstemp(suffix='.docx')
        os.close(fd)
        output_path = temp_path

    output_path = Path(output_path).resolve()

    prog_id, _ = detect_office_app()
    if not prog_id:
        raise RuntimeError("未检测到 WPS 或 Microsoft Office，无法转换 .doc/.wps 文件")

    app = None
    doc = None
    pythoncom.CoInitialize()
    try:
        app = _create_app(prog_id)
        doc = app.Documents.Open(str(input_path))
        doc.SaveAs2(str(output_path), FileFormat=16)
        return str(output_path)
    finally:
        _safe_close(doc)
        _safe_quit(app)
        pythoncom.CoUninitialize()


def convert_from_docx(input_path, output_path, format='doc'):
    """
    将 .docx 转换回 .doc 或 .wps

    处理逻辑：
    - format='doc' → 用任意可用的 Office 应用保存
    - format='wps' → 专门尝试用 WPS 保存；如果 WPS 不可用，回退为 .doc

    Returns:
        str: 实际输出的文件路径（如果回退了格式，路径后缀会变）
    """
    _ensure_windows()
    try:
        import pythoncom
    except ModuleNotFoundError as e:
        raise RuntimeError("缺少 pywin32（pythoncom）。请重新下载最新版 EXE 或安装 pywin32 后再试。") from e

    # .wps 格式必须用 WPS Office 来保存
    if format == 'wps':
        prog_id, app_name = detect_office_app(prefer_wps=True)
        if not prog_id:
            raise RuntimeError("未检测到 Office/WPS")

        if app_name != 'WPS':
            # 系统只有 Word，没有 WPS → 回退为 .doc
            format = 'doc'
            output_path = str(Path(output_path).with_suffix('.doc'))
    else:
        prog_id, app_name = detect_office_app()
        if not prog_id:
            raise RuntimeError("未检测到 Office/WPS")

    # FileFormat 常量
    format_map = {
        'doc': 0,    # wdFormatDocument (.doc)
        'wps': 0,    # WPS 用 FileFormat=0 保存，扩展名决定实际格式
    }

    app = None
    doc = None
    pythoncom.CoInitialize()
    try:
        time.sleep(1)

        app = _create_app(prog_id)
        doc = app.Documents.Open(str(Path(input_path).resolve()))
        doc.SaveAs2(str(Path(output_path).resolve()), FileFormat=format_map.get(format, 0))

        return output_path
    finally:
        _safe_close(doc)
        _safe_quit(app)
        pythoncom.CoUninitialize()
