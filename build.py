#!/usr/bin/env python3
"""
打包脚本 - 生成 Windows/Linux/macOS 可执行文件
用法：python build.py [windows|linux|macos|all|clean]
"""

import os
import re
import sys
import json
import shutil
import subprocess
import platform
import plistlib
import time
from pathlib import Path


def _configure_console_encoding():
    """Keep GitHub Windows runners from failing on Chinese build logs."""
    for stream_name in ("stdout", "stderr"):
        stream = getattr(sys, stream_name, None)
        if hasattr(stream, "reconfigure"):
            try:
                stream.reconfigure(encoding="utf-8", errors="replace")
            except Exception:
                pass


_configure_console_encoding()

# 配置
APP_NAME = "公文格式处理工具"
APP_NAME_EN = "DocFormatter"
VERSION = "1.8.8.4"
MAIN_SCRIPT = "docformat_gui.py"
MACOS_APP_BUNDLE_NAME = os.environ.get("MACOS_APP_BUNDLE_NAME", "公文格式处理助手").strip()

# 输出目录
DIST_DIR = Path("dist")
BUILD_DIR = Path("build")
ASSETS_DIR = Path("assets")
WINDOWS_ICON = ASSETS_DIR / "icon.ico"
MACOS_ICON = ASSETS_DIR / "icon.icns"

# macOS 签名 / 公证配置（通过环境变量提供；本地开发可不设，会自动退回 ad-hoc 签名）
#   MACOS_SIGN_IDENTITY   —— "Developer ID Application: Your Name (TEAMID)" 或其证书指纹
#   MACOS_NOTARY_PROFILE  —— 预先用 notarytool store-credentials 存好的 keychain profile 名
#   （或改用 MACOS_NOTARY_APPLE_ID / MACOS_NOTARY_TEAM_ID / MACOS_NOTARY_PASSWORD 三件套）
#   MACOS_NOTARIZE_APP    —— 设为 1 时额外公证 .app；默认只公证最终 .dmg，构建更快
MACOS_ENTITLEMENTS_CANDIDATES = [
    Path("packaging/macos/entitlements.plist"),
    Path("entitlements.plist"),
]
MACOS_REQUIRE_NOTARIZATION = os.environ.get("MACOS_REQUIRE_NOTARIZATION", "").strip() == "1"
MACOS_NOTARIZE_APP = os.environ.get("MACOS_NOTARIZE_APP", "").strip() == "1"
MACOS_NOTARY_WAIT_TIMEOUT = os.environ.get("MACOS_NOTARY_WAIT_TIMEOUT", "15m").strip() or "15m"
MACOS_NOTARY_WAIT_RETRIES = int(os.environ.get("MACOS_NOTARY_WAIT_RETRIES", "4") or "4")


def _get_macos_entitlements_path():
    """返回 macOS 签名用 entitlements 文件路径。"""
    for path in MACOS_ENTITLEMENTS_CANDIDATES:
        if path.exists():
            return path
    return MACOS_ENTITLEMENTS_CANDIDATES[0]


def check_pyinstaller():
    """检查 PyInstaller 是否安装"""
    try:
        import PyInstaller
        print(f"✓ PyInstaller {PyInstaller.__version__} 已安装")
        return True
    except ImportError:
        print("✗ PyInstaller 未安装")
        print("  请运行: pip install pyinstaller")
        return False


def _add_pyinstaller_icon(cmd, icon_path):
    """如果图标文件存在，将其加入 PyInstaller 参数。"""
    icon = Path(icon_path)
    if icon.exists():
        cmd[-1:-1] = [f"--icon={icon}"]
    else:
        print(f"  [提示] 未找到图标文件，跳过应用图标: {icon}")


def clean():
    """清理构建目录"""
    print("\n清理旧构建文件...")
    for d in [DIST_DIR, BUILD_DIR]:
        if d.exists():
            shutil.rmtree(d)
            print(f"  删除 {d}/")
    
    # 删除 spec 文件
    for f in Path(".").glob("*.spec"):
        f.unlink()
        print(f"  删除 {f}")


def _get_docx_templates_path():
    """获取 python-docx 模板文件路径"""
    try:
        import docx
        templates_dir = Path(docx.__file__).parent / "templates"
        if templates_dir.exists():
            return str(templates_dir)
    except ImportError:
        pass
    return None


def _get_tkinterdnd_pyinstaller_args(platform_name):
    """收集 tkinterdnd2 的 tkdnd 运行库，确保打包后拖拽可用。"""
    try:
        from PyInstaller.utils.hooks import collect_data_files, collect_dynamic_libs
    except Exception as e:
        print(f"  [警告] 无法加载 PyInstaller hook 工具，跳过 tkinterdnd2 资源收集: {e}")
        return []

    try:
        data_files = collect_data_files("tkinterdnd2")
        dynamic_libs = collect_dynamic_libs("tkinterdnd2")
    except Exception as e:
        print(f"  [警告] 无法收集 tkinterdnd2 资源，拖拽功能在打包版中可能不可用: {e}")
        return []

    sep = ";" if platform_name == "windows" else ":"
    args = ["--hidden-import=tkinterdnd2"]

    for src, dest in data_files:
        args.append(f"--add-data={src}{sep}{dest}")
    for src, dest in dynamic_libs:
        args.append(f"--add-binary={src}{sep}{dest}")

    if data_files or dynamic_libs:
        print(f"  ✓ 已收集 tkinterdnd2 资源: data={len(data_files)} binary={len(dynamic_libs)}")

    return args


def _patch_macos_info_plist(app_path):
    """修正 macOS app bundle 的基础元数据。"""
    plist_path = Path(app_path) / "Contents" / "Info.plist"
    if not plist_path.exists():
        return

    with open(plist_path, "rb") as f:
        data = plistlib.load(f)

    macos_name = MACOS_APP_BUNDLE_NAME or APP_NAME
    data["CFBundleName"] = macos_name
    data["CFBundleDisplayName"] = macos_name
    data["CFBundleIdentifier"] = "com.kagurananaga.docformat-gui"
    data["CFBundleDevelopmentRegion"] = "zh_CN"
    data["CFBundleLocalizations"] = ["zh_CN", "en"]

    with open(plist_path, "wb") as f:
        plistlib.dump(data, f)


def _codesign_macos(app_path):
    """对 .app 进行签名。

    若设置了 MACOS_SIGN_IDENTITY，则用真实 Developer ID 证书签名，并启用
    hardened runtime + entitlements + secure timestamp（公证的硬性前提）；
    否则退回 ad-hoc 签名（仅供本地自测，无法通过公证 / 仍会被门禁拦截）。

    返回 True 表示使用了可公证的真实签名。
    """
    identity = _find_macos_sign_identity()
    if not identity:
        print("  [提示] 未设置 MACOS_SIGN_IDENTITY，使用 ad-hoc 签名（不会通过公证）")
        result = subprocess.run(
            ["codesign", "--force", "--deep", "--sign", "-", str(app_path)],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            print(f"  [警告] ad-hoc 签名失败：{result.stderr.strip() or result.stdout.strip()}")
        return False

    print(f"  正在用 Developer ID 证书签名: {identity}")
    cmd = [
        "codesign", "--force", "--deep",
        "--options", "runtime",          # 启用 hardened runtime（公证必需）
        "--timestamp",                    # 安全时间戳（公证必需）
        "--sign", identity,
        str(app_path),
    ]
    entitlements_path = _get_macos_entitlements_path()
    if entitlements_path.exists():
        # PyInstaller 的 Python 程序在 hardened runtime 下需要 entitlements，否则运行时崩溃
        cmd[1:1] = ["--entitlements", str(entitlements_path)]
    else:
        print(f"  [警告] 找不到 entitlements 文件: {entitlements_path}，公证后可能无法运行")
    result = subprocess.run(cmd, capture_output=True, text=True)
    if result.returncode != 0:
        print(f"  [错误] 证书签名失败：{result.stderr.strip() or result.stdout.strip()}")
        return False
    print("  ✓ 证书签名完成（hardened runtime + timestamp）")
    return True


def _find_macos_sign_identity():
    """返回 Developer ID Application 签名身份；未显式配置时尝试从钥匙串自动识别。"""
    identity = os.environ.get("MACOS_SIGN_IDENTITY", "").strip()
    if identity:
        return identity
    if sys.platform != "darwin":
        return ""

    result = subprocess.run(
        ["security", "find-identity", "-v", "-p", "codesigning"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        return ""

    for line in result.stdout.splitlines():
        if "Developer ID Application:" not in line:
            continue
        match = re.search(r'"([^"]*Developer ID Application:[^"]+)"', line)
        if match:
            identity = match.group(1)
            print(f"  ✓ 自动识别 Developer ID 签名身份: {identity}")
            return identity
    return ""


def _notarytool_auth_args():
    """组装 notarytool 的鉴权参数。优先用 keychain profile，其次用 Apple ID 三件套。"""
    profile = os.environ.get("MACOS_NOTARY_PROFILE", "").strip()
    if profile:
        return ["--keychain-profile", profile]
    apple_id = os.environ.get("MACOS_NOTARY_APPLE_ID", "").strip()
    team_id = os.environ.get("MACOS_NOTARY_TEAM_ID", "").strip()
    password = os.environ.get("MACOS_NOTARY_PASSWORD", "").strip()
    if apple_id and team_id and password:
        return ["--apple-id", apple_id, "--team-id", team_id, "--password", password]
    return None


def _notary_json(text):
    """Best-effort JSON parser for notarytool output."""
    text = (text or "").strip()
    if not text:
        return {}
    try:
        return json.loads(text)
    except json.JSONDecodeError:
        return {}


def _notary_submission_id(*texts):
    """Extract a notary submission UUID from JSON or fallback text."""
    for text in texts:
        data = _notary_json(text)
        submission_id = data.get("id") if isinstance(data, dict) else None
        if submission_id:
            return submission_id
    pattern = re.compile(
        r"\b[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}\b"
    )
    for text in texts:
        match = pattern.search(text or "")
        if match:
            return match.group(0)
    return ""


def _notary_status(*texts):
    """Extract notary status from JSON or plain notarytool text."""
    for text in texts:
        data = _notary_json(text)
        status = data.get("status") if isinstance(data, dict) else None
        if status:
            return str(status)
    for text in texts:
        match = re.search(r"status:\s*([A-Za-z]+)", text or "")
        if match:
            return match.group(1)
    return ""


def _wait_for_notary_submission(submission_id, auth):
    """Wait for an existing notary submission, retrying transient network failures."""
    for attempt in range(1, MACOS_NOTARY_WAIT_RETRIES + 1):
        print(f"  等待 Apple 公证结果（第 {attempt}/{MACOS_NOTARY_WAIT_RETRIES} 次）: {submission_id}")
        wait = subprocess.run(
            [
                "xcrun", "notarytool", "wait", submission_id,
                "--timeout", MACOS_NOTARY_WAIT_TIMEOUT,
                "--output-format", "json",
            ] + auth,
            capture_output=True, text=True,
        )
        wait_stdout = wait.stdout.strip()
        wait_stderr = wait.stderr.strip()
        if wait_stdout:
            print(wait_stdout)
        if wait_stderr:
            print(wait_stderr)

        status = _notary_status(wait_stdout, wait_stderr)
        if wait.returncode == 0 and status == "Accepted":
            return True
        if status in {"Invalid", "Rejected"}:
            print(f"  [错误] Apple 公证拒绝：{status}")
            return False

        if attempt < MACOS_NOTARY_WAIT_RETRIES:
            delay = min(120, 20 * attempt)
            print(f"  [警告] 等待公证结果失败或超时，{delay} 秒后重试。")
            time.sleep(delay)
            continue

        detail = wait_stderr or wait_stdout or f"notarytool wait exited with code {wait.returncode}"
        print(f"  [错误] 公证等待失败：{detail}")
        return False
    return False


def _staple_with_retry(target):
    """Staple notarization ticket, retrying short-lived Apple/CDN failures."""
    for attempt in range(1, 4):
        staple = subprocess.run(
            ["xcrun", "stapler", "staple", str(target)],
            capture_output=True, text=True,
        )
        if staple.returncode == 0:
            print(f"  ✓ 公证 + 钉票据完成: {target.name}")
            return True
        detail = staple.stderr.strip() or staple.stdout.strip()
        if attempt < 3:
            print(f"  [警告] 钉票据失败，准备重试：{detail}")
            time.sleep(10 * attempt)
            continue
        print(f"  [错误] 钉票据失败：{detail}")
        return False
    return False


def _notarize_and_staple(target_path):
    """把 target_path（.app 或 .dmg）提交公证并钉上票据。返回 True 表示成功。"""
    auth = _notarytool_auth_args()
    if auth is None:
        print("  [提示] 未配置公证凭据（MACOS_NOTARY_*），跳过公证")
        return False

    target = Path(target_path)
    submit_path = target
    tmp_zip = None
    # .app 是目录，notarytool 需要先压成 zip 再提交（.dmg 可直接提交）
    if target.is_dir():
        tmp_zip = target.with_suffix(".notarize.zip")
        zip_res = subprocess.run(
            ["ditto", "-c", "-k", "--keepParent", str(target), str(tmp_zip)],
            capture_output=True, text=True,
        )
        if zip_res.returncode != 0:
            print(f"  [错误] 打包待公证 zip 失败：{zip_res.stderr.strip()}")
            return False
        submit_path = tmp_zip

    print(f"  正在提交公证: {target.name}")
    submit = subprocess.run(
        ["xcrun", "notarytool", "submit", str(submit_path), "--output-format", "json"] + auth,
        capture_output=True, text=True,
    )
    if tmp_zip is not None:
        tmp_zip.unlink(missing_ok=True)
    submit_stdout = submit.stdout.strip()
    submit_stderr = submit.stderr.strip()
    if submit_stdout:
        print(submit_stdout)
    if submit_stderr:
        print(submit_stderr)

    submission_id = _notary_submission_id(submit_stdout, submit_stderr)
    if submit.returncode != 0 or not submission_id:
        print(f"  [错误] 公证提交失败：{submit_stderr or submit_stdout}")
        return False

    if not _wait_for_notary_submission(submission_id, auth):
        return False

    # 钉票据：必须钉在原始的 .app / .dmg 上（不是提交用的 zip）
    return _staple_with_retry(target)


def _create_macos_installer_dmg(source_path, dmg_path, volume_name="DocFormatter", app_bundle_name=None):
    """生成带 Applications 入口的 macOS 安装 DMG。"""
    source = Path(source_path)
    dmg = Path(dmg_path)
    staging_dir = BUILD_DIR / f"dmg_staging_{dmg.stem}"

    if staging_dir.exists():
        shutil.rmtree(staging_dir)
    staging_dir.mkdir(parents=True, exist_ok=True)

    try:
        if source.is_dir():
            target_name = app_bundle_name or source.name
            if not target_name.endswith(".app"):
                target_name = f"{target_name}.app"
            target = staging_dir / target_name
            # ditto 更适合复制 .app bundle，可保留 macOS 元数据。
            copy_result = subprocess.run(
                ["ditto", str(source), str(target)],
                capture_output=True, text=True,
            )
            if copy_result.returncode != 0:
                print(f"  [错误] 复制 .app 到 DMG 暂存目录失败：{copy_result.stderr.strip()}")
                return False
        else:
            shutil.copy2(source, staging_dir / source.name)

        applications_link = staging_dir / "Applications"
        if applications_link.exists() or applications_link.is_symlink():
            applications_link.unlink()
        applications_link.symlink_to("/Applications", target_is_directory=True)

        dmg_cmd = [
            "hdiutil", "create",
            "-volname", volume_name,
            "-srcfolder", str(staging_dir),
            "-ov", "-format", "UDZO",
            str(dmg),
        ]
        print("  正在生成 DMG（含 Applications 拖拽入口）...")
        dmg_result = subprocess.run(dmg_cmd, capture_output=True, text=True)
        if dmg_result.returncode != 0:
            print(f"  [错误] DMG 生成失败：{dmg_result.stderr.strip() or dmg_result.stdout.strip()}")
            return False
        return dmg.exists()
    finally:
        shutil.rmtree(staging_dir, ignore_errors=True)


def build_windows():
    """构建 Windows 版本"""
    print("\n" + "=" * 50)
    print("构建 Windows 版本")
    print("=" * 50)
    
    output_name = os.environ.get("WINDOWS_OUTPUT_NAME", "").strip() or "docformat_windows"
    
    # 获取 docx 模板路径
    docx_tpl = _get_docx_templates_path()
    dnd_args = _get_tkinterdnd_pyinstaller_args("windows")
    
    cmd = [
        "pyinstaller",
        "--onefile",           # 单文件
        "--windowed",          # 无控制台窗口
        f"--name={output_name}",
        "--clean",
        # 添加数据文件
        "--add-data=scripts;scripts",
        "--add-data=LICENSE;.",
        "--add-data=LICENSE-HISTORY.md;.",
        # python-docx 模板文件（页眉页脚等必需）
        f"--add-data={docx_tpl};docx/templates" if docx_tpl else "--collect-data=docx",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        "--hidden-import=pythoncom",
        "--hidden-import=pywintypes",
        MAIN_SCRIPT
    ]
    cmd[-1:-1] = dnd_args
    _add_pyinstaller_icon(cmd, WINDOWS_ICON)
    
    print(f"运行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode == 0:
        exe_path = DIST_DIR / f"{output_name}.exe"
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n✓ Windows 版本构建成功!")
            print(f"  文件: {exe_path}")
            print(f"  大小: {size_mb:.1f} MB")
            return True
    
    print("\n✗ Windows 版本构建失败")
    return False


def build_linux():
    """构建 Linux 版本"""
    print("\n" + "=" * 50)
    print("构建 Linux 版本")
    print("=" * 50)
    
    output_name = f"docformat_linux"
    
    # 获取 docx 模板路径
    docx_tpl = _get_docx_templates_path()
    dnd_args = _get_tkinterdnd_pyinstaller_args("linux")
    
    cmd = [
        "pyinstaller",
        "--onefile",
        f"--name={output_name}",
        "--clean",
        "--add-data=scripts:scripts",
        "--add-data=LICENSE:.",
        "--add-data=LICENSE-HISTORY.md:.",
        # python-docx 模板文件
        f"--add-data={docx_tpl}:docx/templates" if docx_tpl else "--collect-data=docx",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        MAIN_SCRIPT
    ]
    cmd[-1:-1] = dnd_args
    
    print(f"运行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode == 0:
        exe_path = DIST_DIR / output_name
        if exe_path.exists():
            size_mb = exe_path.stat().st_size / (1024 * 1024)
            print(f"\n✓ Linux 版本构建成功!")
            print(f"  文件: {exe_path}")
            print(f"  大小: {size_mb:.1f} MB")
            return True
    
    print("\n✗ Linux 版本构建失败")
    return False


def build_macos():
    """构建 macOS 版本"""
    print("\n" + "=" * 50)
    print("构建 macOS 版本")
    print("=" * 50)
    
    machine = platform.machine().lower()
    is_apple_silicon = any(token in machine for token in ("arm64", "aarch64"))
    output_name = os.environ.get("MACOS_OUTPUT_NAME", "").strip()
    if not output_name:
        output_name = "docformat_macos_apple_silicon" if is_apple_silicon else "docformat_macos_intel"
    
    # 获取 docx 模板路径
    docx_tpl = _get_docx_templates_path()
    # 注意：macOS 打包版当前默认关闭拖拽（见 docformat_gui._should_enable_drag_drop），
    # 因此不收集 tkdnd 运行库，避免把用不到的二进制打进 .app。

    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",          # macOS 生成 .app bundle
        f"--name={output_name}",
        "--clean",
        "--osx-bundle-identifier=com.kagurananaga.docformat-gui",
        # macOS 路径分隔符与 Linux 相同
        "--add-data=scripts:scripts",
        "--add-data=LICENSE:.",
        "--add-data=LICENSE-HISTORY.md:.",
        # python-docx 模板文件
        f"--add-data={docx_tpl}:docx/templates" if docx_tpl else "--collect-data=docx",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        MAIN_SCRIPT
    ]
    _add_pyinstaller_icon(cmd, MACOS_ICON)
    
    print(f"运行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode == 0:
        # --windowed 在 macOS 上会生成 .app 目录
        app_path = DIST_DIR / f"{output_name}.app"
        bin_path = DIST_DIR / output_name
        
        if app_path.exists():
            print(f"\n✓ macOS 版本构建成功!")
            print(f"  文件: {app_path}")
            _patch_macos_info_plist(app_path)

            # 1) 签名（有证书则启用 hardened runtime，可公证；否则 ad-hoc）
            real_signed = _codesign_macos(app_path)
            if MACOS_REQUIRE_NOTARIZATION and not real_signed:
                print("  [错误] 当前构建要求 macOS 公证，但没有可用的 Developer ID Application 签名身份")
                return False

            # 2) 默认跳过 .app 单独公证，只公证最终 .dmg，避免每个架构等待两次 Apple。
            #    若需要最严格的本地 .app 票据，可设置 MACOS_NOTARIZE_APP=1。
            if real_signed and MACOS_REQUIRE_NOTARIZATION:
                if MACOS_NOTARIZE_APP:
                    if not _notarize_and_staple(app_path):
                        return False
                else:
                    print("  [提示] 跳过 .app 单独公证，只公证最终 DMG（设置 MACOS_NOTARIZE_APP=1 可恢复）")

            # 3) 生成 DMG（此时里面装的是已签名的 .app）
            dmg_path = DIST_DIR / f"{output_name}.dmg"
            if _create_macos_installer_dmg(app_path, dmg_path, app_bundle_name=MACOS_APP_BUNDLE_NAME):
                size_mb = dmg_path.stat().st_size / (1024 * 1024)
                print(f"  DMG: {dmg_path} ({size_mb:.1f} MB)")
                # 4) DMG 容器本身公证 + 钉票据，用户从网上下载双击才不被拦
                if real_signed and not _notarize_and_staple(dmg_path) and MACOS_REQUIRE_NOTARIZATION:
                    return False
            else:
                return False
            return True
        elif bin_path.exists():
            size_mb = bin_path.stat().st_size / (1024 * 1024)
            print(f"\n✓ macOS 版本构建成功!")
            print(f"  文件: {bin_path}")
            print(f"  大小: {size_mb:.1f} MB")
            return True
    
    print("\n✗ macOS 版本构建失败")
    return False


def create_release_notes():
    """生成发布说明"""
    notes = f"""# {APP_NAME} v{VERSION}

## 下载

- **Windows**: `docformat_windows.exe` - 双击运行
- **Linux (麒麟/UOS)**: `docformat_linux` - 添加执行权限后运行
- **macOS (Intel)**: `docformat_macos_intel.dmg` - 双击挂载后拖入应用程序文件夹
- **macOS (Apple Silicon)**: `docformat_macos_apple_silicon.dmg` - 双击挂载后拖入应用程序文件夹

## 功能

### v1.8.8.4 更新

- ✅ 界面升级为统一的暖米色现代布局，并完善高 DPI 与窗口适配
- ✅ 新增自定义段落类型与本地版式规则、从样例学习样式
- ✅ 增强自动编号、表格、页面、页眉页脚和页码排版能力
- ✅ 支持粘贴文本或 Markdown 生成 Word
- ✅ 修复文件选择后无法开始处理，并显示批量文件数量与清单

- ✅ 一键处理（标点修复 + 格式统一）
- ✅ 格式诊断
- ✅ 标点符号修复
- ✅ 支持 GB/T 公文标准、学术论文、法律文书格式
- ✅ 中文禁则、图片附件保护和可自定义标题中英文字体
- ✅ 免费开源版非商业许可提示

## 系统要求

- Windows 10/11 或
- 麒麟 V10 / 统信 UOS 或其他 Linux 发行版 或
- macOS 12 (Monterey) 或更高版本

## 使用说明

1. 下载对应系统的文件
2. 双击运行（Linux 需先添加执行权限）
3. 选择要处理的 .docx 文件
4. 点击「开始处理」

## 注意事项

- 仅支持 .docx 格式，不支持旧版 .doc
- 处理后的文件会另存为新文件，不会覆盖原文件
- macOS 和 Linux 版本不支持 .doc/.wps 格式转换
"""
    
    release_file = DIST_DIR / "RELEASE_NOTES.md"
    release_file.write_text(notes, encoding="utf-8")
    print(f"\n✓ 发布说明已生成: {release_file}")


def main():
    """主函数"""
    print(f"""
╔══════════════════════════════════════════╗
║     {APP_NAME} 打包工具            ║
║     版本: {VERSION}                          ║
╚══════════════════════════════════════════╝
    """)
    
    # 检查依赖
    if not check_pyinstaller():
        sys.exit(1)
    
    # 检查主脚本
    if not Path(MAIN_SCRIPT).exists():
        print(f"✗ 找不到主脚本: {MAIN_SCRIPT}")
        sys.exit(1)
    
    # 解析参数
    target = sys.argv[1] if len(sys.argv) > 1 else "all"
    
    if target not in ["windows", "linux", "macos", "all", "clean"]:
        print(f"用法: python {sys.argv[0]} [windows|linux|macos|all|clean]")
        sys.exit(1)
    
    # 清理
    if target == "clean":
        clean()
        return
    
    clean()
    
    # 构建
    success = True
    
    if target in ["windows", "all"]:
        if sys.platform == "win32":
            success = build_windows() and success
        else:
            print("\n⚠ 跳过 Windows 构建（需要在 Windows 系统上执行）")
    
    if target in ["linux", "all"]:
        if sys.platform.startswith("linux"):
            success = build_linux() and success
        else:
            print("\n⚠ 跳过 Linux 构建（需要在 Linux 系统上执行）")
    
    if target in ["macos", "all"]:
        if sys.platform == "darwin":
            success = build_macos() and success
        else:
            print("\n⚠ 跳过 macOS 构建（需要在 macOS 系统上执行）")
    
    # 生成发布说明
    if DIST_DIR.exists():
        create_release_notes()
    
    # 总结
    print("\n" + "=" * 50)
    if success:
        print("✓ 构建完成!")
        print(f"\n输出目录: {DIST_DIR.absolute()}")
        if DIST_DIR.exists():
            print("\n生成的文件:")
            for f in DIST_DIR.iterdir():
                if f.is_file():
                    size = f.stat().st_size / (1024 * 1024)
                    print(f"  - {f.name} ({size:.1f} MB)")
    else:
        print("✗ 构建过程中出现错误")
        sys.exit(1)


if __name__ == "__main__":
    main()
