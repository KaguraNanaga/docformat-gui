#!/usr/bin/env python3
"""
打包脚本 - 生成 Windows/Linux/macOS 可执行文件
用法：python build.py [windows|linux|macos|all|clean]
"""

import os
import sys
import shutil
import subprocess
from pathlib import Path

# 配置
APP_NAME = "公文格式处理工具"
APP_NAME_EN = "DocFormatter"
VERSION = "1.0.0"
MAIN_SCRIPT = "docformat_gui.py"

# 输出目录
DIST_DIR = Path("dist")
BUILD_DIR = Path("build")


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


def build_windows():
    """构建 Windows 版本"""
    print("\n" + "=" * 50)
    print("构建 Windows 版本")
    print("=" * 50)
    
    output_name = f"docformat_windows"
    
    cmd = [
        "pyinstaller",
        "--onefile",           # 单文件
        "--windowed",          # 无控制台窗口
        f"--name={output_name}",
        "--clean",
        # 添加数据文件
        "--add-data=scripts;scripts",
        # 隐藏导入
        "--hidden-import=docx",
        "--hidden-import=lxml",
        MAIN_SCRIPT
    ]
    
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
    
    cmd = [
        "pyinstaller",
        "--onefile",
        f"--name={output_name}",
        "--clean",
        # Linux 路径分隔符
        "--add-data=scripts:scripts",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        MAIN_SCRIPT
    ]
    
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
    
    output_name = "docformat_macos"
    
    cmd = [
        "pyinstaller",
        "--onefile",
        "--windowed",          # macOS 生成 .app bundle
        f"--name={output_name}",
        "--clean",
        # macOS 路径分隔符与 Linux 相同
        "--add-data=scripts:scripts",
        "--hidden-import=docx",
        "--hidden-import=lxml",
        MAIN_SCRIPT
    ]
    
    print(f"运行: {' '.join(cmd)}")
    result = subprocess.run(cmd, capture_output=False)
    
    if result.returncode == 0:
        # --windowed 在 macOS 上会生成 .app 目录
        app_path = DIST_DIR / f"{output_name}.app"
        bin_path = DIST_DIR / output_name
        
        if app_path.exists():
            print(f"\n✓ macOS 版本构建成功!")
            print(f"  文件: {app_path}")
            # 生成 DMG
            dmg_path = DIST_DIR / f"{output_name}.dmg"
            dmg_cmd = [
                "hdiutil", "create",
                "-volname", "DocFormatter",
                "-srcfolder", str(app_path),
                "-ov", "-format", "UDZO",
                str(dmg_path)
            ]
            print(f"  正在生成 DMG...")
            dmg_result = subprocess.run(dmg_cmd, capture_output=True)
            if dmg_result.returncode == 0 and dmg_path.exists():
                size_mb = dmg_path.stat().st_size / (1024 * 1024)
                print(f"  DMG: {dmg_path} ({size_mb:.1f} MB)")
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
- **macOS**: `docformat_macos.dmg` - 双击挂载后拖入应用程序文件夹

## 功能

- ✅ 智能一键处理（标点修复 + 格式统一）
- ✅ 格式诊断
- ✅ 标点符号修复
- ✅ 支持 GB/T 公文标准、学术论文、法律文书格式

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
