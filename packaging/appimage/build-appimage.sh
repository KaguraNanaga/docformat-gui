#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"

BINARY="${1:-dist/docformat_linux}"
ICON="${2:-$REPO_ROOT/assets/icon.png}"
OUTPUT_NAME="${3:-docformat_linux}"

# ── 验证输入 ──
if [ ! -f "$BINARY" ]; then
  echo "✗ 找不到 PyInstaller 产物: $BINARY"
  exit 1
fi

OUTPUT="$REPO_ROOT/dist/${OUTPUT_NAME}.AppImage"
APPDIR="$REPO_ROOT/dist/AppDir"

echo "构建 AppImage"
echo "  二进制: $BINARY"
echo "  图标: $ICON"
echo "  输出: $OUTPUT"

# ── 准备 AppDir ──
rm -rf "$APPDIR"
mkdir -p "$APPDIR/usr/bin"
mkdir -p "$APPDIR/usr/share/icons/hicolor/256x256/apps"

cp "$BINARY" "$APPDIR/usr/bin/docformat_linux"
chmod 755 "$APPDIR/usr/bin/docformat_linux"

[ -f "$ICON" ] && cp "$ICON" "$APPDIR/usr/share/icons/hicolor/256x256/apps/docformat.png"
[ -f "$ICON" ] && cp "$ICON" "$APPDIR/docformat.png"

# ── .desktop ──
cat > "$APPDIR/docformat.desktop" <<'EOF'
[Desktop Entry]
Name=公文格式处理工具
Comment=GB/T 9704-2012 公文格式一键修复
Exec=docformat_linux
Icon=docformat
Type=Application
Categories=Office;WordProcessor;
Terminal=false
EOF

# ── AppRun ──
cat > "$APPDIR/AppRun" <<'APPRUN'
#!/bin/sh
SELF="$(readlink -f "$0")"
HERE="${SELF%/*}"
export PATH="$HERE/usr/bin:$PATH"
export LD_LIBRARY_PATH="$HERE/usr/lib:$HERE/usr/lib/x86_64-linux-gnu:$HERE/usr/lib/aarch64-linux-gnu:${LD_LIBRARY_PATH:-}"
exec "$HERE/usr/bin/docformat_linux" "$@"
APPRUN
chmod 755 "$APPDIR/AppRun"

# ── 下载 appimagetool ──
if ! command -v appimagetool >/dev/null 2>&1; then
  ARCH_VAL="$(uname -m)"
  echo "下载 appimagetool ($ARCH_VAL)..."
  wget -q -O appimagetool \
    "https://github.com/AppImage/appimagetool/releases/download/continuous/appimagetool-${ARCH_VAL}.AppImage"
  chmod +x appimagetool
  # CI 容器无 FUSE，直接解压执行
  ./appimagetool --appimage-extract >/dev/null 2>&1
  APPIMAGETOOL="./squashfs-root/AppRun"
else
  APPIMAGETOOL="appimagetool"
fi

# ── 打包 ──
echo "生成 AppImage..."
ARCH="$(uname -m)" $APPIMAGETOOL "$APPDIR" "$OUTPUT"

rm -rf "$APPDIR"
[ -f appimagetool ] && rm -f appimagetool
[ -d squashfs-root ] && rm -rf squashfs-root

if [ -f "$OUTPUT" ]; then
  SIZE_MB=$(du -m "$OUTPUT" | cut -f1)
  echo "✓ AppImage 构建成功: $OUTPUT (${SIZE_MB} MB)"
else
  echo "✗ AppImage 构建失败"
  exit 1
fi
