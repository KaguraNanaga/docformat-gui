#!/usr/bin/env bash
set -euo pipefail

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
REPO_ROOT="$(cd "$SCRIPT_DIR/../.." && pwd)"
APPIMAGETOOL_VERSION=1.9.1

BINARY="${1:-$REPO_ROOT/dist/docformat_linux}"
ICON="${2:-$REPO_ROOT/assets/icon.png}"
OUTPUT_NAME="${3:-docformat_linux}"
OUTPUT="$REPO_ROOT/dist/${OUTPUT_NAME}.AppImage"
APPDIR="$(mktemp -d "$REPO_ROOT/dist/appdir.XXXXXX")"
TOOL_DIR="$(mktemp -d)"

cleanup() {
  rm -rf "$APPDIR" "$TOOL_DIR"
}
trap cleanup EXIT

if [ ! -f "$BINARY" ]; then
  echo "✗ 找不到 PyInstaller 产物: $BINARY"
  exit 1
fi
if [ ! -f "$ICON" ]; then
  echo "✗ 找不到应用图标: $ICON"
  exit 1
fi

ARCH_VALUE="$(uname -m)"
case "$ARCH_VALUE" in
  x86_64|aarch64) ;;
  *)
    echo "✗ 不支持的 AppImage 架构: $ARCH_VALUE"
    exit 1
    ;;
esac

mkdir -p "$APPDIR/usr/bin" "$APPDIR/usr/share/icons/hicolor/256x256/apps"
install -m 755 "$BINARY" "$APPDIR/usr/bin/docformat_linux"
install -m 644 "$ICON" "$APPDIR/usr/share/icons/hicolor/256x256/apps/docformat.png"
install -m 644 "$ICON" "$APPDIR/docformat.png"
install -m 644 "$SCRIPT_DIR/docformat.desktop" "$APPDIR/docformat.desktop"

cat > "$APPDIR/AppRun" <<'APPRUN'
#!/bin/sh
HERE="$(dirname "$(readlink -f "$0")")"
export PATH="$HERE/usr/bin:$PATH"
exec "$HERE/usr/bin/docformat_linux" "$@"
APPRUN
chmod 755 "$APPDIR/AppRun"

TOOL="$TOOL_DIR/appimagetool"
TOOL_URL="https://github.com/AppImage/appimagetool/releases/download/${APPIMAGETOOL_VERSION}/appimagetool-${ARCH_VALUE}.AppImage"
echo "下载 appimagetool ${APPIMAGETOOL_VERSION} ($ARCH_VALUE)..."
wget -q -O "$TOOL" "$TOOL_URL"
chmod 755 "$TOOL"

# GitHub Actions container has no FUSE. Extract the tool before invoking it.
(
  cd "$TOOL_DIR"
  "$TOOL" --appimage-extract >/dev/null
)

rm -f "$OUTPUT"
echo "生成 AppImage: $OUTPUT"
ARCH="$ARCH_VALUE" "$TOOL_DIR/squashfs-root/AppRun" "$APPDIR" "$OUTPUT"

test -f "$OUTPUT"
echo "✓ AppImage 构建成功: $OUTPUT"
