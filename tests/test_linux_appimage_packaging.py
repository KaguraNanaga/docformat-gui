"""Release workflow safeguards for portable Linux artifacts."""

from pathlib import Path


PROJECT_ROOT = Path(__file__).parent.parent


def test_linux_workflow_builds_appimages_with_the_community_notice_asset():
    workflow = (PROJECT_ROOT / ".github" / "workflows" / "build.yml").read_text(encoding="utf-8")

    assert workflow.count("image: almalinux:8") == 2
    assert workflow.count('assets/xianyu_qr.png:assets') >= 2
    assert "packaging/appimage/build-appimage.sh" in workflow
    assert "docformat_linux_amd64.AppImage" in workflow
    assert "docformat_linux_aarch64.AppImage" in workflow


def test_appimage_builder_uses_a_pinned_tool_release_and_cleans_only_temp_dirs():
    script = (PROJECT_ROOT / "packaging" / "appimage" / "build-appimage.sh").read_text(encoding="utf-8")

    assert "APPIMAGETOOL_VERSION=1.9.1" in script
    assert "releases/download/${APPIMAGETOOL_VERSION}" in script
    assert "mktemp -d" in script
    assert "trap cleanup EXIT" in script
    assert "docformat.desktop" in script
