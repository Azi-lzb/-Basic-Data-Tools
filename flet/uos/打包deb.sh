#!/usr/bin/env bash
# 把 flet build linux 产出的发行目录打包成 deb。
# 用法：
#   ./打包deb.sh --build      # 先 flet build linux，再打包（推荐）
#   ./打包deb.sh              # 仅用已有产物打包
#   VERSION=1.1.0 ./打包deb.sh
#
# 说明：flet build 产物是一套自包含应用（Flutter 运行时 + libpython + pyc），
# 打进 /opt/base-audit-uos；业务文件（模板/历史审核说明/data）不进包。
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_BIN="${PYTHON_BIN:-$ROOT_DIR/.venv/bin/python}"
VERSION="${VERSION:-1.1.0}"
ARCH="${ARCH:-amd64}"
PKG="base-audit-uos"
BUILD_OUT="${BUILD_OUT:-/tmp/flet_build_test}"
WORK="${ROOT_DIR}/deb/${PKG}_${VERSION}_${ARCH}"

# flet build 入口命令（venv 无 bin/flet 脚本，用 python -m flet.cli）
flet_build() {
  "$PYTHON_BIN" -m flet.cli build linux "$ROOT_DIR" --yes \
    --output "$BUILD_OUT" \
    --project "$PKG" \
    --product "基础数据审核工具（统信版）" \
    --artifact "基础数据审核工具-统信" \
    --skip-flutter-doctor
}

if [ "${1:-}" = "--build" ] || [ ! -d "$BUILD_OUT" ]; then
  echo "==> flet build linux ..."
  rm -rf "$BUILD_OUT"
  flet_build
fi

# 定位产物可执行入口
APP_SRC=""
for cand in "$BUILD_OUT/基础数据审核工具-统信" \
            "$BUILD_OUT/linux/基础数据审核工具-统信" \
            "$BUILD_OUT/app/基础数据审核工具-统信"; do
  [ -f "$cand" ] && APP_SRC="$cand" && break
done
if [ -z "$APP_SRC" ]; then
  echo "未找到 flet 可执行入口。产物目录:" >&2; ls "$BUILD_OUT" >&2
  exit 1
fi
echo "==> 应用入口: $APP_SRC"

# 组装安装树到 /opt/base-audit-uos
rm -rf "$WORK"
mkdir -p "$WORK/opt/base-audit-uos" "$WORK/usr/bin" "$WORK/usr/share/applications" "$WORK/DEBIAN"
# 拷整个 flet 产物（含 lib/data/flutter_assets/python 运行时与可执行）
cp -a "$BUILD_OUT/." "$WORK/opt/base-audit-uos/"
# 清理分发时不应携带的开发/业务路径数据：data 目录只留空的默认配置，
# 用户设置与模板索引在目标机首次使用时重建（避免泄漏开发机绝对路径）。
rm -rf "$WORK/opt/base-audit-uos/app/.venv" "$WORK/opt/base-audit-uos/.venv" "$WORK/opt/base-audit-uos/app/tests"
if [ -d "$WORK/opt/base-audit-uos/app/data" ]; then
  rm -f "$WORK/opt/base-audit-uos/app/data/用户设置.json" "$WORK/opt/base-audit-uos/app/data/模板索引.json"
fi
if [ -d "$WORK/opt/base-audit-uos/data" ]; then
  rm -f "$WORK/opt/base-audit-uos/data/用户设置.json" "$WORK/opt/base-audit-uos/data/模板索引.json"
fi

# 启动包装：定位可执行（不同 flet 版本产物位置略有差异）
# 虚拟显卡(VMware/QXL/VirtIO)或无 GPU 时用软件渲染，避免 Flutter 硬件 GL
# 段错误/卡死；真实 Intel/AMD/NVIDIA 显卡走硬件加速。
cat > "$WORK/opt/base-audit-uos/启动基础数据审核工具.sh" <<'SH'
#!/usr/bin/env bash
set -euo pipefail
APP_DIR="$(cd "$(dirname "$0")" && pwd)"
LOG_DIR="${XDG_STATE_HOME:-$HOME/.local/state}/base-audit-uos"
mkdir -p "$LOG_DIR"
export NO_AT_BRIDGE="${NO_AT_BRIDGE:-1}"
export GTK_USE_PORTAL="${GTK_USE_PORTAL:-0}"
export GDK_BACKEND="${GDK_BACKEND:-x11}"

BIN=""
for cand in "$APP_DIR/基础数据审核工具-统信" \
            "$APP_DIR/linux/基础数据审核工具-统信" \
            "$APP_DIR/app/基础数据审核工具-统信"; do
  [ -f "$cand" ] && BIN="$cand" && break
done
if [ -z "$BIN" ]; then echo "未找到应用入口: $APP_DIR" >&2; exit 1; fi

# 判断是否为虚拟/无硬件加速显卡。常见虚拟 vendor：VMware=0x15ad、
# QXL/RedHat=0x1af4、VirtualBox=0x80ee、Microsoft Hyper-V=0x1414。
_VENDOR="$(cat /sys/class/drm/card0/device/vendor 2>/dev/null || echo '')"
case "$_VENDOR" in
  0x15ad|0x1af4|0x80ee|0x1414)
    export LIBGL_ALWAYS_SOFTWARE=1 ;;
esac
# 环境变量可强制覆盖（真机 GL 异常时手动开软件渲染）
if [ "${BASE_AUDIT_FORCE_SOFTWARE:-0}" = "1" ]; then
  export LIBGL_ALWAYS_SOFTWARE=1
fi

exec "$BIN" "$@" 2>>"$LOG_DIR/runtime.log"
SH
chmod +x "$WORK/opt/base-audit-uos/启动基础数据审核工具.sh"

cat > "$WORK/usr/bin/base-audit-uos" <<'SH'
#!/bin/sh
exec /opt/base-audit-uos/启动基础数据审核工具.sh "$@"
SH
chmod +x "$WORK/usr/bin/base-audit-uos"

ICON_LINE=""
ICON_CAND=""
for cand in "$WORK/opt/base-audit-uos/icons/base-audit-uos.png" \
            "$WORK/opt/base-audit-uos/assets/icon.png" \
            "$WORK/opt/base-audit-uos/app/assets/icon.png"; do
  [ -f "$cand" ] && ICON_CAND="$cand" && break
done
if [ -n "$ICON_CAND" ]; then
  ICON_LINE="Icon=$ICON_CAND"
fi

cat > "$WORK/usr/share/applications/base-audit-uos.desktop" <<DE
[Desktop Entry]
Type=Application
Name=基础数据审核工具（统信版）
Name[zh_CN]=基础数据审核工具
Comment=基础数据审核程序（统信版）
Exec=/usr/bin/base-audit-uos
${ICON_LINE}
Terminal=false
Categories=Office;Utility;
StartupWMClass=base_audit
DE

cat > "$WORK/DEBIAN/control" <<CTL
Package: $PKG
Version: $VERSION
Section: utils
Priority: optional
Architecture: $ARCH
Installed-Size: $(du -sk "$WORK/opt" | cut -f1)
Maintainer: 基础数据审核团队
Depends: libgtk-3-0, libreoffice-calc, libreoffice-core
Description: 基础数据审核工具（统信版）
 基于 Flet 的桌面审核程序，支持流程配置、公式校验与条件格式提取、
 历史审核说明维护与多模板联合汇总。终端用户无需安装 Python。
CTL

cat > "$WORK/DEBIAN/postinst" <<'SH'
#!/bin/sh
set -e
chmod +x /opt/base-audit-uos/启动基础数据审核工具.sh /usr/bin/base-audit-uos 2>/dev/null || true
update-desktop-database /usr/share/applications 2>/dev/null || true
exit 0
SH
chmod +x "$WORK/DEBIAN/postinst"

echo "==> 打包 ..."
dpkg-deb --build --root-owner-group "$WORK" "${ROOT_DIR}/deb/${PKG}_${VERSION}_${ARCH}.deb"
DEB_FINAL="${ROOT_DIR}/deb/${PKG}_${VERSION}_${ARCH}.deb"
echo "==> 完成: $DEB_FINAL"
echo
echo "=================================================================="
echo "  安装依赖提示（重要，内网离线机器务必看）"
echo "=================================================================="
echo "  本包 Depends 需要以下系统包（均需目标机已安装，或随包离线提供）："
DEPS_LINE="$(dpkg-deb -f "$DEB_FINAL" Depends)"
echo "    $DEPS_LINE"
echo
echo "  内网离线安装请执行："
echo "    ./收集依赖.sh $DEB_FINAL   # 在本机能连源处生成 deps/(全部依赖 .deb)"
echo "    然后把 deb/ 整个目录拷到内网机，按 deb/内网安装说明.txt 安装。"
echo "  在线机器直接："
echo "    sudo apt install ./$PKG"
echo "=================================================================="
