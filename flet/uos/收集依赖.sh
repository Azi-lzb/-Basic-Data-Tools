#!/usr/bin/env bash
# 为内网离线部署收集 base-audit-uos deb 的全部运行依赖(.deb)。
# 用法：
#   ./收集依赖.sh               # 收集 deb/ 下最新 deb 的依赖到 deb/deps/
#   ./收集依赖.sh <deb路径>
#
# 前提：本机能访问 UOS 软件源，且软件源与目标内网机一致(V20/V25 源不同，
#       需在对应版本的联网机器上收集)。
# 产物：deb/deps/ = 全部直接+传递依赖的 .deb + 主 deb + 依赖清单.txt
set -euo pipefail

ROOT_DIR="$(cd "$(dirname "$0")" && pwd)"
DEB="${1:-$(ls -t "$ROOT_DIR"/deb/base-audit-uos_*.deb 2>/dev/null | head -1)}"
if [ -z "$DEB" ] || [ ! -f "$DEB" ]; then
  echo "未找到 deb，请先 ./打包deb.sh 或传入 deb 路径" >&2
  exit 1
fi
DEB="$(realpath "$DEB")"
OUT="$ROOT_DIR/deb/deps"
rm -rf "$OUT"; mkdir -p "$OUT"
echo "==> 收集依赖 for: $(basename "$DEB")"

# 递归解析 Depends 闭包，只保留真实包名（剔除 <库文件> 依赖、:any）
CLOSURE=/tmp/base-audit-deps-closure.txt
: > "$CLOSURE"
resolve() {
  local pkg="$1" deps d
  if grep -qxF "$pkg" "$CLOSURE" 2>/dev/null; then return 0; fi
  echo "$pkg" >> "$CLOSURE"
  deps="$(apt-cache depends --no-recommends --no-suggests --no-conflicts \
          --no-breaks --no-replaces --no-enhances "$pkg" 2>/dev/null \
          | grep -E '^  (依赖|Depends):' \
          | sed -E 's/^  [^:]+: *([^ <(]+).*/\1/' \
          | grep -vE '<|:any|\.so($|\.)' || true)"
  for d in $deps; do
    [ -z "$d" ] && continue
    [ "$d" = "$pkg" ] && continue
    resolve "$d"
  done
  return 0
}

# 主 deb 的直接依赖作为起点
DIRECT="$(dpkg-deb -f "$DEB" Depends | tr ',' '\n' \
          | sed -E 's/^ *//; s/ \([^)]*\)//g; s/:any//' \
          | grep -E '^[a-z0-9.+_-]+$' || true)"
echo "直接依赖: $(echo $DIRECT | tr '\n' ' ')"
for p in $DIRECT; do resolve "$p"; done

TOTAL="$(wc -l < "$CLOSURE")"
echo "==> 依赖闭包共 $TOTAL 个包，开始下载 ..."

# 源里没有候选的包先剔除（记录到清单说明）
VALID="$OUT/.valid"; : > "$VALID"
MISSING=""
while read -r p; do
  [ -z "$p" ] && continue
  if apt-cache policy "$p" 2>/dev/null | grep -q "候选"; then
    echo "$p" >> "$VALID"
  else
    MISSING="$MISSING $p"
  fi
done < "$CLOSURE"

# 一次性批量下载（apt 对缺包会整体失败，故先剔除无候选的）
cd "$OUT"
downloaded=0
mapfile -t pkgs < "$VALID"
for start in $(seq 0 30 $((${#pkgs[@]} - 1))); do
  chunk=("${pkgs[@]:start:30}")
  if apt-get download "${chunk[@]}" >>"$OUT/.dl.log" 2>&1; then
    downloaded=$((downloaded + ${#chunk[@]}))
  else
    # 分批内个别失败：逐一下载补齐
    for p in "${chunk[@]}"; do
      if ! ls "${p}"_*.deb >/dev/null 2>&1 && apt-get download "$p" >>"$OUT/.dl.log" 2>&1; then
        downloaded=$((downloaded + 1))
      fi
    done
  fi
done
rm -f "$OUT/.valid" "$OUT/.dl.log"
cp "$DEB" "$OUT/"

{
  echo "依赖收集时间: $(date '+%Y-%m-%d %H:%M')"
  echo "目标 deb: $(basename "$DEB")"
  echo "主 deb 直接依赖: $(dpkg-deb -f "$DEB" Depends)"
  echo "成功下载依赖 .deb 数: $(ls "$OUT"/*.deb 2>/dev/null | wc -l)"
  [ -n "${MISSING:-}" ] && echo "源中无候选的包(需人工找): $MISSING"
  echo
  echo "内网离线安装（目标机与收集源同版本 UOS）："
  echo "  把整个 deb/ 目录拷到内网机，然后："
  echo "  cd deps && sudo dpkg -i *.deb"
  echo "  若偶发顺序缺依赖再执行一次 sudo dpkg -i *.deb"
  echo "  启动：开始菜单「基础数据审核工具（统信版）」"
} > "$OUT/依赖清单.txt"

echo "==> 完成: $OUT/  依赖清单见 依赖清单.txt"
