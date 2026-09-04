#!/bin/sh
# 打包 Flask/UOS 版：组装源码发行包 + zip + SHA256。
# 产物：dist/基础数据审核工具_Flask_UOS_<日期>.zip（及 .sha256）
#
# 包内布局（app.py 同时兼容开发布局与这种打包布局）：
#   基础数据审核工具_Flask_UOS_<日期>/
#   ├─ app.py, run.py, requirements.txt, README.md, 启动Flask-UOS.sh
#   └─ shared/src + shared/frontend   ← 与 Windows 版共用的业务核心
#
# 注意：不打包 历史审核配置.xlsx（首次运行自动按标准表头创建，绝不覆盖用户历史）
#      和 shared/tests（测试属于开发基线，不进发行包）。

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SHARED="$SCRIPT_DIR/../shared"
DIST="$SCRIPT_DIR/dist"
VERSION="$(date +%Y%m%d)"
PKG_NAME="基础数据审核工具_Flask_UOS_$VERSION"
PKG="$DIST/$PKG_NAME"

[ -d "$SHARED/src/base_audit" ] || { echo "[错误] 未找到共享核心：$SHARED"; exit 1; }
[ -f "$SHARED/frontend/web/index.html" ] || { echo "[错误] 未找到共享前端 index.html"; exit 1; }

rm -rf "$PKG"
mkdir -p "$PKG"

# 外壳文件
cp "$SCRIPT_DIR/app.py" "$PKG/"
cp "$SCRIPT_DIR/run.py" "$PKG/"
cp "$SCRIPT_DIR/requirements.txt" "$PKG/"
cp "$SCRIPT_DIR/README.md" "$PKG/"
cp "$SCRIPT_DIR/启动Flask-UOS.sh" "$PKG/"
chmod +x "$PKG/启动Flask-UOS.sh"

# 共享核心：后端 + 前端（不含 tests / __pycache__）
mkdir -p "$PKG/shared"
cp -r "$SHARED/src" "$PKG/shared/src"
cp -r "$SHARED/frontend" "$PKG/shared/frontend"
find "$PKG" -name "__pycache__" -type d -exec rm -rf {} + 2>/dev/null || true
find "$PKG" -name "*.pyc" -delete 2>/dev/null || true

# 压缩包 + SHA256
cd "$DIST"
if command -v zip >/dev/null 2>&1; then
    zip -r -q "$PKG_NAME.zip" "$PKG_NAME"
else
    tar -czf "$PKG_NAME.tar.gz" "$PKG_NAME"
    PKG_NAME="$PKG_NAME"  # 显示用；下面按实际存在的文件计算哈希
fi
ARTIFACT="$(ls -t "$PKG_NAME".zip "$PKG_NAME".tar.gz 2>/dev/null | head -1)"
if command -v sha256sum >/dev/null 2>&1; then
    sha256sum "$ARTIFACT" > "$ARTIFACT.sha256"
else
    python3 -c "import hashlib,sys;print(hashlib.sha256(open(sys.argv[1],'rb').read()).hexdigest()+'  '+sys.argv[1],end='')" "$ARTIFACT" > "$ARTIFACT.sha256"
fi

SIZE="$(du -sh "$ARTIFACT" | cut -f1)"
echo ""
echo "打包完成：dist/$ARTIFACT（$SIZE）"
cat "$ARTIFACT.sha256"
echo ""
echo "目标机器要求：python3 + pip（pip3 install -r requirements.txt）、LibreOffice Calc（计算引擎适配完成后）。"
