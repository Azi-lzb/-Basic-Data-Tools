"""Flask 版基础数据审核工具（统一外壳：Windows + UOS/麒麟 共用一份）。

复用 unified/core 的唯一业务核心（base_audit 后端 + frontend 前端）：
- 前端 index.html 不做修改，由本模块在 ``<head>`` 注入一个 ``pywebview.api``
  兼容层，把 ``pywebview.api.方法名(...)`` 代理为 ``POST /api/方法名``。
- 后端直接继承 ``base_audit.web_app.WebApi``，只把 pywebview 的文件选择
  对话框替换为 tkinter 本机对话框，窗口控制按钮改为无操作（浏览器窗口由
  用户自己管理）。
- 平台差异只有“打开文件/目录”：Windows 用 os.startfile，UOS/麒麟 用
  xdg-open；计算引擎选项由 core/base_audit/engines 按操作系统提供
  （Windows：Excel/WPS COM；UOS：LibreOffice Calc）。
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
from pathlib import Path
from typing import Any

from flask import Flask, jsonify, request

# 冻结态（PyInstaller）__file__ 指向临时解包目录，须按 exe 位置定位 core。
ROOT = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
# 打包布局优先：exe 同级的 core/；开发布局：shell-flask 平级的 ../core。
CORE = ROOT / "core"
if not (CORE / "src" / "base_audit").is_dir():
    CORE = ROOT.parent / "core"
if str(CORE / "src") not in sys.path:
    sys.path.insert(0, str(CORE / "src"))

from base_audit.web_app import WebApi  # noqa: E402


PROJECT_ROOT = CORE
FRONTEND_DIR = CORE / "frontend"

# 在服务线程里弹 tkinter 对话框；同一时刻只允许一个对话框，避免并发请求
# 同时创建多个 Tk 根窗口。
_DIALOG_LOCK = threading.Lock()


def _tk_dialog(open_dialog: Any) -> Any:
    """Run one tkinter dialog call in a dedicated Tk root and return its result."""
    import tkinter as tk
    from tkinter import filedialog

    with _DIALOG_LOCK:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        try:
            return open_dialog(root, filedialog)
        finally:
            root.destroy()


def _open_with_default_app(path: Path) -> None:
    """用系统默认程序打开文件/目录：Windows 用 os.startfile，其余用 xdg-open。"""
    if sys.platform == "win32":
        os.startfile(str(path))  # noqa: S606 - 本机默认程序打开
    else:
        subprocess.Popen(["xdg-open", str(path)])


class FlaskApi(WebApi):
    """WebApi 的 Flask 版：文件对话框用 tkinter，按平台打开文件，窗口控制为无操作。"""

    def _open_path(self, path: Path, what: str) -> None:
        if not path.exists():
            self.state["status"] = f"{what}不存在"
            self._log(f"未找到{what}：{path}")
            return
        try:
            _open_with_default_app(path)
        except Exception as exc:
            self.state["status"] = f"无法打开{what}"
            self._log(f"打开{what}失败：{exc}")
            return
        self._log(f"已打开{what}：{path}")

    def open_history_explanation(self) -> dict[str, Any]:
        self._open_path(self.history_path, "历史审核配置")
        return self.state

    def open_user_guide(self) -> dict[str, Any]:
        self._open_path(self.project_root / "基础数据审核工具使用说明.docx", "使用说明")
        return self.state

    def open_path(self, path: str) -> dict[str, Any]:
        self._open_path(Path(path), "路径")
        return self.state

    def choose_folder(self, field: str) -> str:
        current = self.state.get(field) or str(self.project_root)
        value = _tk_dialog(
            lambda root, fd: fd.askdirectory(parent=root, initialdir=str(current))
        ) or ""
        if value:
            previous = self.state.get(field, "")
            self.state[field] = value
            if field == "templateDir" and value != previous:
                self.state["templateManual"] = False
            if field in {"input", "templateDir"}:
                if field == "input" and not self.state["outputPinned"]:
                    self.state["output"] = str(Path(value) / "执行结果")
                    self.state["outputAuto"] = True
                # 与 pywebview 版一致：仅选择源数据目录触发一次模板推荐。
                self._recognize(allow_template_auto=(field == "input"))
            elif field == "output":
                self.state["outputAuto"] = False
                self.state["outputPinned"] = True
        return value

    def choose_file(self, field: str) -> str:
        current = self.state.get(field) or str(self.project_root)
        start_dir = str(Path(current).parent) if Path(current).is_file() else str(current)
        value = _tk_dialog(
            lambda root, fd: fd.askopenfilename(
                parent=root,
                initialdir=start_dir,
                filetypes=[("Excel 文件", "*.xlsx *.xls")],
            )
        ) or ""
        if value:
            self.state[field] = value
            if field == "template":
                self.state["templateManual"] = True
        return value

    def choose_history_config(self) -> str:
        if self.state["busy"]:
            return str(self.history_path)
        current = self.history_path if self.history_path.is_file() else self.project_root
        start_dir = str(current.parent) if Path(current).is_file() else str(current)
        value = _tk_dialog(
            lambda root, fd: fd.askopenfilename(
                parent=root,
                initialdir=start_dir,
                filetypes=[("Excel 文件", "*.xlsx *.xls")],
            )
        ) or ""
        if value:
            self.history_path = Path(value)
            self.state["historyConfig"] = value
            self.settings.history_config = value
            self.settings_store.save(self.settings)
            self._log(f"已选择历史审核配置：{value}")
        return value

    def create_template_merge(self) -> bool:
        if self.state["busy"]:
            return False
        excel_types = [("Excel 文件", "*.xlsx *.xlsm *.xls")]
        base_dir = self.state.get("templateDir") or str(self.project_root)
        self.state["status"] = "第 1 步：请选择基准模板"
        self._log(
            "第 1 步/2：请选择基准模板（它会另存为联合模板；"
            "应包含集中系统数据、参照表等公共依赖工作表）"
        )
        base = _tk_dialog(
            lambda root, fd: fd.askopenfilename(
                parent=root, initialdir=str(base_dir), filetypes=excel_types
            )
        )
        if not base:
            self._log("已取消制作联合模板：未选择底稿模板")
            return False
        base_template = Path(base).resolve()
        self.state["status"] = "第 2 步：请选择要并入的模板"
        self._log(
            f"已选择基准模板：{base_template.name}。第 2 步/2："
            "请批量选择要并入的其他模板（不要重复选择基准模板）"
        )
        picked = _tk_dialog(
            lambda root, fd: fd.askopenfilenames(
                parent=root, initialdir=str(base_template.parent), filetypes=excel_types
            )
        )
        if not picked:
            self._log("已取消制作联合模板：未选择待复制工作簿")
            return False
        source_templates = [Path(path).resolve() for path in picked]
        if not [path for path in source_templates if path != base_template]:
            self._log("已取消制作联合模板：待复制工作簿不能只有底稿模板本身")
            return False
        self.state["busy"] = True
        self.state["status"] = "正在制作联合模板，请勿关闭页面……"
        self._log(f"开始制作联合模板：底稿“{base_template.name}”")
        self._log("提示：请将含外部依赖工作表的模板选作底稿；原始文件不会修改")
        threading.Thread(
            target=self._template_merge_worker,
            args=(base_template, source_templates),
            daemon=True,
        ).start()
        return True

    # 浏览器没有自绘标题栏；顶栏按钮在 Flask 版中无操作。
    def win_minimize(self) -> None:
        return None

    def win_maximize(self, restore: bool = False) -> None:
        return None

    def win_close(self) -> None:
        return None


# 把 pywebview.api.method(...) 代理到 POST /api/method 的兼容层。
# 页面其余脚本一行不改；pywebviewready 事件在页面加载完成后派发。
_BRIDGE_JS = """
<script>
(function () {
  function callApi(method, args) {
    return fetch('/api/' + method, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(args || [])
    }).then(function (response) {
      if (!response.ok) { throw new Error('HTTP ' + response.status); }
      return response.json().then(function (data) { return data.result; });
    });
  }
  var api = new Proxy({}, {
    get: function (target, name) {
      if (typeof name !== 'string') { return undefined; }
      return function () {
        return callApi(name, Array.prototype.slice.call(arguments));
      };
    }
  });
  window.pywebview = { api: api };
  window.addEventListener('load', function () {
    window.dispatchEvent(new Event('pywebviewready'));
  });
})();
</script>
"""


def create_app() -> Flask:
    app = Flask(__name__)
    api = FlaskApi(PROJECT_ROOT)

    def as_json(value: Any):
        # WebApi 各方法返回 dict/list/bool/str，均可 JSON 化；None 也按原样返回。
        return jsonify({"result": value})

    @app.get("/")
    def index():
        html = (FRONTEND_DIR / "web" / "index.html").read_text(encoding="utf-8")
        if "</head>" not in html:
            raise RuntimeError("本地界面文件缺失 <head>，无法注入 Flask 桥接层")
        return html.replace("</head>", _BRIDGE_JS + "</head>", 1)

    @app.post("/api/<method>")
    def call(method: str):
        handler = getattr(api, method, None)
        if handler is None or not callable(handler):
            return jsonify({"error": f"未知接口：{method}"}), 404
        args = request.get_json(silent=True) or []
        if not isinstance(args, list):
            return jsonify({"error": "请求体必须是 JSON 数组"}), 400
        return as_json(handler(*args))

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=8750, debug=False)
