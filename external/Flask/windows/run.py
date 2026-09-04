from __future__ import annotations

import argparse
import threading
import webbrowser


def main() -> int:
    parser = argparse.ArgumentParser(description="基础数据审核工具（Flask 版）")
    parser.add_argument("--port", type=int, default=8750, help="本地服务端口")
    parser.add_argument("--no-browser", action="store_true", help="不自动打开浏览器窗口")
    args = parser.parse_args()

    from app import app

    url = f"http://127.0.0.1:{args.port}/"
    if not args.no_browser:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()
    app.run(host="127.0.0.1", port=args.port, debug=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
