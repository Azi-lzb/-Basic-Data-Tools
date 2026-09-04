#!/usr/bin/env python3
"""Batch-download every dependency .deb for offline install.

用法: python3 下载依赖.py <闭包清单> <输出目录> [批量大小]
批量调用 `apt-get download pkg1 pkg2 ...`，失败包自动重试(至多 3 轮)。
"""
import subprocess
import sys
import os
from pathlib import Path


def main() -> int:
    if len(sys.argv) < 3:
        print("用法: 下载依赖.py <闭包清单> <输出目录> [批量大小]")
        return 2
    closure = Path(sys.argv[1])
    out_dir = Path(sys.argv[2])
    batch = int(sys.argv[3]) if len(sys.argv) > 3 else 20
    out_dir.mkdir(parents=True, exist_ok=True)
    packages = [line.strip() for line in closure.read_text().splitlines() if line.strip()]
    print(f"共 {len(packages)} 个包，批量={batch}，输出到 {out_dir}")

    old_cwd = Path.cwd()
    os.chdir(out_dir)
    pending = list(packages)
    ok: set[str] = set()
    try:
        for round_no in range(1, 4):  # 至多 3 轮重试
            if not pending:
                break
            print(f"\n--- 第 {round_no} 轮，剩余 {len(pending)} 个 ---")
            next_pending = []
            for start in range(0, len(pending), batch):
                chunk = pending[start:start + batch]
                # 跳过本轮已下载的
                todo = [p for p in chunk if not (out_dir / f"{p}*.deb").exists() and p not in ok]
                # exists 用 glob 匹配不精确，直接都试，apt 对已有包会跳过/报已存在
                todo = chunk
                proc = subprocess.run(
                    ["apt-get", "download", *todo],
                    stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True,
                )
                if proc.returncode != 0:
                    # 找出本轮没成功的（重下失败的）
                    err = proc.stderr
                    # 记录失败的：回看哪些 .deb 未生成
                    for p in todo:
                        if not list(out_dir.glob(f"{p}_*.deb")):
                            next_pending.append(p)
                        else:
                            ok.add(p)
                else:
                    for p in todo:
                        if list(out_dir.glob(f"{p}_*.deb")):
                            ok.add(p)
                        else:
                            next_pending.append(p)
            pending = next_pending
    finally:
        os.chdir(old_cwd)

    fail = [p for p in packages if p not in ok]
    report = out_dir / "下载结果.txt"
    report.write_text(
        "批量下载依赖完成\n"
        f"成功 {len(ok)} 个；失败 {len(fail)} 个\n\n"
        "失败包(需到能连源机器单独补下或人工找):\n" + "\n".join(fail) + "\n",
        encoding="utf-8",
    )
    print(f"\n完成：成功 {len(ok)}，失败 {len(fail)}")
    if fail:
        print("失败包:", " ".join(fail))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
