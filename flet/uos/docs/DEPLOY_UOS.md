# 统信 UOS 原生版部署与打包

## 系统依赖

统信版不使用 Windows 兼容引擎、WebView2、.NET、pywebview、WebKit 或
Python GTK 绑定。Flet 的 Flutter Linux 壳在构建时需要 GTK 3 开发库，运行时
也依赖系统 GTK 运行库；这与旧 pywebview/GTK 前端路线不同。建议安装：

```bash
sudo apt update
sudo apt install python3 python3-venv libreoffice-calc libreoffice-core \
  python3-uno patchelf libgtk-3-0
```

- `python3-uno` 仅在流程包含“条件格式结果提取”时必需；
- LibreOffice Calc 是本发行版唯一的公式计算和条件格式渲染引擎；
- 若 `soffice` 不在 PATH，可设置 `BASE_AUDIT_SOFFICE=/实际/soffice/路径`。

## 开发测试

```bash
cd flet/uos
# 保留系统通过 apt 安装的 python3-uno，供条件格式提取使用。
python3 -m venv --system-site-packages .venv
. .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python run.py --flet
```

也可直接执行项目自带脚本：

```bash
chmod +x 安装依赖Flet-UOS.sh 启动Flet-UOS.sh
./安装依赖Flet-UOS.sh
./启动Flet-UOS.sh
```

若条件格式提取提示找不到 `uno`，先确认已安装 `python3-uno`，再确认工作台由上述 `.venv`（含 `--system-site-packages`）启动。

配置、模板、报送数据和 `历史审核说明.xlsx` 应放在业务目录，不能打进程序包。

## 打包

必须在目标 CPU 架构的统信系统上打包，不能从 Windows 交叉生成 Linux 成品：

```bash
chmod +x 打包Flet-UOS.sh
./打包Flet-UOS.sh
```

该脚本在目标统信系统使用官方 `flet build linux`，将 Flutter/Flet 桌面运行时
打进发行目录。构建机还需要 Flet 官方列出的 GTK 3 开发包、clang、cmake、
ninja、pkg-config、llvm/lld 等工具链；最稳妥的做法是先在目标统信系统按
Flet 的 Linux 构建依赖清单安装。终端用户不需要 Python、pywebview 或
WebKit，但应具备系统 GTK 运行库；LibreOffice Calc 仍应由系统安装，作为
复杂公式与条件格式的兼容计算引擎。

构建机可先尝试安装下面的最小桌面构建工具链（统信不同版本的软件源包名可能
略有差异）：

```bash
sudo apt install binutils clang cmake llvm lld ninja-build pkg-config \
  libgtk-3-dev libsecret-1-0 libsecret-1-dev libunwind-dev
```

第一次执行 `flet build linux` 可能会下载与当前 Flet 版本匹配的 Flutter SDK，
因此构建机需要临时具备访问 Python 软件源和 Flutter 下载源的网络条件。构建
完成后，分发目录可离线复制；业务模板、历史审核说明和 `data/` 用户设置仍应
作为程序外部文件保留，不应直接覆盖。

## 上线前核验

1. 以真实模板逐项核对 `SUMIFS`、`VLOOKUP`、`INDEX`、`MATCH`、命名区域与错误值；
2. 检查 LibreOffice 对真实模板公式的结果与 Windows Excel 基准一致；
3. 检查 LibreOffice 保存后的公式缓存是否与 Windows Excel 结果一致；
4. 核对条件格式实际填充色、历史审核说明、汇总和联合模板结果；
5. 不要将程序包升级覆盖业务目录的模板、历史说明或 JSON 设置。
