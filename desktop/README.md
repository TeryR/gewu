# 格物 Gewu · 桌面壳（Tauri 2）

把 Web 工作台包成原生 Windows 应用：双击 xlsx 或拖文件到图标 → 自动拉起本地引擎 → 应用窗口打开确认页。无地址栏、无浏览器依赖、独立任务栏图标。

## 结构
- `src-tauri/`：Rust 壳。启动时读取命令行的 xlsx 路径（拖拽/双击文件关联传入），以无窗口方式拉起 `python serve.py <file>`，轮询 8765 端口就绪后把主窗口导航到确认页；退出时杀掉 Python 子进程。
- `ui/`：启动占位页（等待引擎就绪时显示）。

## 构建
```bash
cd desktop/src-tauri
cargo build --release          # 产物 target/release/gewu.exe
```
首次构建会下载 Rust 依赖（几分钟）。

## 使用
- 命令行：`gewu.exe C:\path\文件.xlsx`
- 拖拽：把 xlsx 拖到 gewu.exe 图标上
- 安装文件关联（.xlsx 双击用格物打开）需要打安装包（`cargo tauri bundle`，待图标资源就绪后启用）

## 依赖
本机需有 Python 及项目依赖（openpyxl 等）与 llm.json（可选）。桌面壳不内置 Python 运行时——后续如需单文件分发可再打 sidecar。
