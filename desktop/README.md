# 格物 Gewu · 桌面壳（Tauri 2）

把 Web 工作台包成原生 Windows 应用：双击 xlsx 或拖文件到图标 → 自动拉起本地引擎 → 应用窗口打开确认页。无地址栏、无浏览器依赖、独立任务栏图标。

## 功能

- **拖文件到 exe 图标** → 启动并打开该文件（Windows 传参）
- **拖文件进已打开的窗口** → 在窗口内打开该文件（原生拖放，写回原文件）
- **托盘模式**（可选）：⚙️ 设置面板勾选后，关闭窗口缩到系统托盘，引擎保持运行，左键托盘秒回窗口
- **端口健康握手**：启动时探测 8765~8769 端口是否为格物本人（验证属主指纹头），不是自己人不复用
- **动态选道**：8765 被占自动尝试 8766~8769，全忙才报失败
- **看门狗**：Python 引擎监测父进程，桌面壳被强杀也零孤儿进程
- **快启**：同文件重开 0.5 秒（指纹匹配跳过解析）

## 构建

```bash
cd desktop/src-tauri
cargo build --release
# 产物：target/release/gewu.exe（约 8MB）
```

首次构建需下载 Rust 依赖（约 2 分钟）。需要 Rust 1.75+ 和 Cargo。

## 使用

```bash
# 命令行
gewu.exe C:\path\文件.xlsx

# 或拖 xlsx 到 gewu.exe 图标
# 或拖 xlsx 进已打开的 gewu 窗口
```

建议创建桌面快捷方式，日常拖文件到快捷方式即可。

## 依赖

- 目标机器需有 Python 3.12+ 及项目依赖（`pip install -r requirements.txt`）
- 桌面壳不内置 Python——它启动 `python serve.py` 作为本地引擎
- 如需免 Python 分发，后续可打 PyInstaller sidecar

## 技术说明

- **Tauri 2** + WebView2（Windows 自带，无需额外安装）
- 托盘：`tray-icon` feature，设置存 `branding.json` 的 `close_to_tray` 字段
- 端口探测：HTTP GET `/api/status` 验证响应头 `X-Gewu-Owner`（sha256 属主指纹）
- 子进程管理：`CREATE_NO_WINDOW` 启动 Python + `GEWU_PARENT_PID` 环境变量注入看门狗
- 图标：`make_icon.py` 纯代码生成（无设计工具依赖）
