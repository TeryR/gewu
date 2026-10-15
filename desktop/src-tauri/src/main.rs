// 格物 Gewu 桌面壳：拉起本地 Python 引擎 → 就绪后窗口导航到确认页；退出时带走子进程。
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::net::TcpStream;
use std::path::PathBuf;
use std::process::{Child, Command};
use std::sync::Mutex;
use std::time::{Duration, Instant};

use tauri::Manager;

struct Engine {
    child: Mutex<Option<Child>>,
}

fn engine_dir() -> PathBuf {
    // exe 位于 desktop/src-tauri/target/[debug|release]/，项目根在其上四级
    let mut d = std::env::current_exe().expect("无法定位 exe");
    for _ in 0..5 {
        d.pop();
    }
    d
}

fn spawn_engine_with_port(file: &Option<String>, port: u16) -> Option<Child> {
    let mut py = Command::new("python");
    py.current_dir(engine_dir())
        .arg("serve.py");
    if let Some(f) = file {
        py.arg(f); // serve.py <文件> <端口>：文件在前
    }
    py.arg(port.to_string())
        .env("PYTHONIOENCODING", "utf-8")
        .env("GEWU_PARENT_PID", std::process::id().to_string());
    #[cfg(windows)]
    {
        use std::os::windows::process::CommandExt;
        const CREATE_NO_WINDOW: u32 = 0x0800_0000;
        py.creation_flags(CREATE_NO_WINDOW);
    }
    py.spawn().ok()
}

fn probe_gewu(port: u16) -> bool {
    // 健康握手：端口上必须是格物（响应带属主指纹头）——防止复用陌生人占的口
    use std::io::{Read, Write};
    let Ok(mut s) = TcpStream::connect(("127.0.0.1", port)) else { return false };
    let _ = s.write_all(
        b"GET /api/status HTTP/1.1

Host: 127.0.0.1

Connection: close



",
    );
    let mut buf = String::new();
    let _ = s.read_to_string(&mut buf);
    buf.contains("X-Gewu-Owner")
}

fn pick_port() -> (u16, bool) {
    // 返回 (端口, 是否复用已有格物实例)；全忙返回 (0,false)
    for p in 8765..=8769u16 {
        if probe_gewu(p) {
            return (p, true);
        }
        if TcpStream::connect(("127.0.0.1", p)).is_err() {
            return (p, false); // 空口
        }
    }
    (0, false)
}

static TRAY_ENABLED: std::sync::atomic::AtomicBool = std::sync::atomic::AtomicBool::new(false);
static APP_PORT: std::sync::atomic::AtomicU16 = std::sync::atomic::AtomicU16::new(8765);

fn app_port() -> u16 {
    APP_PORT.load(std::sync::atomic::Ordering::Relaxed)
}

fn wait_port(secs: u32) -> bool {
    let deadline = Instant::now() + Duration::from_secs(secs as u64);
    while Instant::now() < deadline {
        if TcpStream::connect(("127.0.0.1", app_port())).is_ok() {
            return true;
        }
        std::thread::sleep(Duration::from_millis(400));
    }
    false
}

fn http_open_path(path: &str) {
    // 零依赖 HTTP POST /api/open（本机回环，手写请求足够）
    use std::io::Write;
    let body = format!("{{\"path\":{}}}", serde_json::to_string(path).unwrap_or_default());
    let req = format!(
        "POST /api/open HTTP/1.1\r\nHost: 127.0.0.1\r\nContent-Type: application/json\r\nContent-Length: {}\r\nConnection: close\r\n\r\n{}",
        body.len(),
        body
    );
    if let Ok(mut stream) = std::net::TcpStream::connect(("127.0.0.1", app_port())) {
        let _ = stream.write_all(req.as_bytes());
        let _ = std::io::Read::read_to_end(&mut stream, &mut Vec::new());
    }
}


fn read_tray_setting() -> bool {
    let p = engine_dir().join("branding.json");
    std::fs::read_to_string(p)
        .ok()
        .and_then(|s| serde_json::from_str::<serde_json::Value>(&s).ok())
        .and_then(|v| v.get("close_to_tray").and_then(|b| b.as_bool()))
        .unwrap_or(false)
}

fn main() {
    // 命令行/拖拽传入的 xlsx 路径（跳过 exe 自身与 tauri flag）
    let file = std::env::args()
        .skip(1)
        .find(|a| !a.starts_with('-') && a.to_lowercase().ends_with(".xlsx"))
        .or_else(|| {
            // exe 同名文件拖入时 Windows 有时传相对路径，规范化一下
            std::env::args()
                .skip(1)
                .find(|a| std::path::Path::new(a).is_file())
        });

    let tray_on = read_tray_setting();
    TRAY_ENABLED.store(tray_on, std::sync::atomic::Ordering::Relaxed);
    println!("托盘模式: {}", tray_on);

    let mut builder = tauri::Builder::default()
        .manage(Engine { child: Mutex::new(None) })
        .setup(move |app| {
            let handle = app.handle().clone();
            let file = file.clone();
            std::thread::spawn(move || {
                let state = handle.state::<Engine>();
                let (port, reuse) = pick_port();
                APP_PORT.store(if port == 0 { 8765 } else { port },
                               std::sync::atomic::Ordering::Relaxed);
                let mut ready = port != 0;
                if ready && !reuse {
                    let mut guard = state.child.lock().unwrap();
                    *guard = spawn_engine_with_port(&file, port);
                }
                if ready && !reuse {
                    ready = wait_port(60);
                }
                if let Some(w) = handle.get_webview_window("main") {
                    let url = if ready {
                        format!("http://127.0.0.1:{}/", app_port())
                    } else {
                        "gewu-loading.html?failed=1".to_string()
                    };
                    let _ = w.eval(&format!("location.replace('{url}')"));
                }
            });
            // 主窗口由 tauri.conf.json 的 windows 配置创建（label=main），不再重复 build
            Ok(())
        })
        .on_window_event(|window, event| {
            if let tauri::WindowEvent::CloseRequested { api, .. } = event {
                if TRAY_ENABLED.load(std::sync::atomic::Ordering::Relaxed) {
                    api.prevent_close();
                    let _ = window.hide();
                    return;
                }
            }
            if let tauri::WindowEvent::DragDrop(tauri::DragDropEvent::Drop { paths, .. }) = event {
                if let Some(x) = paths.iter().find(|p| {
                    p.to_string_lossy().to_lowercase().ends_with(".xlsx")
                }) {
                    let p = x.to_string_lossy().to_string();
                    http_open_path(&p);
                    if let Some(w) = window.webview_windows().get("main") {
                        let _ = w.eval("setTimeout(()=>location.reload(),600)");
                    }
                }
            }
            if let tauri::WindowEvent::Destroyed = event {
                if let Some(state) = window.try_state::<Engine>() {
                    if let Ok(mut guard) = state.child.lock() {
                        if let Some(c) = guard.as_mut() {
                            let _ = c.kill();
                        }
                    }
                }
            }
        })

        .on_tray_icon_event(|tray, event| {
            if let tauri::tray::TrayIconEvent::Click{ button: tauri::tray::MouseButton::Left, .. } = event {
                if let Some(w) = tray.app_handle().get_webview_window("main") {
                    let _ = w.show();
                    let _ = w.set_focus();
                }
            }
        })
        .run(tauri::generate_context!())
        .expect("格物桌面壳运行失败");
}
