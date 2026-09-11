#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

/// Re-applies the traffic-light placement from tauri.macos.conf.json
/// (trafficLightPosition) — macOS resets the buttons to their default spot
/// (or hides them) when the window loses focus, resizes, or changes theme,
/// so the same geometry is applied again on those events.
#[cfg(target_os = "macos")]
mod traffic {
    use objc2::msg_send;
    use objc2::runtime::AnyObject;
    use objc2_foundation::{NSPoint, NSRect};

    // keep in sync with trafficLightPosition in tauri.macos.conf.json
    const X: f64 = 16.0;
    const Y: f64 = 24.0;

    pub fn apply(ns_window: *mut std::ffi::c_void) {
        unsafe {
            let win = ns_window as *mut AnyObject;
            let close: *mut AnyObject = msg_send![win, standardWindowButton: 0u64];
            let mini: *mut AnyObject = msg_send![win, standardWindowButton: 1u64];
            let zoom: *mut AnyObject = msg_send![win, standardWindowButton: 2u64];
            if close.is_null() || mini.is_null() || zoom.is_null() {
                return;
            }
            let titlebar: *mut AnyObject = msg_send![close, superview];
            if titlebar.is_null() {
                return;
            }
            let container: *mut AnyObject = msg_send![titlebar, superview];
            if container.is_null() {
                return;
            }
            let frame_view: *mut AnyObject = msg_send![container, superview];
            if frame_view.is_null() {
                return;
            }

            // same math as tao's inset handling: grow the titlebar container
            // downward, buttons keep their offset within it
            let close_rect: NSRect = msg_send![close, frame];
            let container_h = close_rect.size.height + Y;
            let frame_rect: NSRect = msg_send![frame_view, frame];
            let mut c_rect: NSRect = msg_send![container, frame];
            c_rect.size.height = container_h;
            c_rect.origin.y = frame_rect.size.height - container_h;
            let _: () = msg_send![container, setFrame: c_rect];

            let mini_rect: NSRect = msg_send![mini, frame];
            let spacing = mini_rect.origin.x - close_rect.origin.x;
            for (i, b) in [close, mini, zoom].into_iter().enumerate() {
                let r: NSRect = msg_send![b, frame];
                let o = NSPoint {
                    x: X + spacing * i as f64,
                    y: r.origin.y,
                };
                let _: () = msg_send![b, setFrameOrigin: o];
                let _: () = msg_send![b, setHidden: false];
            }
        }
    }
}


/// AI assistant ↔ Claude Code bridge. The "Claude subscription" provider runs
/// `claude -p` (headless Claude Code, billed to the user's Claude plan) as a
/// child process speaking stream-json on stdin/stdout; the webview owns the
/// protocol, this module only spawns, pipes lines and kills.
mod ai {
    use std::collections::HashMap;
    use std::io::{BufRead, BufReader, Write};
    use std::process::{Child, ChildStdin, Command, Stdio};
    use std::sync::Mutex;
    use tauri::{AppHandle, Emitter, Manager, State};

    #[derive(Default)]
    pub struct Procs(pub Mutex<HashMap<u32, (Child, ChildStdin)>>);

    #[derive(Clone, serde::Serialize)]
    struct Line {
        id: u32,
        kind: &'static str,
        line: String,
    }

    fn home() -> String {
        std::env::var("HOME")
            .or_else(|_| std::env::var("USERPROFILE"))
            .unwrap_or_default()
    }

    /// 0 = a real executable, 1 = a batch wrapper, 9 = something CreateProcess
    /// cannot run (the extensionless Unix shim npm writes, a .ps1, …).
    #[cfg(target_os = "windows")]
    fn exe_rank(p: &str) -> u8 {
        let l = p.to_ascii_lowercase();
        if l.ends_with(".exe") { 0 } else if l.ends_with(".cmd") || l.ends_with(".bat") { 1 } else { 9 }
    }

    /// npm's claude.cmd only launches the native binary shipped inside the
    /// package. Return that binary so nothing goes through cmd.exe: Rust refuses
    /// to hand a batch file any argument with quotes or newlines, and the
    /// assistant's system prompt has both.
    #[cfg(target_os = "windows")]
    fn unwrap_npm_shim(p: &str) -> Option<String> {
        if exe_rank(p) != 1 {
            return None;
        }
        let native = std::path::Path::new(p)
            .parent()?
            .join("node_modules")
            .join("@anthropic-ai")
            .join("claude-code")
            .join("bin")
            .join("claude.exe");
        if native.is_file() { Some(native.to_string_lossy().into_owned()) } else { None }
    }

    /// Every `claude` the login shell / PATH knows about, in PATH order.
    fn login_shell_lookup() -> Vec<String> {
        #[cfg(target_os = "windows")]
        {
            // `where` lists every PATH hit; with nvm-for-windows the first is the
            // extensionless Unix shim, which CreateProcess cannot run (error 193)
            let Ok(out) = Command::new("where").arg("claude").output() else { return Vec::new() };
            let s = String::from_utf8_lossy(&out.stdout);
            return s
                .lines()
                .map(str::trim)
                .filter(|l| !l.is_empty() && exe_rank(l) < 9)
                .map(String::from)
                .collect();
        }
        #[cfg(not(target_os = "windows"))]
        {
            let shell = std::env::var("SHELL").unwrap_or_else(|_| "/bin/sh".into());
            let Ok(out) = Command::new(shell).args(["-lc", "command -v claude"]).output() else { return Vec::new() };
            let s = String::from_utf8_lossy(&out.stdout);
            s.lines().map(str::trim).filter(|l| !l.is_empty()).map(String::from).collect()
        }
    }

    /// The best of several found paths. On Windows a native .exe beats a batch
    /// wrapper (unwrapped to its bundled .exe when possible); earlier hits win
    /// within a rank. Elsewhere the first hit wins, as before.
    fn best_claude(found: Vec<String>) -> Option<String> {
        #[cfg(target_os = "windows")]
        {
            let mut all: Vec<String> = Vec::new();
            for p in found {
                if let Some(native) = unwrap_npm_shim(&p) {
                    all.push(native);
                }
                all.push(p);
            }
            all.retain(|p| exe_rank(p) < 9 && std::path::Path::new(p).is_file());
            all.sort_by_key(|p| exe_rank(p));
            return all.into_iter().next();
        }
        #[cfg(not(target_os = "windows"))]
        {
            found.into_iter().find(|p| std::path::Path::new(p).is_file())
        }
    }

    /// Locate the Claude Code CLI: an explicit path first, then the usual
    /// install spots, then whatever the user's login shell resolves (GUI apps
    /// start with a bare PATH).
    #[tauri::command]
    pub fn ai_claude_path(custom: String) -> Option<String> {
        let c = custom.trim();
        if !c.is_empty() {
            #[cfg(target_os = "windows")]
            {
                // a saved path to the npm shell shim (no extension) or to
                // claude.cmd resolves to the runnable binary beside / inside it
                let mut tries = vec![c.to_string()];
                if std::path::Path::new(c).extension().is_none() {
                    for ext in ["exe", "cmd", "bat"] {
                        tries.push(format!("{c}.{ext}"));
                    }
                }
                return best_claude(tries);
            }
            #[cfg(not(target_os = "windows"))]
            {
                return if std::path::Path::new(c).is_file() { Some(c.to_string()) } else { None };
            }
        }
        let h = home();
        let mut candidates = vec![
            format!("{h}/.local/bin/claude"),
            format!("{h}/.claude/local/claude"),
            "/opt/homebrew/bin/claude".to_string(),
            "/usr/local/bin/claude".to_string(),
        ];
        if let Ok(appdata) = std::env::var("APPDATA") {
            candidates.push(format!("{appdata}\\npm\\claude.cmd"));
        }
        // nvm-for-windows links the active node (and its npm bins) here
        if let Ok(sym) = std::env::var("NVM_SYMLINK") {
            candidates.push(format!("{sym}\\claude.cmd"));
        }
        if let Ok(local) = std::env::var("LOCALAPPDATA") {
            candidates.push(format!("{local}\\Programs\\claude\\claude.exe"));
            // WinGet's portable install: its Links shim, or the package folder
            // itself — GUI apps may start with a PATH that has neither
            candidates.push(format!("{local}\\Microsoft\\WinGet\\Links\\claude.exe"));
            if let Ok(entries) = std::fs::read_dir(format!("{local}\\Microsoft\\WinGet\\Packages")) {
                for e in entries.flatten() {
                    if e.file_name().to_string_lossy().starts_with("Anthropic.ClaudeCode_") {
                        candidates.push(e.path().join("claude.exe").to_string_lossy().into_owned());
                    }
                }
            }
        }
        let mut found: Vec<String> = candidates.into_iter().filter(|p| std::path::Path::new(p).is_file()).collect();
        found.extend(login_shell_lookup());
        best_claude(found)
    }

    #[tauri::command]
    pub fn ai_spawn(app: AppHandle, procs: State<Procs>, bin: String, args: Vec<String>) -> Result<u32, String> {
        // never spawn through cmd.exe: unwrap an npm batch wrapper to the native
        // binary it launches, and refuse a bare .cmd with a fix the user can apply
        #[cfg(target_os = "windows")]
        let bin = unwrap_npm_shim(&bin).unwrap_or(bin);
        #[cfg(target_os = "windows")]
        if exe_rank(&bin) == 1 {
            return Err(format!(
                "{bin} is a batch wrapper and Windows cannot pass this prompt through cmd.exe. \
                 In AI settings set Claude Code path to a claude.exe (the native install, or \
                 node_modules\\@anthropic-ai\\claude-code\\bin\\claude.exe inside your npm folder)."
            ));
        }
        let mut cmd = Command::new(&bin);
        cmd.args(&args)
            .current_dir(home())
            .stdin(Stdio::piped())
            .stdout(Stdio::piped())
            .stderr(Stdio::piped());
        #[cfg(target_os = "windows")]
        {
            use std::os::windows::process::CommandExt;
            cmd.creation_flags(0x0800_0000); // CREATE_NO_WINDOW
        }
        let mut child = cmd.spawn().map_err(|e| format!("could not start {bin}: {e}"))?;
        let stdin = child.stdin.take().ok_or("no stdin")?;
        let stdout = child.stdout.take().ok_or("no stdout")?;
        let stderr = child.stderr.take().ok_or("no stderr")?;
        let id = child.id();
        procs.0.lock().map_err(|e| e.to_string())?.insert(id, (child, stdin));

        let app_out = app.clone();
        std::thread::spawn(move || {
            for line in BufReader::new(stdout).lines().map_while(Result::ok) {
                let _ = app_out.emit("ai-proc", Line { id, kind: "out", line });
            }
            // stdout closed: the process is done — reap it and tell the page
            if let Some(procs) = app_out.try_state::<Procs>() {
                if let Ok(mut m) = procs.0.lock() {
                    if let Some((mut child, _)) = m.remove(&id) {
                        let _ = child.wait();
                    }
                }
            }
            let _ = app_out.emit("ai-proc", Line { id, kind: "exit", line: String::new() });
        });
        let app_err = app.clone();
        std::thread::spawn(move || {
            for line in BufReader::new(stderr).lines().map_while(Result::ok) {
                let _ = app_err.emit("ai-proc", Line { id, kind: "err", line });
            }
        });
        Ok(id)
    }

    #[tauri::command]
    pub fn ai_write(procs: State<Procs>, id: u32, line: String) -> Result<(), String> {
        let mut m = procs.0.lock().map_err(|e| e.to_string())?;
        let (_, stdin) = m.get_mut(&id).ok_or("process is gone")?;
        stdin.write_all(line.as_bytes()).map_err(|e| e.to_string())?;
        stdin.write_all(b"\n").map_err(|e| e.to_string())?;
        stdin.flush().map_err(|e| e.to_string())
    }

    #[tauri::command]
    pub fn ai_kill(procs: State<Procs>, id: u32) -> Result<(), String> {
        let mut m = procs.0.lock().map_err(|e| e.to_string())?;
        if let Some((mut child, stdin)) = m.remove(&id) {
            drop(stdin);
            let _ = child.kill();
            let _ = child.wait();
        }
        Ok(())
    }
}

fn main() {
    tauri::Builder::default()
        .manage(ai::Procs::default())
        .invoke_handler(tauri::generate_handler![ai::ai_claude_path, ai::ai_spawn, ai::ai_write, ai::ai_kill])
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
        .plugin(tauri_plugin_opener::init())
        .plugin(tauri_plugin_http::init())
        .on_window_event(|_window, _event| {
            #[cfg(target_os = "macos")]
            {
                use tauri::WindowEvent;
                if matches!(
                    _event,
                    WindowEvent::Focused(_)
                        | WindowEvent::Resized(_)
                        | WindowEvent::ThemeChanged(_)
                ) {
                    if let Ok(ns) = _window.ns_window() {
                        traffic::apply(ns);
                    }
                    // AppKit re-lays the titlebar out again *after* this event
                    // on focus changes (which is what hid the buttons on
                    // blur), so re-apply once its pass has finished too
                    for delay_ms in [50u64, 250, 600] {
                        let w = _window.clone();
                        std::thread::spawn(move || {
                            std::thread::sleep(std::time::Duration::from_millis(delay_ms));
                            let w2 = w.clone();
                            let _ = w.run_on_main_thread(move || {
                                if let Ok(ns) = w2.ns_window() {
                                    traffic::apply(ns);
                                }
                            });
                        });
                    }
                }
            }
        })
        .run(tauri::generate_context!())
        .expect("error while running Headway");
}
