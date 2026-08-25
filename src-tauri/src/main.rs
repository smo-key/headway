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
    const Y: f64 = 22.0;

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

fn main() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .plugin(tauri_plugin_updater::Builder::new().build())
        .plugin(tauri_plugin_process::init())
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
