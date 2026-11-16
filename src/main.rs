//! Deployment Editor — visual packaging editor for PSAppDeployToolkit (PSADT).
//! Rust rewrite with an Office-style ribbon UI.  (c) TUGI.CH
#![cfg_attr(all(windows, not(debug_assertions)), windows_subsystem = "windows")]

mod app;
mod catalog;
mod condition;
mod detection;
mod generator;
mod highlight;
mod intune;
mod intunewin;
mod logs;
mod paths;
mod pdf;
mod pdf_metrics;
mod plugins;
mod project;
mod remote;
mod report;
mod system;
mod theme;
mod validate;
mod ui;
mod widgets;
mod winget;

use eframe::egui;

fn load_icon() -> Option<egui::IconData> {
    let bytes = include_bytes!("../assets/icon.png");
    let img = image::load_from_memory(bytes).ok()?.into_rgba8();
    let (w, h) = img.dimensions();
    Some(egui::IconData { rgba: img.into_raw(), width: w, height: h })
}

/// Which renderer to use: `--renderer gl|dx` on the command line or `DE_RENDERER`.
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
enum Backend {
    /// wgpu: DirectX 12 / Vulkan on Windows; Windows always provides a software
    /// DirectX 12 (WARP), so this also starts in VMs, Windows Sandbox and RDP.
    Wgpu,
    /// OpenGL 2.0+ (egui_glow).
    Gl,
}

impl Backend {
    fn parse(s: &str) -> Option<Backend> {
        match s.trim().to_ascii_lowercase().as_str() {
            "gl" | "opengl" | "glow" => Some(Backend::Gl),
            "dx" | "dx12" | "directx" | "wgpu" | "vulkan" => Some(Backend::Wgpu),
            _ => None,
        }
    }
    fn name(&self) -> &'static str {
        match self {
            Backend::Wgpu => "DirectX / wgpu",
            Backend::Gl => "OpenGL",
        }
    }
}

fn requested_backend() -> Option<Backend> {
    let mut args = std::env::args().skip(1);
    while let Some(a) = args.next() {
        if a == "--renderer" {
            return args.next().as_deref().and_then(Backend::parse);
        }
        if let Some(v) = a.strip_prefix("--renderer=") {
            return Backend::parse(v);
        }
    }
    std::env::var("DE_RENDERER").ok().as_deref().and_then(Backend::parse)
}

/// Rank the graphics adapters: a real (or virtual VM) GPU through any API beats
/// software rendering. wgpu's default order would pick the software DirectX 12
/// (WARP) even when the VM driver offers hardware-accelerated OpenGL.
fn adapter_score(info: &eframe::wgpu::AdapterInfo) -> i32 {
    use eframe::wgpu::{Backend as B, DeviceType as D};
    let device = match info.device_type {
        D::DiscreteGpu => 400,
        D::IntegratedGpu => 300,
        D::VirtualGpu => 200,
        D::Other => 100,
        D::Cpu => 0,
    };
    let api = match info.backend {
        B::Dx12 => 30,
        B::Vulkan => 20,
        B::Gl => 10,
        _ => 0,
    };
    // Some GL drivers report "Other" for software renderers.
    let name = info.name.to_ascii_lowercase();
    let software = ["llvmpipe", "softpipe", "basic render", "swiftshader", "gdi generic", "lavapipe"]
        .iter()
        .any(|s| name.contains(s));
    if software { api } else { device + api }
}

fn wgpu_options() -> eframe::egui_wgpu::WgpuConfiguration {
    use eframe::egui_wgpu::{WgpuConfiguration, WgpuSetup, WgpuSetupCreateNew};
    let mut create = WgpuSetupCreateNew::without_display_handle();
    create.native_adapter_selector = Some(std::sync::Arc::new(|adapters, surface| {
        let mut best: Option<(&eframe::wgpu::Adapter, i32)> = None;
        for a in adapters {
            if let Some(s) = surface {
                if !a.is_surface_supported(s) {
                    continue;
                }
            }
            let info = a.get_info();
            let score = adapter_score(&info);
            eprintln!("  adapter: {:?} {:?} \"{}\" score {score}", info.backend, info.device_type, info.name);
            if best.map(|(_, b)| score > b).unwrap_or(true) {
                best = Some((a, score));
            }
        }
        best.map(|(a, _)| a.clone()).ok_or_else(|| "No graphics adapter can draw to this window.".to_string())
    }));
    WgpuConfiguration { wgpu_setup: WgpuSetup::CreateNew(create), ..Default::default() }
}

fn run(backend: Backend) -> Result<(), String> {
    let mut viewport = egui::ViewportBuilder::default()
        .with_title("Deployment Editor — TUGI.CH")
        .with_inner_size([1440.0, 900.0])
        .with_min_inner_size([980.0, 620.0])
        .with_drag_and_drop(true);
    if let Some(icon) = load_icon() {
        viewport = viewport.with_icon(icon);
    }
    let options = eframe::NativeOptions {
        viewport,
        persist_window: true,
        renderer: match backend {
            Backend::Wgpu => eframe::Renderer::Wgpu,
            Backend::Gl => eframe::Renderer::Glow,
        },
        wgpu_options: wgpu_options(),
        ..Default::default()
    };
    eprintln!("Deployment Editor: starting with the {} renderer", backend.name());
    // Some driver problems panic instead of returning an error.
    let res = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
        eframe::run_native(
            "DeploymentEditor",
            options,
            Box::new(|cc| Ok(Box::new(app::DeploymentEditorApp::new(cc)))),
        )
    }));
    match res {
        Ok(Ok(())) => Ok(()),
        Ok(Err(e)) => Err(e.to_string()),
        Err(p) => Err(p
            .downcast_ref::<String>()
            .cloned()
            .or_else(|| p.downcast_ref::<&str>().map(|s| s.to_string()))
            .unwrap_or_else(|| "unknown error".into())),
    }
}

/// eframe can't open a second window in the same process after a failed start,
/// so the fallback restarts the editor with `--renderer gl`.
fn restart_with(backend: Backend) -> bool {
    let Ok(exe) = std::env::current_exe() else { return false };
    let mut args: Vec<String> = Vec::new();
    let mut it = std::env::args().skip(1);
    while let Some(a) = it.next() {
        if a == "--renderer" {
            it.next();
        } else if !a.starts_with("--renderer=") {
            args.push(a);
        }
    }
    args.push("--renderer".into());
    args.push(if backend == Backend::Gl { "gl".into() } else { "dx".into() });
    std::process::Command::new(exe)
        .args(args)
        .env_remove("DE_RENDERER")
        .env("DE_RENDERER_FALLBACK", "1")
        .spawn()
        .is_ok()
}

fn show_start_error(text: &str) {
    eprintln!("{text}");
    let _ = rfd::MessageDialog::new()
        .set_level(rfd::MessageLevel::Error)
        .set_title("Deployment Editor")
        .set_description(text)
        .set_buttons(rfd::MessageButtons::Ok)
        .show();
}

fn main() {
    let forced = requested_backend();
    let first = forced.unwrap_or(Backend::Wgpu);
    let err = match run(first) {
        Ok(()) => return,
        Err(e) => e,
    };
    // Automatic fallback: DirectX/wgpu failed -> try OpenGL once.
    if forced.is_none() && std::env::var_os("DE_RENDERER_FALLBACK").is_none() && restart_with(Backend::Gl) {
        eprintln!("{} renderer failed ({err}), restarting with OpenGL", first.name());
        return;
    }
    let tried = if std::env::var_os("DE_RENDERER_FALLBACK").is_some() {
        "Neither DirectX (wgpu) nor OpenGL could be started.".to_string()
    } else {
        format!("The {} renderer could not be started.", first.name())
    };
    show_start_error(&format!(
        "{tried}\n\nError: {err}\n\n\
         Try:\n\
         • Install or update the graphics driver (in a VM: VMware Tools / Guest Additions and 3D acceleration).\n\
         • Start with the other renderer: DeploymentEditor.exe --renderer gl  or  --renderer dx\n\
         • Software OpenGL: copy opengl32.dll from Mesa3D for Windows ({}) next to DeploymentEditor.exe.",
        if cfg!(target_arch = "aarch64") { "arm64 folder" } else { "x64 folder" }
    ));
    std::process::exit(1);
}
