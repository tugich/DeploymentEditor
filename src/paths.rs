//! Locations of the bundled runtime resources (databases, templates, plugins, ...).
//!
//! The layout is identical to the original PureBasic release, so the Rust binary
//! can simply be dropped into an existing DeploymentEditor folder.

use std::path::{Path, PathBuf};

#[derive(Clone, Debug)]
pub struct AppPaths {
    pub root: PathBuf,
}

impl AppPaths {
    /// Resolve the application root:
    /// 1. `DEPLOYMENT_EDITOR_HOME` environment variable
    /// 2. the folder that contains the executable
    /// 3. the current working directory
    /// The first candidate that contains `Databases/PSADT.sqlite` wins.
    pub fn detect() -> Self {
        let mut candidates: Vec<PathBuf> = Vec::new();
        if let Ok(p) = std::env::var("DEPLOYMENT_EDITOR_HOME") {
            candidates.push(PathBuf::from(p));
        }
        if let Ok(exe) = std::env::current_exe() {
            if let Some(dir) = exe.parent() {
                candidates.push(dir.to_path_buf());
                // `cargo run` puts the binary in target/<profile>/
                if let Some(up) = dir.parent().and_then(|p| p.parent()) {
                    candidates.push(up.to_path_buf());
                }
            }
        }
        if let Ok(cwd) = std::env::current_dir() {
            candidates.push(cwd.clone());
            if let Some(parent) = cwd.parent() {
                candidates.push(parent.to_path_buf());
            }
        }
        // Development: the resources live next to the crate in `<repo>/rust`
        // (older checkouts had them in `<repo>`).
        if cfg!(debug_assertions) {
            let manifest = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
            candidates.push(manifest.clone());
            if let Some(parent) = manifest.parent() {
                candidates.push(parent.to_path_buf());
            }
        }
        let root = candidates
            .iter()
            .find(|c| c.join("Databases").join("PSADT.sqlite").is_file())
            .cloned()
            .unwrap_or_else(|| candidates.last().cloned().unwrap_or_else(|| PathBuf::from(".")));
        Self { root }
    }

    #[cfg_attr(not(test), allow(dead_code))]
    pub fn with_root(root: impl AsRef<Path>) -> Self {
        Self { root: root.as_ref().to_path_buf() }
    }

    pub fn psadt_database(&self) -> PathBuf {
        self.root.join("Databases").join("PSADT.sqlite")
    }
    pub fn template_script(&self) -> PathBuf {
        self.root.join("Templates").join("Invoke-AppDeployToolkit.ps1")
    }
    pub fn template_database(&self) -> PathBuf {
        self.root.join("Templates").join("Invoke-AppDeployToolkit.db")
    }
    pub fn sandbox_template(&self) -> PathBuf {
        self.root.join("Templates").join("Windows Sandbox.wsb")
    }
    pub fn psadt_framework(&self) -> PathBuf {
        self.root.join("ThirdParty").join("PSAppDeployToolkit")
    }
    pub fn plugins(&self) -> PathBuf {
        self.root.join("Plugins")
    }
    pub fn snippets(&self) -> PathBuf {
        self.root.join("Snippets")
    }
    pub fn scripts(&self) -> PathBuf {
        self.root.join("Scripts")
    }
    pub fn demo_project(&self) -> PathBuf {
        self.root.join("Test")
    }
}
