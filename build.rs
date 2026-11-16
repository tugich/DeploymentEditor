fn main() {
    // Embed the application icon and version info into the Windows executable.
    #[cfg(windows)]
    {
        if std::env::var("CARGO_CFG_TARGET_OS").as_deref() == Ok("windows") {
            let mut res = winresource::WindowsResource::new();
            res.set_icon("assets/icon.ico");
            res.set("ProductName", "Deployment Editor");
            res.set("FileDescription", "Deployment Editor - visual packaging editor for PSAppDeployToolkit");
            res.set("CompanyName", "TUGI.CH");
            res.set("LegalCopyright", "TUGI.CH");
            if let Err(e) = res.compile() {
                println!("cargo:warning=Could not embed Windows resources: {e}");
            }
        }
    }
}
