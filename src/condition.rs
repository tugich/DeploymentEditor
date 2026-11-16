//! Visual conditions ("Run only if") — a small tree of conditions that is
//! written to / read back from the PowerShell expression stored in
//! `Actions.Conditions`. Anything the builder does not recognise is kept as a
//! "PowerShell expression" leaf, so hand-written conditions are never lost.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum GroupMode {
    All,
    Any,
    None,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum Op {
    Eq,
    Ne,
    Gt,
    Ge,
    Lt,
    Le,
    Like,
    NotLike,
}

impl Op {
    pub const ALL: [Op; 8] = [Op::Eq, Op::Ne, Op::Gt, Op::Ge, Op::Lt, Op::Le, Op::Like, Op::NotLike];
    pub const ORDER: [Op; 6] = [Op::Eq, Op::Ne, Op::Gt, Op::Ge, Op::Lt, Op::Le];
    pub fn ps(&self) -> &'static str {
        match self {
            Op::Eq => "-eq",
            Op::Ne => "-ne",
            Op::Gt => "-gt",
            Op::Ge => "-ge",
            Op::Lt => "-lt",
            Op::Le => "-le",
            Op::Like => "-like",
            Op::NotLike => "-notlike",
        }
    }
    pub fn label(&self) -> &'static str {
        match self {
            Op::Eq => "equals",
            Op::Ne => "not equals",
            Op::Gt => "greater than",
            Op::Ge => "greater than or equals",
            Op::Lt => "less than",
            Op::Le => "less than or equals",
            Op::Like => "is like (wildcard)",
            Op::NotLike => "is not like (wildcard)",
        }
    }
    fn parse(s: &str) -> Option<Op> {
        let s = s.to_ascii_lowercase();
        let s = s.strip_prefix("-c").or_else(|| s.strip_prefix("-i")).map(|r| format!("-{r}")).unwrap_or(s);
        Op::ALL.into_iter().find(|o| o.ps() == s)
    }
}

/// Built-in PSADT booleans for the "Operating system" condition.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum SystemFlag {
    Is64Bit,
    IsWorkstationOS,
    IsServerOS,
    IsMachinePartOfDomain,
    IsMultiSessionOS,
    IsTerminalServer,
}

impl SystemFlag {
    pub const ALL: [SystemFlag; 6] = [
        SystemFlag::Is64Bit,
        SystemFlag::IsWorkstationOS,
        SystemFlag::IsServerOS,
        SystemFlag::IsMachinePartOfDomain,
        SystemFlag::IsMultiSessionOS,
        SystemFlag::IsTerminalServer,
    ];
    pub fn var(&self) -> &'static str {
        match self {
            SystemFlag::Is64Bit => "$Is64Bit",
            SystemFlag::IsWorkstationOS => "$IsWorkstationOS",
            SystemFlag::IsServerOS => "$IsServerOS",
            SystemFlag::IsMachinePartOfDomain => "$IsMachinePartOfDomain",
            SystemFlag::IsMultiSessionOS => "$IsMultiSessionOS",
            SystemFlag::IsTerminalServer => "$IsTerminalServer",
        }
    }
    /// (text when true, text when false)
    pub fn texts(&self) -> (&'static str, &'static str) {
        match self {
            SystemFlag::Is64Bit => ("Windows is 64-bit", "Windows is 32-bit"),
            SystemFlag::IsWorkstationOS => ("Device is a workstation", "Device is not a workstation"),
            SystemFlag::IsServerOS => ("Device is a server", "Device is not a server"),
            SystemFlag::IsMachinePartOfDomain => ("Device is domain joined", "Device is not domain joined"),
            SystemFlag::IsMultiSessionOS => ("Windows is multi-session (AVD)", "Windows is not multi-session"),
            SystemFlag::IsTerminalServer => ("Device is a terminal server", "Device is not a terminal server"),
        }
    }
    fn from_var(s: &str) -> Option<SystemFlag> {
        SystemFlag::ALL.into_iter().find(|f| f.var().eq_ignore_ascii_case(s))
    }
}

/// Windows builds offered in the version picker.
pub const WINDOWS_VERSIONS: [(&str, &str); 8] = [
    ("10.0.19045", "Windows 10 22H2"),
    ("10.0.22000", "Windows 11 21H2"),
    ("10.0.22621", "Windows 11 22H2"),
    ("10.0.22631", "Windows 11 23H2"),
    ("10.0.26100", "Windows 11 24H2"),
    ("10.0.26200", "Windows 11 25H2"),
    ("10.0.17763", "Windows Server 2019"),
    ("10.0.20348", "Windows Server 2022"),
];

pub fn windows_name(version: &str) -> Option<&'static str> {
    WINDOWS_VERSIONS.iter().find(|(v, _)| *v == version.trim()).map(|(_, n)| *n)
}

pub const DEPLOY_MODES: [&str; 3] = ["Silent", "NonInteractive", "Interactive"];

/// Variables offered in the variable picker: (name, description).
pub const KNOWN_VARIABLES: [(&str, &str); 10] = [
    ("$envOSArchitecture", "X64, X86 or Arm64"),
    ("$envOSVersionMajor", "e.g. 10"),
    ("$envOSVersionBuild", "e.g. 22631"),
    ("$envOSName", "e.g. Microsoft Windows 11 Enterprise"),
    ("$envOSProductType", "1 workstation, 2 DC, 3 server"),
    ("$envComputerName", "Computer name"),
    ("$envMachineDNSDomain", "DNS domain of the device"),
    ("$envUserName", "User running the deployment"),
    ("$adtSession.DeploymentType", "Install, Uninstall or Repair"),
    ("$adtSession.AppVersion", "Version from the project settings"),
];

/// Folder variables offered when typing paths.
pub const PATH_VARIABLES: [(&str, &str); 10] = [
    ("$envProgramFiles", "C:\\Program Files"),
    ("$envProgramFilesX86", "C:\\Program Files (x86)"),
    ("$envProgramData", "C:\\ProgramData"),
    ("$envCommonDesktop", "Public desktop"),
    ("$envCommonStartMenuPrograms", "Start menu programs (all users)"),
    ("$envSystemRoot", "C:\\Windows"),
    ("$envSystemDrive", "C:"),
    ("$envTemp", "Temp folder"),
    ("$envLocalAppData", "Local AppData of the user"),
    ("$envAppData", "Roaming AppData of the user"),
];

#[derive(Clone, Debug, PartialEq)]
pub enum Cond {
    /// `$name -op value`
    Variable { name: String, op: Op, value: String },
    /// `$Is64Bit` / `-not $Is64Bit`
    System { flag: SystemFlag, is: bool },
    /// `[version]$envOSVersion -ge [version]'10.0.22000'`
    OsVersion { op: Op, version: String },
    /// `$adtSession.DeployMode -eq 'Silent'`
    DeployMode { mode: String, is: bool },
    /// `Test-Path -LiteralPath '…' -PathType Leaf|Container`
    File { path: String, folder: bool, exists: bool },
    /// `Test-Path -LiteralPath 'HKLM:\…'`
    RegistryKey { key: String, exists: bool },
    /// `(Get-ItemProperty -LiteralPath 'k' -Name 'n' -ErrorAction SilentlyContinue).'n' -eq 'v'`
    /// or `$null -ne (Get-ItemProperty …)` when `op` is None (value exists).
    RegistryValue { key: String, name: String, op: Option<Op>, data: String, exists: bool },
    /// `Get-ADTApplication -Name '…' [-NameMatch 'Exact']`
    App { name: String, exact: bool, installed: bool },
    /// `Get-Process -Name '…' -ErrorAction SilentlyContinue`
    Process { name: String, running: bool },
    /// Anything else, kept as typed.
    Expression { code: String },
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum CondKind {
    Variable,
    System,
    OsVersion,
    DeployMode,
    File,
    RegistryKey,
    RegistryValue,
    App,
    Process,
    Expression,
}

impl CondKind {
    pub const ALL: [CondKind; 10] = [
        CondKind::System,
        CondKind::OsVersion,
        CondKind::App,
        CondKind::File,
        CondKind::RegistryKey,
        CondKind::RegistryValue,
        CondKind::Process,
        CondKind::Variable,
        CondKind::DeployMode,
        CondKind::Expression,
    ];
    pub fn label(&self) -> &'static str {
        match self {
            CondKind::Variable => "Variable",
            CondKind::System => "Operating system",
            CondKind::OsVersion => "Windows version",
            CondKind::DeployMode => "Deploy mode",
            CondKind::File => "File or folder",
            CondKind::RegistryKey => "Registry key",
            CondKind::RegistryValue => "Registry value",
            CondKind::App => "Installed application",
            CondKind::Process => "Running process",
            CondKind::Expression => "PowerShell expression",
        }
    }
    pub fn new_cond(&self) -> Cond {
        match self {
            CondKind::Variable => Cond::Variable { name: String::new(), op: Op::Eq, value: String::new() },
            CondKind::System => Cond::System { flag: SystemFlag::Is64Bit, is: true },
            CondKind::OsVersion => Cond::OsVersion { op: Op::Ge, version: "10.0.22000".into() },
            CondKind::DeployMode => Cond::DeployMode { mode: "Interactive".into(), is: true },
            CondKind::File => Cond::File { path: String::new(), folder: false, exists: true },
            CondKind::RegistryKey => Cond::RegistryKey { key: String::new(), exists: true },
            CondKind::RegistryValue => {
                Cond::RegistryValue { key: String::new(), name: String::new(), op: None, data: String::new(), exists: true }
            }
            CondKind::App => Cond::App { name: String::new(), exact: false, installed: true },
            CondKind::Process => Cond::Process { name: String::new(), running: true },
            CondKind::Expression => Cond::Expression { code: String::new() },
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub enum Node {
    Group(Group),
    Leaf(Cond),
}

#[derive(Clone, Debug, PartialEq)]
pub struct Group {
    pub mode: GroupMode,
    pub children: Vec<Node>,
}

impl Default for Group {
    fn default() -> Self {
        Group { mode: GroupMode::All, children: Vec::new() }
    }
}

// ------------------------------------------------------------------ quoting

/// Quote a value for PowerShell: single quotes, or double quotes when it uses `$variables`.
pub fn quote(s: &str) -> String {
    if s.contains('$') {
        format!("\"{}\"", s.replace('`', "``").replace('"', "`\"").replace("``$", "`$"))
    } else {
        format!("'{}'", s.replace('\'', "''"))
    }
}

fn unquote(s: &str) -> Option<String> {
    let s = s.trim();
    if s.len() >= 2 && s.starts_with('\'') && s.ends_with('\'') {
        Some(s[1..s.len() - 1].replace("''", "'"))
    } else if s.len() >= 2 && s.starts_with('"') && s.ends_with('"') {
        let inner = &s[1..s.len() - 1];
        let mut out = String::new();
        let mut chars = inner.chars().peekable();
        while let Some(c) = chars.next() {
            if c == '`' {
                match chars.next() {
                    Some('"') => out.push('"'),
                    Some('`') => out.push('`'),
                    Some(o) => {
                        out.push('`');
                        out.push(o);
                    }
                    None => out.push('`'),
                }
            } else {
                out.push(c);
            }
        }
        Some(out)
    } else {
        None
    }
}

/// A comparison value: keep `$vars`, numbers and booleans as typed, quote text.
fn value_ps(v: &str) -> String {
    let t = v.trim();
    if t.is_empty() {
        return "''".into();
    }
    let raw = t.starts_with('$') && !t.contains(char::is_whitespace)
        || t.parse::<f64>().is_ok() && !t.contains(char::is_whitespace)
        || (t.starts_with('\'') && t.ends_with('\'') && t.len() >= 2)
        || (t.starts_with('"') && t.ends_with('"') && t.len() >= 2);
    if raw {
        t.to_string()
    } else {
        quote(t)
    }
}

fn value_from_ps(token: &str) -> String {
    // Keep quotes when the text would otherwise be read as number / variable.
    match unquote(token) {
        Some(u) if value_ps(&u) == quote(&u) => u,
        Some(_) => token.to_string(),
        None => token.to_string(),
    }
}

/// `HKEY_LOCAL_MACHINE\…` / `HKLM\…` → `HKLM:\…`
pub fn registry_ps_path(key: &str) -> String {
    let k = key.trim().trim_start_matches("Registry::");
    for (long, short) in [
        ("HKEY_LOCAL_MACHINE", "HKLM"),
        ("HKEY_CURRENT_USER", "HKCU"),
        ("HKEY_CLASSES_ROOT", "HKCR"),
        ("HKEY_USERS", "HKU"),
    ] {
        for prefix in [long, short] {
            if k.len() >= prefix.len() && k[..prefix.len()].eq_ignore_ascii_case(prefix) {
                let rest = &k[prefix.len()..];
                let rest = rest.trim_start_matches(':').trim_start_matches('\\');
                if short == "HKCR" || short == "HKU" {
                    return format!("Registry::{long}\\{rest}");
                }
                return format!("{short}:\\{rest}");
            }
        }
    }
    k.to_string()
}

fn process_name(n: &str) -> String {
    let n = n.trim();
    if n.len() > 4 && n[n.len() - 4..].eq_ignore_ascii_case(".exe") {
        n[..n.len() - 4].to_string()
    } else {
        n.to_string()
    }
}

// ------------------------------------------------------------------ to PowerShell

fn negate(positive: String, is: bool) -> String {
    if is {
        positive
    } else if positive.contains(' ') {
        format!("-not ({positive})")
    } else {
        format!("-not {positive}")
    }
}

impl Cond {
    pub fn kind(&self) -> CondKind {
        match self {
            Cond::Variable { .. } => CondKind::Variable,
            Cond::System { .. } => CondKind::System,
            Cond::OsVersion { .. } => CondKind::OsVersion,
            Cond::DeployMode { .. } => CondKind::DeployMode,
            Cond::File { .. } => CondKind::File,
            Cond::RegistryKey { .. } => CondKind::RegistryKey,
            Cond::RegistryValue { .. } => CondKind::RegistryValue,
            Cond::App { .. } => CondKind::App,
            Cond::Process { .. } => CondKind::Process,
            Cond::Expression { .. } => CondKind::Expression,
        }
    }

    pub fn to_ps(&self) -> String {
        match self {
            Cond::Variable { name, op, value } => format!("{} {} {}", name.trim(), op.ps(), value_ps(value)),
            Cond::System { flag, is } => negate(flag.var().to_string(), *is),
            Cond::OsVersion { op, version } => format!("[version]$envOSVersion {} [version]{}", op.ps(), quote(version.trim())),
            Cond::DeployMode { mode, is } => {
                format!("$adtSession.DeployMode {} {}", if *is { "-eq" } else { "-ne" }, quote(mode.trim()))
            }
            Cond::File { path, folder, exists } => negate(
                format!("Test-Path -LiteralPath {} -PathType {}", quote(path.trim()), if *folder { "Container" } else { "Leaf" }),
                *exists,
            ),
            Cond::RegistryKey { key, exists } => negate(format!("Test-Path -LiteralPath {}", quote(&registry_ps_path(key))), *exists),
            Cond::RegistryValue { key, name, op, data, exists } => {
                let get = format!(
                    "Get-ItemProperty -LiteralPath {} -Name {} -ErrorAction SilentlyContinue",
                    quote(&registry_ps_path(key)),
                    quote(name.trim())
                );
                match op {
                    None => format!("$null {} ({get})", if *exists { "-ne" } else { "-eq" }),
                    Some(op) => format!("({get}).{} {} {}", quote(name.trim()), op.ps(), value_ps(data)),
                }
            }
            Cond::App { name, exact, installed } => {
                let mut s = format!("Get-ADTApplication -Name {}", quote(name.trim()));
                if *exact {
                    s.push_str(" -NameMatch 'Exact'");
                }
                negate(s, *installed)
            }
            Cond::Process { name, running } => {
                negate(format!("Get-Process -Name {} -ErrorAction SilentlyContinue", quote(&process_name(name))), *running)
            }
            Cond::Expression { code } => code.trim().to_string(),
        }
    }

    /// Plain-language description for the tree and the PDF.
    pub fn describe(&self) -> String {
        let or_blank = |s: &str| if s.trim().is_empty() { "…".to_string() } else { s.trim().to_string() };
        match self {
            Cond::Variable { name, op, value } => format!("{} {} {}", or_blank(name), op.label(), or_blank(value)),
            Cond::System { flag, is } => {
                let (t, f) = flag.texts();
                (if *is { t } else { f }).to_string()
            }
            Cond::OsVersion { op, version } => {
                let name = windows_name(version).map(|n| format!(" ({n})")).unwrap_or_default();
                format!("Windows version {} {}{name}", op.label(), or_blank(version))
            }
            Cond::DeployMode { mode, is } => format!("Deploy mode {} {}", if *is { "is" } else { "is not" }, or_blank(mode)),
            Cond::File { path, folder, exists } => format!(
                "{} {} {}",
                if *folder { "Folder" } else { "File" },
                or_blank(path),
                if *exists { "exists" } else { "does not exist" }
            ),
            Cond::RegistryKey { key, exists } => {
                format!("Registry key {} {}", or_blank(key), if *exists { "exists" } else { "does not exist" })
            }
            Cond::RegistryValue { key, name, op, data, exists } => match op {
                None => format!(
                    "Registry value {}\\{} {}",
                    or_blank(key),
                    or_blank(name),
                    if *exists { "exists" } else { "does not exist" }
                ),
                Some(op) => format!("Registry value {}\\{} {} {}", or_blank(key), or_blank(name), op.label(), or_blank(data)),
            },
            Cond::App { name, installed, .. } => {
                format!("{} {}", or_blank(name), if *installed { "is installed" } else { "is not installed" })
            }
            Cond::Process { name, running } => {
                format!("Process {} {}", or_blank(name), if *running { "is running" } else { "is not running" })
            }
            Cond::Expression { code } => format!("PowerShell: {}", or_blank(code)),
        }
    }

    pub fn problems(&self) -> Vec<String> {
        let mut out = Vec::new();
        let empty = |s: &String| s.trim().is_empty();
        match self {
            Cond::Variable { name, .. } => {
                let n = name.trim();
                if n.is_empty() {
                    out.push("Choose or type a variable.".into());
                } else if !n.starts_with('$') || n.contains(char::is_whitespace) {
                    out.push("A variable starts with $ and contains no spaces, e.g. $envOSArchitecture.".into());
                }
            }
            Cond::OsVersion { version, .. } => {
                let ok = version.trim().split('.').all(|p| p.parse::<u32>().is_ok()) && version.trim().split('.').count() >= 2;
                if !ok {
                    out.push("Enter a version like 10.0.22631.".into());
                }
            }
            Cond::DeployMode { mode, .. } if empty(mode) => out.push("Choose a deploy mode.".into()),
            Cond::File { path, .. } if empty(path) => out.push("Enter a path.".into()),
            Cond::RegistryKey { key, .. } if empty(key) => out.push("Enter a registry key, e.g. HKLM\\SOFTWARE\\Vendor.".into()),
            Cond::RegistryValue { key, name, .. } => {
                if empty(key) {
                    out.push("Enter a registry key, e.g. HKLM\\SOFTWARE\\Vendor.".into());
                }
                if empty(name) {
                    out.push("Enter the name of the registry value.".into());
                }
            }
            Cond::App { name, .. } if empty(name) => out.push("Enter the application name as shown in Apps & features.".into()),
            Cond::Process { name, .. } if empty(name) => out.push("Enter a process name, e.g. excel.".into()),
            Cond::Expression { code } if empty(code) => out.push("Enter a PowerShell expression.".into()),
            _ => {}
        }
        if let Cond::RegistryKey { key, .. } | Cond::RegistryValue { key, .. } = self {
            let k = registry_ps_path(key);
            if !k.is_empty() && !(k.starts_with("HKLM:\\") || k.starts_with("HKCU:\\") || k.starts_with("Registry::")) {
                out.push("The key must start with HKLM, HKCU, HKCR or HKU.".into());
            }
        }
        out
    }
}

impl Node {
    pub fn to_ps(&self) -> String {
        match self {
            Node::Leaf(c) => c.to_ps(),
            Node::Group(g) => g.to_ps(),
        }
    }
}

impl Group {
    pub fn to_ps(&self) -> String {
        let parts: Vec<String> = self
            .children
            .iter()
            .filter(|c| !matches!(c, Node::Group(g) if g.children.is_empty()))
            .map(|c| c.to_ps())
            .filter(|s| !s.is_empty())
            .collect();
        if parts.is_empty() {
            return String::new();
        }
        let wrap = |p: &String| if needs_parens(p) { format!("({p})") } else { p.clone() };
        match self.mode {
            GroupMode::All | GroupMode::Any if parts.len() == 1 => parts[0].clone(),
            GroupMode::All => parts.iter().map(wrap).collect::<Vec<_>>().join(" -and "),
            GroupMode::Any => parts.iter().map(wrap).collect::<Vec<_>>().join(" -or "),
            GroupMode::None => {
                let inner = if parts.len() == 1 { parts[0].clone() } else { parts.iter().map(wrap).collect::<Vec<_>>().join(" -or ") };
                format!("-not ({inner})")
            }
        }
    }

    pub fn describe(&self) -> String {
        let parts: Vec<String> = self
            .children
            .iter()
            .map(|c| match c {
                Node::Leaf(c) => c.describe(),
                Node::Group(g) => format!("({})", g.describe()),
            })
            .collect();
        match (self.mode, parts.len()) {
            (_, 0) => String::new(),
            (GroupMode::None, 1) => format!("not: {}", parts[0]),
            (_, 1) => parts[0].clone(),
            (GroupMode::All, _) => parts.join(" and "),
            (GroupMode::Any, _) => parts.join(" or "),
            (GroupMode::None, _) => format!("none of: {}", parts.join("; ")),
        }
    }

    pub fn leaves(&self) -> Vec<&Cond> {
        let mut out = Vec::new();
        for c in &self.children {
            match c {
                Node::Leaf(l) => out.push(l),
                Node::Group(g) => out.extend(g.leaves()),
            }
        }
        out
    }

    pub fn node(&self, path: &[usize]) -> Option<&Node> {
        let (first, rest) = path.split_first()?;
        let n = self.children.get(*first)?;
        if rest.is_empty() {
            Some(n)
        } else if let Node::Group(g) = n {
            g.node(rest)
        } else {
            None
        }
    }

    pub fn node_mut(&mut self, path: &[usize]) -> Option<&mut Node> {
        let (first, rest) = path.split_first()?;
        let n = self.children.get_mut(*first)?;
        if rest.is_empty() {
            Some(n)
        } else if let Node::Group(g) = n {
            g.node_mut(rest)
        } else {
            None
        }
    }

    pub fn group_mut(&mut self, path: &[usize]) -> Option<&mut Group> {
        if path.is_empty() {
            return Some(self);
        }
        match self.node_mut(path)? {
            Node::Group(g) => Some(g),
            Node::Leaf(_) => None,
        }
    }

    /// Insert next to the selection: into a selected group, or after a selected leaf.
    /// Returns the path of the new node.
    pub fn insert(&mut self, selected: &[usize], node: Node) -> Vec<usize> {
        let target_is_group = selected.is_empty() || matches!(self.node(selected), Some(Node::Group(_)));
        if target_is_group {
            if let Some(g) = self.group_mut(selected) {
                g.children.push(node);
                let mut p = selected.to_vec();
                p.push(g.children.len() - 1);
                return p;
            }
        }
        let (parent, idx) = selected.split_at(selected.len() - 1);
        if let Some(g) = self.group_mut(parent) {
            let at = (idx[0] + 1).min(g.children.len());
            g.children.insert(at, node);
            let mut p = parent.to_vec();
            p.push(at);
            return p;
        }
        self.children.push(node);
        vec![self.children.len() - 1]
    }

    /// Remove the node at `path`; returns the path that should be selected afterwards.
    pub fn remove(&mut self, path: &[usize]) -> Vec<usize> {
        let Some((idx, parent)) = path.split_last() else { return Vec::new() };
        if let Some(g) = self.group_mut(parent) {
            if *idx < g.children.len() {
                g.children.remove(*idx);
                if g.children.is_empty() {
                    return parent.to_vec();
                }
                let mut p = parent.to_vec();
                p.push((*idx).min(g.children.len() - 1));
                return p;
            }
        }
        Vec::new()
    }

    /// Move a node one position up / down within its group.
    pub fn move_by(&mut self, path: &[usize], delta: i32) -> Vec<usize> {
        let Some((idx, parent)) = path.split_last() else { return Vec::new() };
        if let Some(g) = self.group_mut(parent) {
            let to = *idx as i32 + delta;
            if to >= 0 && (to as usize) < g.children.len() {
                g.children.swap(*idx, to as usize);
                let mut p = parent.to_vec();
                p.push(to as usize);
                return p;
            }
        }
        path.to_vec()
    }
}

// ------------------------------------------------------------------ parsing

fn needs_parens(s: &str) -> bool {
    s.contains(char::is_whitespace) && !is_wrapped(s)
}

/// True when the whole string is enclosed in one pair of parentheses.
fn is_wrapped(s: &str) -> bool {
    let s = s.trim();
    if !(s.starts_with('(') && s.ends_with(')')) {
        return false;
    }
    matching_close(s, 0) == Some(s.len() - 1)
}

/// Byte index of the bracket closing the one at `open`, honouring quotes.
fn matching_close(s: &str, open: usize) -> Option<usize> {
    let mut depth = 0i32;
    let mut quote: Option<char> = None;
    let mut escape = false;
    for (i, c) in s.char_indices().skip_while(|(i, _)| *i < open) {
        if let Some(q) = quote {
            if q == '"' && escape {
                escape = false;
                continue;
            }
            if q == '"' && c == '`' {
                escape = true;
            } else if c == q {
                quote = None;
            }
            continue;
        }
        match c {
            '\'' | '"' => quote = Some(c),
            '(' | '[' | '{' => depth += 1,
            ')' | ']' | '}' => {
                depth -= 1;
                if depth == 0 {
                    return Some(i);
                }
            }
            _ => {}
        }
    }
    None
}

/// Split at top-level ` -and ` / ` -or `. Returns the parts and the operators between them.
fn split_top(s: &str) -> (Vec<String>, Vec<String>) {
    let mut parts = Vec::new();
    let mut ops = Vec::new();
    let mut depth = 0i32;
    let mut quote: Option<char> = None;
    let mut escape = false;
    let mut start = 0;
    let lower = s.to_ascii_lowercase();
    let bytes = s.as_bytes();
    let mut i = 0;
    while i < s.len() {
        let c = bytes[i] as char;
        if !s.is_char_boundary(i) {
            i += 1;
            continue;
        }
        if let Some(q) = quote {
            if q == '"' && escape {
                escape = false;
            } else if q == '"' && c == '`' {
                escape = true;
            } else if c == q {
                quote = None;
            }
            i += 1;
            continue;
        }
        match c {
            '\'' | '"' => quote = Some(c),
            '(' | '[' | '{' => depth += 1,
            ')' | ']' | '}' => depth -= 1,
            c if c.is_whitespace() && depth == 0 => {
                for op in ["-and", "-or"] {
                    let rest = &lower[i + 1..];
                    if rest.starts_with(op) && rest[op.len()..].starts_with(char::is_whitespace) {
                        parts.push(s[start..i].trim().to_string());
                        ops.push(op.to_string());
                        start = i + 1 + op.len();
                        i = start;
                        break;
                    }
                }
            }
            _ => {}
        }
        i += 1;
    }
    parts.push(s[start..].trim().to_string());
    (parts, ops)
}

/// Split a command line into tokens at whitespace outside of quotes / brackets.
fn tokens(s: &str) -> Vec<String> {
    let mut out = Vec::new();
    let mut cur = String::new();
    let mut quote: Option<char> = None;
    let mut depth = 0i32;
    let mut escape = false;
    for c in s.chars() {
        if let Some(q) = quote {
            cur.push(c);
            if q == '"' && escape {
                escape = false;
            } else if q == '"' && c == '`' {
                escape = true;
            } else if c == q {
                quote = None;
            }
            continue;
        }
        match c {
            '\'' | '"' => {
                quote = Some(c);
                cur.push(c);
            }
            '(' | '[' | '{' => {
                depth += 1;
                cur.push(c);
            }
            ')' | ']' | '}' => {
                depth -= 1;
                cur.push(c);
            }
            c if c.is_whitespace() && depth == 0 => {
                if !cur.is_empty() {
                    out.push(std::mem::take(&mut cur));
                }
            }
            _ => cur.push(c),
        }
    }
    if !cur.is_empty() {
        out.push(cur);
    }
    out
}

fn eq(a: &str, b: &str) -> bool {
    a.eq_ignore_ascii_case(b)
}

fn is_variable(s: &str) -> bool {
    let Some(rest) = s.strip_prefix('$') else { return false };
    !rest.is_empty() && rest.chars().all(|c| c.is_alphanumeric() || c == '_' || c == ':' || c == '.')
}

fn strip_wrap(s: &str) -> &str {
    let mut s = s.trim();
    while is_wrapped(s) {
        s = s[1..s.len() - 1].trim();
    }
    s
}

/// Parse one leaf in its positive form.
fn parse_positive(s: &str) -> Option<Cond> {
    let s = strip_wrap(s);
    if let Some(flag) = SystemFlag::from_var(s) {
        return Some(Cond::System { flag, is: true });
    }
    // Registry value: "(Get-ItemProperty …).'name' -op value" or "$null -ne (Get-ItemProperty …)"
    if let Some(c) = parse_registry_value(s) {
        return Some(c);
    }
    let t = tokens(s);
    let t: Vec<&str> = t.iter().map(|x| x.as_str()).collect();
    match t.as_slice() {
        [cmd, lp, path, pt, kind] if eq(cmd, "Test-Path") && eq(lp, "-LiteralPath") && eq(pt, "-PathType") => {
            let folder = if eq(kind, "Leaf") { false } else if eq(kind, "Container") { true } else { return None };
            Some(Cond::File { path: unquote(path)?, folder, exists: true })
        }
        [cmd, lp, path] if eq(cmd, "Test-Path") && (eq(lp, "-LiteralPath") || eq(lp, "-Path")) => {
            let p = unquote(path)?;
            let k = registry_ps_path(&p);
            (k.starts_with("HKLM:") || k.starts_with("HKCU:") || k.starts_with("Registry::"))
                .then_some(Cond::RegistryKey { key: p, exists: true })
        }
        [cmd, n, name] if eq(cmd, "Get-ADTApplication") && eq(n, "-Name") => {
            Some(Cond::App { name: unquote(name)?, exact: false, installed: true })
        }
        [cmd, n, name, nm, exact] if eq(cmd, "Get-ADTApplication") && eq(n, "-Name") && eq(nm, "-NameMatch") => {
            (unquote(exact)?.eq_ignore_ascii_case("Exact")).then_some(())?;
            Some(Cond::App { name: unquote(name)?, exact: true, installed: true })
        }
        [cmd, n, name, ea, sc] if eq(cmd, "Get-Process") && eq(n, "-Name") && eq(ea, "-ErrorAction") && eq(sc, "SilentlyContinue") => {
            Some(Cond::Process { name: unquote(name)?, running: true })
        }
        [lhs, op, rhs] if eq(lhs, "[version]$envOSVersion") => {
            let v = rhs.strip_prefix("[version]").or_else(|| rhs.strip_prefix("[Version]"))?;
            Some(Cond::OsVersion { op: Op::parse(op)?, version: unquote(v)? })
        }
        [lhs, op, rhs] if eq(lhs, "$adtSession.DeployMode") && (eq(op, "-eq") || eq(op, "-ne")) => {
            Some(Cond::DeployMode { mode: unquote(rhs)?, is: eq(op, "-eq") })
        }
        [lhs, op, rhs] if is_variable(lhs) && !lhs.eq_ignore_ascii_case("$null") => {
            Some(Cond::Variable { name: lhs.to_string(), op: Op::parse(op)?, value: value_from_ps(rhs) })
        }
        _ => None,
    }
}

fn parse_get_item_property(inner: &str) -> Option<(String, String)> {
    let t = tokens(inner);
    let t: Vec<&str> = t.iter().map(|x| x.as_str()).collect();
    match t.as_slice() {
        [cmd, lp, key, n, name, ea, sc]
            if eq(cmd, "Get-ItemProperty") && eq(lp, "-LiteralPath") && eq(n, "-Name") && eq(ea, "-ErrorAction") && eq(sc, "SilentlyContinue") =>
        {
            Some((unquote(key)?, unquote(name)?))
        }
        _ => None,
    }
}

fn parse_registry_value(s: &str) -> Option<Cond> {
    let lower = s.to_ascii_lowercase();
    // $null -ne (Get-ItemProperty …)
    for (prefix, exists) in [("$null -ne ", true), ("$null -eq ", false)] {
        if lower.starts_with(prefix) {
            let rest = s[prefix.len()..].trim();
            if is_wrapped(rest) {
                let (key, name) = parse_get_item_property(&rest[1..rest.len() - 1])?;
                return Some(Cond::RegistryValue { key, name, op: None, data: String::new(), exists });
            }
        }
    }
    // (Get-ItemProperty …).'name' -op value
    if s.starts_with('(') && lower[1..].trim_start().starts_with("get-itemproperty") {
        let close = matching_close(s, 0)?;
        let (key, name) = parse_get_item_property(&s[1..close])?;
        let rest = s[close + 1..].strip_prefix('.')?;
        let t = tokens(rest);
        if let [prop, op, value] = t.as_slice() {
            let prop = unquote(prop).unwrap_or_else(|| prop.clone());
            if prop.eq_ignore_ascii_case(&name) {
                return Some(Cond::RegistryValue { key, name, op: Some(Op::parse(op)?), data: value_from_ps(value), exists: true });
            }
        }
    }
    None
}

fn parse_leaf(s: &str) -> Option<Cond> {
    let s = strip_wrap(s);
    if s.len() > 5 && s[..5].eq_ignore_ascii_case("-not ") {
        let inner = s[5..].trim();
        let mut c = parse_positive(inner)?;
        match &mut c {
            Cond::System { is, .. } | Cond::DeployMode { is, .. } => *is = !*is,
            Cond::File { exists, .. } | Cond::RegistryKey { exists, .. } => *exists = !*exists,
            Cond::App { installed, .. } => *installed = !*installed,
            Cond::Process { running, .. } => *running = !*running,
            _ => return None,
        }
        return Some(c);
    }
    parse_positive(s)
}

fn parse_node(s: &str) -> Node {
    let s = strip_wrap(s);
    if let Some(c) = parse_leaf(s) {
        return Node::Leaf(c);
    }
    let (parts, ops) = split_top(s);
    if !ops.is_empty() && parts.iter().all(|p| !p.is_empty()) {
        let mode = if ops.iter().all(|o| o == "-and") {
            Some(GroupMode::All)
        } else if ops.iter().all(|o| o == "-or") {
            Some(GroupMode::Any)
        } else {
            None
        };
        if let Some(mode) = mode {
            return Node::Group(Group { mode, children: parts.iter().map(|p| parse_node(p)).collect() });
        }
    }
    // -not ( a -or b )  →  "none of"
    if s.len() > 5 && s[..5].eq_ignore_ascii_case("-not ") && is_wrapped(s[5..].trim()) {
        match parse_node(s[5..].trim()) {
            Node::Group(g) if g.mode == GroupMode::Any => return Node::Group(Group { mode: GroupMode::None, children: g.children }),
            Node::Group(g) if g.mode == GroupMode::All => {}
            other @ Node::Group(_) => return Node::Group(Group { mode: GroupMode::None, children: vec![other] }),
            Node::Leaf(Cond::Expression { .. }) => {}
            leaf => return Node::Group(Group { mode: GroupMode::None, children: vec![leaf] }),
        }
    }
    Node::Leaf(Cond::Expression { code: s.to_string() })
}

/// Read a stored condition into the builder tree (the root is always a group).
pub fn parse(conditions: &str) -> Group {
    let s = conditions.trim();
    if s.is_empty() {
        return Group::default();
    }
    match parse_node(s) {
        Node::Group(g) => g,
        leaf => Group { mode: GroupMode::All, children: vec![leaf] },
    }
}

/// Human-readable description of a stored condition (PDF, tooltips).
pub fn describe(conditions: &str) -> String {
    parse(conditions).describe()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn roundtrip(g: &Group) {
        let ps = g.to_ps();
        let back = parse(&ps);
        assert_eq!(back.to_ps(), ps, "re-serialised text differs for {ps}");
        assert_eq!(&back, g, "tree differs for {ps}");
    }

    #[test]
    fn every_kind_roundtrips() {
        let leaves = vec![
            Cond::Variable { name: "$envOSArchitecture".into(), op: Op::Eq, value: "X64".into() },
            Cond::Variable { name: "$existing".into(), op: Op::Ne, value: "$null".into() },
            Cond::Variable { name: "$envOSVersionBuild".into(), op: Op::Ge, value: "22631".into() },
            Cond::Variable { name: "$envOSName".into(), op: Op::Like, value: "*Enterprise*".into() },
            Cond::System { flag: SystemFlag::Is64Bit, is: true },
            Cond::System { flag: SystemFlag::IsServerOS, is: false },
            Cond::OsVersion { op: Op::Ge, version: "10.0.22000".into() },
            Cond::DeployMode { mode: "Silent".into(), is: false },
            Cond::File { path: "$envProgramFiles\\7-Zip\\7z.exe".into(), folder: false, exists: true },
            Cond::File { path: "C:\\Temp\\It's here".into(), folder: true, exists: false },
            Cond::RegistryKey { key: "HKLM:\\SOFTWARE\\7-Zip".into(), exists: true },
            Cond::RegistryValue {
                key: "HKLM:\\SOFTWARE\\7-Zip".into(),
                name: "Path".into(),
                op: None,
                data: String::new(),
                exists: false,
            },
            Cond::RegistryValue {
                key: "HKLM:\\SOFTWARE\\7-Zip".into(),
                name: "Version".into(),
                op: Some(Op::Eq),
                data: "25.01".into(),
                exists: true,
            },
            Cond::App { name: "7-Zip".into(), exact: true, installed: false },
            Cond::App { name: "Adobe Acrobat".into(), exact: false, installed: true },
            Cond::Process { name: "excel".into(), running: true },
            Cond::Expression { code: "(Get-Date) -gt (Get-Date '2026-01-01')".into() },
        ];
        for l in &leaves {
            roundtrip(&Group { mode: GroupMode::All, children: vec![Node::Leaf(l.clone())] });
        }
        let all = Group { mode: GroupMode::All, children: leaves.iter().cloned().map(Node::Leaf).collect() };
        roundtrip(&all);
        let nested = Group {
            mode: GroupMode::All,
            children: vec![
                Node::Leaf(Cond::System { flag: SystemFlag::Is64Bit, is: true }),
                Node::Group(Group {
                    mode: GroupMode::Any,
                    children: vec![
                        Node::Leaf(Cond::Process { name: "excel".into(), running: false }),
                        Node::Leaf(Cond::Process { name: "winword".into(), running: false }),
                    ],
                }),
                Node::Group(Group {
                    mode: GroupMode::None,
                    children: vec![
                        Node::Leaf(Cond::App { name: "Old App".into(), exact: false, installed: true }),
                        Node::Leaf(Cond::RegistryKey { key: "HKLM:\\SOFTWARE\\Old".into(), exists: true }),
                    ],
                }),
            ],
        };
        roundtrip(&nested);
        let any = Group { mode: GroupMode::Any, children: nested.children.clone() };
        roundtrip(&any);
    }

    #[test]
    fn output_is_readable_powershell() {
        let g = Group {
            mode: GroupMode::All,
            children: vec![
                Node::Leaf(Cond::System { flag: SystemFlag::Is64Bit, is: true }),
                Node::Leaf(Cond::File { path: "$envProgramFiles\\7-Zip\\7z.exe".into(), folder: false, exists: false }),
                Node::Leaf(Cond::RegistryKey { key: "HKEY_LOCAL_MACHINE\\SOFTWARE\\7-Zip".into(), exists: true }),
            ],
        };
        assert_eq!(
            g.to_ps(),
            "$Is64Bit -and (-not (Test-Path -LiteralPath \"$envProgramFiles\\7-Zip\\7z.exe\" -PathType Leaf)) -and (Test-Path -LiteralPath 'HKLM:\\SOFTWARE\\7-Zip')"
        );
        assert_eq!(
            g.describe(),
            "Windows is 64-bit and File $envProgramFiles\\7-Zip\\7z.exe does not exist and Registry key HKEY_LOCAL_MACHINE\\SOFTWARE\\7-Zip exists"
        );
        let none = Group {
            mode: GroupMode::None,
            children: vec![Node::Leaf(Cond::Process { name: "excel.exe".into(), running: true })],
        };
        assert_eq!(none.to_ps(), "-not (Get-Process -Name 'excel' -ErrorAction SilentlyContinue)");
    }

    #[test]
    fn hand_written_conditions() {
        // The demo project's condition
        let g = parse("$existing -ne $null");
        assert_eq!(g.children, vec![Node::Leaf(Cond::Variable { name: "$existing".into(), op: Op::Ne, value: "$null".into() })]);
        // Mixed -and / -or is kept as an expression, unchanged
        let s = "$a -eq 1 -and $b -eq 2 -or $c";
        assert_eq!(parse(s).to_ps(), s);
        // Unbalanced text survives as an expression
        let s = "($existing -ne $null";
        assert_eq!(parse(s).to_ps(), s);
        // Common hand-written forms
        assert!(matches!(parse("$Is64Bit").children[0], Node::Leaf(Cond::System { is: true, .. })));
        assert!(matches!(parse("-not $Is64Bit").children[0], Node::Leaf(Cond::System { is: false, .. })));
        assert!(matches!(parse("Test-Path -Path 'HKLM:\\SOFTWARE\\X'").children[0], Node::Leaf(Cond::RegistryKey { .. })));
        let g = parse("($envOSArchitecture -eq 'X64') -and (Get-ADTApplication -Name '7-Zip')");
        assert_eq!(g.mode, GroupMode::All);
        assert_eq!(g.children.len(), 2);
        // Text values with spaces get quotes, numbers and variables do not
        assert_eq!(value_ps("Hello World"), "'Hello World'");
        assert_eq!(value_ps("42"), "42");
        assert_eq!(value_ps("$true"), "$true");
        assert_eq!(value_ps("'42'"), "'42'");
        assert_eq!(registry_ps_path("HKEY_CURRENT_USER\\Software\\X"), "HKCU:\\Software\\X");
        assert_eq!(registry_ps_path("HKLM\\Software\\X"), "HKLM:\\Software\\X");
    }

    #[test]
    fn tree_editing() {
        let mut g = Group::default();
        let a = g.insert(&[], Node::Leaf(CondKind::System.new_cond()));
        assert_eq!(a, vec![0]);
        let grp = g.insert(&a, Node::Group(Group { mode: GroupMode::Any, children: vec![] }));
        assert_eq!(grp, vec![1]);
        let inner = g.insert(&grp, Node::Leaf(CondKind::Process.new_cond()));
        assert_eq!(inner, vec![1, 0]);
        let after = g.insert(&inner, Node::Leaf(CondKind::App.new_cond()));
        assert_eq!(after, vec![1, 1]);
        assert_eq!(g.move_by(&after, -1), vec![1, 0]);
        assert_eq!(g.remove(&[1, 0]), vec![1, 0]);
        assert_eq!(g.remove(&[1, 0]), vec![1]);
        assert_eq!(g.leaves().len(), 1);
        // Empty groups are not written
        assert_eq!(g.to_ps(), "$Is64Bit");
    }
}
