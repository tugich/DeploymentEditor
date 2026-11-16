//! A DeploymentEditor project: a folder with the PSADT framework and an
//! `Invoke-AppDeployToolkit.db` SQLite file that stores the action sequence.
//!
//! The schema is 100 % compatible with projects created by the PureBasic version.

use anyhow::{bail, Context, Result};
use rusqlite::{params, Connection, OptionalExtension};
use std::collections::BTreeMap;
use std::fs;
use std::path::{Path, PathBuf};

use crate::catalog::CUSTOM_SCRIPT;
use crate::paths::AppPaths;

pub const DATABASE_FILE: &str = "Invoke-AppDeployToolkit.db";
pub const DEPLOYMENT_FILE: &str = "Invoke-AppDeployToolkit.ps1";
/// PSADT launcher used as setup file / install command.
pub const SETUP_FILE: &str = "Invoke-AppDeployToolkit.exe";
pub const DATABASE_VERSION: &str = "1.0.4";

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, serde::Serialize, serde::Deserialize)]
pub enum DeploymentType {
    Installation,
    Uninstall,
    Repair,
}

impl DeploymentType {
    pub const ALL: [DeploymentType; 3] =
        [DeploymentType::Installation, DeploymentType::Uninstall, DeploymentType::Repair];

    /// Value stored in the `Actions.DeploymentType` column.
    pub fn as_db(&self) -> &'static str {
        match self {
            DeploymentType::Installation => "Installation",
            DeploymentType::Uninstall => "Uninstall",
            DeploymentType::Repair => "Repair",
        }
    }

    /// Value passed to `Invoke-AppDeployToolkit.exe -DeploymentType`.
    pub fn as_psadt(&self) -> &'static str {
        match self {
            DeploymentType::Installation => "Install",
            DeploymentType::Uninstall => "Uninstall",
            DeploymentType::Repair => "Repair",
        }
    }
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct Action {
    pub id: i64,
    pub step: i64,
    pub command: String,
    pub name: String,
    pub disabled: bool,
    pub description: String,
    pub continue_on_error: bool,
    pub deployment_type: String,
    pub variable_name: String,
    pub conditions: String,
    /// (parameter, value) in the order they were stored
    pub values: Vec<(String, String)>,
}

impl Action {
    pub fn is_custom_script(&self) -> bool {
        self.command == CUSTOM_SCRIPT
    }
    pub fn value(&self, parameter: &str) -> &str {
        self.values
            .iter()
            .rev() // the original took the last row when duplicates existed
            .find(|(p, _)| p == parameter)
            .map(|(_, v)| v.as_str())
            .unwrap_or("")
    }
    pub fn set_value(&mut self, parameter: &str, value: impl Into<String>) {
        let value = value.into();
        if let Some(slot) = self.values.iter_mut().find(|(p, _)| p == parameter) {
            slot.1 = value;
        } else {
            self.values.push((parameter.to_string(), value));
        }
    }
}

/// Well-known keys in the `Settings` table.
pub mod keys {
    pub const PROJECT_NAME: &str = "Project_Name";
    pub const APP_NAME: &str = "App_Name";
    pub const APP_VERSION: &str = "App_Version";
    pub const APP_VENDOR: &str = "App_Vendor";
    pub const APP_ARCH: &str = "App_Architecture";
    pub const APP_LANGUAGE: &str = "App_Language";
    pub const APP_AUTHOR: &str = "App_Author";
    pub const DATABASE_VERSION: &str = "Database_Version";
}

#[derive(Clone, Debug, Default, PartialEq)]
pub struct ProjectSettings {
    pub project_name: String,
    pub app_name: String,
    pub app_version: String,
    pub app_vendor: String,
    pub app_architecture: String,
    pub app_language: String,
    pub app_author: String,
}

pub struct Project {
    pub folder: PathBuf,
    pub database: PathBuf,
    conn: Connection,
    settings: BTreeMap<String, String>,
}

impl Project {
    /// Open an existing project database (`.db`).
    pub fn open(database: &Path) -> Result<Self> {
        if !database.is_file() {
            bail!("Project database not found: {}", database.display());
        }
        let conn = Connection::open(database)
            .with_context(|| format!("Can't open the database: {}", database.display()))?;
        let folder = database
            .parent()
            .map(Path::to_path_buf)
            .unwrap_or_else(|| PathBuf::from("."));
        let mut p = Self { folder, database: database.to_path_buf(), conn, settings: BTreeMap::new() };
        p.migrate()?;
        p.reload_settings()?;
        Ok(p)
    }

    /// Create a new project in `folder`: copies the PSADT framework and an empty database.
    /// Existing files with the same names are overwritten (same as the original editor).
    pub fn create(paths: &AppPaths, folder: &Path) -> Result<Self> {
        fs::create_dir_all(folder)?;
        let framework = paths.psadt_framework();
        if framework.is_dir() {
            copy_dir_all(&framework, folder)
                .with_context(|| format!("Copying the PSADT framework from {}", framework.display()))?;
        } else {
            bail!("PSADT framework folder is missing: {}", framework.display());
        }
        let db = folder.join(DATABASE_FILE);
        fs::copy(paths.template_database(), &db).with_context(|| "Copying the empty project database")?;
        let mut p = Self::open(&db)?;
        p.set_setting(keys::DATABASE_VERSION, DATABASE_VERSION)?;
        Ok(p)
    }

    /// Bring older project databases up to the current schema (idempotent).
    fn migrate(&mut self) -> Result<()> {
        let has_actions: bool = self
            .conn
            .query_row("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='Actions'", [], |r| {
                r.get::<_, i64>(0)
            })?
            > 0;
        if !has_actions {
            self.conn.execute_batch(
                "CREATE TABLE IF NOT EXISTS Actions (ID INTEGER NOT NULL UNIQUE, Step INTEGER, Command TEXT, Name TEXT, Disabled INTEGER, Description TEXT, ContinueOnError INTEGER, DeploymentType TEXT, VariableName TEXT, Conditions TEXT, PRIMARY KEY (ID AUTOINCREMENT));
                 CREATE TABLE IF NOT EXISTS Actions_Values (Action INTEGER, Parameter TEXT, Value TEXT, FOREIGN KEY(Action) REFERENCES Actions(ID));
                 CREATE TABLE IF NOT EXISTS Settings (ID INTEGER NOT NULL UNIQUE, Name TEXT NOT NULL UNIQUE, Value TEXT, File BLOB, PRIMARY KEY(ID AUTOINCREMENT));",
            )?;
        }
        let cols: Vec<String> = {
            let mut stmt = self.conn.prepare("PRAGMA table_info(Actions)")?;
            let rows = stmt.query_map([], |r| r.get::<_, String>(1))?;
            rows.collect::<std::result::Result<_, _>>()?
        };
        for (col, ty) in [
            ("Description", "TEXT"),
            ("ContinueOnError", "INTEGER"),
            ("DeploymentType", "TEXT"),
            ("VariableName", "TEXT"),
            ("Conditions", "TEXT"),
        ] {
            if !cols.iter().any(|c| c.eq_ignore_ascii_case(col)) {
                self.conn.execute(&format!("ALTER TABLE Actions ADD COLUMN {col} {ty}"), [])?;
            }
        }
        self.conn.execute_batch(
            "CREATE VIEW IF NOT EXISTS View_Sequence AS SELECT * FROM Actions LEFT JOIN Actions_Values ON Actions.ID = Actions_Values.Action;",
        )?;
        Ok(())
    }

    pub fn deployment_file(&self) -> PathBuf {
        self.folder.join(DEPLOYMENT_FILE)
    }
    pub fn deployment_template_file(&self) -> PathBuf {
        self.folder.join(format!("{DEPLOYMENT_FILE}.template"))
    }
    pub fn files_folder(&self) -> PathBuf {
        self.folder.join("Files")
    }
    pub fn support_files_folder(&self) -> PathBuf {
        self.folder.join("SupportFiles")
    }

    // ------------------------------------------------------------------ settings

    pub fn reload_settings(&mut self) -> Result<()> {
        let mut stmt = self.conn.prepare("SELECT Name, IFNULL(Value,'') FROM Settings")?;
        let rows = stmt.query_map([], |r| Ok((r.get::<_, String>(0)?, r.get::<_, String>(1)?)))?;
        let mut map = BTreeMap::new();
        for r in rows {
            let (k, v) = r?;
            map.insert(k, v);
        }
        self.settings = map;
        Ok(())
    }

    pub fn setting(&self, key: &str) -> &str {
        self.settings.get(key).map(String::as_str).unwrap_or("")
    }

    pub fn set_setting(&mut self, key: &str, value: &str) -> Result<()> {
        self.conn.execute("INSERT OR IGNORE INTO Settings (Name) VALUES (?1)", params![key])?;
        self.conn.execute("UPDATE Settings SET Value = ?1 WHERE Name = ?2", params![value, key])?;
        self.settings.insert(key.to_string(), value.to_string());
        Ok(())
    }

    pub fn settings(&self) -> ProjectSettings {
        ProjectSettings {
            project_name: self.setting(keys::PROJECT_NAME).into(),
            app_name: self.setting(keys::APP_NAME).into(),
            app_version: self.setting(keys::APP_VERSION).into(),
            app_vendor: self.setting(keys::APP_VENDOR).into(),
            app_architecture: self.setting(keys::APP_ARCH).into(),
            app_language: self.setting(keys::APP_LANGUAGE).into(),
            app_author: self.setting(keys::APP_AUTHOR).into(),
        }
    }

    pub fn save_settings(&mut self, s: &ProjectSettings) -> Result<()> {
        let tx_items = [
            (keys::PROJECT_NAME, &s.project_name),
            (keys::APP_NAME, &s.app_name),
            (keys::APP_VERSION, &s.app_version),
            (keys::APP_VENDOR, &s.app_vendor),
            (keys::APP_ARCH, &s.app_architecture),
            (keys::APP_LANGUAGE, &s.app_language),
            (keys::APP_AUTHOR, &s.app_author),
        ];
        for (k, v) in tx_items {
            self.set_setting(k, v)?;
        }
        Ok(())
    }

    pub fn display_name(&self) -> String {
        let n = self.setting(keys::PROJECT_NAME).trim();
        if n.is_empty() {
            self.folder
                .file_name()
                .map(|s| s.to_string_lossy().to_string())
                .unwrap_or_else(|| "Untitled project".into())
        } else {
            n.to_string()
        }
    }

    // ------------------------------------------------------------------ actions

    /// All actions of a deployment type ordered by step, with their parameter values.
    pub fn actions(&self, dt: DeploymentType) -> Result<Vec<Action>> {
        let mut stmt = self.conn.prepare(
            "SELECT ID, IFNULL(Step,0), IFNULL(Command,''), IFNULL(Name,''), IFNULL(Disabled,0), IFNULL(Description,''), \
             IFNULL(ContinueOnError,0), IFNULL(DeploymentType,''), IFNULL(VariableName,''), IFNULL(Conditions,'') \
             FROM Actions WHERE DeploymentType = ?1 ORDER BY Step ASC, ID ASC",
        )?;
        let rows = stmt.query_map(params![dt.as_db()], |r| {
            Ok(Action {
                id: r.get(0)?,
                step: r.get(1)?,
                command: r.get(2)?,
                name: r.get(3)?,
                disabled: r.get::<_, i64>(4)? != 0,
                description: r.get(5)?,
                continue_on_error: r.get::<_, i64>(6)? != 0,
                deployment_type: r.get(7)?,
                variable_name: r.get(8)?,
                conditions: r.get(9)?,
                values: Vec::new(),
            })
        })?;
        let mut actions: Vec<Action> = rows.collect::<std::result::Result<_, _>>()?;

        let mut vstmt = self.conn.prepare(
            "SELECT IFNULL(Parameter,''), IFNULL(Value,'') FROM Actions_Values WHERE Action = ?1 ORDER BY rowid ASC",
        )?;
        for a in actions.iter_mut() {
            let rows = vstmt.query_map(params![a.id], |r| Ok((r.get::<_, String>(0)?, r.get::<_, String>(1)?)))?;
            a.values = rows.collect::<std::result::Result<_, _>>()?;
        }
        Ok(actions)
    }

    /// Actions exactly as the original `BuildScript()` read them: from `View_Sequence`
    /// ordered by `Step`, so the parameter order (and therefore the generated script)
    /// is identical to the PureBasic version. Disabled actions are excluded.
    pub fn actions_for_generation(&self, dt: DeploymentType) -> Result<Vec<Action>> {
        let mut stmt = self.conn.prepare(
            "SELECT Action, IFNULL(Command,''), IFNULL(Parameter,''), IFNULL(Value,''), IFNULL(Description,''), \
             IFNULL(ContinueOnError,0), IFNULL(VariableName,''), IFNULL(Conditions,''), ID, IFNULL(Name,'') \
             FROM View_Sequence WHERE Disabled=0 AND DeploymentType=?1 ORDER BY Step ASC",
        )?;
        let mut rows = stmt.query(params![dt.as_db()])?;
        let mut out: Vec<Action> = Vec::new();
        while let Some(r) = rows.next()? {
            let action_id: Option<i64> = r.get(0)?;
            let id: i64 = r.get(8)?;
            let parameter: String = r.get(2)?;
            let value: String = r.get(3)?;
            let same = out.last().map(|a| a.id == id).unwrap_or(false);
            if !same {
                out.push(Action {
                    id,
                    command: r.get(1)?,
                    description: r.get(4)?,
                    continue_on_error: r.get::<_, i64>(5)? != 0,
                    variable_name: r.get(6)?,
                    conditions: r.get(7)?,
                    name: r.get(9)?,
                    deployment_type: dt.as_db().into(),
                    ..Default::default()
                });
            }
            if action_id.is_some() {
                out.last_mut().unwrap().values.push((parameter, value));
            }
        }
        Ok(out)
    }

    pub fn action(&self, id: i64) -> Result<Option<Action>> {
        let dt: Option<String> = self
            .conn
            .query_row("SELECT DeploymentType FROM Actions WHERE ID = ?1", params![id], |r| r.get(0))
            .optional()?;
        let Some(dt) = dt else { return Ok(None) };
        let dt = match dt.as_str() {
            "Uninstall" => DeploymentType::Uninstall,
            "Repair" => DeploymentType::Repair,
            _ => DeploymentType::Installation,
        };
        Ok(self.actions(dt)?.into_iter().find(|a| a.id == id))
    }

    fn next_step(&self, dt: DeploymentType) -> Result<i64> {
        let max: Option<i64> = self.conn.query_row(
            "SELECT MAX(Step) FROM Actions WHERE DeploymentType = ?1",
            params![dt.as_db()],
            |r| r.get(0),
        )?;
        Ok(max.unwrap_or(0) + 1)
    }

    /// Append a new action at the end of the sequence and return its ID.
    pub fn add_action(
        &mut self,
        dt: DeploymentType,
        command: &str,
        name: &str,
        description: &str,
        values: &[(&str, &str)],
    ) -> Result<i64> {
        let step = self.next_step(dt)?;
        let tx = self.conn.transaction()?;
        tx.execute(
            "INSERT INTO Actions (Step, Command, Name, Disabled, Description, ContinueOnError, DeploymentType, VariableName, Conditions) \
             VALUES (?1, ?2, ?3, 0, ?4, 0, ?5, '', '')",
            params![step, command, name, description, dt.as_db()],
        )?;
        let id = tx.last_insert_rowid();
        for (p, v) in values {
            tx.execute("INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (?1, ?2, ?3)", params![id, p, v])?;
        }
        tx.commit()?;
        Ok(id)
    }

    /// Persist the options and all parameter values of an action.
    pub fn save_action(&mut self, a: &Action) -> Result<()> {
        let tx = self.conn.transaction()?;
        tx.execute(
            "UPDATE Actions SET Name=?1, Description=?2, Conditions=?3, VariableName=?4, Disabled=?5, ContinueOnError=?6 WHERE ID=?7",
            params![
                a.name,
                a.description,
                a.conditions,
                a.variable_name,
                a.disabled as i64,
                a.continue_on_error as i64,
                a.id
            ],
        )?;
        for (p, v) in &a.values {
            tx.execute("DELETE FROM Actions_Values WHERE Action=?1 AND Parameter=?2", params![a.id, p])?;
            tx.execute("INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (?1, ?2, ?3)", params![a.id, p, v])?;
        }
        tx.commit()?;
        Ok(())
    }

    pub fn set_disabled(&mut self, id: i64, disabled: bool) -> Result<()> {
        self.conn.execute("UPDATE Actions SET Disabled=?1 WHERE ID=?2", params![disabled as i64, id])?;
        Ok(())
    }

    pub fn remove_action(&mut self, id: i64) -> Result<()> {
        let tx = self.conn.transaction()?;
        tx.execute("DELETE FROM Actions_Values WHERE Action=?1", params![id])?;
        tx.execute("DELETE FROM Actions WHERE ID=?1", params![id])?;
        tx.commit()?;
        self.renumber_steps()
    }

    /// Copy an action directly below the original. Returns the new ID.
    pub fn duplicate_action(&mut self, id: i64) -> Result<Option<i64>> {
        let Some(src) = self.action(id)? else { return Ok(None) };
        let tx = self.conn.transaction()?;
        tx.execute(
            "UPDATE Actions SET Step = Step + 1 WHERE DeploymentType = ?1 AND Step > ?2",
            params![src.deployment_type, src.step],
        )?;
        tx.execute(
            "INSERT INTO Actions (Step, Command, Name, Disabled, Description, ContinueOnError, DeploymentType, VariableName, Conditions) \
             VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9)",
            params![
                src.step + 1,
                src.command,
                format!("{} (Copy)", src.name),
                src.disabled as i64,
                src.description,
                src.continue_on_error as i64,
                src.deployment_type,
                src.variable_name,
                src.conditions
            ],
        )?;
        let new_id = tx.last_insert_rowid();
        for (p, v) in &src.values {
            tx.execute("INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (?1, ?2, ?3)", params![new_id, p, v])?;
        }
        tx.commit()?;
        self.renumber_steps()?;
        Ok(Some(new_id))
    }

    /// Move an action one position up (`-1`) or down (`+1`) within its sequence.
    /// Returns `false` if it is already at the boundary.
    pub fn move_action(&mut self, id: i64, dt: DeploymentType, delta: i32) -> Result<bool> {
        self.renumber_steps()?;
        let ids: Vec<i64> = self.actions(dt)?.iter().map(|a| a.id).collect();
        let Some(pos) = ids.iter().position(|x| *x == id) else { return Ok(false) };
        let target = pos as i64 + delta as i64;
        if target < 0 || target >= ids.len() as i64 {
            return Ok(false);
        }
        self.move_action_to(id, dt, target as usize)
    }

    /// Move an action to an absolute position (0-based) — used by drag & drop.
    pub fn move_action_to(&mut self, id: i64, dt: DeploymentType, index: usize) -> Result<bool> {
        let mut ids: Vec<i64> = self.actions(dt)?.iter().map(|a| a.id).collect();
        let Some(pos) = ids.iter().position(|x| *x == id) else { return Ok(false) };
        let index = index.min(ids.len() - 1);
        if pos == index {
            return Ok(false);
        }
        let item = ids.remove(pos);
        ids.insert(index, item);
        let tx = self.conn.transaction()?;
        for (i, aid) in ids.iter().enumerate() {
            tx.execute("UPDATE Actions SET Step=?1 WHERE ID=?2", params![i as i64 + 1, aid])?;
        }
        tx.commit()?;
        Ok(true)
    }

    /// Equivalent of `UpdateDatabaseStepCount()` — close gaps in the step numbers.
    pub fn renumber_steps(&mut self) -> Result<()> {
        for dt in DeploymentType::ALL {
            let ids: Vec<i64> = {
                let mut stmt = self
                    .conn
                    .prepare("SELECT ID FROM Actions WHERE DeploymentType=?1 ORDER BY Step ASC, ID ASC")?;
                let rows = stmt.query_map(params![dt.as_db()], |r| r.get(0))?;
                rows.collect::<std::result::Result<_, _>>()?
            };
            let tx = self.conn.transaction()?;
            for (i, id) in ids.iter().enumerate() {
                tx.execute("UPDATE Actions SET Step=?1 WHERE ID=?2", params![i as i64 + 1, id])?;
            }
            tx.commit()?;
        }
        Ok(())
    }

    pub fn action_count(&self, dt: DeploymentType) -> usize {
        self.conn
            .query_row("SELECT COUNT(*) FROM Actions WHERE DeploymentType=?1", params![dt.as_db()], |r| {
                r.get::<_, i64>(0)
            })
            .unwrap_or(0) as usize
    }

    // ------------------------------------------------------------------ undo support

    /// Capture the complete editable state (actions, values, settings) for undo.
    pub fn snapshot(&self) -> Result<Snapshot> {
        type ActionRow = (i64, Option<i64>, Option<String>, Option<String>, Option<i64>, Option<String>, Option<i64>, Option<String>, Option<String>, Option<String>);
        let actions: Vec<ActionRow> = {
            let mut stmt = self.conn.prepare(
                "SELECT ID, Step, Command, Name, Disabled, Description, ContinueOnError, DeploymentType, VariableName, Conditions FROM Actions ORDER BY ID",
            )?;
            let rows = stmt.query_map([], |r| {
                Ok((r.get(0)?, r.get(1)?, r.get(2)?, r.get(3)?, r.get(4)?, r.get(5)?, r.get(6)?, r.get(7)?, r.get(8)?, r.get(9)?))
            })?;
            rows.collect::<std::result::Result<_, _>>()?
        };
        let values: Vec<(Option<i64>, Option<String>, Option<String>)> = {
            let mut stmt = self.conn.prepare("SELECT Action, Parameter, Value FROM Actions_Values ORDER BY rowid")?;
            let rows = stmt.query_map([], |r| Ok((r.get(0)?, r.get(1)?, r.get(2)?)))?;
            rows.collect::<std::result::Result<_, _>>()?
        };
        let settings: Vec<(String, Option<String>)> = {
            let mut stmt = self.conn.prepare("SELECT Name, Value FROM Settings ORDER BY ID")?;
            let rows = stmt.query_map([], |r| Ok((r.get(0)?, r.get(1)?)))?;
            rows.collect::<std::result::Result<_, _>>()?
        };
        Ok(Snapshot { actions, values, settings })
    }

    /// Restore a snapshot taken with [`Project::snapshot`]. Action IDs are preserved.
    pub fn restore(&mut self, snap: &Snapshot) -> Result<()> {
        let tx = self.conn.transaction()?;
        tx.execute("DELETE FROM Actions_Values", [])?;
        tx.execute("DELETE FROM Actions", [])?;
        for a in &snap.actions {
            tx.execute(
                "INSERT INTO Actions (ID, Step, Command, Name, Disabled, Description, ContinueOnError, DeploymentType, VariableName, Conditions) \
                 VALUES (?1, ?2, ?3, ?4, ?5, ?6, ?7, ?8, ?9, ?10)",
                params![a.0, a.1, a.2, a.3, a.4, a.5, a.6, a.7, a.8, a.9],
            )?;
        }
        for v in &snap.values {
            tx.execute("INSERT INTO Actions_Values (Action, Parameter, Value) VALUES (?1, ?2, ?3)", params![v.0, v.1, v.2])?;
        }
        // Settings: restore values; keep rows (they may carry a File blob), drop ones added later.
        let names: Vec<&str> = snap.settings.iter().map(|(n, _)| n.as_str()).collect();
        let existing: Vec<String> = {
            let mut stmt = tx.prepare("SELECT Name FROM Settings")?;
            let rows = stmt.query_map([], |r| r.get::<_, String>(0))?;
            rows.collect::<std::result::Result<_, _>>()?
        };
        for n in existing {
            if !names.contains(&n.as_str()) {
                tx.execute("DELETE FROM Settings WHERE Name=?1", params![n])?;
            }
        }
        for (n, v) in &snap.settings {
            tx.execute("INSERT OR IGNORE INTO Settings (Name) VALUES (?1)", params![n])?;
            tx.execute("UPDATE Settings SET Value=?1 WHERE Name=?2", params![v, n])?;
        }
        tx.commit()?;
        self.reload_settings()
    }
}

/// Full copy of a project's editable tables, used by undo / redo.
#[derive(Clone, Debug, PartialEq)]
#[allow(clippy::type_complexity)]
pub struct Snapshot {
    actions: Vec<(i64, Option<i64>, Option<String>, Option<String>, Option<i64>, Option<String>, Option<i64>, Option<String>, Option<String>, Option<String>)>,
    values: Vec<(Option<i64>, Option<String>, Option<String>)>,
    settings: Vec<(String, Option<String>)>,
}

/// Recursively copy `src` into `dst`, overwriting existing files.
pub fn copy_dir_all(src: &Path, dst: &Path) -> std::io::Result<()> {
    fs::create_dir_all(dst)?;
    for entry in fs::read_dir(src)? {
        let entry = entry?;
        let ty = entry.file_type()?;
        let target = dst.join(entry.file_name());
        if ty.is_dir() {
            copy_dir_all(&entry.path(), &target)?;
        } else {
            fs::copy(entry.path(), target)?;
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn repo() -> Option<PathBuf> {
        std::env::var("DE_REPO").ok().map(PathBuf::from).filter(|p| p.join("Templates").is_dir())
    }

    fn fresh(name: &str) -> Option<(AppPaths, Project)> {
        let root = repo()?;
        let paths = AppPaths::with_root(&root);
        let dir = std::env::temp_dir().join(format!("de-proj-{name}"));
        let _ = fs::remove_dir_all(&dir);
        let p = Project::create(&paths, &dir).unwrap();
        Some((paths, p))
    }

    #[test]
    fn create_add_move_duplicate_remove() {
        let Some((_paths, mut p)) = fresh("crud") else { return };
        assert!(p.folder.join("Invoke-AppDeployToolkit.exe").is_file(), "framework copied");
        assert_eq!(p.setting(keys::DATABASE_VERSION), DATABASE_VERSION);
        let dt = DeploymentType::Installation;
        let a = p.add_action(dt, "Start-ADTProcess", "A", "", &[("FilePath", "a.exe")]).unwrap();
        let b = p.add_action(dt, "Copy-ADTFile", "B", "", &[]).unwrap();
        let c = p.add_action(dt, "Remove-ADTFile", "C", "", &[]).unwrap();
        p.add_action(DeploymentType::Uninstall, "Remove-ADTFile", "U", "", &[]).unwrap();
        let names = |p: &Project| p.actions(dt).unwrap().iter().map(|x| x.name.clone()).collect::<Vec<_>>();
        assert_eq!(names(&p), ["A", "B", "C"]);
        assert!(p.move_action(c, dt, -1).unwrap());
        assert_eq!(names(&p), ["A", "C", "B"]);
        assert!(!p.move_action(a, dt, -1).unwrap());
        p.move_action_to(a, dt, 2).unwrap();
        assert_eq!(names(&p), ["C", "B", "A"]);
        let d = p.duplicate_action(b).unwrap().unwrap();
        assert_eq!(names(&p), ["C", "B", "B (Copy)", "A"]);
        let mut act = p.action(a).unwrap().unwrap();
        assert_eq!(act.value("FilePath"), "a.exe");
        act.set_value("FilePath", "b.exe");
        act.set_value("ArgumentList", "/S");
        act.conditions = "$true".into();
        act.continue_on_error = true;
        p.save_action(&act).unwrap();
        let again = p.action(a).unwrap().unwrap();
        assert_eq!(again.value("FilePath"), "b.exe");
        assert_eq!(again.value("ArgumentList"), "/S");
        assert!(again.continue_on_error);
        assert_eq!(again.values.iter().filter(|(k, _)| k == "FilePath").count(), 1);
        p.remove_action(d).unwrap();
        p.remove_action(c).unwrap();
        assert_eq!(names(&p), ["B", "A"]);
        let steps: Vec<i64> = p.actions(dt).unwrap().iter().map(|x| x.step).collect();
        assert_eq!(steps, [1, 2]);
        assert_eq!(p.action_count(DeploymentType::Uninstall), 1);
        // generation order follows the steps
        let gen = p.actions_for_generation(dt).unwrap();
        assert_eq!(gen.iter().map(|x| x.name.as_str()).collect::<Vec<_>>(), ["B", "A"]);
    }

    #[test]
    fn snapshot_restore_roundtrip() {
        let Some((_paths, mut p)) = fresh("undo") else { return };
        let dt = DeploymentType::Installation;
        let a = p.add_action(dt, "Start-ADTProcess", "A", "", &[("FilePath", "a.exe")]).unwrap();
        p.set_setting(keys::APP_NAME, "Before").unwrap();
        let snap = p.snapshot().unwrap();
        p.add_action(dt, "Copy-ADTFile", "B", "", &[]).unwrap();
        p.remove_action(a).unwrap();
        p.set_setting(keys::APP_NAME, "After").unwrap();
        p.set_setting("Something_New", "x").unwrap();
        p.restore(&snap).unwrap();
        assert_eq!(p.snapshot().unwrap(), snap);
        let acts = p.actions(dt).unwrap();
        assert_eq!(acts.len(), 1);
        assert_eq!(acts[0].id, a);
        assert_eq!(acts[0].value("FilePath"), "a.exe");
        assert_eq!(p.setting(keys::APP_NAME), "Before");
        assert_eq!(p.setting("Something_New"), "");
        // New IDs continue after the restored ones (AUTOINCREMENT)
        let c = p.add_action(dt, "Remove-ADTFile", "C", "", &[]).unwrap();
        assert!(c > a);
    }

    #[test]
    fn settings_roundtrip_and_reopen() {
        let Some((_paths, mut p)) = fresh("settings") else { return };
        let mut s = p.settings();
        s.app_name = "7-Zip".into();
        s.app_version = "25.01".into();
        s.project_name = "Zip".into();
        p.save_settings(&s).unwrap();
        let db = p.database.clone();
        drop(p);
        let p2 = Project::open(&db).unwrap();
        assert_eq!(p2.settings().app_name, "7-Zip");
        assert_eq!(p2.display_name(), "Zip");
    }
}
