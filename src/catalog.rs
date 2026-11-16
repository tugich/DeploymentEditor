//! The PSADT command catalog (`Databases/PSADT.sqlite`).

use anyhow::{Context, Result};
use rusqlite::{Connection, OpenFlags};
use std::collections::HashMap;
use std::path::Path;

pub const CUSTOM_SCRIPT: &str = "#CustomScript";

#[derive(Clone, Debug)]
#[allow(dead_code)] // `category` is part of the database schema, kept for future grouping
pub struct CommandInfo {
    pub name: String,
    pub command: String,
    pub category: String,
    pub description: String,
}

#[derive(Clone, Debug)]
pub struct ParameterInfo {
    pub name: String,
    pub ptype: String,
    pub description: String,
    pub required: bool,
    /// `String`, `Checkbox` or `Script`
    pub control: String,
}

impl ParameterInfo {
    pub fn is_checkbox(&self) -> bool {
        self.control.eq_ignore_ascii_case("Checkbox")
    }
}

#[derive(Default, Debug)]
pub struct Catalog {
    pub commands: Vec<CommandInfo>,
    /// keyed by lower-case command (e.g. `start-adtprocess`), sorted by SortIndex
    params: HashMap<String, Vec<ParameterInfo>>,
}

impl Catalog {
    pub fn load(path: &Path) -> Result<Self> {
        let conn = Connection::open_with_flags(path, OpenFlags::SQLITE_OPEN_READ_ONLY)
            .with_context(|| format!("Can't open the PSADT database: {}", path.display()))?;

        let mut commands = Vec::new();
        {
            let mut stmt = conn.prepare(
                "SELECT Name, Command, IFNULL(Category,''), IFNULL(Description,'') FROM Commands ORDER BY Name ASC",
            )?;
            let rows = stmt.query_map([], |r| {
                Ok(CommandInfo {
                    name: r.get(0)?,
                    command: r.get(1)?,
                    category: r.get(2)?,
                    description: r.get(3)?,
                })
            })?;
            for r in rows {
                commands.push(r?);
            }
        }

        let mut params: HashMap<String, Vec<ParameterInfo>> = HashMap::new();
        {
            let mut stmt = conn.prepare(
                "SELECT C.Command, P.Parameter, IFNULL(P.Type,'String'), IFNULL(P.Description,''), \
                 IFNULL(P.Required,0), IFNULL(P.Control,'String') \
                 FROM Parameters AS P LEFT JOIN Commands AS C ON P.Command = C.ID \
                 ORDER BY C.Command, P.SortIndex ASC, P.ID ASC",
            )?;
            let rows = stmt.query_map([], |r| {
                let cmd: Option<String> = r.get(0)?;
                Ok((
                    cmd.unwrap_or_default(),
                    ParameterInfo {
                        name: r.get(1)?,
                        ptype: r.get(2)?,
                        description: r.get(3)?,
                        required: r.get::<_, i64>(4)? != 0,
                        control: r.get(5)?,
                    },
                ))
            })?;
            for r in rows {
                let (cmd, p) = r?;
                params.entry(cmd.to_ascii_lowercase()).or_default().push(p);
            }
        }

        Ok(Self { commands, params })
    }

    /// Commands that can be picked from the library (the `#...` pseudo commands are hidden).
    pub fn visible_commands(&self) -> impl Iterator<Item = &CommandInfo> {
        self.commands.iter().filter(|c| !c.command.contains('#'))
    }

    pub fn command(&self, command: &str) -> Option<&CommandInfo> {
        self.commands.iter().find(|c| c.command.eq_ignore_ascii_case(command))
    }

    pub fn parameters(&self, command: &str) -> &[ParameterInfo] {
        self.params
            .get(&command.to_ascii_lowercase())
            .map(|v| v.as_slice())
            .unwrap_or(&[])
    }

    pub fn parameter(&self, command: &str, parameter: &str) -> Option<&ParameterInfo> {
        // PSADT parameter names are case-insensitive in PowerShell, but the original
        // editor matched exactly; try exact first, then case-insensitive.
        let list = self.parameters(command);
        list.iter()
            .find(|p| p.name == parameter)
            .or_else(|| list.iter().find(|p| p.name.eq_ignore_ascii_case(parameter)))
    }

    /// Equivalent of `ParameterTypeByCommand()` in the PureBasic version.
    pub fn parameter_type(&self, command: &str, parameter: &str) -> &str {
        self.parameter(command, parameter).map(|p| p.ptype.as_str()).unwrap_or("")
    }

    /// Online help URL for a PSADT command.
    pub fn help_url(command: &str) -> String {
        format!(
            "https://psappdeploytoolkit.com/docs/reference/functions/{}",
            command
        )
    }
}
