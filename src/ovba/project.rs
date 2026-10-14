//! A VBA project (`vbaProject.bin`, [MS-OVBA] §2.2–2.3): a compound file holding the `PROJECT`
//! and `PROJECTwm` streams and the `VBA` storage (`dir`, `_VBA_PROJECT`, one stream per module).
//!
//! Reading takes the module names, kinds and sources. Writing keeps what isn't the modules —
//! the `dir` stream's project information and references, the `PROJECT` stream's other lines,
//! other storages (UserForm designers) — and writes the modules from their sources, with no
//! performance cache (`_VBA_PROJECT` version 0xFFFF: the host recompiles from the sources).

use std::io::{Cursor, Read, Write};

use super::codepage::{self, WINDOWS_1252};
use super::compress::{compress, decompress};

/// Largest project read (bytes).
const MAX_PROJECT: usize = 128 << 20;
/// Most modules read.
const MAX_MODULES: usize = 4096;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ModuleKind {
    /// A standard (procedural) module.
    Standard,
    /// A class module (or a designer's: a UserForm).
    Class,
    /// A document module (`ThisWorkbook`, `Sheet1`).
    Document,
}

impl ModuleKind {
    pub fn as_str(self) -> &'static str {
        match self {
            ModuleKind::Standard => "standard",
            ModuleKind::Class => "class",
            ModuleKind::Document => "document",
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct Module {
    pub name: String,
    pub stream_name: String,
    pub kind: ModuleKind,
    /// The module's source, `Attribute` lines included, lines ending in CRLF.
    pub source: String,
}

#[derive(Clone, Debug, Default)]
pub struct Project {
    pub code_page: u16,
    pub name: String,
    pub modules: Vec<Module>,
    /// The decompressed `dir` stream before `PROJECTMODULES` (information and references).
    info: Vec<u8>,
    /// The `PROJECT` stream's lines other than the module lines.
    project_lines: Vec<String>,
    /// Streams outside the `VBA` storage and the project streams, kept as they were.
    other: Vec<(String, Vec<u8>)>,
    /// Storages to recreate around `other` (designers).
    other_storages: Vec<String>,
}

fn u16_at(b: &[u8], at: usize) -> Option<u16> {
    Some(u16::from_le_bytes([*b.get(at)?, *b.get(at + 1)?]))
}

fn u32_at(b: &[u8], at: usize) -> Option<u32> {
    Some(u32::from_le_bytes([*b.get(at)?, *b.get(at + 1)?, *b.get(at + 2)?, *b.get(at + 3)?]))
}

/// One `dir` record: its id, data, and where it starts.
struct Record<'a> {
    id: u16,
    data: &'a [u8],
    start: usize,
}

fn records(dir: &[u8]) -> Result<Vec<Record<'_>>, String> {
    let mut out = Vec::new();
    let mut at = 0;
    while at + 6 <= dir.len() {
        let id = u16_at(dir, at).unwrap_or(0);
        let size = u32_at(dir, at + 2).unwrap_or(0) as usize;
        // PROJECTVERSION's size field is fixed at 4 but 6 bytes follow it ([MS-OVBA] 2.3.4.2.1.10).
        let len = if id == 0x0009 { 6 } else { size };
        let data = dir.get(at + 6..at + 6 + len).ok_or("the dir stream is cut short")?;
        out.push(Record { id, data, start: at });
        at += 6 + len;
        if id == 0x0010 {
            break;
        }
    }
    Ok(out)
}

fn stream(cf: &mut cfb::CompoundFile<Cursor<Vec<u8>>>, path: &str) -> Result<Vec<u8>, String> {
    let mut s = cf.open_stream(path).map_err(|e| format!("{path}: {e}"))?;
    let mut out = Vec::new();
    Read::by_ref(&mut s).take(MAX_PROJECT as u64).read_to_end(&mut out).map_err(|e| format!("{path}: {e}"))?;
    Ok(out)
}

impl Project {
    /// A new project with no modules.
    pub fn new() -> Project {
        Project { code_page: WINDOWS_1252, name: "VBAProject".into(), ..Default::default() }
    }

    /// Reads `vbaProject.bin`.
    pub fn read(bytes: &[u8]) -> Result<Project, String> {
        if bytes.len() > MAX_PROJECT {
            return Err("the VBA project is larger than 128 MiB".into());
        }
        // The compound file reader is a dependency: a malformed file must be an error, never a crash.
        let owned = bytes.to_vec();
        match std::panic::catch_unwind(move || Project::read_inner(owned)) {
            Ok(r) => r,
            Err(_) => Err("the VBA project is damaged".into()),
        }
    }

    fn read_inner(bytes: Vec<u8>) -> Result<Project, String> {
        let mut cf = cfb::CompoundFile::open(Cursor::new(bytes)).map_err(|e| format!("not a VBA project: {e}"))?;
        let dir = decompress(&stream(&mut cf, "/VBA/dir")?)?;
        let recs = records(&dir)?;
        let mut p = Project::new();
        let modules_at = recs.iter().find(|r| r.id == 0x000F).map(|r| r.start).unwrap_or(dir.len());
        p.info = dir.get(..modules_at).unwrap_or(&[]).to_vec();
        for r in &recs {
            match r.id {
                0x0003 => p.code_page = u16_at(r.data, 0).unwrap_or(WINDOWS_1252),
                0x0004 => p.name = codepage::decode(r.data, p.code_page),
                _ => {}
            }
        }
        let text = String::from_utf8_lossy(&stream(&mut cf, "/PROJECT").unwrap_or_default()).into_owned();
        let text = codepage::decode(text.as_bytes(), codepage::UTF8);
        let mut kinds: Vec<(String, ModuleKind)> = Vec::new();
        for line in text.split("\r\n").flat_map(|l| l.split('\n')) {
            let line = line.trim_end_matches('\r');
            let kind = if let Some(v) = line.strip_prefix("Document=") {
                Some((v.split('/').next().unwrap_or(v), ModuleKind::Document))
            } else if let Some(v) = line.strip_prefix("Module=") {
                Some((v, ModuleKind::Standard))
            } else if let Some(v) = line.strip_prefix("Class=") {
                Some((v, ModuleKind::Class))
            } else {
                line.strip_prefix("BaseClass=").map(|v| (v, ModuleKind::Class))
            };
            match kind {
                Some((name, k)) => kinds.push((name.to_string(), k)),
                None => p.project_lines.push(line.to_string()),
            }
        }
        while p.project_lines.last().is_some_and(|l| l.is_empty()) {
            p.project_lines.pop();
        }
        // The modules: MODULENAME starts one, the terminator ends it.
        let mut cur: Option<(String, String, Option<String>, u32, bool)> = None;
        for r in recs.iter().filter(|r| r.start >= modules_at) {
            match r.id {
                0x0019 => cur = Some((codepage::decode(r.data, p.code_page), String::new(), None, 0, true)),
                0x0047 => {
                    if let Some(c) = cur.as_mut() {
                        c.2 = Some(codepage::from_utf16le(r.data));
                    }
                }
                0x001A => {
                    if let Some(c) = cur.as_mut() {
                        c.1 = codepage::decode(r.data, p.code_page);
                    }
                }
                0x0031 => {
                    if let Some(c) = cur.as_mut() {
                        c.3 = u32_at(r.data, 0).unwrap_or(0);
                    }
                }
                0x0021 => {
                    if let Some(c) = cur.as_mut() {
                        c.4 = true;
                    }
                }
                0x0022 => {
                    if let Some(c) = cur.as_mut() {
                        c.4 = false;
                    }
                }
                0x002B => {
                    if let Some((name, stream_name, unicode, offset, procedural)) = cur.take() {
                        if p.modules.len() >= MAX_MODULES {
                            return Err("too many modules".into());
                        }
                        let name = unicode.filter(|u| !u.is_empty()).unwrap_or(name);
                        let stream_name = if stream_name.is_empty() { name.clone() } else { stream_name };
                        let raw = stream(&mut cf, &format!("/VBA/{stream_name}"))?;
                        let src = raw.get(offset as usize..).ok_or_else(|| format!("module {name}: its source offset is past its stream"))?;
                        let source = codepage::decode(&decompress(src)?, p.code_page);
                        let kind = if procedural {
                            ModuleKind::Standard
                        } else {
                            kinds.iter().find(|(n, _)| n.eq_ignore_ascii_case(&name)).map(|(_, k)| *k).unwrap_or(ModuleKind::Class)
                        };
                        p.modules.push(Module { name, stream_name, kind, source });
                    }
                }
                _ => {}
            }
        }
        // Everything outside the VBA storage and the project streams is kept (designers).
        let entries: Vec<(String, bool)> = cf.walk().map(|e| (e.path().to_string_lossy().replace('\\', "/"), e.is_stream())).collect();
        for (path, is_stream) in entries {
            let skip = path == "/" || path == "/PROJECT" || path == "/PROJECTwm" || path == "/VBA" || path.starts_with("/VBA/");
            if skip {
                continue;
            }
            if is_stream {
                let data = stream(&mut cf, &path)?;
                p.other.push((path, data));
            } else {
                p.other_storages.push(path);
            }
        }
        Ok(p)
    }

    pub fn module(&self, name: &str) -> Option<&Module> {
        self.modules.iter().find(|m| m.name.eq_ignore_ascii_case(name))
    }

    /// Adds a standard module or replaces a module's source. A source without its `Attribute
    /// VB_Name` line gets one.
    pub fn set_module(&mut self, name: &str, source: &str) -> Result<(), String> {
        let valid = name.chars().next().is_some_and(char::is_alphabetic) && name.len() <= 31 && name.chars().all(|c| c.is_alphanumeric() || c == '_');
        if !valid {
            return Err(format!("`{name}` isn't a valid module name (a letter, then letters, digits or _; at most 31)"));
        }
        let mut src: String = source.replace("\r\n", "\n").replace('\r', "\n").replace('\n', "\r\n");
        if !src.lines().any(|l| l.trim_start().to_ascii_lowercase().starts_with("attribute vb_name")) {
            src = format!("Attribute VB_Name = \"{name}\"\r\n{src}");
        }
        if !src.ends_with("\r\n") {
            src.push_str("\r\n");
        }
        match self.modules.iter_mut().find(|m| m.name.eq_ignore_ascii_case(name)) {
            Some(m) => m.source = src,
            None => self.modules.push(Module { name: name.to_string(), stream_name: name.to_string(), kind: ModuleKind::Standard, source: src }),
        }
        Ok(())
    }

    pub fn remove_module(&mut self, name: &str) -> bool {
        let before = self.modules.len();
        self.modules.retain(|m| !m.name.eq_ignore_ascii_case(name) || m.kind == ModuleKind::Document);
        self.modules.len() != before
    }

    /// Writes `vbaProject.bin`.
    pub fn write(&self) -> Result<Vec<u8>, String> {
        let cp = self.code_page;
        let mut cf = cfb::CompoundFile::create(Cursor::new(Vec::new())).map_err(|e| e.to_string())?;
        cf.create_storage("/VBA").map_err(|e| e.to_string())?;
        // dir
        let mut dir = if self.info.is_empty() { default_info(&self.name, cp) } else { self.info.clone() };
        rec(&mut dir, 0x000F, &(self.modules.len() as u16).to_le_bytes());
        rec(&mut dir, 0x0013, &0xFFFFu16.to_le_bytes());
        for m in &self.modules {
            rec(&mut dir, 0x0019, &codepage::encode(&m.name, cp));
            rec(&mut dir, 0x0047, &codepage::utf16le(&m.name));
            rec(&mut dir, 0x001A, &codepage::encode(&m.stream_name, cp));
            rec(&mut dir, 0x0032, &codepage::utf16le(&m.stream_name));
            rec(&mut dir, 0x001C, &[]);
            rec(&mut dir, 0x0048, &[]);
            rec(&mut dir, 0x0031, &0u32.to_le_bytes());
            rec(&mut dir, 0x001E, &0u32.to_le_bytes());
            rec(&mut dir, 0x002C, &0xFFFFu16.to_le_bytes());
            rec(&mut dir, if m.kind == ModuleKind::Standard { 0x0021 } else { 0x0022 }, &[]);
            rec(&mut dir, 0x002B, &[]);
        }
        rec(&mut dir, 0x0010, &[]);
        put(&mut cf, "/VBA/dir", &compress(&dir))?;
        // _VBA_PROJECT: no performance cache ([MS-OVBA] 2.3.4.1: version 0xFFFF on write).
        put(&mut cf, "/VBA/_VBA_PROJECT", &[0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])?;
        for m in &self.modules {
            put(&mut cf, &format!("/VBA/{}", m.stream_name), &compress(&codepage::encode(&m.source, cp)))?;
        }
        // PROJECT
        let mut lines: Vec<String> = Vec::new();
        let base = if self.project_lines.is_empty() { default_project_lines() } else { self.project_lines.clone() };
        let mut rest = base.into_iter().peekable();
        // The ID line first, then the modules, then the rest as it was.
        if let Some(first) = rest.peek()
            && first.starts_with("ID=")
            && let Some(id) = rest.next()
        {
            lines.push(id);
        }
        for m in &self.modules {
            lines.push(match m.kind {
                ModuleKind::Standard => format!("Module={}", m.name),
                ModuleKind::Class => format!("Class={}", m.name),
                ModuleKind::Document => format!("Document={}/&H00000000", m.name),
            });
        }
        let mut in_workspace = false;
        for l in rest {
            if l.starts_with('[') {
                in_workspace = l.eq_ignore_ascii_case("[Workspace]");
            } else if in_workspace
                && let Some((name, _)) = l.split_once('=')
                && self.module(name).is_none()
            {
                // A deleted module's window position.
                continue;
            }
            lines.push(l);
        }
        let mut text = lines.join("\r\n");
        text.push_str("\r\n");
        put(&mut cf, "/PROJECT", &codepage::encode(&text, cp))?;
        // PROJECTwm: each module's name in the code page, then in UTF-16.
        let mut wm = Vec::new();
        for m in &self.modules {
            wm.extend_from_slice(&codepage::encode(&m.name, cp));
            wm.push(0);
            wm.extend_from_slice(&codepage::utf16le(&m.name));
            wm.extend_from_slice(&[0, 0]);
        }
        wm.extend_from_slice(&[0, 0]);
        put(&mut cf, "/PROJECTwm", &wm)?;
        for s in &self.other_storages {
            cf.create_storage_all(s).map_err(|e| format!("{s}: {e}"))?;
        }
        for (path, data) in &self.other {
            put(&mut cf, path, data)?;
        }
        cf.flush().map_err(|e| e.to_string())?;
        Ok(cf.into_inner().into_inner())
    }
}

fn put(cf: &mut cfb::CompoundFile<Cursor<Vec<u8>>>, path: &str, data: &[u8]) -> Result<(), String> {
    let mut s = cf.create_stream(path).map_err(|e| format!("{path}: {e}"))?;
    s.write_all(data).map_err(|e| format!("{path}: {e}"))
}

fn rec(out: &mut Vec<u8>, id: u16, data: &[u8]) {
    out.extend_from_slice(&id.to_le_bytes());
    out.extend_from_slice(&(data.len() as u32).to_le_bytes());
    out.extend_from_slice(data);
}

/// The references a new project gets: the VBA and Excel libraries and OLE Automation.
const REFERENCES: [(&str, &str); 3] = [
    ("VBA", "*\\G{000204EF-0000-0000-C000-000000000046}#4.2#9#VBE7.DLL#Visual Basic For Applications"),
    ("Excel", "*\\G{00020813-0000-0000-C000-000000000046}#1.9#0#EXCEL.EXE#Microsoft Excel Object Library"),
    ("stdole", "*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#stdole2.tlb#OLE Automation"),
];

/// A new project's information and references records ([MS-OVBA] 2.3.4.2.1–2).
fn default_info(name: &str, cp: u16) -> Vec<u8> {
    let mut d = Vec::new();
    rec(&mut d, 0x0001, &1u32.to_le_bytes()); // SYSKIND: 32-bit Windows
    rec(&mut d, 0x0002, &0x0409u32.to_le_bytes()); // LCID
    rec(&mut d, 0x0014, &0x0409u32.to_le_bytes()); // LCIDINVOKE
    rec(&mut d, 0x0003, &cp.to_le_bytes()); // CODEPAGE
    rec(&mut d, 0x0004, &codepage::encode(name, cp)); // NAME
    rec(&mut d, 0x0005, &[]); // DOCSTRING
    rec(&mut d, 0x0040, &[]);
    rec(&mut d, 0x0006, &[]); // HELPFILEPATH
    rec(&mut d, 0x003D, &[]);
    rec(&mut d, 0x0007, &0u32.to_le_bytes()); // HELPCONTEXT
    rec(&mut d, 0x0008, &0u32.to_le_bytes()); // LIBFLAGS
    // VERSION: the size field is 4, then the major (4 bytes) and minor (2) versions.
    d.extend_from_slice(&0x0009u16.to_le_bytes());
    d.extend_from_slice(&4u32.to_le_bytes());
    d.extend_from_slice(&1u32.to_le_bytes());
    d.extend_from_slice(&0u16.to_le_bytes());
    rec(&mut d, 0x000C, &[]); // CONSTANTS
    rec(&mut d, 0x003C, &[]);
    for (name, libid) in REFERENCES {
        rec(&mut d, 0x0016, &codepage::encode(name, cp));
        rec(&mut d, 0x003E, &codepage::utf16le(name));
        // REFERENCEREGISTERED: the size covers the libid's length, the libid and 6 reserved bytes.
        let lib = codepage::encode(libid, cp);
        let mut body = (lib.len() as u32).to_le_bytes().to_vec();
        body.extend_from_slice(&lib);
        body.extend_from_slice(&0u32.to_le_bytes());
        body.extend_from_slice(&0u16.to_le_bytes());
        rec(&mut d, 0x000D, &body);
    }
    d
}

/// A new project's `PROJECT` stream ([MS-OVBA] 2.3.1), without the module lines.
fn default_project_lines() -> Vec<String> {
    let id = new_guid();
    vec![
        format!("ID=\"{id}\""),
        "Name=\"VBAProject\"".into(),
        "HelpContextID=\"0\"".into(),
        "VersionCompatible32=\"393222000\"".into(),
        format!("CMG=\"{}\"", encrypt(&id, &0u32.to_le_bytes())),
        format!("DPB=\"{}\"", encrypt(&id, &[0])),
        format!("GC=\"{}\"", encrypt(&id, &[0xFF])),
        String::new(),
        "[Host Extender Info]".into(),
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000".into(),
    ]
}

/// A `{XXXXXXXX-...}` project id from the clock (it only has to differ between projects).
fn new_guid() -> String {
    let t = std::time::SystemTime::now().duration_since(std::time::UNIX_EPOCH).map(|d| d.as_nanos()).unwrap_or(0);
    let mut x = (t as u64) ^ 0x9E37_79B9_7F4A_7C15;
    let mut next = || {
        x ^= x << 13;
        x ^= x >> 7;
        x ^= x << 17;
        x
    };
    let (a, b) = (next(), next());
    format!("{{{:08X}-{:04X}-{:04X}-{:04X}-{:012X}}}", a >> 32, (a >> 16) & 0xFFFF, 0x4000 | (a & 0x0FFF), 0x8000 | (b >> 48 & 0x3FFF), b & 0xFFFF_FFFF_FFFF)
}

/// The data encryption of [MS-OVBA] 2.4.3 (the `CMG`, `DPB` and `GC` values), as hex.
pub fn encrypt(project_id: &str, data: &[u8]) -> String {
    let seed: u8 = (project_id.bytes().fold(0u32, |a, b| a.wrapping_mul(31).wrapping_add(u32::from(b))) & 0xFF) as u8;
    let version: u8 = 2;
    let proj_key: u8 = project_id.bytes().fold(0u8, |a, b| a.wrapping_add(b));
    let mut out = vec![seed, seed ^ version, seed ^ proj_key];
    let mut unencrypted1 = proj_key;
    let mut encrypted1 = seed ^ proj_key;
    let mut encrypted2 = seed ^ version;
    let mut push = |b: u8, out: &mut Vec<u8>| {
        let enc = b ^ encrypted2.wrapping_add(unencrypted1);
        out.push(enc);
        encrypted2 = encrypted1;
        encrypted1 = enc;
        unencrypted1 = b;
    };
    let ignored = (seed & 6) / 2;
    for _ in 0..ignored {
        push(0x07, &mut out);
    }
    for b in (data.len() as u32).to_le_bytes().into_iter().chain(data.iter().copied()) {
        push(b, &mut out);
    }
    out.iter().map(|b| format!("{b:02X}")).collect()
}

/// [`encrypt`]'s inverse: the data, or `None` when the value is malformed.
pub fn decrypt(hex: &str) -> Option<Vec<u8>> {
    let bytes: Vec<u8> = (0..hex.len()).step_by(2).map(|i| u8::from_str_radix(hex.get(i..i + 2)?, 16).ok()).collect::<Option<_>>()?;
    let (&seed, rest) = bytes.split_first()?;
    let (&version_enc, rest) = rest.split_first()?;
    let (&proj_key_enc, rest) = rest.split_first()?;
    if seed ^ version_enc != 2 {
        return None;
    }
    let mut unencrypted1 = seed ^ proj_key_enc;
    let mut encrypted1 = proj_key_enc;
    let mut encrypted2 = version_enc;
    let mut plain = Vec::new();
    for &enc in rest {
        let b = enc ^ encrypted2.wrapping_add(unencrypted1);
        plain.push(b);
        encrypted2 = encrypted1;
        encrypted1 = enc;
        unencrypted1 = b;
    }
    let ignored = usize::from((seed & 6) / 2);
    let len = u32_at(&plain, ignored)? as usize;
    plain.get(ignored + 4..ignored + 4 + len).map(<[u8]>::to_vec)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn a_new_project_round_trips() {
        let mut p = Project::new();
        p.set_module("Module1", "Sub Hello()\n    MsgBox \"Café\"\nEnd Sub").unwrap();
        p.set_module("Tools", "Attribute VB_Name = \"Tools\"\r\nFunction Twice(x)\r\n    Twice = x * 2\r\nEnd Function\r\n").unwrap();
        let bytes = p.write().unwrap();
        let q = Project::read(&bytes).unwrap();
        assert_eq!(q.name, "VBAProject");
        assert_eq!(q.code_page, 1252);
        assert_eq!(q.modules.len(), 2);
        let m = q.module("module1").unwrap();
        assert_eq!(m.kind, ModuleKind::Standard);
        assert!(m.source.starts_with("Attribute VB_Name = \"Module1\"\r\nSub Hello()\r\n"), "{:?}", m.source);
        assert!(m.source.contains("Café"));
        assert_eq!(q.module("Tools").unwrap().source.matches("Attribute VB_Name").count(), 1);
        // Read again after another write: what isn't a module is kept.
        let mut q2 = q.clone();
        assert!(q2.remove_module("Tools"));
        q2.set_module("Module2", "Sub B()\nEnd Sub").unwrap();
        let r = Project::read(&q2.write().unwrap()).unwrap();
        let names: Vec<&str> = r.modules.iter().map(|m| m.name.as_str()).collect();
        assert_eq!(names, ["Module1", "Module2"]);
        assert_eq!(r.project_lines.iter().find(|l| l.starts_with("ID=")), q.project_lines.iter().find(|l| l.starts_with("ID=")));
        assert!(!r.project_lines.iter().any(|l| l.starts_with("Module=")), "module lines are regenerated");
    }

    #[test]
    fn document_and_class_modules_keep_their_kind() {
        let mut p = Project::new();
        p.set_module("Module1", "Sub A()\nEnd Sub").unwrap();
        p.modules.push(Module {
            name: "ThisWorkbook".into(),
            stream_name: "ThisWorkbook".into(),
            kind: ModuleKind::Document,
            source: "Attribute VB_Name = \"ThisWorkbook\"\r\n".into(),
        });
        p.modules.push(Module {
            name: "Counter".into(),
            stream_name: "Counter".into(),
            kind: ModuleKind::Class,
            source: "Attribute VB_Name = \"Counter\"\r\n".into(),
        });
        let q = Project::read(&p.write().unwrap()).unwrap();
        assert_eq!(q.module("ThisWorkbook").unwrap().kind, ModuleKind::Document);
        assert_eq!(q.module("Counter").unwrap().kind, ModuleKind::Class);
        let mut q = q;
        assert!(!q.remove_module("ThisWorkbook"), "document modules belong to the workbook");
        assert!(q.set_module("1bad", "").is_err());
        assert!(q.set_module("has space", "").is_err());
    }

    #[test]
    fn protection_values_decrypt() {
        let id = "{917DED54-440B-4FD1-A5C1-74ACF261E600}";
        for data in [&[0u8, 0, 0, 0][..], &[0], &[0xFF], b"longer data here"] {
            assert_eq!(decrypt(&encrypt(id, data)).as_deref(), Some(data));
        }
        assert!(decrypt("zz").is_none());
        assert!(decrypt("").is_none());
    }

    #[test]
    fn damaged_projects_are_errors() {
        assert!(Project::read(b"not a compound file").is_err());
        assert!(Project::read(&[]).is_err());
        let good = {
            let mut p = Project::new();
            p.set_module("M", "Sub A()\nEnd Sub").unwrap();
            p.write().unwrap()
        };
        for cut in [100, 600, 1500, good.len() - 1] {
            let _ = Project::read(good.get(..cut).unwrap_or(&good));
        }
        let mut flipped = good.clone();
        for i in (0..flipped.len()).step_by(97) {
            flipped[i] ^= 0x5A;
        }
        let _ = Project::read(&flipped);
    }
}
