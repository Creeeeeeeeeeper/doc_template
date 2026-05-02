use std::collections::BTreeSet;
use std::fs;
use std::path::PathBuf;

use ttf_parser::Face;
use walkdir::WalkDir;

#[tauri::command]
fn read_file_bytes(path: String) -> Result<Vec<u8>, String> {
    fs::read(&path).map_err(|e| format!("read {}: {}", path, e))
}

#[tauri::command]
fn save_bytes(path: String, bytes: Vec<u8>) -> Result<(), String> {
    fs::write(&path, &bytes).map_err(|e| format!("write {}: {}", path, e))
}

// ----------------------------------------------------------------
// Font enumeration
// ----------------------------------------------------------------
// Replaces window.queryLocalFonts() so we don't need a user-activation
// permission gesture, and so we can pull localized family names (Chinese,
// Japanese, etc.) directly out of each font file's `name` table — that is
// the same source Word uses, so anything Word shows we should show too.

fn font_directories() -> Vec<PathBuf> {
    let mut dirs: Vec<PathBuf> = Vec::new();

    #[cfg(target_os = "windows")]
    {
        if let Some(windir) = std::env::var_os("WINDIR") {
            dirs.push(PathBuf::from(windir).join("Fonts"));
        } else {
            dirs.push(PathBuf::from(r"C:\Windows\Fonts"));
        }
        // Per-user fonts (installed without admin rights since Win10 1809)
        if let Some(local) = std::env::var_os("LOCALAPPDATA") {
            dirs.push(PathBuf::from(local).join(r"Microsoft\Windows\Fonts"));
        }
    }

    #[cfg(target_os = "macos")]
    {
        dirs.push(PathBuf::from("/System/Library/Fonts"));
        dirs.push(PathBuf::from("/Library/Fonts"));
        if let Some(home) = std::env::var_os("HOME") {
            dirs.push(PathBuf::from(home).join("Library/Fonts"));
        }
    }

    #[cfg(target_os = "linux")]
    {
        dirs.push(PathBuf::from("/usr/share/fonts"));
        dirs.push(PathBuf::from("/usr/local/share/fonts"));
        if let Some(home) = std::env::var_os("HOME") {
            dirs.push(PathBuf::from(home).join(".fonts"));
            dirs.push(PathBuf::from(home).join(".local/share/fonts"));
        }
    }

    dirs
}

fn is_font_extension(ext: &str) -> bool {
    matches!(
        ext.to_ascii_lowercase().as_str(),
        "ttf" | "otf" | "ttc" | "otc"
    )
}

// nameID 1 = Family name. nameID 16 = Typographic / Preferred family
// (used by modern fonts that split their styles into multiple families).
// Both should be exposed so users see exactly what Word shows them.
const NAME_ID_FAMILY: u16 = 1;
const NAME_ID_PREFERRED_FAMILY: u16 = 16;

fn extract_family_names(bytes: &[u8], face_index: u32, into: &mut BTreeSet<String>) {
    let face = match Face::parse(bytes, face_index) {
        Ok(f) => f,
        Err(_) => return,
    };
    let names = face.names();
    for i in 0..names.len() {
        let Some(record) = names.get(i) else { continue };
        if record.name_id != NAME_ID_FAMILY && record.name_id != NAME_ID_PREFERRED_FAMILY {
            continue;
        }
        // ttf-parser decodes per platform/encoding; UTF-16BE for Windows
        // platform records (which is where the Chinese names live).
        if let Some(s) = record.to_string() {
            let trimmed = s.trim();
            if !trimmed.is_empty() {
                into.insert(trimmed.to_string());
            }
        }
    }
}

#[tauri::command]
fn list_fonts() -> Vec<String> {
    let mut all: BTreeSet<String> = BTreeSet::new();

    for dir in font_directories() {
        if !dir.is_dir() {
            continue;
        }
        for entry in WalkDir::new(&dir)
            .max_depth(4)
            .into_iter()
            .filter_map(Result::ok)
        {
            let path = entry.path();
            let Some(ext) = path.extension().and_then(|s| s.to_str()) else {
                continue;
            };
            if !is_font_extension(ext) {
                continue;
            }
            let Ok(bytes) = fs::read(path) else { continue };
            // TTC / OTC: enumerate every face; otherwise face_index = 0.
            let face_count = ttf_parser::fonts_in_collection(&bytes).unwrap_or(1);
            for i in 0..face_count {
                extract_family_names(&bytes, i, &mut all);
            }
        }
    }

    all.into_iter().collect()
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .invoke_handler(tauri::generate_handler![
            read_file_bytes,
            save_bytes,
            list_fonts
        ])
        .run(tauri::generate_context!())
        .expect("error while running tauri application");
}
