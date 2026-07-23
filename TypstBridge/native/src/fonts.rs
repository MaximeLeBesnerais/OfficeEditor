use std::collections::HashMap;
use std::fs;
use std::path::{Path, PathBuf};
use std::sync::{Arc, Mutex, OnceLock};

use typst::foundations::Bytes;
use typst::text::{Font, FontBook};
use typst_utils::LazyHash;

pub struct BridgeFonts {
    pub book: LazyHash<FontBook>,
    pub fonts: Vec<Font>,
}

/// Process-wide cache of parsed font sets.
///
/// Keyed by the caller-supplied font-path list (sorted) plus the working
/// directory and the embedded-fonts flag, so a cache hit skips directory
/// scanning, `fs::read` of font files, AND `Font` parsing entirely.
///
/// Tradeoff: the cache is never evicted. In practice its size is bounded by
/// the number of distinct font-path lists a process compiles with (one per
/// deck/profile in the preview server), and each entry is shared by `Arc`, so
/// steady-state memory stays flat. Entries that became unreachable on disk are
/// still served; callers that need invalidation must restart the process.
static FONT_CACHE: OnceLock<Mutex<HashMap<String, Arc<BridgeFonts>>>> = OnceLock::new();

fn font_cache() -> &'static Mutex<HashMap<String, Arc<BridgeFonts>>> {
    FONT_CACHE.get_or_init(|| Mutex::new(HashMap::new()))
}

/// `include_system_fonts` stays disabled for deterministic output; the flag is
/// part of the cache key so a future caller-facing toggle cannot collide with
/// entries built under a different flag state.
const INCLUDE_SYSTEM_FONTS: bool = false;

pub fn load_font_paths(paths: &[String], working_dir: &Path) -> Result<Arc<BridgeFonts>, String> {
    let key = cache_key(paths, working_dir);
    if let Some(cached) = lock_cache().get(&key) {
        return Ok(Arc::clone(cached));
    }

    let loaded = Arc::new(load_font_paths_uncached(paths, working_dir)?);
    let mut cache = lock_cache();
    let cached = cache.entry(key).or_insert_with(|| Arc::clone(&loaded));
    Ok(Arc::clone(cached))
}

fn cache_key(paths: &[String], working_dir: &Path) -> String {
    let mut sorted = paths.to_vec();
    sorted.sort();

    let mut key = String::new();
    key.push_str(if INCLUDE_SYSTEM_FONTS { "sys:1" } else { "sys:0" });
    key.push('\u{1f}');
    key.push_str(&working_dir.to_string_lossy());
    for path in &sorted {
        key.push('\u{1f}');
        key.push_str(path);
    }
    key
}

fn lock_cache() -> std::sync::MutexGuard<'static, HashMap<String, Arc<BridgeFonts>>> {
    font_cache()
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner())
}

fn load_font_paths_uncached(paths: &[String], working_dir: &Path) -> Result<BridgeFonts, String> {
    let mut book = FontBook::new();
    let mut fonts = Vec::new();

    for path in collect_font_files(paths, working_dir)? {
        let data = fs::read(&path)
            .map_err(|error| format!("Failed to read font '{}': {error}", path.display()))?;
        let bytes = Bytes::new(data);
        for font in Font::iter(bytes) {
            book.push(font.info().clone());
            fonts.push(font);
        }
    }

    // Keep system fonts disabled for deterministic behavior, but include Typst's
    // embedded fallback fonts so minimal documents can compile without a font path.
    for (font, info) in typst_kit::fonts::embedded() {
        book.push(info);
        fonts.push(font);
    }

    Ok(BridgeFonts {
        book: LazyHash::new(book),
        fonts,
    })
}

fn collect_font_files(paths: &[String], working_dir: &Path) -> Result<Vec<PathBuf>, String> {
    let mut files = Vec::new();
    for path in paths {
        let path = PathBuf::from(path);
        let path = if path.is_absolute() {
            path
        } else {
            working_dir.join(path)
        };
        let metadata = fs::symlink_metadata(&path)
            .map_err(|error| format!("Invalid font path '{}': {error}", path.display()))?;
        if metadata.file_type().is_symlink() {
            continue;
        } else if metadata.is_file() {
            files.push(path);
        } else if metadata.is_dir() {
            collect_font_dir(&path, &mut files)?;
        } else {
            return Err(format!(
                "Invalid font path '{}': expected file or directory",
                path.display()
            ));
        }
    }

    files.sort();
    Ok(files)
}

fn collect_font_dir(dir: &Path, files: &mut Vec<PathBuf>) -> Result<(), String> {
    let mut entries = fs::read_dir(dir)
        .map_err(|error| format!("Failed to read font directory '{}': {error}", dir.display()))?
        .collect::<Result<Vec<_>, _>>()
        .map_err(|error| format!("Failed to read font directory '{}': {error}", dir.display()))?;

    entries.sort_by_key(|entry| entry.path());
    for entry in entries {
        let path = entry.path();
        let metadata = fs::symlink_metadata(&path)
            .map_err(|error| format!("Failed to read font path '{}': {error}", path.display()))?;
        if metadata.file_type().is_symlink() {
            continue;
        } else if metadata.is_file() {
            files.push(path);
        } else if metadata.is_dir() {
            collect_font_dir(&path, files)?;
        }
    }

    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::time::{SystemTime, UNIX_EPOCH};

    fn temp_dir(name: &str) -> PathBuf {
        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let dir = std::env::temp_dir().join(format!("typst-bridge-fonts-{name}-{unique}"));
        fs::create_dir(&dir).unwrap();
        dir
    }

    #[cfg(unix)]
    #[test]
    fn collect_font_files_skips_symlinked_files_and_dirs() {
        use std::os::unix::fs::symlink;

        let root = temp_dir("root");
        let outside = temp_dir("outside");
        fs::write(root.join("local.ttf"), b"local").unwrap();
        fs::write(outside.join("outside.ttf"), b"outside").unwrap();

        if symlink(outside.join("outside.ttf"), root.join("linked.ttf")).is_err()
            || symlink(&outside, root.join("linked-dir")).is_err()
        {
            fs::remove_dir_all(&root).unwrap();
            fs::remove_dir_all(&outside).unwrap();
            return;
        }

        let files = collect_font_files(&[root.to_string_lossy().to_string()], Path::new(""))
            .expect("font files");
        assert_eq!(files, vec![root.join("local.ttf")]);

        fs::remove_dir_all(root).unwrap();
        fs::remove_dir_all(outside).unwrap();
    }

    #[test]
    fn relative_font_paths_are_resolved_against_working_dir() {
        let root = temp_dir("relative");
        fs::write(root.join("local.ttf"), b"local").unwrap();

        let files = collect_font_files(&["local.ttf".to_owned()], &root).expect("font files");
        assert_eq!(files, vec![root.join("local.ttf")]);

        fs::remove_dir_all(root).unwrap();
    }

    #[test]
    fn load_font_paths_memoizes_parsed_font_set() {
        let first = load_font_paths(&[], Path::new(".")).expect("first load");
        let second = load_font_paths(&[], Path::new(".")).expect("second load");

        // Cache hit returns the same shared entry: no fs::read or Font parsing.
        assert!(Arc::ptr_eq(&first, &second));
        assert!(!first.fonts.is_empty(), "embedded fallback fonts are present");
    }
}
