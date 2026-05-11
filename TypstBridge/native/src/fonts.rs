use std::fs;
use std::path::{Path, PathBuf};

use typst::foundations::Bytes;
use typst::text::{Font, FontBook};
use typst_kit::fonts::FontSearcher;

pub struct BridgeFonts {
    pub book: FontBook,
    pub fonts: Vec<Font>,
}

pub fn load_font_paths(paths: &[String], working_dir: &Path) -> Result<BridgeFonts, String> {
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
    let embedded = FontSearcher::new().include_system_fonts(false).search();
    for slot in embedded.fonts {
        if let Some(font) = slot.get() {
            book.push(font.info().clone());
            fonts.push(font);
        }
    }

    Ok(BridgeFonts { book, fonts })
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
}
