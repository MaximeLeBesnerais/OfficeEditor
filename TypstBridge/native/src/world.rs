use std::fs;
use std::path::{Path, PathBuf};
use std::sync::Arc;

use typst::diag::{FileError, FileResult};
use typst::foundations::{Bytes, Datetime};
use typst::syntax::{FileId, Source, VirtualPath};
use typst::text::{Font, FontBook};
use typst::{Library, LibraryExt, World};
use typst_utils::LazyHash;

use crate::fonts::BridgeFonts;

pub struct BridgeWorld {
    library: LazyHash<Library>,
    fonts: Arc<BridgeFonts>,
    main_id: FileId,
    main_source: Source,
    working_dir: PathBuf,
    canonical_working_dir: PathBuf,
}

impl BridgeWorld {
    pub fn new(
        source_text: String,
        working_dir: PathBuf,
        root_name: &str,
        fonts: Arc<BridgeFonts>,
    ) -> Self {
        let canonical_working_dir =
            fs::canonicalize(&working_dir).unwrap_or_else(|_| working_dir.clone());

        let mut world = Self {
            library: LazyHash::new(Library::default()),
            fonts,
            main_id: FileId::new(None, VirtualPath::new("main.typ")),
            main_source: Source::new(
                FileId::new(None, VirtualPath::new("main.typ")),
                String::new(),
            ),
            working_dir,
            canonical_working_dir,
        };
        world.set_source(source_text, root_name);
        world
    }

    /// Replaces the main source in place, keeping the library, font set, and
    /// working-directory resolution hot. Used by persistent sessions so an
    /// edited-slide recompile keeps comemo-memoized layout work reachable.
    pub fn set_source(&mut self, source_text: String, root_name: &str) {
        let root_name = if root_name.is_empty() {
            "main.typ"
        } else {
            root_name
        };
        self.main_id = FileId::new(None, VirtualPath::new(root_name));
        self.main_source = Source::new(self.main_id, source_text);
    }
}

impl World for BridgeWorld {
    fn library(&self) -> &LazyHash<Library> {
        &self.library
    }

    fn book(&self) -> &LazyHash<FontBook> {
        &self.fonts.book
    }

    fn main(&self) -> FileId {
        self.main_id
    }

    fn source(&self, id: FileId) -> FileResult<Source> {
        if id == self.main_id {
            return Ok(self.main_source.clone());
        }

        let path = resolve(&self.working_dir, &self.canonical_working_dir, id)?;
        if path.is_dir() {
            return Err(FileError::IsDirectory);
        }

        let text = fs::read_to_string(&path).map_err(|error| FileError::from_io(error, &path))?;
        Ok(Source::new(id, text))
    }

    fn file(&self, id: FileId) -> FileResult<Bytes> {
        let path = resolve(&self.working_dir, &self.canonical_working_dir, id)?;
        if path.is_dir() {
            return Err(FileError::IsDirectory);
        }

        let data = fs::read(&path).map_err(|error| FileError::from_io(error, &path))?;
        Ok(Bytes::new(data))
    }

    fn font(&self, index: usize) -> Option<Font> {
        self.fonts.fonts.get(index).cloned()
    }

    fn today(&self, _offset: Option<i64>) -> Option<Datetime> {
        Datetime::from_ymd(1970, 1, 1)
    }
}

fn resolve(root: &Path, canonical_root: &Path, id: FileId) -> FileResult<PathBuf> {
    if id.package().is_some() {
        return Err(FileError::NotFound(
            id.vpath().as_rootless_path().to_owned(),
        ));
    }

    let path = id
        .vpath()
        .resolve(root)
        .ok_or_else(|| FileError::NotFound(id.vpath().as_rootless_path().to_owned()))?;

    let canonical_path =
        fs::canonicalize(&path).map_err(|error| FileError::from_io(error, &path))?;
    if !canonical_path.starts_with(canonical_root) {
        return Err(FileError::NotFound(
            id.vpath().as_rootless_path().to_owned(),
        ));
    }

    Ok(canonical_path)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[cfg(unix)]
    #[test]
    fn resolve_rejects_symlink_escape() {
        use std::os::unix::fs::symlink;
        use std::time::{SystemTime, UNIX_EPOCH};

        let unique = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .unwrap()
            .as_nanos();
        let root = std::env::temp_dir().join(format!("typst-bridge-world-root-{unique}"));
        let outside = std::env::temp_dir().join(format!("typst-bridge-world-outside-{unique}"));
        fs::create_dir(&root).unwrap();
        fs::create_dir(&outside).unwrap();
        fs::write(outside.join("secret.txt"), "secret").unwrap();

        if symlink(outside.join("secret.txt"), root.join("link.txt")).is_err() {
            fs::remove_dir_all(&root).unwrap();
            fs::remove_dir_all(&outside).unwrap();
            return;
        }

        let id = FileId::new(None, VirtualPath::new("link.txt"));
        let canonical_root = fs::canonicalize(&root).unwrap();
        assert!(matches!(
            resolve(&root, &canonical_root, id),
            Err(FileError::NotFound(_))
        ));

        fs::remove_dir_all(root).unwrap();
        fs::remove_dir_all(outside).unwrap();
    }
}
