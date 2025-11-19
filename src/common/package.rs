use crate::error::{OfficeError, Result};
use std::fs::File;
use std::io::{Read, Seek};
use std::path::Path;
use zip::ZipArchive;

/// Represents an open Office package (a zip archive).
pub struct OfficePackage<R: Read + Seek> {
    pub archive: ZipArchive<R>,
}

impl<R: Read + Seek> OfficePackage<R> {
    /// Creates a new OfficePackage from a reader.
    pub fn new(reader: R) -> Result<Self> {
        let archive = ZipArchive::new(reader)?;
        Ok(OfficePackage { archive })
    }

    /// Reads a file from the package by its path.
    pub fn read_file_by_path(&mut self, file_path: &str) -> Result<String> {
        let mut file = self
            .archive
            .by_name(file_path)
            .map_err(|_| OfficeError::FileNotFoundInArchive(file_path.to_string()))?;

        let mut content = String::new();
        file.read_to_string(&mut content)?;

        Ok(content)
    }

    /// Checks if a file exists in the package.
    pub fn has_file(&mut self, file_path: &str) -> bool {
        self.archive.by_name(file_path).is_ok()
    }
}

/// Opens an Office file from the given path.
pub fn open_package<P: AsRef<Path>>(path: P) -> Result<OfficePackage<File>> {
    let file = File::open(path)?;
    OfficePackage::new(file)
}