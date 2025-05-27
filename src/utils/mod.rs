//! Utility functions for the application.
//! This module contains helper functions for file path handling and other utilities.

use std::path::{Path, PathBuf};
use anyhow::{Result, Context};
use log::debug;

/// Ensures that a directory exists, creating it if necessary.
///
/// # Arguments
///
/// * `dir_path` - Path to the directory
///
/// # Returns
///
/// * `Result<()>` - Success or an error
pub fn ensure_dir_exists(dir_path: &Path) -> Result<()> {
    if !dir_path.exists() {
        debug!("Creating directory: {}", dir_path.display());
        std::fs::create_dir_all(dir_path)
            .context(format!("Failed to create directory: {}", dir_path.display()))?;
    }
    Ok(())
}

/// Validates that a path exists and is a directory.
///
/// # Arguments
///
/// * `dir_path` - Path to validate
///
/// # Returns
///
/// * `Result<()>` - Success or an error
pub fn validate_directory(dir_path: &Path) -> Result<()> {
    if !dir_path.exists() {
        anyhow::bail!("Directory does not exist: {}", dir_path.display());
    }
    
    if !dir_path.is_dir() {
        anyhow::bail!("Path is not a directory: {}", dir_path.display());
    }
    
    Ok(())
}

/// Resolves a relative path to an absolute path.
///
/// # Arguments
///
/// * `path` - Path to resolve
///
/// # Returns
///
/// * `Result<PathBuf>` - Absolute path or an error
pub fn resolve_path(path: &Path) -> Result<PathBuf> {
    let absolute_path = if path.is_absolute() {
        path.to_path_buf()
    } else {
        std::env::current_dir()
            .context("Failed to get current directory")?
            .join(path)
    };
    
    debug!("Resolved path: {} -> {}", path.display(), absolute_path.display());
    Ok(absolute_path)
}

/// Gets a list of supported file extensions.
///
/// # Returns
///
/// * `Vec<&'static str>` - List of supported file extensions
pub fn get_supported_extensions() -> Vec<&'static str> {
    vec!["docx", "xlsx", "xls"]
}

/// Checks if a file has a supported extension.
///
/// # Arguments
///
/// * `path` - Path to the file
///
/// # Returns
///
/// * `bool` - True if the file has a supported extension
pub fn is_supported_file(path: &Path) -> bool {
    if let Some(extension) = path.extension() {
        let ext_str = extension.to_string_lossy().to_lowercase();
        get_supported_extensions().contains(&ext_str.as_str())
    } else {
        false
    }
}