//! Common utilities for integration tests
//! This file contains shared functions and mock implementations for tests

use std::fs;
use std::path::{Path, PathBuf};
use tempfile::TempDir;
use anyhow::Result;

/// Creates a temporary test environment with input and output directories
pub fn setup_test_env() -> Result<(TempDir, PathBuf, PathBuf)> {
    let temp_dir = tempfile::tempdir()?;
    let input_dir = temp_dir.path().join("input");
    let output_dir = temp_dir.path().join("output");
    
    fs::create_dir_all(&input_dir)?;
    fs::create_dir_all(&output_dir)?;
    
    Ok((temp_dir, input_dir, output_dir))
}

/// Creates a mock DOCX file for testing
pub fn create_mock_docx(dir: &Path, filename: &str) -> Result<PathBuf> {
    let file_path = dir.join(format!("{}.docx", filename));
    fs::write(&file_path, b"Mock DOCX content")?;
    Ok(file_path)
}

/// Creates a mock XLSX file for testing
pub fn create_mock_xlsx(dir: &Path, filename: &str) -> Result<PathBuf> {
    let file_path = dir.join(format!("{}.xlsx", filename));
    fs::write(&file_path, b"Mock XLSX content")?;
    Ok(file_path)
}

/// Verifies that a PDF file exists with the expected name
pub fn verify_pdf_output(output_dir: &Path, expected_name: &str) -> Result<PathBuf> {
    let pdf_path = output_dir.join(format!("{}.pdf", expected_name));
    if !pdf_path.exists() {
        anyhow::bail!("Expected PDF file does not exist: {}", pdf_path.display());
    }
    Ok(pdf_path)
}

/// Verifies that a Markdown file exists with the expected name
pub fn verify_markdown_output(output_dir: &Path, expected_name: &str) -> Result<PathBuf> {
    let md_path = output_dir.join(format!("{}.md", expected_name));
    if !md_path.exists() {
        anyhow::bail!("Expected Markdown file does not exist: {}", md_path.display());
    }
    Ok(md_path)
}