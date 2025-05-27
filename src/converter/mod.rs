//! Converter module for handling document conversions.
//! This module contains functionality for converting various document formats to PDF.

pub mod docx_reader;
pub mod xlsx_reader;
pub mod pdf_writer;

use std::path::{Path, PathBuf};
use anyhow::{Result, Context};
use log::{info, error};

/// Converts a document to PDF format.
/// 
/// # Arguments
/// 
/// * `input_path` - Path to the input document
/// * `output_dir` - Directory where the output PDF will be saved
/// 
/// # Returns
/// 
/// * `Result<PathBuf>` - Path to the generated PDF file or an error
pub fn convert_to_pdf(input_path: &Path, output_dir: &Path) -> Result<PathBuf> {
    let file_name = input_path.file_name()
        .context("Failed to get file name")?
        .to_string_lossy();
    
    let extension = input_path.extension()
        .map(|ext| ext.to_string_lossy().to_lowercase())
        .unwrap_or_default();
    
    info!("Converting file: {}", file_name);
    
    let result = match extension.as_ref() {
        "docx" => {
            info!("Detected Word document");
            let content = docx_reader::extract_content(input_path)?;
            pdf_writer::create_pdf_from_docx(&content, input_path, output_dir)?
        },
        "xlsx" | "xls" => {
            info!("Detected Excel spreadsheet");
            let sheets = xlsx_reader::extract_sheets(input_path)?;
            pdf_writer::create_pdf_from_xlsx(&sheets, input_path, output_dir)?
        },
        _ => {
            error!("Unsupported file format: {}", extension);
            anyhow::bail!("Unsupported file format: {}", extension);
        }
    };
    
    info!("Successfully converted {} to PDF", file_name);
    Ok(result)
}

/// Batch converts all supported documents in a directory to PDF.
/// 
/// # Arguments
/// 
/// * `input_dir` - Directory containing documents to convert
/// * `output_dir` - Directory where the output PDFs will be saved
/// 
/// # Returns
/// 
/// * `Result<Vec<PathBuf>>` - Paths to the generated PDF files or an error
pub fn batch_convert(input_dir: &Path, output_dir: &Path) -> Result<Vec<PathBuf>> {
    info!("Starting batch conversion from {} to {}", 
          input_dir.display(), output_dir.display());
    
    // Create output directory if it doesn't exist
    if !output_dir.exists() {
        std::fs::create_dir_all(output_dir)
            .context("Failed to create output directory")?;
    }
    
    let mut results = Vec::new();
    
    // Walk through the input directory
    for entry in walkdir::WalkDir::new(input_dir)
        .follow_links(true)
        .into_iter()
        .filter_map(|e| e.ok()) {
            
        let path = entry.path();
        
        // Skip directories
        if path.is_dir() {
            continue;
        }
        
        // Check if file extension is supported
        if let Some(ext) = path.extension() {
            let ext_str = ext.to_string_lossy().to_lowercase();
            if ext_str == "docx" || ext_str == "xlsx" || ext_str == "xls" {
                match convert_to_pdf(path, output_dir) {
                    Ok(pdf_path) => {
                        results.push(pdf_path);
                    },
                    Err(err) => {
                        error!("Failed to convert {}: {}", path.display(), err);
                    }
                }
            }
        }
    }
    
    info!("Batch conversion completed. Converted {} files.", results.len());
    Ok(results)
}