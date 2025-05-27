//! Module for generating Markdown files from extracted document content.

use std::path::{Path, PathBuf};
use std::fs::File;
use std::io::Write;
use anyhow::{Result, Context};
use log::{info, debug};

use crate::converter::docx_reader::DocxContent;
use crate::converter::xlsx_reader::Sheet;

/// Creates a Markdown file from Word document content
///
/// # Arguments
///
/// * `content` - The extracted content from a Word document
/// * `input_path` - Path to the original Word document
/// * `output_dir` - Directory where the Markdown will be saved
///
/// # Returns
///
/// * `Result<PathBuf>` - Path to the generated Markdown file or an error
pub fn create_markdown_from_docx(
    content: &DocxContent,
    input_path: &Path,
    output_dir: &Path,
) -> Result<PathBuf> {
    let output_filename = generate_output_filename(input_path, output_dir)?;
    info!("Creating Markdown from Word document: {}", output_filename.display());

    let mut markdown_content = String::new();

    // Add title based on filename
    if let Some(file_stem) = input_path.file_stem() {
        let title = file_stem.to_string_lossy();
        markdown_content.push_str(&format!("# {}\n\n", title));
    }

    // Add paragraphs
    for paragraph in &content.paragraphs {
        markdown_content.push_str(&format!("{}\n\n", paragraph));
    }

    // Add tables
    for table_data in &content.tables {
        if !table_data.is_empty() {
            // Create table header based on first row
            if let Some(first_row) = table_data.first() {
                // Table header
                markdown_content.push_str("|");
                for cell in first_row {
                    markdown_content.push_str(&format!(" {} |", cell));
                }
                markdown_content.push_str("\n|");

                // Table separator
                for _ in first_row {
                    markdown_content.push_str(" --- |");
                }
                markdown_content.push_str("\n");

                // Table rows (skip first row if it was used as header)
                let data_rows = if table_data.len() > 1 { &table_data[1..] } else { &[] };
                for row in data_rows {
                    markdown_content.push_str("|");
                    for cell in row {
                        // Escape pipe characters in cell content
                        let escaped_cell = cell.replace("|", "\\|");
                        markdown_content.push_str(&format!(" {} |", escaped_cell));
                    }
                    markdown_content.push_str("\n");
                }
                markdown_content.push_str("\n");
            }
        }
    }

    // Write to file
    let mut file = File::create(&output_filename)
        .context(format!("Failed to create Markdown file: {}", output_filename.display()))?;
    
    file.write_all(markdown_content.as_bytes())
        .context(format!("Failed to write to Markdown file: {}", output_filename.display()))?;

    info!("Successfully created Markdown: {}", output_filename.display());
    Ok(output_filename)
}

/// Creates a Markdown file from Excel spreadsheet content
///
/// # Arguments
///
/// * `sheets` - The extracted sheets from an Excel workbook
/// * `input_path` - Path to the original Excel file
/// * `output_dir` - Directory where the Markdown will be saved
///
/// # Returns
///
/// * `Result<PathBuf>` - Path to the generated Markdown file or an error
pub fn create_markdown_from_xlsx(
    sheets: &[Sheet],
    input_path: &Path,
    output_dir: &Path,
) -> Result<PathBuf> {
    let output_filename = generate_output_filename(input_path, output_dir)?;
    info!("Creating Markdown from Excel spreadsheet: {}", output_filename.display());

    let mut markdown_content = String::new();

    // Add title based on filename
    if let Some(file_stem) = input_path.file_stem() {
        let title = file_stem.to_string_lossy();
        markdown_content.push_str(&format!("# {}\n\n", title));
    }

    // Process each sheet
    for (i, sheet) in sheets.iter().enumerate() {
        // Add sheet name as heading
        markdown_content.push_str(&format!("## Sheet: {}\n\n", sheet.name));

        if !sheet.data.is_empty() {
            // Create table header based on first row
            if let Some(first_row) = sheet.data.first() {
                // Table header
                markdown_content.push_str("|");
                for cell in first_row {
                    // Escape pipe characters in cell content
                    let escaped_cell = cell.replace("|", "\\|");
                    markdown_content.push_str(&format!(" {} |", escaped_cell));
                }
                markdown_content.push_str("\n|");

                // Table separator
                for _ in first_row {
                    markdown_content.push_str(" --- |");
                }
                markdown_content.push_str("\n");

                // Table rows (skip first row if it was used as header)
                let data_rows = if sheet.data.len() > 1 { &sheet.data[1..] } else { &[] };
                for row in data_rows {
                    markdown_content.push_str("|");
                    for cell in row {
                        // Escape pipe characters in cell content
                        let escaped_cell = cell.replace("|", "\\|");
                        markdown_content.push_str(&format!(" {} |", escaped_cell));
                    }
                    markdown_content.push_str("\n");
                }
                markdown_content.push_str("\n");
            }
        } else {
            markdown_content.push_str("*(Empty sheet)*\n\n");
        }

        // Add separator between sheets (except for the last one)
        if i < sheets.len() - 1 {
            markdown_content.push_str("---\n\n");
        }
    }

    // Write to file
    let mut file = File::create(&output_filename)
        .context(format!("Failed to create Markdown file: {}", output_filename.display()))?;
    
    file.write_all(markdown_content.as_bytes())
        .context(format!("Failed to write to Markdown file: {}", output_filename.display()))?;

    info!("Successfully created Markdown: {}", output_filename.display());
    Ok(output_filename)
}

/// Generates an output filename for the Markdown based on the input file
///
/// # Arguments
///
/// * `input_path` - Path to the input document
/// * `output_dir` - Directory where the Markdown will be saved
///
/// # Returns
///
/// * `Result<PathBuf>` - The generated output path or an error
fn generate_output_filename(input_path: &Path, output_dir: &Path) -> Result<PathBuf> {
    let file_stem = input_path.file_stem()
        .context("Failed to get file name")?;

    // Create a new path with the file stem and .md extension
    let mut output_filename = output_dir.join(file_stem);
    output_filename.set_extension("md");

    debug!("Generated output filename: {}", output_filename.display());

    Ok(output_filename)
}