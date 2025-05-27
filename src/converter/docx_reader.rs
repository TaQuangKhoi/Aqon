//! Module for reading and extracting content from Word (.docx) documents.

use std::path::Path;
use anyhow::{Result, Context};
use docx_rs::{DocumentChild, ParagraphChild, RunChild, TableChild, TableRowChild, TableCellContent};
use log::{info, debug, warn};

/// Represents the content extracted from a Word document
#[derive(Debug, Default)]
pub struct DocxContent {
    /// Paragraphs of text from the document
    pub paragraphs: Vec<String>,
    /// Tables extracted from the document
    pub tables: Vec<Vec<Vec<String>>>, // Tables -> Rows -> Cells
}

/// Extracts content from a Word document
///
/// # Arguments
///
/// * `path` - Path to the Word document
///
/// # Returns
///
/// * `Result<DocxContent>` - Extracted content or an error
pub fn extract_content(path: &Path) -> Result<DocxContent> {
    info!("Extracting content from Word document: {}", path.display());

    let buf = std::fs::read(path)
        .context(format!("Failed to read file: {}", path.display()))?;

    let docx = docx_rs::read_docx(&buf)
        .context("Failed to parse DOCX file")?;

    let document = docx.document;
    let mut content = DocxContent::default();

    // Process document body
    for child in document.children {
        match child {
            DocumentChild::Paragraph(paragraph) => {
                let mut paragraph_text = String::new();

                for child in paragraph.children {
                    if let ParagraphChild::Run(run) = child {
                        for child in run.children {
                            if let RunChild::Text(text) = child {
                                paragraph_text.push_str(&text.text);
                            }
                        }
                    }
                }

                if !paragraph_text.trim().is_empty() {
                    debug!("Extracted paragraph: {}", paragraph_text);
                    content.paragraphs.push(paragraph_text);
                }
            },
            DocumentChild::Table(table) => {
                let mut table_data = Vec::new();

                for row_child in &table.rows {
                    if let TableChild::TableRow(row) = row_child {
                        let mut row_data = Vec::new();

                        for cell_child in &row.cells {
                            if let TableRowChild::TableCell(cell) = cell_child {
                                let mut cell_text = String::new();

                                for content in &cell.children {
                                    // Match on TableCellContent instead of DocumentChild
                                    if let docx_rs::TableCellContent::Paragraph(paragraph) = content {
                                        for child in &paragraph.children {
                                            if let ParagraphChild::Run(run) = child {
                                                for child in &run.children {
                                                    if let RunChild::Text(text) = child {
                                                        cell_text.push_str(&text.text);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                row_data.push(cell_text);
                            }
                        }

                        if !row_data.is_empty() {
                            table_data.push(row_data);
                        }
                    }
                }

                if !table_data.is_empty() {
                    debug!("Extracted table with {} rows", table_data.len());
                    content.tables.push(table_data);
                }
            },
            _ => {
                // Ignore other document elements
            }
        }
    }

    info!("Extracted {} paragraphs and {} tables from document", 
          content.paragraphs.len(), content.tables.len());

    if content.paragraphs.is_empty() && content.tables.is_empty() {
        warn!("No content extracted from document");
    }

    Ok(content)
}
