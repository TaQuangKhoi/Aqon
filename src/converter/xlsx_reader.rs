//! Module for reading and extracting data from Excel (.xlsx/.xls) spreadsheets.

use std::path::Path;
use anyhow::{Result, Context};
use calamine::{Reader, open_workbook, Xlsx, Range, DataType};
use log::{info, debug, warn};

/// Represents a sheet in an Excel workbook
#[derive(Debug)]
pub struct Sheet {
    /// Name of the sheet
    pub name: String,
    /// Data in the sheet (rows and columns)
    pub data: Vec<Vec<String>>,
}

/// Extracts data from all sheets in an Excel workbook
///
/// # Arguments
///
/// * `path` - Path to the Excel file
///
/// # Returns
///
/// * `Result<Vec<Sheet>>` - Vector of extracted sheets or an error
pub fn extract_sheets(path: &Path) -> Result<Vec<Sheet>> {
    info!("Extracting data from Excel file: {}", path.display());
    
    let mut workbook: Xlsx<_> = open_workbook(path)
        .context(format!("Failed to open Excel file: {}", path.display()))?;
    
    let sheet_names = workbook.sheet_names().to_vec();
    info!("Found {} sheets in workbook", sheet_names.len());
    
    let mut sheets = Vec::new();
    
    for sheet_name in sheet_names {
        debug!("Processing sheet: {}", sheet_name);
        
        if let Some(Ok(range)) = workbook.worksheet_range(&sheet_name) {
            let sheet_data = process_range(range);
            
            if !sheet_data.is_empty() {
                debug!("Extracted {} rows from sheet '{}'", sheet_data.len(), sheet_name);
                sheets.push(Sheet {
                    name: sheet_name,
                    data: sheet_data,
                });
            } else {
                warn!("Sheet '{}' appears to be empty", sheet_name);
            }
        } else {
            warn!("Failed to read sheet: {}", sheet_name);
        }
    }
    
    if sheets.is_empty() {
        warn!("No data extracted from Excel file");
    } else {
        info!("Successfully extracted data from {} sheets", sheets.len());
    }
    
    Ok(sheets)
}

/// Processes a range of cells from an Excel sheet
///
/// # Arguments
///
/// * `range` - The range of cells to process
///
/// # Returns
///
/// * `Vec<Vec<String>>` - The processed data as rows of strings
fn process_range(range: Range<DataType>) -> Vec<Vec<String>> {
    let height = range.height();
    let width = range.width();
    
    if height == 0 || width == 0 {
        return Vec::new();
    }
    
    debug!("Processing range with dimensions: {}x{}", width, height);
    
    let mut data = Vec::with_capacity(height);
    
    for row_index in 0..height {
        let mut row = Vec::with_capacity(width);
        
        for col_index in 0..width {
            let cell_value = match range.get_value((row_index as u32, col_index as u32)) {
                Some(value) => value.to_string(),
                None => String::new(),
            };
            
            row.push(cell_value);
        }
        
        // Skip completely empty rows
        if row.iter().any(|cell| !cell.is_empty()) {
            data.push(row);
        }
    }
    
    data
}