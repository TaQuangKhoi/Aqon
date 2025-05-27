//! Module for generating PDF files from extracted document content.

use std::path::{Path, PathBuf};
use anyhow::{Result, Context};
use genpdf::{elements, fonts, style, Element};
use log::{info, debug, warn};

use crate::converter::docx_reader::DocxContent;
use crate::converter::xlsx_reader::Sheet;

/// Default font to use in generated PDFs
const DEFAULT_FONT_NAME: &str = "Roboto";

/// Creates a PDF file from Word document content
///
/// # Arguments
///
/// * `content` - The extracted content from a Word document
/// * `input_path` - Path to the original Word document
/// * `output_dir` - Directory where the PDF will be saved
///
/// # Returns
///
/// * `Result<PathBuf>` - Path to the generated PDF file or an error
pub fn create_pdf_from_docx(
    content: &DocxContent,
    input_path: &Path,
    output_dir: &Path,
) -> Result<PathBuf> {
    let output_filename = generate_output_filename(input_path, output_dir)?;
    info!("Creating PDF from Word document: {}", output_filename.display());
    
    // Load default font
    let font_family = load_default_font()?;
    
    // Create PDF document
    let mut doc = genpdf::Document::new(font_family);
    doc.set_title(input_path.file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "Converted Document".to_string()));
    
    // Set default margins
    let mut decorator = genpdf::SimplePageDecorator::new();
    decorator.set_margins(20);
    doc.set_page_decorator(decorator);
    
    // Add paragraphs
    for paragraph in &content.paragraphs {
        doc.push(elements::Paragraph::new(paragraph));
        doc.push(elements::Break::new(1));
    }
    
    // Add tables
    for table_data in &content.tables {
        if !table_data.is_empty() {
            // Determine column widths based on the first row
            if let Some(first_row) = table_data.first() {
                let col_count = first_row.len();
                let widths = vec![1; col_count]; // Equal width for all columns
                let mut table = elements::TableLayout::new(widths);
                
                // Add table data
                for row in table_data {
                    // Create a vector of cell elements first
                    let cell_elements: Vec<elements::Paragraph> = row.iter()
                        .map(|cell| elements::Paragraph::new(cell))
                        .collect();
                    
                    // Then add them to a row
                    let mut table_row = table.row();
                    for cell_element in cell_elements {
                        table_row = table_row.element(cell_element);
                    }
                    table_row.push().unwrap();
                }
                
                doc.push(table);
                doc.push(elements::Break::new(1));
            }
        }
    }
    
    // Generate PDF
    doc.render_to_file(&output_filename)
        .context(format!("Failed to generate PDF file: {}", output_filename.display()))?;
    
    info!("Successfully created PDF: {}", output_filename.display());
    Ok(output_filename)
}

/// Creates a PDF file from Excel spreadsheet content
///
/// # Arguments
///
/// * `sheets` - The extracted sheets from an Excel workbook
/// * `input_path` - Path to the original Excel file
/// * `output_dir` - Directory where the PDF will be saved
///
/// # Returns
///
/// * `Result<PathBuf>` - Path to the generated PDF file or an error
pub fn create_pdf_from_xlsx(
    sheets: &[Sheet],
    input_path: &Path,
    output_dir: &Path,
) -> Result<PathBuf> {
    let output_filename = generate_output_filename(input_path, output_dir)?;
    info!("Creating PDF from Excel spreadsheet: {}", output_filename.display());
    
    // Load default font
    let font_family = load_default_font()?;
    
    // Create PDF document
    let mut doc = genpdf::Document::new(font_family);
    doc.set_title(input_path.file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "Converted Spreadsheet".to_string()));
    
    // Set default margins
    let mut decorator = genpdf::SimplePageDecorator::new();
    decorator.set_margins(20);
    doc.set_page_decorator(decorator);
    
    // Process each sheet
    for (i, sheet) in sheets.iter().enumerate() {
        // Add sheet name as heading
        let sheet_title = format!("Sheet: {}", sheet.name);
        let heading = elements::Paragraph::new(&sheet_title)
            .styled(style::Style::new().bold());
        doc.push(heading);
        doc.push(elements::Break::new(1));
        
        if !sheet.data.is_empty() {
            // Determine column widths based on the first row
            let col_count = sheet.data.first().map_or(0, |row| row.len());
            if col_count > 0 {
                let widths = vec![1; col_count]; // Equal width for all columns
                let mut table = elements::TableLayout::new(widths);
                
                // Add table data
                for row in &sheet.data {
                    // Create a vector of cell elements first
                    let cell_elements: Vec<elements::Paragraph> = row.iter()
                        .map(|cell| elements::Paragraph::new(cell))
                        .collect();
                    
                    // Then add them to a row
                    let mut table_row = table.row();
                    for cell_element in cell_elements {
                        table_row = table_row.element(cell_element);
                    }
                    table_row.push().unwrap();
                }
                
                doc.push(table);
            }
        } else {
            doc.push(elements::Paragraph::new("(Empty sheet)"));
        }
        
        // Add page break between sheets (except for the last one)
        if i < sheets.len() - 1 {
            doc.push(elements::PageBreak::new());
        }
    }
    
    // Generate PDF
    doc.render_to_file(&output_filename)
        .context(format!("Failed to generate PDF file: {}", output_filename.display()))?;
    
    info!("Successfully created PDF: {}", output_filename.display());
    Ok(output_filename)
}

/// Loads the default font for PDF generation
///
/// # Returns
///
/// * `Result<fonts::FontFamily<fonts::FontData>>` - The loaded font family or an error
fn load_default_font() -> Result<fonts::FontFamily<fonts::FontData>> {
    debug!("Loading default font: {}", DEFAULT_FONT_NAME);
    
    // Try to load custom fonts, but fall back to built-in fonts if that fails
    match load_custom_fonts() {
        Ok(fonts) => {
            debug!("Successfully loaded custom fonts");
            Ok(fonts)
        },
        Err(err) => {
            debug!("Failed to load custom fonts: {}. Using built-in fonts instead.", err);
            // Create a basic font family with empty data
            // In a real implementation, we would use a built-in font
            let empty_data = fonts::FontData::new(Vec::new(), None)
                .context("Failed to create empty font data")?;
            
            let font_family = fonts::FontFamily {
                regular: empty_data.clone(),
                bold: empty_data.clone(),
                italic: empty_data.clone(),
                bold_italic: empty_data,
            };
            
            warn!("Using empty font data. In a real implementation, use built-in fonts.");
            Ok(font_family)
        }
    }
}

/// Attempts to load custom fonts from the resources directory
///
/// # Returns
///
/// * `Result<fonts::FontFamily<fonts::FontData>>` - The loaded font family or an error
fn load_custom_fonts() -> Result<fonts::FontFamily<fonts::FontData>> {
    // In a real implementation, these would be actual font files
    // For this example, we're using placeholders, so this will likely fail
    // and the application will fall back to built-in fonts
    
    let font_data = fonts::FontData::new(
        std::fs::read("resources/fonts/Roboto/static/Roboto-Regular.ttf")
            .context("Failed to read regular font file")?,
        None,
    )
    .context("Failed to load regular font data")?;
    
    let font_data_bold = fonts::FontData::new(
        std::fs::read("resources/fonts/Roboto/static/Roboto-Bold.ttf")
            .context("Failed to read bold font file")?,
        None,
    )
    .context("Failed to load bold font data")?;
    
    let font_data_italic = fonts::FontData::new(
        std::fs::read("resources/fonts/Roboto/static/Roboto-Italic.ttf")
            .context("Failed to read italic font file")?,
        None,
    )
    .context("Failed to load italic font data")?;
    
    let font_data_bold_italic = fonts::FontData::new(
        std::fs::read("resources/fonts/Roboto/static/Roboto-BoldItalic.ttf")
            .context("Failed to read bold italic font file")?,
        None,
    )
    .context("Failed to load bold italic font data")?;
    
    let font_family = fonts::FontFamily {
        regular: font_data,
        bold: font_data_bold,
        italic: font_data_italic,
        bold_italic: font_data_bold_italic,
    };
    
    Ok(font_family)
}

/// Generates an output filename for the PDF based on the input file
///
/// # Arguments
///
/// * `input_path` - Path to the input document
/// * `output_dir` - Directory where the PDF will be saved
///
/// # Returns
///
/// * `Result<PathBuf>` - The generated output path or an error
fn generate_output_filename(input_path: &Path, output_dir: &Path) -> Result<PathBuf> {
    let file_stem = input_path.file_stem()
        .context("Failed to get file name")?
        .to_string_lossy();
    
    let output_filename = output_dir.join(format!("{}.pdf", file_stem));
    debug!("Generated output filename: {}", output_filename.display());
    
    Ok(output_filename)
}