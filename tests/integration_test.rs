//! Integration tests for the Aqon document converter
//! These tests verify the main functionality of the application

use std::path::Path;
use anyhow::Result;

// Import the common test utilities
mod common;

// Import the crate to test
use Aqon::converter;
use Aqon::utils;

#[test]
fn test_convert_docx_to_pdf() -> Result<()> {
    // Set up test environment
    let (temp_dir, input_dir, output_dir) = common::setup_test_env()?;
    
    // Create a mock DOCX file
    let docx_path = common::create_mock_docx(&input_dir, "test_document")?;
    
    // Convert the DOCX to PDF
    let result = converter::convert_to_pdf(&docx_path, &output_dir);
    
    // Verify the result
    assert!(result.is_ok(), "Conversion failed: {:?}", result.err());
    let pdf_path = result.unwrap();
    assert!(pdf_path.exists(), "PDF file does not exist");
    assert_eq!(pdf_path.extension().unwrap(), "pdf", "Output file is not a PDF");
    
    // Verify the output file name
    let expected_name = docx_path.file_stem().unwrap().to_string_lossy();
    let actual_name = pdf_path.file_stem().unwrap().to_string_lossy();
    assert_eq!(actual_name, expected_name, "Output file name does not match input file name");
    
    Ok(())
}

#[test]
fn test_convert_xlsx_to_pdf() -> Result<()> {
    // Set up test environment
    let (temp_dir, input_dir, output_dir) = common::setup_test_env()?;
    
    // Create a mock XLSX file
    let xlsx_path = common::create_mock_xlsx(&input_dir, "test_spreadsheet")?;
    
    // Convert the XLSX to PDF
    let result = converter::convert_to_pdf(&xlsx_path, &output_dir);
    
    // Verify the result
    assert!(result.is_ok(), "Conversion failed: {:?}", result.err());
    let pdf_path = result.unwrap();
    assert!(pdf_path.exists(), "PDF file does not exist");
    assert_eq!(pdf_path.extension().unwrap(), "pdf", "Output file is not a PDF");
    
    // Verify the output file name
    let expected_name = xlsx_path.file_stem().unwrap().to_string_lossy();
    let actual_name = pdf_path.file_stem().unwrap().to_string_lossy();
    assert_eq!(actual_name, expected_name, "Output file name does not match input file name");
    
    Ok(())
}

#[test]
fn test_convert_docx_to_markdown() -> Result<()> {
    // Set up test environment
    let (temp_dir, input_dir, output_dir) = common::setup_test_env()?;
    
    // Create a mock DOCX file
    let docx_path = common::create_mock_docx(&input_dir, "test_document")?;
    
    // Convert the DOCX to Markdown
    let result = converter::convert_to_markdown(&docx_path, &output_dir);
    
    // Verify the result
    assert!(result.is_ok(), "Conversion failed: {:?}", result.err());
    let md_path = result.unwrap();
    assert!(md_path.exists(), "Markdown file does not exist");
    assert_eq!(md_path.extension().unwrap(), "md", "Output file is not a Markdown file");
    
    // Verify the output file name
    let expected_name = docx_path.file_stem().unwrap().to_string_lossy();
    let actual_name = md_path.file_stem().unwrap().to_string_lossy();
    assert_eq!(actual_name, expected_name, "Output file name does not match input file name");
    
    Ok(())
}

#[test]
fn test_batch_convert() -> Result<()> {
    // Set up test environment
    let (temp_dir, input_dir, output_dir) = common::setup_test_env()?;
    
    // Create multiple mock files
    let docx_path1 = common::create_mock_docx(&input_dir, "document1")?;
    let docx_path2 = common::create_mock_docx(&input_dir, "document2")?;
    let xlsx_path = common::create_mock_xlsx(&input_dir, "spreadsheet")?;
    
    // Perform batch conversion
    let result = converter::batch_convert(&input_dir, &output_dir);
    
    // Verify the result
    assert!(result.is_ok(), "Batch conversion failed: {:?}", result.err());
    let pdf_paths = result.unwrap();
    
    // We should have 3 PDF files
    assert_eq!(pdf_paths.len(), 3, "Expected 3 PDF files, got {}", pdf_paths.len());
    
    // Verify each output file
    common::verify_pdf_output(&output_dir, "document1")?;
    common::verify_pdf_output(&output_dir, "document2")?;
    common::verify_pdf_output(&output_dir, "spreadsheet")?;
    
    Ok(())
}

#[test]
fn test_utils_directory_functions() -> Result<()> {
    // Set up test environment
    let (temp_dir, input_dir, output_dir) = common::setup_test_env()?;
    
    // Test ensure_dir_exists
    let new_dir = temp_dir.path().join("new_directory");
    utils::ensure_dir_exists(&new_dir)?;
    assert!(new_dir.exists(), "Directory was not created");
    
    // Test validate_directory
    let result = utils::validate_directory(&input_dir);
    assert!(result.is_ok(), "Valid directory failed validation");
    
    let invalid_dir = temp_dir.path().join("non_existent_dir");
    let result = utils::validate_directory(&invalid_dir);
    assert!(result.is_err(), "Invalid directory passed validation");
    
    // Test resolve_path
    let relative_path = Path::new("relative/path");
    let absolute_path = utils::resolve_path(relative_path)?;
    assert!(absolute_path.is_absolute(), "Path was not resolved to absolute");
    
    Ok(())
}

#[test]
fn test_supported_file_extensions() -> Result<()> {
    // Set up test environment
    let (temp_dir, input_dir, _) = common::setup_test_env()?;
    
    // Create files with different extensions
    let docx_path = common::create_mock_docx(&input_dir, "document")?;
    let xlsx_path = common::create_mock_xlsx(&input_dir, "spreadsheet")?;
    let unsupported_path = input_dir.join("unsupported.txt");
    std::fs::write(&unsupported_path, b"Unsupported file")?;
    
    // Test is_supported_file
    assert!(utils::is_supported_file(&docx_path), "DOCX file should be supported");
    assert!(utils::is_supported_file(&xlsx_path), "XLSX file should be supported");
    assert!(!utils::is_supported_file(&unsupported_path), "TXT file should not be supported");
    
    // Test get_supported_extensions
    let extensions = utils::get_supported_extensions();
    assert!(extensions.contains(&"docx"), "DOCX should be in supported extensions");
    assert!(extensions.contains(&"xlsx"), "XLSX should be in supported extensions");
    assert!(extensions.contains(&"xls"), "XLS should be in supported extensions");
    
    Ok(())
}