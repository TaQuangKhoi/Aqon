use std::path::PathBuf;
use anyhow::{Result, Context};
use clap::Parser;
use log::{info, error, LevelFilter};
use env_logger::Builder;

mod converter;
mod utils;

/// A CLI tool for batch conversion of Word and Excel documents to PDF
#[derive(Parser, Debug)]
#[clap(author, version, about, long_about = None)]
struct Args {
    /// Input directory containing documents to convert
    #[clap(short, long, value_parser)]
    input: PathBuf,

    /// Output directory for generated PDFs
    #[clap(short, long, value_parser)]
    output: PathBuf,

    /// Enable verbose logging
    #[clap(short, long, action)]
    verbose: bool,
}

fn main() -> Result<()> {
    // Parse command line arguments
    let args = Args::parse();

    // Initialize logger
    let mut builder = Builder::new();
    builder.filter_level(if args.verbose {
        LevelFilter::Debug
    } else {
        LevelFilter::Info
    });
    builder.init();

    info!("Starting Aqon document converter");

    // Validate and resolve paths
    let input_dir = utils::resolve_path(&args.input)
        .context("Failed to resolve input directory path")?;
    let output_dir = utils::resolve_path(&args.output)
        .context("Failed to resolve output directory path")?;

    // Validate input directory
    utils::validate_directory(&input_dir)
        .context("Invalid input directory")?;

    // Ensure output directory exists
    utils::ensure_dir_exists(&output_dir)
        .context("Failed to create output directory")?;

    info!("Input directory: {}", input_dir.display());
    info!("Output directory: {}", output_dir.display());

    // Perform batch conversion
    match converter::batch_convert(&input_dir, &output_dir) {
        Ok(converted_files) => {
            if converted_files.is_empty() {
                info!("No files were converted. Check if the input directory contains supported documents.");
            } else {
                info!("Successfully converted {} files:", converted_files.len());
                for file in &converted_files {
                    info!("  - {}", file.display());
                }
            }
        },
        Err(err) => {
            error!("Batch conversion failed: {}", err);
            return Err(err);
        }
    }

    info!("Conversion completed successfully");
    Ok(())
}
