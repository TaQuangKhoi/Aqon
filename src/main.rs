use std::path::PathBuf;
use std::time::Duration;
use std::sync::mpsc;
use anyhow::{Result, Context, anyhow};
use clap::{Parser, Subcommand};
use log::{info, error, warn, debug, LevelFilter};
use env_logger::Builder;
use colored::Colorize;
use indicatif::{ProgressBar, ProgressStyle, MultiProgress};
use notify::{Watcher, RecursiveMode, Event, EventKind, Config};

mod converter;
mod utils;

/// A CLI tool for batch conversion of Word and Excel documents to PDF
#[derive(Parser, Debug)]
#[clap(author, version, about, long_about = None)]
struct Cli {
    #[command(subcommand)]
    command: Commands,

    /// Enable verbose logging
    #[clap(short, long, action, global = true)]
    verbose: bool,
}

#[derive(Subcommand, Debug)]
enum Commands {
    /// Convert documents from input directory to output directory
    Convert {
        /// Input directory containing documents to convert
        #[clap(short, long, value_parser)]
        input: PathBuf,

        /// Output directory for generated PDFs
        #[clap(short, long, value_parser)]
        output: PathBuf,

        /// Only convert files of specified type (docx, xlsx, xls)
        #[clap(short, long, value_parser)]
        r#type: Option<String>,
    },
    /// Watch a directory and automatically convert new documents
    Watch {
        /// Input directory to watch for new documents
        #[clap(short, long, value_parser)]
        input: PathBuf,

        /// Output directory for generated PDFs
        #[clap(short, long, value_parser)]
        output: PathBuf,

        /// Only convert files of specified type (docx, xlsx, xls)
        #[clap(short, long, value_parser)]
        r#type: Option<String>,
    },
}

fn main() -> Result<()> {
    // Parse command line arguments
    let cli = Cli::parse();

    // Initialize logger
    let mut builder = Builder::new();
    builder.filter_level(if cli.verbose {
        LevelFilter::Debug
    } else {
        LevelFilter::Info
    });
    builder.init();

    println!("{}", "Starting Aqon document converter".bright_green());

    match &cli.command {
        Commands::Convert { input, output, r#type } => {
            convert_command(input, output, r#type)?;
        },
        Commands::Watch { input, output, r#type } => {
            watch_command(input, output, r#type)?;
        }
    }

    Ok(())
}

/// Handle the convert command
fn convert_command(input: &PathBuf, output: &PathBuf, file_type: &Option<String>) -> Result<()> {
    // Validate and resolve paths
    let input_dir = utils::resolve_path(input)
        .context("Failed to resolve input directory path")?;
    let output_dir = utils::resolve_path(output)
        .context("Failed to resolve output directory path")?;

    // Validate input directory
    utils::validate_directory(&input_dir)
        .context("Invalid input directory")?;

    // Ensure output directory exists
    utils::ensure_dir_exists(&output_dir)
        .context("Failed to create output directory")?;

    println!("{} {}", "Input directory:".blue(), input_dir.display());
    println!("{} {}", "Output directory:".blue(), output_dir.display());

    if let Some(t) = file_type {
        println!("{} {}", "File type filter:".blue(), t);
    }

    // Get list of files to convert
    let files = get_files_to_convert(&input_dir, file_type)?;

    if files.is_empty() {
        println!("{}", "No files found to convert.".yellow());
        return Ok(());
    }

    println!("{} {} {}", "Found".blue(), files.len(), "files to convert".blue());

    // Create progress bar
    let progress = ProgressBar::new(files.len() as u64);
    progress.set_style(
        ProgressStyle::default_bar()
            .template("{spinner:.green} [{elapsed_precise}] [{bar:40.cyan/blue}] {pos}/{len} ({eta})")
            .unwrap()
            .progress_chars("#>-")
    );

    let mut converted_files = Vec::new();

    for file_path in files {
        let file_name = file_path.file_name().unwrap_or_default().to_string_lossy();
        progress.set_message(format!("Converting {}", file_name));

        match converter::convert_to_pdf(&file_path, &output_dir) {
            Ok(pdf_path) => {
                converted_files.push(pdf_path);
                progress.inc(1);
            },
            Err(err) => {
                progress.suspend(|| {
                    eprintln!("{} {} - {}", "Error converting".red(), file_name, err);
                });
                progress.inc(1);
            }
        }
    }

    progress.finish_with_message("Conversion completed");

    if converted_files.is_empty() {
        println!("{}", "No files were successfully converted.".yellow());
    } else {
        println!("{} {} {}", "Successfully converted".green(), converted_files.len(), "files:".green());
        for file in &converted_files {
            println!("  - {}", file.display());
        }
    }

    Ok(())
}

/// Handle the watch command
fn watch_command(input: &PathBuf, output: &PathBuf, file_type: &Option<String>) -> Result<()> {
    // Validate and resolve paths
    let input_dir = utils::resolve_path(input)
        .context("Failed to resolve input directory path")?;
    let output_dir = utils::resolve_path(output)
        .context("Failed to resolve output directory path")?;

    // Validate input directory
    utils::validate_directory(&input_dir)
        .context("Invalid input directory")?;

    // Ensure output directory exists
    utils::ensure_dir_exists(&output_dir)
        .context("Failed to create output directory")?;

    println!("{} {}", "Watching directory:".blue(), input_dir.display());
    println!("{} {}", "Output directory:".blue(), output_dir.display());

    if let Some(t) = file_type {
        println!("{} {}", "File type filter:".blue(), t);
    }

    println!("{}", "Press Ctrl+C to stop watching".yellow());

    // Create channel for watcher events
    let (tx, rx) = mpsc::channel();

    // Create watcher
    let mut watcher = notify::recommended_watcher(tx)?;

    // Start watching the directory
    watcher.watch(&input_dir, RecursiveMode::Recursive)?;

    // Process events
    for res in rx {
        match res {
            Ok(event) => {
                // Only process file creation or modification events
                if let EventKind::Create(_) | EventKind::Modify(_) = event.kind {
                    for path in event.paths {
                        // Skip directories
                        if path.is_dir() {
                            continue;
                        }

                        // Check if file matches the type filter
                        if !is_file_type_match(&path, file_type) {
                            continue;
                        }

                        let file_name = path.file_name().unwrap_or_default().to_string_lossy();
                        println!("{} {}", "New file detected:".blue(), file_name);

                        // Convert the file
                        match converter::convert_to_pdf(&path, &output_dir) {
                            Ok(pdf_path) => {
                                println!("{} {} -> {}", "Successfully converted".green(), file_name, pdf_path.display());
                            },
                            Err(err) => {
                                eprintln!("{} {} - {}", "Error converting".red(), file_name, err);
                            }
                        }
                    }
                }
            },
            Err(e) => {
                error!("Watch error: {:?}", e);
            }
        }
    }

    Ok(())
}

/// Get list of files to convert based on the file type filter
fn get_files_to_convert(input_dir: &PathBuf, file_type: &Option<String>) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();

    for entry in walkdir::WalkDir::new(input_dir)
        .follow_links(true)
        .into_iter()
        .filter_map(|e| e.ok()) {

        let path = entry.path().to_path_buf();

        // Skip directories
        if path.is_dir() {
            continue;
        }

        // Check if file matches the type filter
        if is_file_type_match(&path, file_type) {
            files.push(path);
        }
    }

    Ok(files)
}

/// Check if a file matches the specified type filter
fn is_file_type_match(path: &PathBuf, file_type: &Option<String>) -> bool {
    if let Some(filter) = file_type {
        if let Some(ext) = path.extension() {
            let ext_str = ext.to_string_lossy().to_lowercase();
            return ext_str == filter.to_lowercase();
        }
        return false;
    }

    // If no filter is specified, check if it's a supported file type
    if let Some(ext) = path.extension() {
        let ext_str = ext.to_string_lossy().to_lowercase();
        return ext_str == "docx" || ext_str == "xlsx" || ext_str == "xls";
    }

    false
}
