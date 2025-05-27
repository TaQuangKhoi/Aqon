# Aqon

A CLI tool for batch conversion of Word and Excel documents to PDF.

## Description

Aqon is a command-line utility 
that simplifies the process of converting 
Microsoft Office documents (DOCX, XLSX, XLS) to PDF format. 
It supports both batch conversion of existing files and watching directories for new files to convert automatically.

Aqon was created to make it easy to push data files into Google's NotebookLM. 
By converting Word and Excel documents to PDF format, they become compatible with NotebookLM's document analysis capabilities.

## Features

- Convert DOCX (Word) documents to PDF
- Convert XLSX/XLS (Excel) spreadsheets to PDF
- Batch process entire directories of documents
- Watch mode to automatically convert new files as they appear
- Progress indicators for batch operations
- Colorized terminal output
- Detailed logging with verbose option
- Cross-platform support

## Installation

### Prerequisites

- Rust and Cargo (1.70 or newer recommended)

### Building from Source

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/Aqon.git
   cd Aqon
   ```

2. Build the project:
   ```
   cargo build --release
   ```

3. The compiled binary will be available at `target/release/Aqon`

## Usage

### Converting Documents

To convert documents from an input directory to an output directory:

```
Aqon convert --input <INPUT_DIR> --output <OUTPUT_DIR>
```

Options:
- `--input`, `-i`: Input directory containing documents to convert
- `--output`, `-o`: Output directory for generated PDFs
- `--type`, `-t`: (Optional) Only convert files of specified type (docx, xlsx, xls)
- `--verbose`, `-v`: Enable verbose logging

### Watching a Directory

To watch a directory and automatically convert new documents as they appear:

```
Aqon watch --input <INPUT_DIR> --output <OUTPUT_DIR>
```

Options:
- `--input`, `-i`: Input directory to watch for new documents
- `--output`, `-o`: Output directory for generated PDFs
- `--type`, `-t`: (Optional) Only convert files of specified type (docx, xlsx, xls)
- `--verbose`, `-v`: Enable verbose logging

### Examples

Convert all supported documents in the "documents" folder to PDFs in the "output" folder:
```
Aqon convert --input documents --output output
```

Convert only Word documents:
```
Aqon convert --input documents --output output --type docx
```

Watch a directory for new Excel files and convert them automatically:
```
Aqon watch --input documents --output output --type xlsx
```

## Google NotebookLM Integration

[Google NotebookLM](https://notebooklm.google/) is an AI-powered note-taking tool that can analyze documents and help you work with their content. NotebookLM accepts PDF files as input for its document analysis.

Aqon was specifically created to streamline the workflow of preparing documents for NotebookLM:

1. Convert your Word documents and Excel spreadsheets to PDF format using Aqon
2. Upload the resulting PDF files to NotebookLM
3. Let NotebookLM analyze your documents and extract insights

This integration helps you quickly transform your existing document library into a format that can be processed by NotebookLM's AI capabilities.

## Dependencies

Aqon relies on the following Rust crates:
- clap: Command-line argument parsing
- docx-rs: Reading Word documents
- calamine: Reading Excel spreadsheets
- genpdf: PDF generation
- anyhow: Error handling
- walkdir: Directory traversal
- log & env_logger: Logging
- indicatif: Progress indicators
- colored: Terminal coloring
- notify: File system notifications

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add some amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- The Rust community for providing excellent libraries
- All contributors who have helped shape this project
