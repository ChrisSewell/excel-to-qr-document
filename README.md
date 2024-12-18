# QR Code Generator for Excel Data

A Python script that generates QR codes from Excel data and arranges them in a document. Supports both DOCX and PDF output formats.

## Features

- Interactive prompts for all options
- Supports both DOCX and PDF output
- Parallel processing for faster generation
- Preview data before column selection
- Configurable grid layout
- Progress tracking with elapsed time
- Detailed or minimal progress display

## Installation

1. Clone this repository or download the script
2. Install required packages:
```bash
pip install -r requirements.txt
```

## Usage

Basic usage with interactive prompts:
```bash
python qr_code_generator.py
```

Command line options:
```bash
python qr_code_generator.py [-h] [-i INPUT] [-c COLUMN] [-w WIDTH] [-t {docx,pdf}] [--header] [-v]
```

Options:
- `-i, --input`: Specify Excel file to use
- `-c, --column`: Select column number (1-based)
- `-w, --wide`: Number of QR codes per row (1-10)
- `-t, --type`: Output format (docx/pdf)
- `--header`: Skip header row prompt
- `-v, --verbose`: Show detailed progress

## Examples

```bash
# Use specific file and column
python qr_code_generator.py -i data.xlsx -c 2

# Create 4-wide grid in PDF format
python qr_code_generator.py -w 4 -t pdf

# Skip header row prompt and show detailed progress
python qr_code_generator.py --header -v
```

## Notes

- Output files are saved to the 'output' directory
- Temporary files are automatically cleaned up
- Use Ctrl+C to cancel at any time
- Supports Windows, macOS, and Linux
- Handles various system fonts for text rendering

## Requirements

- Python 3.7+
- See requirements.txt for package dependencies