# PDF_Ctrl-f

## Description

`PDF_Ctrl-f` is a Python script that searches for terms within one or more PDF files and outputs the results to an Excel file. The script provides two modes of operation:
- Mark the presence of search terms with an 'X'.
- Count the occurrences of each search term in the PDF(s).

The script can process individual PDFs or all PDFs in a directory, using either a single search term or a file containing multiple search terms.

## Features

- **Single or batch mode**: Search through one PDF or multiple PDFs in a directory.
- **Customizable search**: Use a single term or a text file with multiple terms.
- **Flexible output**: Output an Excel file with either a presence indicator ('X') or the number of occurrences of each term.
- **Incremental search**: Automatically transverses subdirectories to find all PDF files.
- **Count mode**: With the `-c` option, count occurrences of each term instead of marking with 'X'.

## Requirements

- Python 3.x
- Required modules (install via `pip`):
  - `PyMuPDF` (for reading PDFs)
  - `pandas` (for handling Excel outputs)
  - `openpyxl` (for writing Excel files)


## Environment Setup
Create a virtual environment (Optional)
```bash
python -m virtualenv .venv
# Activate the virtual environment (*nix)
source .venv/bin/activate
# Activate the virtual environment (Windows)
.venv\Scripts\activate
```
Install using the requirements file:
```bash
pip install -r requirements.txt
```
Alternatively, install the required modules manually:
```bash
pip install pymupdf pandas openpyxl
```

## Usage

```bash
python PDF_Ctrl-f.py [-f SEARCH_TERM | -t TERMS_FILE] [-p PDF_PATH | -P PDF_PARENT_PATH] [-o OUTPUT_FILE] [-c]
```

## Arguments
- `-f`, `--search_term`: Single search term to look for in the PDF(s).
- `-t`, `--terms_path`: Text file containing multiple search terms (one per line).
- `-p`, `--pdf_path`: Path to a single PDF file.
- `-P`, `--pdf_parent_path`: Path to a directory containing PDF files.
- `-o`, `--output_file`: Path to the output Excel file. Default is `Term_Usage_by_PDF.xlsx`.
- `-c`, `--count`: Count the occurrences of each term in the PDF(s) instead of an `X` in each cell.

## Examples
1. Search for the term 'Python' in a single PDF file:
```bash
python PDF_Ctrl-f.py -f "Python" -p "example.pdf" -o "output.xlsx"
```
2. Search for multiple terms in a directory of PDF files:
```bash
python PDF_Ctrl-f.py -t "search_terms.txt" -P "/path/to/pdf_directory" -o "output.xlsx"
```
3. Count the occurrences of each term in a single PDF file:
```bash
python PDF_Ctrl-f.py -f "Python" -p "example.pdf" -o "output.xlsx" -c
```
4. Count the occurrences of multiple terms in a directory of PDF files:
```bash
python PDF_Ctrl-f.py -t "search_terms.txt" -P "/path/to/pdf_directory" -o "output.xlsx" -c
```
## Example Output

### Example Search Terms:
- Python
- Programming
- Data Analysis

### Example Files:
- `example1.pdf`
- `example2.pdf`

### Output (Mode: `Presence`)

| Term            | example1.pdf | example2.pdf |
|-----------------|--------------|--------------|
| Python          | X            |              |
| Programming     | X            | X            |
| Data Analysis   |              | X            |

### Output (Mode: `Count`)

| Term            | example1.pdf | example2.pdf |
|-----------------|--------------|--------------|
| Python          | 3            | 0            |
| Programming     | 5            | 2            |
| Data Analysis   | 0            | 4            |

In `Presence` mode, the script will place an `X` in the cell if the term is found in the corresponding PDF. In `Count` mode, the script will count how many times the term appears in each PDF file.


## License
MIT License.
