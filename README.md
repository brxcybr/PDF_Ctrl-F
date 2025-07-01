# PDF_Ctrl-f

## Description

`PDF_Ctrl-f` is a Python script that searches for terms within one or more PDF files and outputs the results to an Excel file. The script provides two modes of operation:
- Count the occurrences of each search term in the PDF(s).
- Mark the presence of search terms with an 'X'.

The script can process individual PDFs or all PDFs in a directory, using either a single search term or regex pattern or a file containing multiple search terms and patterns.

## Features

- **Single or batch mode**: Search through one PDF or multiple PDFs in a directory.
- **Customizable search**: Use a single term or a text file with multiple terms or regex patterns.
- **Flexible output**: Output to the console or an Excel file sorted by number of occurrences or a presence indicator ('X') for each term using the `-x` option.
- **Incremental search**: Automatically transverses subdirectories to find all PDF files.
- **Verbose mode**: Use the `-v` option for detailed output during processing.
- **Comprehensive results**: Optionally include files searched with zero results in the output using the `-n` option.

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
python PDF_Ctrl-f.py [-t SEARCH_TERM | -T TERMS_FILE] [-p PDF_PATH | -P PDF_PARENT_PATH] [-o | -o OUTPUT_FILE] [-x] [-n] [-v]
```

## Arguments
- `-n`, `--include-null`: Include terms with zero results in the output.
- `-o`, `--output_file`: Path to the output Excel file. Default is `Term_Usage_by_PDF.xlsx`. Omitting this argument will print the results to the console.
- `-p`, `--pdf_path`: Path to a single PDF file.
- `-P`, `--pdf_parent_path`: Path to a directory containing PDF files.
- `-t`, `--term`: Single search term to look for in the PDF(s).
- `-T`, `--terms-file`: Text file containing multiple search terms (one per line).
- `-x`, `--no-count`: Mark occurrences of each term with an 'X' instead of showing counts.
- `-v`, `--verbose`: Enable verbose output for debugging.

## Examples
1. Search for the term 'Python' in a single PDF file and output results to the console:
```bash
python PDF_Ctrl-f.py -t "Python" -p "example.pdf" 
```
2. Search for multiple terms in a directory of PDF files and output results to an Excel file:
```bash
python PDF_Ctrl-f.py -T "search_terms.txt" -P "/path/to/pdf_directory" -o "output.xlsx"
```
3. Count the occurrences of each term in a single PDF file:
```bash
python PDF_Ctrl-f.py -t "Python" -p "example.pdf" -o "output.xlsx" -c
```
4. Count the occurrences of multiple terms in a directory of PDF files:
```bash
python PDF_Ctrl-f.py -T "search_terms.txt" -P "/path/to/pdf_directory" -o "output.xlsx" -c
```
5. Search using a regex pattern to match terms starting with “Data” and output results to the console:
```bash
python PDF_Ctrl-f.py -t "^Data.*" -P "/path/to/pdf_directory"
```

## Example Output

### Example Search Terms:
- Python
- Programming
- Data Analysis

### Example Files:
- `example1.pdf`
- `example2.pdf`

### Console Output (Mode: `Count/Default`)
```bash
$ python PDF_Ctrl-f.py -T search_terms.txt 
Processing file 2/2... Done!

Found 'Python' 8 times in 2 PDF file(s):
        Count   Path
        5       example2.pdf
        2       example1.pdf

The following terms had no results:
        'Programming'
        'Data Analysis'
```
### Excel Output (Mode: `Count/Default`)

| Term            | example1.pdf | example2.pdf |
|-----------------|--------------|--------------|
| Python          | 3            | 0            |
| Programming     | 5            | 2            |
| Data Analysis   | 0            | 4            |

### File Output (Mode: `Presence`)

| Term            | example1.pdf | example2.pdf |
|-----------------|--------------|--------------|
| Python          | X            |              |
| Programming     | X            | X            |
| Data Analysis   |              | X            |

In `Presence` mode, the script will place an `X` in the cell if the term is found in the corresponding PDF. In `Count` mode, the script will count how many times the term appears in each PDF file.

## License
MIT License.
