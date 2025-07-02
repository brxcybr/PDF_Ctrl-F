#!/usr/bin/env python3

# PDF_Ctrl-f.py
# Author: brx.cybr@gmail.com
# Creation Date: 2024-09-01
# Last Modified: 2025-07-01
# Search for terms in one or more PDF files and save the results to an Excel file.
# Usage: python PDF_Ctrl-f.py -p "path/to/pdf_file.pdf" -t "search_term" -o "output_file.xlsx"
# Usage: python PDF_Ctrl-f.py -P "path/to/pdf_directory" -T "path/to/terms_file.txt" -o "output_file.xlsx"
# Usage: python PDF_Ctrl-f.py -P "path/to/pdf_directory" -T "path/to/terms_file.txt" -c # Count occurrences of each term

import re
import fitz  # PyMuPDF for extracting the text from the PDF
import os
import pandas as pd
import sys
import argparse

# Parse input
def parse_args():
    parser = argparse.ArgumentParser(description="Search for terms in one or more PDF files.")
    parser.add_argument("-t", "--term", dest="search_term", type=str, help="Search for a single term.")
    parser.add_argument("-p", "--pdf_path", type=str, help="Path to the PDF file to search.")
    parser.add_argument("-P", "--pdf_parent_path", type=str, help="Path to the parent directory containing PDF files.")
    parser.add_argument("-T", "--terms-file", dest="terms_path", type=str, help="Path to the file containing search terms.")
    parser.add_argument("-o", "--output_file", nargs="?", const="Term_Usage_by_PDF.xlsx", help="Optional: Path to the output Excel file. If this argument is used with no filename provided, uses a default filename. Omitting this argument prints results to the console.") # Defaults to "Term_Usage_by_PDF.xlsx" 
    parser.add_argument("-n", "--include-null", action="store_true", help="Include empty results in the output.")
    parser.add_argument("-r", "--regex", action="store_true", help="Treat search terms as regular expressions.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose output for debugging.")
    parser.add_argument("-x", "--no-count", action="store_true", help="Mark occurrences of each term instead with an 'X' instead of count.")

    options = parser.parse_args(sys.argv[1:])
    
    # Validate arguments
    if not options.pdf_parent_path and not options.pdf_path:
        print("Please provide either a single PDF path or a parent directory containing PDFs.")
        sys.exit(1)

    if options.search_term and options.terms_path:
        print("Please provide either a search term or a file containing search terms, not both.")
        sys.exit(1)

    if options.pdf_path and not os.path.exists(options.pdf_path):
        print(f"PDF file not found at {options.pdf_path}.")
        sys.exit(1)

    if not options.pdf_parent_path and options.pdf_path:
        options.pdf_parent_path = os.path.dirname(options.pdf_path)
    
    if options.terms_path and not os.path.exists(options.terms_path):
        print(f"Search terms file not found at {options.terms_path}.")
        sys.exit(1)

    # Set defaults for output and terms
    # If output_file is None, print to console. If -o is used without argument, output_file is set to const value.
    if options.search_term:
        options.terms = [options.search_term]
    elif options.terms_path:
        options.terms = load_terms_from_file(options.terms_path)
    else:
        print("No search term or terms file provided.")
        sys.exit(1)

    return options

# Load terms from the file
def load_terms_from_file(terms_path):
    try:
        with open(terms_path, 'r') as f:
            terms = [term.strip() for term in f if term.strip()]  # Remove empty lines
        return terms
    except Exception as e:
        print(f"Error loading terms from {terms_path}: {e}")
        sys.exit(1)

# Gather all PDF file paths in subdirectories
def gather_files(pdf_parent_path):
    pdf_files = []
    for root, dirs, files in os.walk(pdf_parent_path):
        for file in files:
            if file.endswith(".pdf"):
                pdf_files.append(os.path.join(root, file))
    return pdf_files

# Search for terms in the PDF and return a dictionary of term counts or presence
def find_terms_in_pdf(pdf_file, terms, count=False, regex=False):
    try:
        pdf_document = fitz.open(pdf_file)
    except Exception as e:
        print(f"Error opening PDF {pdf_file}: {e}")
        return {}

    text = ""
    for page_num in range(pdf_document.page_count):
        try:
            page = pdf_document.load_page(page_num)
            text += page.get_text().lower()  # Combine all pages' text into one string
        except Exception as e:
            print(f"Error processing page {page_num + 1} in {pdf_file}: {e}")
    
    pdf_document.close()

    # Search for each term or regex in the document text
    term_found = {}
    for term in terms:
        if regex:
            # Use the term as a regex pattern
            pattern = re.compile(term, re.IGNORECASE)
        else:
            # Escape term for literal search
            pattern = re.compile(r'\b' + re.escape(term.lower()) + r'\b')

        if count:
            term_found[term] = len(pattern.findall(text))
        else:
            term_found[term] = bool(pattern.search(text))
    
    return term_found

# Write results to Excel with terms in column A and PDF names as headers
def export_to_excel(term_results, terms, pdf_files, output_file, no_count=False, include_null=False):
    # Initialize a dictionary where the keys are terms, and values are lists with counts or 'X'
    data = {term: [''] * len(pdf_files) for term in terms}

    # Fill the dictionary with counts or 'X' where the term was found, or leave blank otherwise
    if include_null: # Include empty results if --include-null is specified
        for term in terms:
            if term not in data:
                data[term] = [''] * len(pdf_files)
    else:
        for term in terms:
            if term not in data:
                data[term] = []

    # Iterate through each PDF file and fill the data dictionary
    # If no_count is True, mark occurrences with 'X', otherwise count occurrences
    for pdf_idx, pdf_file in enumerate(pdf_files):
        for term in terms:
            if term_results[pdf_file][term]:
                if no_count:
                    data[term][pdf_idx] = 'X'
                else:
                    data[term][pdf_idx] = term_results[pdf_file][term] if term_results[pdf_file][term] > 0 else ''
            else:
                data[term][pdf_idx] = ''

    # Create a DataFrame where column A is "Term", and each PDF file is a new column
    df = pd.DataFrame(data)
    
    # Transpose to have terms in rows and PDFs in columns
    df = df.T
    df.columns = [os.path.splitext(os.path.basename(pdf))[0] for pdf in pdf_files]
    df.index.name = 'Term'

    # Write the DataFrame to Excel
    df.to_excel(output_file)
    print(f"Results saved to {output_file}")

# Main function to process all files and terms
def main(options):
    # Gather PDF files
    pdf_files = [options.pdf_path] if options.pdf_path else gather_files(options.pdf_parent_path)

    if not pdf_files:
        print(f"No PDF files found in directory {options.pdf_parent_path}.")
        sys.exit(1)

    # Initialize a dictionary to store term search results per PDF
    term_results = {}

    # Process each PDF file and store term results
    for idx, pdf_file in enumerate(pdf_files, 1):
        # Shows verbose output if enabled
        if options.verbose:
            print(f"Processing file ({idx}/{len(pdf_files)}): {pdf_file}")
        else:
            print(f"Processing file {idx}/{len(pdf_files)}...", end='\r')
        term_found = find_terms_in_pdf(pdf_file, options.terms, not options.no_count, options.regex)
        term_results[pdf_file] = term_found
    print(f"Processing file {len(pdf_files)}/{len(pdf_files)}... Done!\n")

    # If no output file is specified, print results to console
    if not options.output_file:
        terms_str = ", ".join([f"'{t}'" for t in options.terms])
        printed_any = False
        not_found = []
        for term in options.terms:
            # Pre-calculate totals
            total = 0
            file_count = 0
            for pdf_file in pdf_files:
                count = term_results[pdf_file].get(term, 0)
                if isinstance(count, int) and count > 0:
                    total += count
                    file_count += 1
                elif isinstance(count, bool) and count:
                    total += 1
                    file_count += 1
            if file_count == 0:
                not_found.append(term)
            if file_count == 0 and not options.include_null:
                continue
            printed_any = True
            # Header with totals
            print(f"Found '{term}' {total} times in {file_count} PDF file(s):")
            print("\tCount\tPath")

            # Detailed lines with sorting by count (highest first)
            results = []
            for pdf_file in pdf_files:
                count = term_results[pdf_file].get(term, 0)
                show = options.include_null or (count if isinstance(count, int) else bool(count))
                if not show:
                    continue
                results.append((count, pdf_file))
            # Sort by count descending
            results.sort(key=lambda x: x[0] if isinstance(x[0], int) else 1, reverse=True)
            for count, pdf_file in results:
                abs_path = os.path.abspath(pdf_file)
                if options.no_count:
                    display = 'X' if bool(count) else 'Not Found'
                else:
                    display = count
                print(f"\t{display}\t{abs_path}")
            print()  # Blank line after each term
        if not printed_any:
            print(f"The terms {{{terms_str}}} were found in 0 PDF file(s).")
        if printed_any and not_found:
            print("The following terms had no results:")
            for term in sorted(not_found, key=str.lower):
                print(f"\t'{term}'")

    # Export all results to Excel
    else:
        export_to_excel(term_results, options.terms, pdf_files, options.output_file, options.no_count, options.include_null)

# Example usage:
if __name__ == "__main__":
    # Get input from command line
    options = parse_args()

    # Run the main function with parsed options
    main(options)
