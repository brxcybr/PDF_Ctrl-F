#!/usr/bin/env python3

# PDF_Ctrl-f.py
# Author: brx.cybr@gmail.com
# Date: 2021-09-01
# Search for terms in one or more PDF files and save the results to an Excel file.
# Usage: python PDF_Ctrl-f.py -p "path/to/pdf_file.pdf" -t "search_term" -o "output_file.xlsx"
# Usage: python PDF_Ctrl-f.py -P "path/to/pdf_directory" -t "path/to/terms_file.txt" -o "output_file.xlsx"
# Usage: python PDF_Ctrl-f.py -P "path/to/pdf_directory" -t "path/to/terms_file.txt" -c # Count occurrences of each term

import re
import fitz  # PyMuPDF for extracting the text from the PDF
import os
import pandas as pd
import sys
import argparse

# Parse input
def parse_args():
    parser = argparse.ArgumentParser(description="Search for terms in one or more PDF files.")
    parser.add_argument("-f", "--search_term", type=str, help="Search for a single term.")
    parser.add_argument("-p", "--pdf_path", type=str, help="Path to the PDF file to search.")
    parser.add_argument("-P", "--pdf_parent_path", type=str, help="Path to the parent directory containing PDF files.")
    parser.add_argument("-t", "--terms_path", type=str, help="Path to the file containing search terms.")
    parser.add_argument("-o", "--output_file", type=str, help="Path to the output Excel file.")
    parser.add_argument("-c", "--count", action="store_true", help="Count occurrences of each term instead of marking with an 'X'.")
    
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
    if not options.output_file:
        options.output_file = "Term_Usage_by_PDF.xlsx"

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
def find_terms_in_pdf(pdf_file, terms, count=False):
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

    # Search for each term in the document text
    term_found = {}
    for term in terms:
        escaped_term = re.escape(term.lower())
        if count:
            term_found[term] = len(re.findall(r'\b' + escaped_term + r'\b', text))
        else:
            term_found[term] = bool(re.search(r'\b' + escaped_term + r'\b', text))
    
    return term_found

# Write results to Excel with terms in column A and PDF names as headers
def export_to_excel(term_results, terms, pdf_files, output_file, count=False):
    # Initialize a dictionary where the keys are terms, and values are lists with counts or 'X'
    data = {term: [''] * len(pdf_files) for term in terms}

    # Fill the dictionary with counts or 'X' where the term was found, or leave blank otherwise
    for pdf_idx, pdf_file in enumerate(pdf_files):
        for term in terms:
            if term_results[pdf_file][term]:
                if count:
                    data[term][pdf_idx] = term_results[pdf_file][term] if term_results[pdf_file][term] > 0 else ''
                else:
                    data[term][pdf_idx] = 'X'
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
    for pdf_file in pdf_files:
        print(f"Processing file: {pdf_file}")
        term_found = find_terms_in_pdf(pdf_file, options.terms, options.count)
        term_results[pdf_file] = term_found

    # Export all results to Excel
    export_to_excel(term_results, options.terms, pdf_files, options.output_file, options.count)
    print("Processing complete.")

# Example usage:
if __name__ == "__main__":
    # Get input from command line
    options = parse_args()

    # Run the main function with parsed options
    main(options)
