#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import PyPDF2
from docx import Document

def extract_text_from_pdf(pdf_path):
    pdf_text = ""
    with open(pdf_path, 'rb') as file:
        pdf_reader = PyPDF2.PdfFileReader(file)
        num_pages = pdf_reader.numPages
        for page_num in range(num_pages):
            page = pdf_reader.getPage(page_num)
            pdf_text += page.extractText()
    return pdf_text

def write_text_to_word(text, word_path):
    document = Document()
    document.add_paragraph(text)
    document.save(word_path)

def main():
    pdf_file = "input.pdf"  # Replace with your PDF file path
    word_file = "output.docx"  # Replace with the desired output Word file path

    text = extract_text_from_pdf(pdf_file)
    write_text_to_word(text, word_file)
    print("Text extracted from PDF has been written to Word successfully!")

if __name__ == "__main__":
    main()
