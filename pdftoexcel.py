import fitz  # PyMuPDF
import pandas as pd
import os
from collections import defaultdict

def extract_tables_from_pdf(pdf_path, output_excel):
    doc = fitz.open(pdf_path)
    all_tables = []

    for page_num, page in enumerate(doc):
        blocks = page.get_text("dict")["blocks"]
        rows = defaultdict(list)

        for block in blocks:
            if "lines" in block:
                for line in block["lines"]:
                    for span in line["spans"]:
                        y = round(span['bbox'][1], 1)  # y0 position
                        text = span['text'].strip()
                        if text:
                            rows[y].append((span['bbox'][0], text))  # x0 and text

        # Sort each row's text by x position
        table_data = []
        for y in sorted(rows.keys()):
            sorted_row = [t[1] for t in sorted(rows[y], key=lambda x: x[0])]
            table_data.append(sorted_row)

        if table_data:
            df = pd.DataFrame(table_data)
            all_tables.append((f"Page_{page_num+1}", df))

    # Save all tables to Excel
    with pd.ExcelWriter(output_excel, engine="openpyxl") as writer:
        for sheet_name, table_df in all_tables:
            table_df.to_excel(writer, sheet_name=sheet_name, index=False, header=False)

    print(f"✅ Extraction complete! Output saved to '{output_excel}'")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Extract tables from system-generated PDF without using Tabula or Camelot.")
    parser.add_argument("pdf_file", help="Path to input PDF file")
    parser.add_argument("--output", help="Output Excel file name", default="extracted_tables.xlsx")

    args = parser.parse_args()
    extract_tables_from_pdf(args.pdf_file, args.output)