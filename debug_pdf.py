import pdfplumber

pdf_path = r"f:\Prog-Projects\AI-Moratbat\ExamplePdf\صرفيه موظف.pdf"
output_path = r"f:\Prog-Projects\PDF-EXTRACTOR\debug_output.txt"

with open(output_path, "w", encoding="utf-8") as out:
    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages[:2]):
            out.write(f"\n{'='*80}\n=== Page {page_num + 1} ===\n{'='*80}\n")
            
            text = page.extract_text()
            if text:
                out.write("\n--- Raw Text ---\n")
                for i, line in enumerate(text.split("\n")[:20]):
                    out.write(f"  Line {i}: [{line}]\n")
                    out.write(f"         repr: {repr(line[:100])}\n")
            
            tables = page.extract_tables()
            out.write(f"\n--- Tables: {len(tables)} ---\n")
            
            for t_idx, table in enumerate(tables):
                cols = len(table[0]) if table else 0
                out.write(f"\n  === Table {t_idx+1} (rows:{len(table)}, cols:{cols}) ===\n")
                for r_idx, row in enumerate(table[:25]):
                    out.write(f"\n    Row {r_idx} ({len(row)} cols):\n")
                    for c_idx, cell in enumerate(row):
                        if cell:
                            out.write(f"      Col {c_idx}: [{str(cell)[:80]}]\n")
                            out.write(f"               repr: {repr(str(cell)[:80])}\n")

    out.write("\n\n=== DONE ===\n")

print(f"Output written to: {output_path}")
