import pdfplumber

pdf_path = r"f:\Prog-Projects\AI-Moratbat\ExamplePdf\صرفيه موظف.pdf"
out_path = r"f:\Prog-Projects\PDF-EXTRACTOR\debug_output.txt"

with open(out_path, "w", encoding="utf-8") as f:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        
        # 1. النص الخام
        txt = page.extract_text() or ""
        f.write("=== RAW TEXT (first 30 lines) ===\n")
        for i, line in enumerate(txt.split("\n")[:30]):
            f.write(f"LINE {i}: [{line}]\n")
            # check each char
            if "اسم" in line or "موظف" in line:
                f.write(f"  >>> MATCH FOUND IN LINE {i}!\n")
                f.write(f"  >>> Chars: {[hex(ord(c)) for c in line[:80]]}\n")
        
        # 2. reversed check
        f.write("\n=== REVERSED CHECK ===\n")
        for i, line in enumerate(txt.split("\n")[:30]):
            rev = line[::-1]
            if "اسم" in rev or "موظف" in rev:
                f.write(f"LINE {i} REVERSED: [{rev}]\n")
                f.write(f"  >>> REVERSED MATCH!\n")

        # 3. Words with coordinates for first 30 unique Y
        f.write("\n=== WORDS (first page, top area) ===\n")
        words = page.extract_words()
        for w in words[:60]:
            f.write(f"  x0={w['x0']:.1f} top={w['top']:.1f} text=[{w['text']}]\n")

print("Done! Check debug_output.txt")
