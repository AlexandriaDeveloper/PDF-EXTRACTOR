import pdfplumber

pdf_path = r"f:\Prog-Projects\AI-Moratbat\ExamplePdf\صرفيه موظف.pdf"

print("--- Start Debugging Page 1 ---")
try:
    with pdfplumber.open(pdf_path) as pdf:
        page = pdf.pages[0]
        words = page.extract_words()
        
        # group by top
        lines = {}
        for w in words:
            # try different rounding to group words in same line
            y = round(w['top'] / 3) * 3
            if y not in lines: lines[y] = []
            lines[y].append(w)
            
        print(f"Total lines found (top 20):")
        for y in sorted(lines.keys())[:20]:
            # LTR
            ltr_words = sorted(lines[y], key=lambda w: w['x0'])
            ltr_str = " | ".join([w['text'] for w in ltr_words])
            
            # RTL
            rtl_words = sorted(lines[y], key=lambda w: w['x0'], reverse=True)
            rtl_str = " | ".join([w['text'] for w in rtl_words])
            
            print(f"Y={y}:")
            print(f"  LTR: {ltr_str}")
            print(f"  RTL: {rtl_str}")
            
except Exception as e:
    print(f"Error: {e}")
print("--- End Debugging ---")
