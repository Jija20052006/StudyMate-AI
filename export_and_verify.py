import os
import win32com.client
import fitz

DOCX_PATH = os.path.abspath(r"c:\Users\DELL\Desktop\StudyMate-AI-main\StudyMate_AI_Project_Report.docx")
PDF_PATH = os.path.abspath(r"c:\Users\DELL\Desktop\StudyMate-AI-main\StudyMate_AI_Project_Report.pdf")

print("Exporting DOCX to PDF via Microsoft Word...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False

try:
    doc = word.Documents.Open(DOCX_PATH)
    # Update all fields in document (page numbers, etc.)
    doc.Fields.Update()
    # Export to PDF (17 = wdExportFormatPDF)
    doc.ExportAsFixedFormat(PDF_PATH, 17)
    print("PDF export successful!")
    doc.Close(False)
except Exception as e:
    print("Export notice / error:", e)
finally:
    try:
        word.Quit()
    except Exception:
        pass

if os.path.exists(PDF_PATH):
    pdf = fitz.open(PDF_PATH)
    print(f"\n==========================================")
    print(f"VERIFICATION: Total PDF Pages = {len(pdf)}")
    print(f"==========================================")
    
    for i in range(len(pdf)):
        page = pdf[i]
        text = page.get_text().strip()
        lines = [l.strip() for l in text.split('\n') if l.strip()]
        first_line = lines[0] if lines else "(empty)"
        last_line = lines[-1] if lines else "(empty)"
        drawings = len(page.get_drawings())
        images = len(page.get_images())
        print(f"Page {i+1:2d}: chars={len(text):4d} | drawings={drawings:2d} | images={images} | First: {first_line[:40]!r} | Last: {last_line[:30]!r}")
    
    pdf.close()
else:
    print("ERROR: PDF was not generated!")
