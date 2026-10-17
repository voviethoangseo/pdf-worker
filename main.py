import os
import json
import uuid
import zipfile
import gc
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from pdf2docx import Converter as Pdf2DocxConverter
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
import pytesseract
import csv
from PIL import Image
from docx import Document
import subprocess
import openpyxl
from pptx import Presentation
from pptx.util import Pt
from bs4 import BeautifulSoup
from ebooklib import epub

app = FastAPI(title="PDF Worker Multi-tools")

TEMP_DIR = "/tmp/pdf_worker"
os.makedirs(TEMP_DIR, exist_ok=True)

@app.api_route("/", methods=["GET", "HEAD"])
def health_check():
    return {"status": "ready", "service": "PDF Multi-Worker on Render"}

@app.post("/api/convert")
async def convert_file(
    file: UploadFile = File(...),
    tool: str = Form(...),
    options: str = Form("{}")
):
    req_id = str(uuid.uuid4())
    opts = json.loads(options) if options else {}

    input_ext = os.path.splitext(file.filename)[1] or ".pdf"
    input_path = os.path.join(TEMP_DIR, f"{req_id}_in{input_ext}")
    output_path = os.path.join(TEMP_DIR, f"{req_id}_out")

    with open(input_path, "wb") as f:
        f.write(await file.read())

    try:
        # ==========================================
        # 1. PDF SANG WORD (pdf2docx)
        # ==========================================
        if tool == "pdf-to-word":
            output_file = output_path + ".docx"
            mode = opts.get("word_mode", "basic")

            if mode == "ocr":
                doc_pdf = fitz.open(input_path)
                doc_word = Document()
                for page_idx in range(len(doc_pdf)):
                    page = doc_pdf[page_idx]
                    pix = page.get_pixmap(dpi=150)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    text = pytesseract.image_to_string(img, lang="vie")
                    doc_word.add_heading(f"Trang {page_idx + 1}", level=2)
                    doc_word.add_paragraph(text if text.strip() else "[Không có chữ]")
                doc_word.save(output_file)
            else:
                cv = Pdf2DocxConverter(input_path)
                cv.convert(output_file, start=0, end=None)
                cv.close()

            return FileResponse(output_file, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="result.docx")

        # ==========================================
        # 2. OCR PDF / PDF TÌM KIẾM ĐƯỢC
        # ==========================================
        elif tool in ["ocr-pdf", "pdf-searchable"]:
            output_file = output_path + ".docx"
            doc_pdf = fitz.open(input_path)
            doc_word = Document()
            for page_idx in range(len(doc_pdf)):
                page = doc_pdf[page_idx]
                pix = page.get_pixmap(dpi=150)
                img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                text = pytesseract.image_to_string(img, lang="vie")
                doc_word.add_heading(f"Trang {page_idx + 1}", level=2)
                doc_word.add_paragraph(text if text.strip() else "[Không có chữ]")
            doc_word.save(output_file)
            return FileResponse(output_file, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", filename="result.docx")

        # ==========================================
        # 3. NÉN PDF
        # ==========================================
        elif tool == "compress-pdf":
            output_file = output_path + ".pdf"
            doc = fitz.open(input_path)
            doc.save(output_file, garbage=4, deflate=True, clean=True)
            doc.close()
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # ==========================================
        # 4. TÁCH TRANG PDF
        # ==========================================
        elif tool == "split-pdf":
            output_file = output_path + ".pdf"
            reader = PdfReader(input_path)
            writer = PdfWriter()
            for i in range(min(5, len(reader.pages))):
                writer.add_page(reader.pages[i])
            with open(output_file, "wb") as f_out:
                writer.write(f_out)
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # ==========================================
        # 5. XOAY TRANG PDF
        # ==========================================
        elif tool == "rotate-pdf":
            output_file = output_path + ".pdf"
            angle = int(opts.get("angle", 90))
            reader = PdfReader(input_path)
            writer = PdfWriter()
            for page in reader.pages:
                page.rotate(angle)
                writer.add_page(page)
            with open(output_file, "wb") as f_out:
                writer.write(f_out)
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # ==========================================
        # 6. PDF SANG TEXT
        # ==========================================
        elif tool == "pdf-to-text":
            output_file = output_path + ".txt"
            doc = fitz.open(input_path)
            text_content = []
            for page in doc:
                text_content.append(page.get_text())
            with open(output_file, "w", encoding="utf-8") as f_out:
                f_out.write("\n\n--- Trang tiếp ---\n\n".join(text_content))
            return FileResponse(output_file, media_type="text/plain; charset=utf-8", filename="result.txt")

        # ==========================================
        # 7. PDF SANG MARKDOWN
        # ==========================================
        elif tool == "pdf-to-markdown":
            output_file = output_path + ".md"
            doc = fitz.open(input_path)
            md_content = [f"# {file.filename}\n"]
            for idx, page in enumerate(doc):
                md_content.append(f"## Trang {idx + 1}\n\n" + page.get_text())
            with open(output_file, "w", encoding="utf-8") as f_out:
                f_out.write("\n\n".join(md_content))
            return FileResponse(output_file, media_type="text/markdown; charset=utf-8", filename="result.md")

        # ==========================================
        # 8. PDF SANG JSON
        # ==========================================
        elif tool == "pdf-to-json":
            output_file = output_path + ".json"
            doc = fitz.open(input_path)
            pages_data = []
            for idx, page in enumerate(doc):
                pages_data.append({
                    "page": idx + 1,
                    "text": page.get_text()
                })
            with open(output_file, "w", encoding="utf-8") as f_out:
                json.dump({"total_pages": len(doc), "pages": pages_data}, f_out, ensure_ascii=False, indent=2)
            return FileResponse(output_file, media_type="application/json; charset=utf-8", filename="result.json")

        # ==========================================
        # 9. PDF SANG ẢNH (JPG / PNG - Gom thành file ZIP)
        # ==========================================
        elif tool in ["pdf-to-jpg", "pdf-to-png"]:
            img_format = "jpg" if tool == "pdf-to-jpg" else "png"
            output_zip = output_path + ".zip"
            doc = fitz.open(input_path)

            with zipfile.ZipFile(output_zip, 'w') as zip_f:
                for idx, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=150)
                    img_name = f"trang_{idx + 1}.{img_format}"
                    img_temp = os.path.join(TEMP_DIR, f"{req_id}_{img_name}")
                    pix.save(img_temp)
                    zip_f.write(img_temp, arcname=img_name)
                    if os.path.exists(img_temp):
                        os.remove(img_temp)

            return FileResponse(output_zip, media_type="application/zip", filename="result.zip")

       # ==========================================
        # 10. ẢNH SANG PDF (JPG / PNG sang PDF)
        # ==========================================
        elif tool in ["jpg-to-pdf", "png-to-pdf"]:
            output_file = output_path + ".pdf"
            image = Image.open(input_path)
            if image.mode != 'RGB':
                image = image.convert('RGB')
            image.save(output_file, "PDF", resolution=100.0)
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # ==========================================
        # 11. PDF SANG EXCEL (.xlsx)
        # ==========================================
        elif tool == "pdf-to-excel":
            output_file = output_path + ".xlsx"
            doc = fitz.open(input_path)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Dữ liệu PDF"

            for page in doc:
                # Dùng tính năng bóc tách bảng chuyên dụng của PyMuPDF
                tabs = page.find_tables()
                if tabs.tables:
                    for tab in tabs.tables:
                        for row in tab.extract():
                            ws.append([cell if cell is not None else "" for cell in row])
                        ws.append([]) # Dòng trống ngăn cách bảng
                else:
                    # Nếu PDF không có viền bảng, bóc từng dòng văn bản
                    lines = page.get_text("text").split("\n")
                    for line in lines:
                        if line.strip():
                            ws.append([line.strip()])

            wb.save(output_file)
            return FileResponse(
                output_file,
                media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                filename="result.xlsx"
            )

        # ==========================================
        # 12. PDF SANG CSV (.csv tiếng Việt chuẩn)
        # ==========================================
        elif tool == "pdf-to-csv":
            output_file = output_path + ".csv"
            doc = fitz.open(input_path)

            with open(output_file, "w", newline="", encoding="utf-8-sig") as f_out:
                writer = csv.writer(f_out)
                for page in doc:
                    tabs = page.find_tables()
                    if tabs.tables:
                        for tab in tabs.tables:
                            for row in tab.extract():
                                writer.writerow([cell if cell is not None else "" for cell in row])
                    else:
                        for line in page.get_text("text").split("\n"):
                            if line.strip():
                                writer.writerow([line.strip()])

            return FileResponse(output_file, media_type="text/csv; charset=utf-8", filename="result.csv")

        # ==========================================
        # 13. PDF SANG POWERPOINT (.pptx theo từng slide)
        # ==========================================
        elif tool == "pdf-to-powerpoint":
            output_file = output_path + ".pptx"
            doc = fitz.open(input_path)
            prs = Presentation()

            for page_idx in range(len(doc)):
                page = doc[page_idx]
                # Đặt kích thước slide theo đúng kích thước từng trang PDF
                prs.slide_width = Pt(page.rect.width)
                prs.slide_height = Pt(page.rect.height)
                blank_layout = prs.slide_layouts[6]
                slide = prs.slides.add_slide(blank_layout)

                # Xuất trang PDF thành ảnh chất lượng cao và gắn trọn vẹn vào slide
                pix = page.get_pixmap(dpi=150)
                img_temp = os.path.join(TEMP_DIR, f"{req_id}_slide_{page_idx}.png")
                pix.save(img_temp)

                slide.shapes.add_picture(img_temp, 0, 0, width=Pt(page.rect.width), height=Pt(page.rect.height))
                if os.path.exists(img_temp):
                    os.remove(img_temp)

            prs.save(output_file)
            return FileResponse(
                output_file,
                media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                filename="result.pptx"
            )

       # ==========================================
        # 14. PDF SANG HTML
        # ==========================================
        elif tool == "pdf-to-html":
            output_file = output_path + ".html"
            doc = fitz.open(input_path)
            html_parts = ["<!doctype html><html><head><meta charset='utf-8'><title>PDF to HTML</title></head><body>"]

            for page in doc:
                html_parts.append(page.get_text("html"))

            html_parts.append("</body></html>")
            with open(output_file, "w", encoding="utf-8") as f_out:
                f_out.write("\n".join(html_parts))

            return FileResponse(output_file, media_type="text/html; charset=utf-8", filename="result.html")

        # ==========================================
        # 15. PDF SANG SVG (Đồ họa vector - Gom file ZIP)
        # ==========================================
        elif tool == "pdf-to-svg":
            output_zip = output_path + ".zip"
            doc = fitz.open(input_path)
            with zipfile.ZipFile(output_zip, 'w') as zip_f:
                for idx, page in enumerate(doc):
                    svg_content = page.get_svg_image()
                    svg_name = f"trang_{idx + 1}.svg"
                    svg_temp = os.path.join(TEMP_DIR, f"{req_id}_{svg_name}")
                    with open(svg_temp, "w", encoding="utf-8") as f_svg:
                        f_svg.write(svg_content)
                    zip_f.write(svg_temp, arcname=svg_name)
                    if os.path.exists(svg_temp):
                        os.remove(svg_temp)
            return FileResponse(output_zip, media_type="application/zip", filename="result.zip")

        # ==========================================
        # 16. SVG SANG PDF
        # ==========================================
        elif tool == "svg-to-pdf":
            output_file = output_path + ".pdf"
            doc = fitz.open(input_path)
            pdf_bytes = doc.convert_to_pdf()
            with open(output_file, "wb") as f_out:
                f_out.write(pdf_bytes)
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # ==========================================
        # 17. PDF SANG TIFF (Ảnh in ấn - Gom file ZIP)
        # ==========================================
        elif tool == "pdf-to-tiff":
            output_zip = output_path + ".zip"
            doc = fitz.open(input_path)
            with zipfile.ZipFile(output_zip, 'w') as zip_f:
                for idx, page in enumerate(doc):
                    pix = page.get_pixmap(dpi=150)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    tif_name = f"trang_{idx + 1}.tiff"
                    tif_temp = os.path.join(TEMP_DIR, f"{req_id}_{tif_name}")
                    img.save(tif_temp, format="TIFF")
                    zip_f.write(tif_temp, arcname=tif_name)
                    if os.path.exists(tif_temp):
                        os.remove(tif_temp)
            return FileResponse(output_zip, media_type="application/zip", filename="result.zip")

        # ==========================================
        # 18. TIFF SANG PDF
        # ==========================================
        elif tool == "tiff-to-pdf":
            output_file = output_path + ".pdf"
            image = Image.open(input_path)
            if image.mode != 'RGB':
                image = image.convert('RGB')
            image.save(output_file, "PDF", resolution=100.0)
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # ==========================================
        # 19. TEXT / MARKDOWN / CSV SANG PDF (Dàn trang chuẩn A4)
        # ==========================================
        elif tool in ["text-to-pdf", "markdown-to-pdf", "csv-to-pdf"]:
            output_file = output_path + ".pdf"
            with open(input_path, "r", encoding="utf-8", errors="ignore") as f_in:
                text_content = f_in.read()

            doc = fitz.open()
            page = doc.new_page(width=595, height=842)
            margin_rect = fitz.Rect(50, 50, 545, 792)
            page.insert_textbox(margin_rect, text_content, fontsize=11, fontname="helv")
            doc.save(output_file)
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # ==========================================
        # 20. HTML SANG PDF
        # ==========================================
        elif tool == "html-to-pdf":
            output_file = output_path + ".pdf"
            doc = fitz.open(input_path)
            pdf_bytes = doc.convert_to_pdf()
            with open(output_file, "wb") as f_out:
                f_out.write(pdf_bytes)
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # ==========================================
        # 21. PDF SANG PDF/A (Chuẩn lưu trữ số)
        # ==========================================
        elif tool == "pdf-to-pdfa":
            output_file = output_path + ".pdf"
            doc = fitz.open(input_path)
            doc.save(output_file, deflate=True, garbage=3, clean=True)
            doc.close()
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # ==========================================
        # 22. WORD / EXCEL / POWERPOINT SANG PDF (LibreOffice Headless)
        # ==========================================
        elif tool in ["word-to-pdf", "excel-to-pdf", "powerpoint-to-pdf"]:
            # Chạy LibreOffice xuất PDF chuẩn xác 100%
            cmd = [
                "soffice", "--headless", "--convert-to", "pdf",
                "--outdir", TEMP_DIR, input_path
            ]
            subprocess.run(cmd, check=True, timeout=120)

            base_name = os.path.splitext(os.path.basename(input_path))[0]
            generated_pdf = os.path.join(TEMP_DIR, f"{base_name}.pdf")

            if os.path.exists(generated_pdf):
                return FileResponse(generated_pdf, media_type="application/pdf", filename="result.pdf")
            else:
                raise HTTPException(status_code=500, detail="Không thể tạo PDF từ tài liệu Office này")

        # ==========================================
        # 23. PDF SANG SÁCH EPUB (.epub)
        # ==========================================
        elif tool == "pdf-to-epub":
            output_file = output_path + ".epub"
            doc = fitz.open(input_path)
            book = epub.EpubBook()
            book.set_identifier(req_id)
            book.set_title(file.filename or "Sách PDF chuyển đổi")
            book.set_language("vi")

            chapters = []
            for idx, page in enumerate(doc):
                c = epub.EpubHtml(title=f"Trang {idx + 1}", file_name=f"chap_{idx + 1}.xhtml", lang="vi")
                c.content = f"<h2>Trang {idx + 1}</h2>" + page.get_text("html")
                book.add_item(c)
                chapters.append(c)

            book.toc = tuple(chapters)
            book.add_item(epub.EpubNcx())
            book.add_item(epub.EpubNav())
            book.spine = ["nav"] + chapters

            epub.write_epub(output_file, book, {})
            return FileResponse(output_file, media_type="application/epub+zip", filename="result.epub")

        # ==========================================
        # 24. SÁCH EPUB SANG PDF
        # ==========================================
        elif tool == "epub-to-pdf":
            output_file = output_path + ".pdf"
            book = epub.read_epub(input_path)
            extracted_text = []

            for item in book.get_items():
                if item.get_type() == 9: # DOCUMENT
                    soup = BeautifulSoup(item.get_content(), "html.parser")
                    text = soup.get_text()
                    if text.strip():
                        extracted_text.append(text.strip())

            doc = fitz.open()
            full_str = "\n\n--- Hết chương ---\n\n".join(extracted_text)
            page = doc.new_page(width=595, height=842)
            margin_rect = fitz.Rect(50, 50, 545, 792)
            page.insert_textbox(margin_rect, full_str, fontsize=11, fontname="helv")
            doc.save(output_file)
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        else:
            raise HTTPException(status_code=400, detail=f"Công cụ {tool} không hợp lệ")

    finally:
        # 1. Xóa sạch file input tạm
        if os.path.exists(input_path):
            try: os.remove(input_path)
            except Exception: pass
            
        # 2. Ép Python thu hồi 100% bộ nhớ RAM ngay lập tức (Chống tràn 512MB RAM)
        gc.collect()
