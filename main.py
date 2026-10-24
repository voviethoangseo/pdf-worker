import os
import io
import json
import uuid
import zipfile
import gc
import asyncio
import subprocess
import csv
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse, Response
from pdf2docx import Converter as Pdf2DocxConverter
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
import pytesseract
from PIL import Image
from docx import Document
import openpyxl
from pptx import Presentation
from pptx.util import Pt
from bs4 import BeautifulSoup
from ebooklib import epub

app = FastAPI(title="PDF Worker Multi-tools")

TEMP_DIR = "/tmp/pdf_worker"
os.makedirs(TEMP_DIR, exist_ok=True)

# Khóa hàng đợi: Chỉ cho phép 1 tác vụ chạy tại 1 thời điểm để bảo vệ 512MB RAM trên Render
queue_lock = asyncio.Lock()


# 1. Chống lỗi 405 của UptimeRobot / Ping máy chủ
@app.middleware("http")
async def handle_head_requests(request, call_next):
    if request.method == "HEAD" and request.url.path == "/":
        return Response(status_code=200)
    return await call_next(request)


# 2. Health check
@app.get("/")
def health_check():
    return {"status": "ready", "service": "PDF Multi-Worker on Render"}


# 3. Thuật toán nén PDF thông minh theo từng mức độ
def compress_pdf_engine(input_path: str, output_path: str, level: str = "balanced"):
    doc = fitz.open(input_path)
    
    if level == "strong":
        # Nén tối đa: Re-render từng trang thành ảnh tối ưu dung lượng (giảm 50-80% file)
        out_doc = fitz.open()
        for page in doc:
            pix = page.get_pixmap(dpi=100)
            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            buf = io.BytesIO()
            img.save(buf, format="JPEG", quality=45, optimize=True)
            
            rect = page.rect
            new_page = out_doc.new_page(width=rect.width, height=rect.height)
            new_page.insert_image(rect, stream=buf.getvalue())
        out_doc.save(output_path, garbage=4, deflate=True, clean=True)
        out_doc.close()
    elif level == "light":
        # Nén nhẹ: Giữ nguyên độ nét cao, dọn dẹp các stream thừa
        doc.save(output_path, garbage=3, deflate=True, deflate_images=True, deflate_fonts=True)
    else:
        # Cân bằng (khuyên dùng): Dọn dẹp metadata, làm sạch đối tượng rác và nén chặt luồng
        doc.save(output_path, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True, clean=True)
    
    doc.close()


@app.post("/api/convert")
async def convert_file(
    file: UploadFile = File(...),
    tool: str = Form(...),
    options: str = Form("{}")
):
    # Khóa hàng đợi: Toàn bộ quá trình xử lý được bọc trong lock để chống tràn 512MB RAM
    async with queue_lock:
        req_id = str(uuid.uuid4())
        opts = json.loads(options) if options else {}

        input_ext = os.path.splitext(file.filename)[1] or ".pdf"
        input_path = os.path.join(TEMP_DIR, f"{req_id}_in{input_ext}")
        output_path = os.path.join(TEMP_DIR, f"{req_id}_out")

        with open(input_path, "wb") as f:
            f.write(await file.read())

        try:
            # ==========================================
            # 1. PDF SANG WORD (3 Chế độ thực thụ)
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
                        text = pytesseract.image_to_string(img, lang="vie+eng")
                        doc_word.add_heading(f"Trang {page_idx + 1}", level=2)
                        doc_word.add_paragraph(text if text.strip() else "[Không tìm thấy chữ]")
                    doc_word.save(output_file)
                    doc_pdf.close()
                elif mode == "advanced":
                    cv = Pdf2DocxConverter(input_path)
                    cv.convert(output_file, multi_processing=False, parse_lattice_table=True, parse_stream_table=True)
                    cv.close()
                else:
                    cv = Pdf2DocxConverter(input_path)
                    cv.convert(output_file, start=0, end=None)
                    cv.close()

                return FileResponse(
                    output_file,
                    media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    filename="result.docx"
                )

            # ==========================================
            # 2. OCR PDF & PDF TÌM KIẾM ĐƯỢC (Xuất Searchable PDF)
            # ==========================================
            elif tool in ["ocr-pdf", "pdf-searchable"]:
                output_file = output_path + ".pdf"
                doc = fitz.open(input_path)
                out_doc = fitz.open()

                for page in doc:
                    pix = page.get_pixmap(dpi=150)
                    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                    pdf_bytes = pytesseract.image_to_pdf_or_hocr(img, extension='pdf', lang='vie+eng')
                    sub_doc = fitz.open("pdf", pdf_bytes)
                    out_doc.insert_pdf(sub_doc)
                    sub_doc.close()

                doc.close()
                out_doc.save(output_file, garbage=3, deflate=True)
                out_doc.close()
                return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

            # ==========================================
            # 3. NÉN PDF
            # ==========================================
            elif tool == "compress-pdf":
                output_file = output_path + ".pdf"
                level = opts.get("compress_level", "balanced")
                compress_pdf_engine(input_path, output_file, level)
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
                text_content = [page.get_text() for page in doc]
                doc.close()
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
                doc.close()
                with open(output_file, "w", encoding="utf-8") as f_out:
                    f_out.write("\n\n".join(md_content))
                return FileResponse(output_file, media_type="text/markdown; charset=utf-8", filename="result.md")

            # ==========================================
            # 8. PDF SANG JSON
            # ==========================================
            elif tool == "pdf-to-json":
                output_file = output_path + ".json"
                doc = fitz.open(input_path)
                pages_data = [{"page": idx + 1, "text": page.get_text()} for idx, page in enumerate(doc)]
                total_pages = len(doc)
                doc.close()
                with open(output_file, "w", encoding="utf-8") as f_out:
                    json.dump({"total_pages": total_pages, "pages": pages_data}, f_out, ensure_ascii=False, indent=2)
                return FileResponse(output_file, media_type="application/json; charset=utf-8", filename="result.json")

            # ==========================================
            # 9. PDF SANG ẢNH (JPG / PNG - ZIP)
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
                doc.close()
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
                    tabs = page.find_tables()
                    if tabs.tables:
                        for tab in tabs.tables:
                            for row in tab.extract():
                                ws.append([cell if cell is not None else "" for cell in row])
                            ws.append([])
                    else:
                        lines = page.get_text("text").split("\n")
                        for line in lines:
                            if line.strip():
                                ws.append([line.strip()])
                doc.close()
                wb.save(output_file)
                return FileResponse(
                    output_file,
                    media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    filename="result.xlsx"
                )

            # ==========================================
            # 12. PDF SANG CSV (.csv)
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
                doc.close()
                return FileResponse(output_file, media_type="text/csv; charset=utf-8", filename="result.csv")

            # ==========================================
            # 13. PDF SANG POWERPOINT (.pptx)
            # ==========================================
            elif tool == "pdf-to-powerpoint":
                output_file = output_path + ".pptx"
                doc = fitz.open(input_path)
                prs = Presentation()

                for page_idx in range(len(doc)):
                    page = doc[page_idx]
                    prs.slide_width = Pt(page.rect.width)
                    prs.slide_height = Pt(page.rect.height)
                    blank_layout = prs.slide_layouts[6]
                    slide = prs.slides.add_slide(blank_layout)

                    pix = page.get_pixmap(dpi=150)
                    img_temp = os.path.join(TEMP_DIR, f"{req_id}_slide_{page_idx}.png")
                    pix.save(img_temp)

                    slide.shapes.add_picture(img_temp, 0, 0, width=Pt(page.rect.width), height=Pt(page.rect.height))
                    if os.path.exists(img_temp):
                        os.remove(img_temp)
                doc.close()
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
                doc.close()

                with open(output_file, "w", encoding="utf-8") as f_out:
                    f_out.write("\n".join(html_parts))
                return FileResponse(output_file, media_type="text/html; charset=utf-8", filename="result.html")

            # ==========================================
            # 15. PDF SANG SVG (ZIP)
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
                doc.close()
                return FileResponse(output_zip, media_type="application/zip", filename="result.zip")

            # ==========================================
            # 16. SVG SANG PDF
            # ==========================================
            elif tool == "svg-to-pdf":
                output_file = output_path + ".pdf"
                doc = fitz.open(input_path)
                pdf_bytes = doc.convert_to_pdf()
                doc.close()
                with open(output_file, "wb") as f_out:
                    f_out.write(pdf_bytes)
                return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

            # ==========================================
            # 17. PDF SANG TIFF (ZIP)
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
                doc.close()
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
            # 19. TEXT / MARKDOWN / CSV SANG PDF
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
                doc.close()
                return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

            # ==========================================
            # 20. HTML SANG PDF
            # ==========================================
            elif tool == "html-to-pdf":
                output_file = output_path + ".pdf"
                doc = fitz.open(input_path)
                pdf_bytes = doc.convert_to_pdf()
                doc.close()
                with open(output_file, "wb") as f_out:
                    f_out.write(pdf_bytes)
                return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

            # ==========================================
            # 21. PDF SANG PDF/A
            # ==========================================
            elif tool == "pdf-to-pdfa":
                output_file = output_path + ".pdf"
                doc = fitz.open(input_path)
                doc.save(output_file, deflate=True, garbage=3, clean=True)
                doc.close()
                return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

            # ==========================================
            # 22. OFFICE (WORD / EXCEL / PPT) SANG PDF
            # ==========================================
            elif tool in ["word-to-pdf", "excel-to-pdf", "powerpoint-to-pdf"]:
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
            # 23. PDF SANG EPUB
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

                doc.close()
                book.toc = tuple(chapters)
                book.add_item(epub.EpubNcx())
                book.add_item(epub.EpubNav())
                book.spine = ["nav"] + chapters

                epub.write_epub(output_file, book, {})
                return FileResponse(output_file, media_type="application/epub+zip", filename="result.epub")

            # ==========================================
            # 24. EPUB SANG PDF
            # ==========================================
            elif tool == "epub-to-pdf":
                output_file = output_path + ".pdf"
                book = epub.read_epub(input_path)
                extracted_text = []

                for item in book.get_items():
                    if item.get_type() == 9:  # DOCUMENT
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
                doc.close()
                return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

            else:
                raise HTTPException(status_code=400, detail=f"Công cụ {tool} không hợp lệ")

        finally:
            # 1. Xóa file đầu vào tạm thời
            if os.path.exists(input_path):
                try:
                    os.remove(input_path)
                except Exception:
                    pass

            # 2. Thu hồi toàn bộ RAM giải phóng bộ nhớ
            gc.collect()
