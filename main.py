import os
import json
import uuid
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.responses import FileResponse
from pdf2docx import Converter as Pdf2DocxConverter
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter
import pytesseract
from PIL import Image
from docx import Document

app = FastAPI(title="PDF Worker")

TEMP_DIR = "/tmp/pdf_worker"
os.makedirs(TEMP_DIR, exist_ok=True)

@app.get("/")
def health_check():
    return {"status": "ready", "service": "PDF Worker on Render"}

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
        # 1. PDF SANG WORD (Bố cục chuẩn bằng pdf2docx)
        if tool == "pdf-to-word":
            output_file = output_path + ".docx"
            mode = opts.get("word_mode", "basic")

            if mode == "ocr":
                # Chế độ OCR tiếng Việt
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
                # Chế độ thường: pdf2docx giữ nguyên bảng, khung, chữ
                cv = Pdf2DocxConverter(input_path)
                cv.convert(output_file, start=0, end=None)
                cv.close()

            return FileResponse(
                output_file, 
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                filename="result.docx"
            )

        # 2. OCR PDF / PDF TÌM KIẾM ĐƯỢC
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
            return FileResponse(
                output_file, 
                media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                filename="result.docx"
            )

        # 3. NÉN PDF
        elif tool == "compress-pdf":
            output_file = output_path + ".pdf"
            doc = fitz.open(input_path)
            doc.save(output_file, garbage=4, deflate=True, clean=True)
            doc.close()
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        # 4. TÁCH TRANG PDF
        elif tool == "split-pdf":
            output_file = output_path + ".pdf"
            reader = PdfReader(input_path)
            writer = PdfWriter()
            for i in range(min(3, len(reader.pages))):
                writer.add_page(reader.pages[i])
            with open(output_file, "wb") as f_out:
                writer.write(f_out)
            return FileResponse(output_file, media_type="application/pdf", filename="result.pdf")

        else:
            raise HTTPException(status_code=400, detail="Tool chưa được hỗ trợ")

    finally:
        if os.path.exists(input_path):
            os.remove(input_path)
