import sys
import os

# Pastikan direktori backend masuk sys.path agar impor lokal selalu berhasil
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from fastapi import FastAPI, UploadFile, File, HTTPException, Form
from fastapi.responses import FileResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware
from starlette.background import BackgroundTask
from contextlib import asynccontextmanager
import os
import uuid
import shutil
import asyncio
from typing import List, Optional
import json
from converter import (
    convert_pdf_to_word_ms_word,
    convert_pdf_to_word_preserve,
    convert_pdf_to_word_ocr,
    convert_pdf_to_word_auto,
    convert_word_to_pdf,
    convert_ppt_to_pdf,
    convert_pdf_to_ppt,
    convert_excel_to_pdf,
    convert_pdf_to_excel,
    convert_images_to_pdf_with_options,
    compress_pdf,
    compress_image,
    compress_image_to_target,
    merge_pdf_files,
    remove_watermark_auto,
    remove_camscanner_watermark,
    remove_watermark_inpaint,
    scan_document,
    get_document_boundary,
    split_pdf_by_ranges,
    split_pdf_every_n,
    zip_files,
    convert_pdf_to_markdown,
    convert_word_to_markdown,
    convert_pdf_to_images
)
from utils import cleanup_old_files

# =============================================
# KONFIGURASI PATH
# =============================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)
FRONTEND_DIR = os.path.abspath(os.path.join(BASE_DIR, "..", "frontend"))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")

# =============================================
# PEMBERSIHAN FILE
# =============================================

def _delete_path(path: str):
    """Hapus satu file atau folder. Dipakai sebagai BackgroundTask setelah response terkirim."""
    try:
        if os.path.isdir(path):
            shutil.rmtree(path, ignore_errors=True)
        elif os.path.exists(path):
            os.remove(path)
    except Exception as e:
        print(f"Gagal menghapus {path}: {e}")

def _delete_paths(*paths: str):
    for p in paths:
        _delete_path(p)

async def _periodic_cleanup():
    while True:
        cleanup_old_files()
        await asyncio.sleep(600)  

@asynccontextmanager
async def lifespan(app: FastAPI):
    os.makedirs(UPLOAD_DIR, exist_ok=True)
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    task = asyncio.create_task(_periodic_cleanup())
    yield
    task.cancel()
    try:
        await task
    except asyncio.CancelledError:
        pass

app = FastAPI(title="Converter App - PDF, Word, Excel, PowerPoint & Images", lifespan=lifespan)

# CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
    expose_headers=["*"],
)

# Serve frontend
app.mount("/static", StaticFiles(directory=FRONTEND_DIR), name="static")

@app.get("/")
async def root():
    return FileResponse(os.path.join(FRONTEND_DIR, "index.html"))

# =============================================
# HELPER KONVERSI GAMBAR KE PDF (REUSABLE)
# =============================================

async def _process_images_to_pdf(
    files: List[UploadFile],
    rotations: Optional[str] = '[]',
    page_size: Optional[str] = 'A4',
    fit: Optional[str] = 'inside',
    output_prefix: str = "converted_images"
) -> FileResponse:
    """Helper pemrosesan gambar ke PDF untuk menghindari duplikasi kode."""
    allowed_extensions = ('.jpg', '.jpeg', '.png')
    image_paths = []

    try:
        rotation_angles = json.loads(rotations)
    except Exception:
        rotation_angles = [0] * len(files)

    if len(rotation_angles) != len(files):
        rotation_angles = [0] * len(files)

    for file in files:
        if not file.filename.lower().endswith(allowed_extensions):
            _delete_paths(*image_paths)
            raise HTTPException(400, f"File {file.filename} bukan gambar (.jpg, .jpeg, .png)!")

        input_filename = f"{uuid.uuid4()}_{file.filename}"
        input_path = os.path.join(UPLOAD_DIR, input_filename)

        with open(input_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)

        image_paths.append(input_path)

    output_filename = f"{uuid.uuid4()}.pdf"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        convert_images_to_pdf_with_options(
            image_paths,
            output_path,
            rotation_angles,
            page_size,
            fit
        )
    except Exception as e:
        _delete_paths(*image_paths)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_paths(*image_paths)

    return FileResponse(
        output_path,
        media_type='application/pdf',
        filename=f"{output_prefix}.pdf",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# PDF TO WORD (DENGAN PILIHAN METODE)
# =============================================

@app.post("/convert/pdf-to-word")
async def pdf_to_word(
    file: UploadFile = File(...),
    method: Optional[str] = Form('auto')
):
    """Convert PDF ke Word dengan pilihan metode"""

    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, "File harus PDF!")

    input_filename = f"{uuid.uuid4()}.pdf"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.docx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        if method == 'ms_word':
            convert_pdf_to_word_ms_word(input_path, output_path)
        elif method == 'pdf2docx':
            convert_pdf_to_word_preserve(input_path, output_path)
        elif method == 'ocr':
            convert_pdf_to_word_ocr(input_path, output_path)
        else:
            convert_pdf_to_word_auto(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_path(input_path)  

    return FileResponse(
        output_path,
        media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        filename=f"converted_{os.path.splitext(file.filename)[0]}.docx",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# WORD TO PDF
# =============================================

@app.post("/convert/word-to-pdf")
async def word_to_pdf(file: UploadFile = File(...)):
    """Convert Word ke PDF"""
    if not file.filename.lower().endswith(('.docx', '.doc')):
        raise HTTPException(400, "File harus Word (.docx atau .doc)!")

    input_filename = f"{uuid.uuid4()}{os.path.splitext(file.filename)[1]}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.pdf"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        convert_word_to_pdf(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='application/pdf',
        filename=f"converted_{os.path.splitext(file.filename)[0]}.pdf",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# PDF TO MARKDOWN (AI / LLM OPTIMIZED)
# =============================================

@app.post("/convert/pdf-to-markdown")
async def pdf_to_markdown(file: UploadFile = File(...)):
    """Convert PDF ke Markdown yang teroptimasi untuk LLM (ChatGPT, Claude, Gemini)"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, "File harus PDF!")

    input_filename = f"{uuid.uuid4()}.pdf"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.md"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        convert_pdf_to_markdown(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi PDF ke Markdown gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='text/markdown; charset=utf-8',
        filename=f"{os.path.splitext(file.filename)[0]}.md",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# WORD TO MARKDOWN (AI / LLM OPTIMIZED)
# =============================================

@app.post("/convert/word-to-markdown")
async def word_to_markdown(file: UploadFile = File(...)):
    """Convert Word (.docx/.doc) ke Markdown yang teroptimasi untuk LLM"""
    if not file.filename.lower().endswith(('.docx', '.doc')):
        raise HTTPException(400, "File harus Word (.docx atau .doc)!")

    input_filename = f"{uuid.uuid4()}{os.path.splitext(file.filename)[1]}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.md"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        convert_word_to_markdown(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi Word ke Markdown gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='text/markdown; charset=utf-8',
        filename=f"{os.path.splitext(file.filename)[0]}.md",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# POWERPOINT TO PDF
# =============================================

@app.post("/convert/ppt-to-pdf")
async def ppt_to_pdf(file: UploadFile = File(...)):
    """Convert PowerPoint ke PDF"""
    if not file.filename.lower().endswith(('.pptx', '.ppt')):
        raise HTTPException(400, "File harus PowerPoint (.pptx atau .ppt)!")

    input_filename = f"{uuid.uuid4()}{os.path.splitext(file.filename)[1]}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.pdf"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        convert_ppt_to_pdf(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='application/pdf',
        filename=f"converted_{os.path.splitext(file.filename)[0]}.pdf",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# PDF TO POWERPOINT
# =============================================

@app.post("/convert/pdf-to-ppt")
async def pdf_to_ppt(file: UploadFile = File(...)):
    """Convert PDF ke PowerPoint (ekstraksi teks & layout)"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, "File harus PDF!")

    input_filename = f"{uuid.uuid4()}.pdf"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.pptx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        convert_pdf_to_ppt(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        filename=f"converted_{os.path.splitext(file.filename)[0]}.pptx",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# EXCEL TO PDF
# =============================================

@app.post("/convert/excel-to-pdf")
async def excel_to_pdf(file: UploadFile = File(...)):
    """Convert Excel ke PDF"""
    if not file.filename.lower().endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "File harus Excel (.xlsx atau .xls)!")

    input_filename = f"{uuid.uuid4()}{os.path.splitext(file.filename)[1]}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.pdf"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        convert_excel_to_pdf(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='application/pdf',
        filename=f"converted_{os.path.splitext(file.filename)[0]}.pdf",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# PDF TO EXCEL
# =============================================

@app.post("/convert/pdf-to-excel")
async def pdf_to_excel(file: UploadFile = File(...)):
    """Convert PDF ke Excel (deteksi tabel)"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, "File harus PDF!")

    input_filename = f"{uuid.uuid4()}.pdf"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.xlsx"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        convert_pdf_to_excel(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=f"converted_{os.path.splitext(file.filename)[0]}.xlsx",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# SPLIT PDF
# =============================================

@app.post("/split/pdf")
async def split_pdf_endpoint(
    file: UploadFile = File(...),
    mode: Optional[str] = Form('range'),
    ranges: Optional[str] = Form(None),
    every_n: Optional[int] = Form(None)
):
    """Split PDF berdasarkan range halaman atau tiap N halaman"""
    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, "File harus PDF!")

    input_filename = f"{uuid.uuid4()}.pdf"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    split_dir = os.path.join(OUTPUT_DIR, f"split_{uuid.uuid4()}")
    os.makedirs(split_dir, exist_ok=True)

    try:
        if mode == 'every':
            if not every_n or every_n < 1:
                raise HTTPException(400, "Jumlah halaman per file tidak valid")
            output_paths = split_pdf_every_n(input_path, split_dir, every_n)
        else:
            if not ranges or not ranges.strip():
                raise HTTPException(400, "Range halaman tidak boleh kosong")
            output_paths = split_pdf_by_ranges(input_path, split_dir, ranges)
    except HTTPException:
        _delete_path(input_path)
        _delete_path(split_dir)
        raise
    except Exception as e:
        _delete_path(input_path)
        _delete_path(split_dir)
        raise HTTPException(500, f"Split gagal: {str(e)}")

    _delete_path(input_path)

    base_name = os.path.splitext(file.filename)[0]

    if len(output_paths) == 1:
        result_path = output_paths[0]
        media_type = 'application/pdf'
        download_name = os.path.basename(result_path)
        cleanup_after_send = [split_dir]
    else:
        zip_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4()}_split.zip")
        zip_files(output_paths, zip_path)
        _delete_path(split_dir)
        result_path = zip_path
        media_type = 'application/zip'
        download_name = f"split_{base_name}.zip"
        cleanup_after_send = [zip_path]

    return FileResponse(
        result_path,
        media_type=media_type,
        filename=download_name,
        background=BackgroundTask(_delete_paths, *cleanup_after_send)
    )

# =============================================
# IMAGE TO PDF
# =============================================

@app.post("/convert/image-to-pdf")
async def image_to_pdf(
    files: List[UploadFile] = File(...),
    rotations: Optional[str] = Form('[]'),
    page_size: Optional[str] = Form('A4'),
    fit: Optional[str] = Form('inside')
):
    """Convert gambar ke PDF dengan sorting & rotasi"""
    return await _process_images_to_pdf(
        files, rotations, page_size, fit, output_prefix="converted_images"
    )

# =============================================
# COMPRESS PDF
# =============================================

@app.post("/compress/pdf")
async def compress_pdf_endpoint(
    file: UploadFile = File(...),
    compression_level: Optional[str] = Form('medium'),
    image_quality: Optional[int] = Form(None)
):
    """Kompres file PDF"""

    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, "File harus PDF!")

    valid_levels = ['low', 'medium', 'high', 'extreme']
    if compression_level not in valid_levels:
        raise HTTPException(400, f"Compression level harus salah satu dari: {', '.join(valid_levels)}")

    input_filename = f"{uuid.uuid4()}.pdf"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}_compressed.pdf"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        result = compress_pdf(
            input_path,
            output_path,
            compression_level,
            image_quality,
            remove_metadata=True
        )
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Kompresi gagal: {str(e)}")

    _delete_path(input_path)

    response = FileResponse(
        output_path,
        media_type='application/pdf',
        filename=f"compressed_{file.filename}",
        background=BackgroundTask(_delete_path, output_path)
    )
    response.headers['X-Original-Size'] = str(result['original_size'])
    response.headers['X-Compressed-Size'] = str(result['compressed_size'])
    response.headers['X-Compression-Ratio'] = str(result['ratio'])

    return response

# =============================================
# COMPRESS IMAGE
# =============================================

@app.post("/compress/image")
async def compress_image_endpoint(
    file: UploadFile = File(...),
    compression_level: Optional[str] = Form('medium'),
    quality: Optional[int] = Form(None)
):
    """Kompres file gambar (JPG, PNG, WEBP)"""

    allowed_exts = ('.jpg', '.jpeg', '.png', '.webp')
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in allowed_exts:
        raise HTTPException(400, "File harus gambar (.jpg, .jpeg, .png, .webp)!")

    valid_levels = ['low', 'medium', 'high', 'extreme']
    if compression_level not in valid_levels:
        raise HTTPException(400, f"Compression level harus salah satu dari: {', '.join(valid_levels)}")

    input_filename = f"{uuid.uuid4()}{ext}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}_compressed{ext}"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        result = compress_image(
            input_path,
            output_path,
            compression_level=compression_level,
            quality=quality
        )
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Kompresi gambar gagal: {str(e)}")

    _delete_path(input_path)

    media_type_map = {
        '.jpg': 'image/jpeg',
        '.jpeg': 'image/jpeg',
        '.png': 'image/png',
        '.webp': 'image/webp'
    }
    media_type = media_type_map.get(ext, 'image/jpeg')

    response = FileResponse(
        output_path,
        media_type=media_type,
        filename=f"compressed_{file.filename}",
        background=BackgroundTask(_delete_path, output_path)
    )
    response.headers['X-Original-Size'] = str(result['original_size'])
    response.headers['X-Compressed-Size'] = str(result['compressed_size'])
    response.headers['X-Compression-Ratio'] = str(result['ratio'])

    return response

# =============================================
# COMPRESS IMAGE TO TARGET SIZE (Standar Administrasi)
# =============================================

@app.post("/compress/image-to-target")
async def compress_image_to_target_endpoint(
    file: UploadFile = File(...),
    target_kb: Optional[int] = Form(1000)
):
    """Kompres gambar hingga di bawah target ukuran tertentu (default 1000 KB = 1 MB)"""

    allowed_exts = ('.jpg', '.jpeg', '.png', '.webp')
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in allowed_exts:
        raise HTTPException(400, "File harus gambar (.jpg, .jpeg, .png, .webp)!")

    if target_kb < 50 or target_kb > 10000:
        raise HTTPException(400, "Target ukuran harus antara 50 KB dan 10000 KB.")

    input_filename = f"{uuid.uuid4()}{ext}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    # Output selalu JPG karena lebih hemat ukuran
    output_filename = f"{uuid.uuid4()}_target.jpg"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        result = compress_image_to_target(
            input_path,
            output_path,
            target_kb=target_kb
        )
        # Fungsi mungkin mengubah output_path (selalu ke .jpg)
        actual_output = result.get('output_path', output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Kompresi ke target ukuran gagal: {str(e)}")

    _delete_path(input_path)

    response = FileResponse(
        actual_output,
        media_type='image/jpeg',
        filename=f"admin_{os.path.splitext(file.filename)[0]}.jpg",
        background=BackgroundTask(_delete_path, actual_output)
    )
    response.headers['X-Original-Size'] = str(result['original_size'])
    response.headers['X-Compressed-Size'] = str(result['compressed_size'])
    response.headers['X-Compression-Ratio'] = str(result['ratio'])

    return response

# =============================================
# MERGE PDF & MERGE IMAGES
# =============================================

@app.post("/merge/pdf")
async def merge_pdf_endpoint(
    files: List[UploadFile] = File(...),
    remove_metadata: Optional[bool] = Form(True)
):
    """Gabungkan multiple file PDF"""

    pdf_paths = []

    for file in files:
        if not file.filename.lower().endswith('.pdf'):
            _delete_paths(*pdf_paths)
            raise HTTPException(400, f"File {file.filename} bukan PDF!")

        input_filename = f"{uuid.uuid4()}_{file.filename}"
        input_path = os.path.join(UPLOAD_DIR, input_filename)

        with open(input_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)

        pdf_paths.append(input_path)

    output_filename = f"{uuid.uuid4()}_merged.pdf"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        merge_pdf_files(pdf_paths, output_path, remove_metadata)
    except Exception as e:
        _delete_paths(*pdf_paths)
        raise HTTPException(500, f"Merge gagal: {str(e)}")

    _delete_paths(*pdf_paths)

    return FileResponse(
        output_path,
        media_type='application/pdf',
        filename=f"merged_{len(files)}_files.pdf",
        background=BackgroundTask(_delete_path, output_path)
    )

@app.post("/merge/images")
async def merge_images_endpoint(
    files: List[UploadFile] = File(...),
    rotations: Optional[str] = Form('[]'),
    page_size: Optional[str] = Form('A4'),
    fit: Optional[str] = Form('inside')
):
    """Gabungkan multiple gambar menjadi satu PDF (menggunakan helper bersama)"""
    return await _process_images_to_pdf(
        files, rotations, page_size, fit, output_prefix=f"merged_{len(files)}_images"
    )

# =============================================
# REMOVE WATERMARK
# =============================================

@app.post("/remove-watermark")
async def remove_watermark_endpoint(
    file: UploadFile = File(...),
    mode: Optional[str] = Form('auto'),
    inpaint_radius: Optional[int] = Form(5)
):
    """Hapus watermark dari PDF/gambar"""

    allowed_exts = ('.pdf', '.jpg', '.jpeg', '.png')
    if not file.filename.lower().endswith(allowed_exts):
        raise HTTPException(400, "File harus PDF, JPG, PNG, atau JPEG!")

    input_filename = f"{uuid.uuid4()}_{file.filename}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    ext = os.path.splitext(file.filename)[1].lower()
    if ext == '.pdf':
        output_ext = '.pdf'
    elif ext in ('.jpg', '.jpeg'):
        output_ext = '.jpg'
    else:
        output_ext = '.png'

    output_filename = f"{uuid.uuid4()}_clean{output_ext}"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    try:
        if mode == 'camscanner':
            remove_camscanner_watermark(input_path, output_path)
        elif mode == 'inpaint':
            remove_watermark_inpaint(input_path, output_path, inpaint_radius=inpaint_radius)
        elif mode in ['bottom', 'top', 'corners', 'border']:
            mode_map = {'corners': 'all_corners'}
            actual_mode = mode_map.get(mode, mode)
            remove_watermark_auto(input_path, output_path, mode=actual_mode)
        else:
            remove_camscanner_watermark(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Gagal hapus watermark: {str(e)}")

    _delete_path(input_path)

    if output_path.endswith('.pdf'):
        media_type = 'application/pdf'
    elif output_path.endswith('.png'):
        media_type = 'image/png'
    else:
        media_type = 'image/jpeg'

    return FileResponse(
        output_path,
        media_type=media_type,
        filename=f"clean_{file.filename}",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# SCAN & DETECT DOCUMENT
# =============================================

@app.post("/scan-document")
async def scan_document_endpoint(
    file: UploadFile = File(...),
    crop_x: Optional[int] = Form(0),
    crop_y: Optional[int] = Form(0),
    crop_w: Optional[int] = Form(0),
    crop_h: Optional[int] = Form(0),
    corners: Optional[str] = Form(None)
):
    """Scan dokumen dengan crop atau 4-point perspective transform"""

    allowed_exts = ('.jpg', '.jpeg', '.png')
    if not file.filename.lower().endswith(allowed_exts):
        raise HTTPException(400, "File harus JPG, PNG, atau JPEG!")

    input_filename = f"{uuid.uuid4()}_{file.filename}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}_scanned.jpg"
    output_path = os.path.join(OUTPUT_DIR, output_filename)

    # Parsing 4 titik sudut (JSON string [[x,y], [x,y], ...])
    corners_list = None
    if corners:
        try:
            import json
            parsed = json.loads(corners)
            if isinstance(parsed, list) and len(parsed) == 4:
                corners_list = [[float(p[0]), float(p[1])] for p in parsed]
        except Exception as e:
            print(f"Format corners tidak valid: {e}")
            corners_list = None

    try:
        scan_document(
            input_path,
            output_path,
            crop_x, crop_y, crop_w, crop_h,
            corners=corners_list
        )
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Scan gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='image/jpeg',
        filename=f"scanned_{file.filename}",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# PDF TO IMAGES
# =============================================

@app.post("/convert/pdf-to-image")
async def pdf_to_image_endpoint(
    file: UploadFile = File(...),
    format: Optional[str] = Form('jpg'),
    dpi: Optional[int] = Form(150),
    pages: Optional[str] = Form('all')
):
    """Konversi halaman PDF menjadi gambar (JPG/PNG). Multi-page otomatis dikemas dalam ZIP."""

    if not file.filename.lower().endswith('.pdf'):
        raise HTTPException(400, "File harus PDF!")

    valid_formats = ['jpg', 'jpeg', 'png']
    if format.lower() not in valid_formats:
        raise HTTPException(400, "Format harus jpg atau png!")

    if dpi not in [72, 96, 150, 200, 300]:
        dpi = 150

    input_filename = f"{uuid.uuid4()}.pdf"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_subdir = os.path.join(OUTPUT_DIR, f"pdf_img_{uuid.uuid4()}")
    os.makedirs(output_subdir, exist_ok=True)

    try:
        output_paths = convert_pdf_to_images(
            input_path,
            output_subdir,
            image_format=format,
            dpi=dpi,
            pages=pages
        )
    except Exception as e:
        _delete_path(input_path)
        _delete_path(output_subdir)
        raise HTTPException(500, f"Konversi PDF ke gambar gagal: {str(e)}")

    _delete_path(input_path)

    base_name = os.path.splitext(file.filename)[0]

    if len(output_paths) == 1:
        result_path = output_paths[0]
        ext = os.path.splitext(result_path)[1].lower()
        media_type = 'image/jpeg' if ext in ('.jpg', '.jpeg') else 'image/png'
        download_name = f"{base_name}_page1{ext}"
        cleanup_paths = [output_subdir]
        return FileResponse(
            result_path,
            media_type=media_type,
            filename=download_name,
            background=BackgroundTask(_delete_paths, *cleanup_paths)
        )
    else:
        zip_path = os.path.join(OUTPUT_DIR, f"{uuid.uuid4()}_{base_name}_images.zip")
        zip_files(output_paths, zip_path)
        _delete_path(output_subdir)
        return FileResponse(
            zip_path,
            media_type='application/zip',
            filename=f"{base_name}_images.zip",
            background=BackgroundTask(_delete_path, zip_path)
        )

@app.post("/detect-boundary")
async def detect_boundary_endpoint(file: UploadFile = File(...)):
    """Deteksi boundary dokumen untuk preview"""

    if not file.filename.lower().endswith(('.jpg', '.jpeg', '.png')):
        raise HTTPException(400, "File harus gambar!")

    input_filename = f"{uuid.uuid4()}_{file.filename}"
    input_path = os.path.join(UPLOAD_DIR, input_filename)

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    try:
        boundary = get_document_boundary(input_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Deteksi gagal: {str(e)}")
    finally:
        import gc
        gc.collect()

    _delete_path(input_path)

    return boundary

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="0.0.0.0", port=8000, reload=True)
