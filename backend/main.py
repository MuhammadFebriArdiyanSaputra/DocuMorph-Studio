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
    convert_pdf_to_word,
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
    compress_image,
    compress_pdf,
    merge_pdf_files,
    merge_images_to_pdf,
    remove_watermark_auto,
    remove_camscanner_watermark,
    remove_watermark_inpaint,
    scan_document,
    get_document_boundary,
    split_pdf_by_ranges,
    split_pdf_every_n,
    zip_files
)
from utils import cleanup_old_files

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
    os.makedirs("uploads", exist_ok=True)
    os.makedirs("outputs", exist_ok=True)
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
)

# Serve frontend
app.mount("/static", StaticFiles(directory="../frontend"), name="static")

@app.get("/")
async def root():
    return FileResponse("../frontend/index.html")

# =============================================
# PDF TO WORD (DENGAN PILIHAN METODE)
# =============================================

@app.post("/convert/pdf-to-word")
async def pdf_to_word(
    file: UploadFile = File(...),
    method: Optional[str] = Form('auto')
):
    """Convert PDF ke Word dengan pilihan metode"""

    if not file.filename.endswith('.pdf'):
        raise HTTPException(400, "File harus PDF!")

    input_filename = f"{uuid.uuid4()}.pdf"
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.docx"
    output_path = f"outputs/{output_filename}"

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
        filename=f"converted_{file.filename.replace('.pdf', '.docx')}",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# WORD TO PDF
# =============================================

@app.post("/convert/word-to-pdf")
async def word_to_pdf(file: UploadFile = File(...)):
    """Convert Word ke PDF"""
    if not file.filename.endswith(('.docx', '.doc')):
        raise HTTPException(400, "File harus Word (.docx atau .doc)!")

    input_filename = f"{uuid.uuid4()}{os.path.splitext(file.filename)[1]}"
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.pdf"
    output_path = f"outputs/{output_filename}"

    try:
        convert_word_to_pdf(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='application/pdf',
        filename=f"converted_{file.filename.replace('.docx', '.pdf').replace('.doc', '.pdf')}",
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
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.pdf"
    output_path = f"outputs/{output_filename}"

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
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.pptx"
    output_path = f"outputs/{output_filename}"

    try:
        convert_pdf_to_ppt(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='application/vnd.openxmlformats-officedocument.presentationml.presentation',
        filename=f"converted_{file.filename.replace('.pdf', '.pptx')}",
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
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.pdf"
    output_path = f"outputs/{output_filename}"

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
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}.xlsx"
    output_path = f"outputs/{output_filename}"

    try:
        convert_pdf_to_excel(input_path, output_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Konversi gagal: {str(e)}")

    _delete_path(input_path)

    return FileResponse(
        output_path,
        media_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        filename=f"converted_{file.filename.replace('.pdf', '.xlsx')}",
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
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    split_dir = f"outputs/split_{uuid.uuid4()}"
    os.makedirs(split_dir, exist_ok=True)

    try:
        if mode == 'every':
            if not every_n or every_n < 1:
                raise HTTPException(400, "Jumlah halaman per file tidak valid")
            output_paths = split_pdf_every_n(input_path, split_dir, every_n)
        else:
            if not ranges:
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
        # split_dir (berisi result_path) dihapus setelah file selesai terkirim
        cleanup_after_send = [split_dir]
    else:
        zip_path = f"outputs/{uuid.uuid4()}_split.zip"
        zip_files(output_paths, zip_path)
        _delete_path(split_dir)  # sudah masuk ke zip, folder asalnya tidak dibutuhkan lagi
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

    allowed_extensions = ('.jpg', '.jpeg', '.png')
    image_paths = []
    rotation_angles = []

    try:
        rotation_angles = json.loads(rotations)
    except:
        rotation_angles = [0] * len(files)

    if len(rotation_angles) != len(files):
        rotation_angles = [0] * len(files)

    for file in files:
        if not file.filename.lower().endswith(allowed_extensions):
            _delete_paths(*image_paths)
            raise HTTPException(400, f"File {file.filename} bukan gambar!")

        input_filename = f"{uuid.uuid4()}_{file.filename}"
        input_path = f"uploads/{input_filename}"

        with open(input_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)

        image_paths.append(input_path)

    output_filename = f"{uuid.uuid4()}.pdf"
    output_path = f"outputs/{output_filename}"

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
        filename=f"converted_images.pdf",
        background=BackgroundTask(_delete_path, output_path)
    )

# =============================================
# COMPRESS PDF
# =============================================

@app.post("/compress/pdf")
async def compress_pdf_endpoint(
    file: UploadFile = File(...),
    compression_level: Optional[str] = Form('medium'),
    image_quality: Optional[int] = Form(70)
):
    """Kompres file PDF"""

    if not file.filename.endswith('.pdf'):
        raise HTTPException(400, "File harus PDF!")

    valid_levels = ['low', 'medium', 'high', 'extreme']
    if compression_level not in valid_levels:
        raise HTTPException(400, f"Compression level harus salah satu dari: {', '.join(valid_levels)}")

    input_filename = f"{uuid.uuid4()}.pdf"
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}_compressed.pdf"
    output_path = f"outputs/{output_filename}"

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
# MERGE PDF
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
        input_path = f"uploads/{input_filename}"

        with open(input_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)

        pdf_paths.append(input_path)

    output_filename = f"{uuid.uuid4()}_merged.pdf"
    output_path = f"outputs/{output_filename}"

    try:
        result = merge_pdf_files(pdf_paths, output_path, remove_metadata)
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
    """Gabungkan multiple gambar menjadi satu PDF"""

    allowed_extensions = ('.jpg', '.jpeg', '.png')
    image_paths = []
    rotation_angles = []

    try:
        rotation_angles = json.loads(rotations)
    except:
        rotation_angles = [0] * len(files)

    if len(rotation_angles) != len(files):
        rotation_angles = [0] * len(files)

    for file in files:
        if not file.filename.lower().endswith(allowed_extensions):
            _delete_paths(*image_paths)
            raise HTTPException(400, f"File {file.filename} bukan gambar!")

        input_filename = f"{uuid.uuid4()}_{file.filename}"
        input_path = f"uploads/{input_filename}"

        with open(input_path, "wb") as buffer:
            content = await file.read()
            buffer.write(content)

        image_paths.append(input_path)

    output_filename = f"{uuid.uuid4()}_merged_images.pdf"
    output_path = f"outputs/{output_filename}"

    try:
        merge_images_to_pdf(
            image_paths,
            output_path,
            page_size,
            fit,
            rotation_angles
        )
    except Exception as e:
        _delete_paths(*image_paths)
        raise HTTPException(500, f"Merge gagal: {str(e)}")

    _delete_paths(*image_paths)

    return FileResponse(
        output_path,
        media_type='application/pdf',
        filename=f"merged_{len(files)}_images.pdf",
        background=BackgroundTask(_delete_path, output_path)
    )

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
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    if file.filename.lower().endswith('.pdf'):
        output_ext = '.pdf'
    else:
        output_ext = os.path.splitext(file.filename)[1].lower()
        if output_ext == '.jpg':
            output_ext = '.jpg'
        elif output_ext == '.jpeg':
            output_ext = '.jpg'
        else:
            output_ext = '.png'

    output_filename = f"{uuid.uuid4()}_clean{output_ext}"
    output_path = f"outputs/{output_filename}"

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

    media_type = 'application/pdf' if output_path.endswith('.pdf') else 'image/jpeg'
    return FileResponse(
        output_path,
        media_type=media_type,
        filename=f"clean_{file.filename}",
        background=BackgroundTask(_delete_path, output_path)
    )

@app.post("/scan-document")
async def scan_document_endpoint(
    file: UploadFile = File(...),
    crop_x: Optional[int] = Form(0),
    crop_y: Optional[int] = Form(0),
    crop_w: Optional[int] = Form(0),
    crop_h: Optional[int] = Form(0)
):
    """Scan dokumen dengan crop"""

    allowed_exts = ('.jpg', '.jpeg', '.png')
    if not file.filename.lower().endswith(allowed_exts):
        raise HTTPException(400, "File harus JPG, PNG, atau JPEG!")

    input_filename = f"{uuid.uuid4()}_{file.filename}"
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    output_filename = f"{uuid.uuid4()}_scanned.jpg"
    output_path = f"outputs/{output_filename}"

    try:
        scan_document(
            input_path,
            output_path,
            crop_x, crop_y, crop_w, crop_h
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

@app.post("/detect-boundary")
async def detect_boundary_endpoint(file: UploadFile = File(...)):
    """Deteksi boundary dokumen untuk preview"""

    if not file.filename.lower().endswith(('.jpg', '.jpeg', '.png')):
        raise HTTPException(400, "File harus gambar!")

    input_filename = f"{uuid.uuid4()}_{file.filename}"
    input_path = f"uploads/{input_filename}"

    with open(input_path, "wb") as buffer:
        content = await file.read()
        buffer.write(content)

    try:
        boundary = get_document_boundary(input_path)
    except Exception as e:
        _delete_path(input_path)
        raise HTTPException(500, f"Deteksi gagal: {str(e)}")

    _delete_path(input_path)

    return boundary

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)