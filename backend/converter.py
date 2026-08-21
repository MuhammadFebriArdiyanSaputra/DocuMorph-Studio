from pdf2docx import Converter
import os
import sys
import pythoncom
import win32com.client
from docx import Document
from docx.shared import Cm, Inches, Emu
from docx.enum.section import WD_ORIENT
from PIL import Image
import img2pdf
import json
from reportlab.lib.pagesizes import A4, letter, legal
from reportlab.lib.utils import ImageReader
from reportlab.pdfgen import canvas
import subprocess
import tempfile
import shutil
import zipfile
import cv2
import numpy as np

# =============================================
# MARGIN & FORMATTING HELPERS
# =============================================

def get_docx_margins(docx_path):
    """Ambil margin dari file DOCX asli"""
    doc = Document(docx_path)
    margins = []
    for section in doc.sections:
        margins.append({
            'top': section.top_margin,
            'bottom': section.bottom_margin,
            'left': section.left_margin,
            'right': section.right_margin,
            'page_width': section.page_width,
            'page_height': section.page_height,
            'orientation': section.orientation
        })
    return margins

def set_docx_margins(docx_path, margins):
    """Set margin pada file DOCX"""
    doc = Document(docx_path)
    for i, section in enumerate(doc.sections):
        if i < len(margins):
            m = margins[i]
            section.top_margin = m['top']
            section.bottom_margin = m['bottom']
            section.left_margin = m['left']
            section.right_margin = m['right']
            section.page_width = m['page_width']
            section.page_height = m['page_height']
            section.orientation = m['orientation']
    doc.save(docx_path)

# =============================================
# PDF TO WORD CONVERTERS (MULTI-METHOD)
# =============================================

def convert_pdf_to_word_preserve(input_path: str, output_path: str):
    """
    Convert PDF ke Word menggunakan pdf2docx
    Metode ini cepat tapi kurang akurat untuk tabel kompleks
    """
    cv = Converter(input_path)
    cv.convert(
        output_path,
        start=0,
        end=None,
        pages=None,
        multi_processing=False,
        min_section_height=20.0,
        debug=False
    )
    cv.close()

    # Normalisasi margin hasil konversi
    doc = Document(output_path)
    for section in doc.sections:
        if section.left_margin < Cm(1):
            section.left_margin = Cm(2.54)
        if section.right_margin < Cm(1):
            section.right_margin = Cm(2.54)

    doc.save(output_path)
    return output_path

def convert_pdf_to_word_ms_word(input_path: str, output_path: str):
    """
    Convert PDF ke Word menggunakan Microsoft Word langsung
    Metode PALING AKURAT untuk Windows + Microsoft Word
    Karena Word punya fitur import PDF bawaan
    """
    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        try:
            doc = word.Documents.Open(os.path.abspath(input_path))
            doc.SaveAs2(
                os.path.abspath(output_path),
                FileFormat=16  # wdFormatXMLDocument (DOCX)
            )
            doc.Close()

        finally:
            word.Quit()
            pythoncom.CoUninitialize()

        return output_path

    except Exception as e:
        error_msg = str(e)
        if "Word.Application" in error_msg:
            raise Exception("Microsoft Word tidak terinstall")
        else:
            raise Exception(f"Gagal konversi dengan Word: {error_msg}")

def convert_pdf_to_word_ocr(input_path: str, output_path: str):
    """
    Convert PDF scan ke Word menggunakan OCR
    """
    try:
        import pytesseract
        from pdf2image import convert_from_path
    except ImportError:
        raise Exception(
            "pytesseract atau pdf2image tidak terinstall. "
            "Jalankan: pip install pytesseract pdf2image"
        )

    images = convert_from_path(input_path, dpi=300)
    doc = Document()

    for i, image in enumerate(images):
        text = pytesseract.image_to_string(image, lang='ind+eng')
        if i > 0:
            doc.add_page_break()
        doc.add_paragraph(text)

    doc.save(output_path)
    return output_path

def convert_pdf_to_word_auto(input_path: str, output_path: str):
    """
    Coba beberapa metode dan pilih yang terbaik
    Prioritas: Microsoft Word > pdf2docx
    """
    errors = []

    if sys.platform == "win32":
        try:
            print("Mencoba metode: Microsoft Word...")
            convert_pdf_to_word_ms_word(input_path, output_path)
            print("Berhasil dengan Microsoft Word!")
            return output_path
        except Exception as e:
            errors.append(f"Word: {str(e)}")
            print(f"Word gagal: {str(e)}")

    try:
        print("Mencoba metode: pdf2docx...")
        convert_pdf_to_word_preserve(input_path, output_path)
        print("Berhasil dengan pdf2docx!")
        return output_path
    except Exception as e:
        errors.append(f"pdf2docx: {str(e)}")
        print(f"pdf2docx gagal: {str(e)}")

    raise Exception("Semua metode konversi gagal:\n" + "\n".join(errors))

def convert_pdf_to_word(input_path: str, output_path: str):
    """Convert PDF ke Word dengan metode terbaik yang tersedia"""
    return convert_pdf_to_word_auto(input_path, output_path)

# =============================================
# WORD TO PDF CONVERTER (WINDOWS)
# =============================================

def convert_word_to_pdf_windows(input_path: str, output_path: str):
    """Convert Word ke PDF menggunakan Microsoft Word COM automation"""
    try:
        pythoncom.CoInitialize()
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False

        try:
            doc = word.Documents.Open(os.path.abspath(input_path))
            doc.ExportAsFixedFormat(
                os.path.abspath(output_path),
                17,  # wdFormatPDF
                OptimizeFor=0,
                CreateBookmarks=1,
                DocStructureTags=True,
                BitmapMissingFonts=True,
                UseISO19005_1=False
            )
            doc.Close()

        finally:
            word.Quit()
            pythoncom.CoUninitialize()

        return output_path

    except Exception as e:
        error_msg = str(e)
        if "No module named 'win32com'" in error_msg:
            raise Exception("pywin32 tidak terinstall. Jalankan: pip install pywin32")
        elif "Word.Application" in error_msg:
            raise Exception("Microsoft Word tidak terinstall")
        else:
            raise Exception(f"Gagal konversi Word ke PDF: {error_msg}")

def convert_word_to_pdf(input_path: str, output_path: str):
    """Convert Word ke PDF - Windows dengan Microsoft Word"""
    if sys.platform != "win32":
        raise Exception("Fitur Word ke PDF saat ini hanya mendukung Windows dengan Microsoft Word")
    return convert_word_to_pdf_windows(input_path, output_path)

# =============================================
# POWERPOINT TO PDF CONVERTER (WINDOWS)
# =============================================

def convert_ppt_to_pdf_windows(input_path: str, output_path: str):
    """Convert PowerPoint ke PDF menggunakan Microsoft PowerPoint COM automation"""
    try:
        pythoncom.CoInitialize()
        powerpoint = win32com.client.Dispatch("PowerPoint.Application")

        try:
            presentation = powerpoint.Presentations.Open(
                os.path.abspath(input_path),
                WithWindow=False
            )
            presentation.SaveAs(os.path.abspath(output_path), 32)  # ppSaveAsPDF
            presentation.Close()

        finally:
            powerpoint.Quit()
            pythoncom.CoUninitialize()

        return output_path

    except Exception as e:
        error_msg = str(e)
        if "PowerPoint.Application" in error_msg:
            raise Exception("Microsoft PowerPoint tidak terinstall")
        else:
            raise Exception(f"Gagal konversi PowerPoint ke PDF: {error_msg}")

def convert_ppt_to_pdf(input_path: str, output_path: str):
    """Convert PowerPoint ke PDF - Windows dengan Microsoft PowerPoint"""
    if sys.platform != "win32":
        raise Exception("Fitur PowerPoint ke PDF saat ini hanya mendukung Windows dengan Microsoft PowerPoint")
    return convert_ppt_to_pdf_windows(input_path, output_path)

# =============================================
# EXCEL TO PDF CONVERTER (WINDOWS)
# =============================================

def convert_excel_to_pdf_windows(input_path: str, output_path: str):
    """Convert Excel ke PDF menggunakan Microsoft Excel COM automation"""
    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        try:
            workbook = excel.Workbooks.Open(os.path.abspath(input_path))
            workbook.ExportAsFixedFormat(
                0,  # xlTypePDF
                os.path.abspath(output_path)
            )
            workbook.Close(False)

        finally:
            excel.Quit()
            pythoncom.CoUninitialize()

        return output_path

    except Exception as e:
        error_msg = str(e)
        if "Excel.Application" in error_msg:
            raise Exception("Microsoft Excel tidak terinstall")
        else:
            raise Exception(f"Gagal konversi Excel ke PDF: {error_msg}")

def convert_excel_to_pdf(input_path: str, output_path: str):
    """Convert Excel ke PDF - Windows dengan Microsoft Excel"""
    if sys.platform != "win32":
        raise Exception("Fitur Excel ke PDF saat ini hanya mendukung Windows dengan Microsoft Excel")
    return convert_excel_to_pdf_windows(input_path, output_path)

# =============================================
# PDF TO POWERPOINT (EKSTRAK TEKS & LAYOUT)
# =============================================

def convert_pdf_to_ppt(input_path: str, output_path: str):
    """
    Convert PDF ke PowerPoint dengan ekstraksi teks & layout (bukan rasterisasi).
    Setiap halaman jadi satu slide; posisi, ukuran font, warna, dan gaya
    (bold/italic) teks diusahakan mengikuti tata letak asli sehingga tetap
    bisa diedit di PowerPoint. Elemen non-teks (gambar/vector kompleks)
    tidak ikut terangkut - hasil terbaik untuk PDF yang didominasi teks.
    Semua halaman memakai ukuran slide dari halaman pertama (keterbatasan
    format .pptx yang tidak mendukung ukuran berbeda per slide).
    """
    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise Exception("PyMuPDF tidak terinstall. Jalankan: pip install PyMuPDF")

    try:
        from pptx import Presentation
        from pptx.util import Emu, Pt
        from pptx.dml.color import RGBColor
    except ImportError:
        raise Exception("python-pptx tidak terinstall. Jalankan: pip install python-pptx")

    EMU_PER_POINT = 12700  # 1 point = 12700 EMU

    doc = fitz.open(input_path)
    if len(doc) == 0:
        raise Exception("PDF tidak memiliki halaman")

    first_page = doc[0]
    page_w_pt = first_page.rect.width
    page_h_pt = first_page.rect.height

    prs = Presentation()
    prs.slide_width = Emu(int(page_w_pt * EMU_PER_POINT))
    prs.slide_height = Emu(int(page_h_pt * EMU_PER_POINT))
    blank_layout = prs.slide_layouts[6]

    for page_index in range(len(doc)):
        page = doc[page_index]
        slide = prs.slides.add_slide(blank_layout)

        text_dict = page.get_text("dict")

        for block in text_dict.get("blocks", []):
            if block.get("type") != 0:
                continue  # lewati blok non-teks (gambar)

            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = span.get("text", "")
                    if not text.strip():
                        continue

                    x0, y0, x1, y1 = span["bbox"]
                    width_pt = max(x1 - x0, 10)
                    height_pt = max(y1 - y0, 10)

                    textbox = slide.shapes.add_textbox(
                        Emu(int(x0 * EMU_PER_POINT)),
                        Emu(int(y0 * EMU_PER_POINT)),
                        Emu(int(width_pt * EMU_PER_POINT)),
                        Emu(int(height_pt * EMU_PER_POINT))
                    )
                    tf = textbox.text_frame
                    tf.word_wrap = True
                    tf.margin_left = 0
                    tf.margin_right = 0
                    tf.margin_top = 0
                    tf.margin_bottom = 0

                    p = tf.paragraphs[0]
                    run = p.add_run()
                    run.text = text

                    font_size = span.get("size", 12)
                    run.font.size = Pt(max(font_size, 6))

                    color_int = span.get("color", 0)
                    r = (color_int >> 16) & 255
                    g = (color_int >> 8) & 255
                    b = color_int & 255
                    run.font.color.rgb = RGBColor(r, g, b)

                    flags = span.get("flags", 0)
                    run.font.bold = bool(flags & (1 << 4))
                    run.font.italic = bool(flags & (1 << 1))

    doc.close()
    prs.save(output_path)
    return output_path

# =============================================
# PDF TO EXCEL (DETEKSI TABEL AKURAT)
# =============================================

def convert_pdf_to_excel(input_path: str, output_path: str):
    """
    Convert PDF ke Excel dengan deteksi tabel (camelot-py) - lebih akurat
    untuk tabel berbingkai jelas dibanding ekstraksi teks polos, namun
    proses lebih lambat. Setiap tabel yang terdeteksi ditulis ke sheet
    terpisah. Butuh Ghostscript terinstall di sistem untuk mode 'lattice'.
    """
    try:
        import camelot
    except ImportError:
        raise Exception(
            "camelot-py tidak terinstall. Jalankan: pip install camelot-py[cv] "
            "(juga butuh Ghostscript terinstall di sistem)"
        )

    from openpyxl import Workbook
    from openpyxl.utils.dataframe import dataframe_to_rows

    tables = camelot.read_pdf(input_path, pages='all', flavor='lattice')

    if len(tables) == 0:
        # Fallback ke flavor 'stream' untuk tabel tanpa garis bingkai jelas
        tables = camelot.read_pdf(input_path, pages='all', flavor='stream')

    if len(tables) == 0:
        raise Exception("Tidak ada tabel yang terdeteksi di dalam PDF")

    wb = Workbook()
    wb.remove(wb.active)

    for i, table in enumerate(tables, start=1):
        sheet_name = f"Tabel_{i}_Hal{table.page}"[:31]  # batas nama sheet Excel
        ws = wb.create_sheet(title=sheet_name)

        for row in dataframe_to_rows(table.df, index=False, header=False):
            ws.append(row)

    wb.save(output_path)
    return output_path

# =============================================
# SPLIT PDF
# =============================================

def split_pdf_by_ranges(input_path: str, output_dir: str, ranges: str):
    """
    Split PDF berdasarkan range halaman.
    ranges: string seperti "1-3,5,7-9" (1-indexed, inklusif)
    Return: list path file PDF hasil split, satu file per range
    """
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    output_paths = []

    groups = [g.strip() for g in ranges.split(',') if g.strip()]
    if not groups:
        raise Exception("Range halaman tidak valid")

    for idx, group in enumerate(groups, start=1):
        try:
            if '-' in group:
                start_str, end_str = group.split('-')
                start, end = int(start_str), int(end_str)
            else:
                start = end = int(group)
        except ValueError:
            raise Exception(f"Format range '{group}' tidak valid")

        start = max(1, start)
        end = min(total_pages, end)
        if start > end:
            continue

        writer = PdfWriter()
        for p in range(start - 1, end):
            writer.add_page(reader.pages[p])

        out_path = os.path.join(output_dir, f"split_{idx}_hal{start}-{end}.pdf")
        with open(out_path, 'wb') as f:
            writer.write(f)
        output_paths.append(out_path)

    if not output_paths:
        raise Exception("Tidak ada halaman valid untuk displit")

    return output_paths

def split_pdf_every_n(input_path: str, output_dir: str, n: int):
    """Split PDF tiap N halaman menjadi file terpisah"""
    from pypdf import PdfReader, PdfWriter

    if n < 1:
        raise Exception("Jumlah halaman per file minimal 1")

    reader = PdfReader(input_path)
    total_pages = len(reader.pages)
    output_paths = []

    part = 1
    for start in range(0, total_pages, n):
        end = min(start + n, total_pages)
        writer = PdfWriter()
        for p in range(start, end):
            writer.add_page(reader.pages[p])

        out_path = os.path.join(output_dir, f"split_bagian{part}.pdf")
        with open(out_path, 'wb') as f:
            writer.write(f)
        output_paths.append(out_path)
        part += 1

    return output_paths

def zip_files(file_paths: list, zip_output_path: str):
    """Gabungkan beberapa file jadi satu ZIP"""
    with zipfile.ZipFile(zip_output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        for fp in file_paths:
            zf.write(fp, arcname=os.path.basename(fp))
    return zip_output_path

# =============================================
# IMAGE TO PDF CONVERTER (HIGH QUALITY)
# =============================================

def convert_images_to_pdf_with_options(
    image_paths: list,
    output_path: str,
    rotation_angles: list = None,
    page_size: str = 'A4',
    fit: str = 'inside'
):
    """Convert multiple gambar ke PDF dengan kualitas tinggi"""
    if rotation_angles is None:
        rotation_angles = [0] * len(image_paths)

    if len(rotation_angles) < len(image_paths):
        rotation_angles += [0] * (len(image_paths) - len(rotation_angles))

    page_sizes = {
        'A4': A4,
        'Letter': letter,
        'Legal': legal
    }

    if page_size == 'original':
        return _convert_images_original_size_high_quality(image_paths, rotation_angles, output_path)

    pagesize = page_sizes.get(page_size, A4)
    c = canvas.Canvas(output_path, pagesize=pagesize)
    page_w, page_h = pagesize

    total_images = len(image_paths)

    for i, img_path in enumerate(image_paths):
        try:
            image = Image.open(img_path)

            if image.mode in ('RGBA', 'LA', 'P'):
                rgb_image = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                rgb_image.paste(
                    image,
                    mask=image.split()[-1] if image.mode in ('RGBA', 'LA') else None
                )
                image = rgb_image
            elif image.mode != 'RGB':
                image = image.convert('RGB')

            rotation = rotation_angles[i] if i < len(rotation_angles) else 0
            if rotation > 0:
                image = image.rotate(-rotation, expand=True)

            img_w, img_h = image.size

            if fit == 'inside':
                ratio = min(page_w / img_w, page_h / img_h)
            elif fit == 'fill':
                ratio = max(page_w / img_w, page_h / img_h)
            elif fit == 'stretch':
                new_w = page_w
                new_h = page_h

                temp_path = img_path + f'_stretch_{i}.png'
                image_resized = image.resize(
                    (int(new_w), int(new_h)),
                    Image.Resampling.LANCZOS
                )
                image_resized.save(temp_path, 'PNG')
                c.drawImage(temp_path, 0, 0, page_w, page_h)

                if os.path.exists(temp_path):
                    os.remove(temp_path)

                if i < total_images - 1:
                    c.showPage()
                continue
            else:
                ratio = min(page_w / img_w, page_h / img_h)

            target_dpi = 150
            dpi_scale = target_dpi / 72.0

            new_w = img_w * ratio * dpi_scale
            new_h = img_h * ratio * dpi_scale

            x = (page_w - new_w / dpi_scale) / 2
            y = (page_h - new_h / dpi_scale) / 2

            temp_path = img_path + f'_processed_{i}.png'

            if img_w * ratio * dpi_scale < img_w:
                image_resized = image.resize(
                    (int(new_w), int(new_h)),
                    Image.Resampling.LANCZOS
                )
            else:
                image_resized = image

            image_resized.save(temp_path, 'PNG')

            c.drawImage(
                temp_path,
                x,
                y,
                new_w / dpi_scale,
                new_h / dpi_scale
            )

            if os.path.exists(temp_path):
                os.remove(temp_path)

            if i < total_images - 1:
                c.showPage()

        except Exception as e:
            print(f"Error processing image {img_path}: {str(e)}")
            continue

    c.save()
    return output_path

def _convert_images_original_size_high_quality(image_paths, rotation_angles, output_path):
    """Convert gambar ke PDF dengan ukuran original dan kualitas TINGGI"""
    processed_images = []

    for i, img_path in enumerate(image_paths):
        try:
            image = Image.open(img_path)

            if image.mode in ('RGBA', 'LA', 'P'):
                rgb_image = Image.new('RGB', image.size, (255, 255, 255))
                if image.mode == 'P':
                    image = image.convert('RGBA')
                rgb_image.paste(
                    image,
                    mask=image.split()[-1] if image.mode in ('RGBA', 'LA') else None
                )
                image = rgb_image
            elif image.mode != 'RGB':
                image = image.convert('RGB')

            rotation = rotation_angles[i] if i < len(rotation_angles) else 0
            if rotation > 0:
                image = image.rotate(-rotation, expand=True)

            temp_path = img_path + f'_original_{i}.png'
            image.save(temp_path, 'PNG')
            processed_images.append(temp_path)

        except Exception as e:
            print(f"Error processing image {img_path}: {str(e)}")
            continue

    if not processed_images:
        raise Exception("Tidak ada gambar yang bisa diproses")

    with open(output_path, 'wb') as f:
        f.write(img2pdf.convert(processed_images))

    for temp_path in processed_images:
        if os.path.exists(temp_path):
            os.remove(temp_path)

    return output_path

def convert_single_image_to_pdf(
    image_path: str,
    output_path: str,
    page_size: str = 'A4',
    fit: str = 'inside'
):
    """Convert satu gambar ke PDF"""
    return convert_images_to_pdf_with_options(
        [image_path],
        output_path,
        [0],
        page_size,
        fit
    )

# =============================================
# IMAGE COMPRESSOR
# =============================================

def compress_image(
    input_path: str,
    output_path: str,
    quality: int = 85,
    max_size: tuple = None
):
    """Kompres gambar"""
    try:
        image = Image.open(input_path)

        if image.mode in ('RGBA', 'LA', 'P'):
            rgb_image = Image.new('RGB', image.size, (255, 255, 255))
            if image.mode == 'P':
                image = image.convert('RGBA')
            rgb_image.paste(
                image,
                mask=image.split()[-1] if image.mode in ('RGBA', 'LA') else None
            )
            image = rgb_image

        if max_size:
            image.thumbnail(max_size, Image.Resampling.LANCZOS)

        ext = output_path.lower()
        if ext.endswith('.jpg') or ext.endswith('.jpeg'):
            image.save(output_path, 'JPEG', quality=quality, optimize=True)
        elif ext.endswith('.png'):
            image.save(output_path, 'PNG', optimize=True)
        elif ext.endswith('.webp'):
            image.save(output_path, 'WEBP', quality=quality)
        else:
            image.save(output_path, 'JPEG', quality=quality, optimize=True)

        return output_path

    except Exception as e:
        raise Exception(f"Gagal kompres gambar: {str(e)}")

def compress_multiple_images(image_paths, output_dir, quality=85, max_size=None):
    """Kompres multiple gambar"""
    os.makedirs(output_dir, exist_ok=True)
    output_paths = []

    for img_path in image_paths:
        filename = os.path.basename(img_path)
        name, ext = os.path.splitext(filename)
        output_path = os.path.join(output_dir, f"{name}_compressed{ext}")

        try:
            compress_image(img_path, output_path, quality, max_size)
            output_paths.append(output_path)
        except Exception as e:
            print(f"Error compressing {filename}: {str(e)}")

    return output_paths

# =============================================
# PDF COMPRESSOR
# =============================================

def compress_pdf(
    input_path: str,
    output_path: str,
    compression_level: str = 'medium',
    image_quality: int = None,
    remove_metadata: bool = True
):
    """
    Kompres file PDF: content stream + gambar embedded direkompresi ulang
    sesuai level. Gambar di dalam PDF diekstrak, di-resize sesuai skala
    level, lalu ditulis ulang dengan kualitas JPEG sesuai level - sehingga
    tiap level (low/medium/high/extreme) menghasilkan ukuran yang benar-benar
    berbeda.
    """
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        raise Exception("pypdf tidak terinstall. Jalankan: pip install --upgrade pypdf")

    from PIL import Image as PILImage

    compression_settings = {
        'low':     {'image_quality': 85, 'scale': 1.0},
        'medium':  {'image_quality': 65, 'scale': 0.85},
        'high':    {'image_quality': 45, 'scale': 0.65},
        'extreme': {'image_quality': 25, 'scale': 0.45},
    }

    settings = compression_settings.get(compression_level, compression_settings['medium'])
    quality = image_quality if image_quality else settings['image_quality']
    scale = settings['scale']

    reader = PdfReader(input_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    for page in writer.pages:
        try:
            page.compress_content_streams()
        except Exception:
            pass

        try:
            for img in page.images:
                pil_img = img.image
                if pil_img.mode not in ('RGB', 'L'):
                    pil_img = pil_img.convert('RGB')
                if scale < 1.0:
                    new_w = max(1, int(pil_img.width * scale))
                    new_h = max(1, int(pil_img.height * scale))
                    pil_img = pil_img.resize((new_w, new_h), PILImage.Resampling.LANCZOS)
                img.replace(pil_img, quality=quality)
        except Exception as e:
            print(f"Gagal kompres gambar di halaman: {e}")

    if remove_metadata or settings.get('remove_metadata', True):
        try:
            writer.add_metadata({})
        except Exception:
            pass

    with open(output_path, 'wb') as output_file:
        writer.write(output_file)

    original_size = os.path.getsize(input_path)
    compressed_size = os.path.getsize(output_path)

    if compressed_size >= original_size:
        shutil.copy2(input_path, output_path)
        compressed_size = original_size
        print("Kompresi tidak mengurangi ukuran, menggunakan file original.")

    ratio = (1 - compressed_size / original_size) * 100 if original_size else 0

    print(f"Kompresi selesai:")
    print(f"  Original: {format_file_size(original_size)}")
    print(f"  Compressed: {format_file_size(compressed_size)}")
    print(f"  Saved: {ratio:.1f}%")

    return {
        'original_size': original_size,
        'compressed_size': compressed_size,
        'ratio': round(ratio, 1)
    }

def get_pdf_info(input_path: str):
    """Dapatkan informasi file PDF"""
    try:
        from pypdf import PdfReader
    except ImportError:
        raise Exception("pypdf tidak terinstall")

    reader = PdfReader(input_path)

    info = {
        'pages': len(reader.pages),
        'size': os.path.getsize(input_path),
        'size_formatted': format_file_size(os.path.getsize(input_path)),
        'metadata': {}
    }

    if reader.metadata:
        for key in ['/Title', '/Author', '/Subject', '/Creator', '/Producer']:
            if key in reader.metadata:
                info['metadata'][key.replace('/', '')] = reader.metadata[key]

    return info

def format_file_size(size_bytes):
    """Format ukuran file"""
    if size_bytes == 0:
        return "0 B"

    units = ['B', 'KB', 'MB', 'GB']
    i = 0
    size = float(size_bytes)

    while size >= 1024 and i < len(units) - 1:
        size /= 1024
        i += 1

    return f"{size:.2f} {units[i]}"

# =============================================
# MERGE PDF FUNCTIONS
# =============================================

def merge_pdf_files(pdf_paths: list, output_path: str, remove_metadata: bool = True):
    """Gabungkan multiple file PDF menjadi satu PDF"""
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        raise Exception("pypdf tidak terinstall")

    writer = PdfWriter()
    total_pages = 0

    for pdf_path in pdf_paths:
        try:
            reader = PdfReader(pdf_path)

            for page in reader.pages:
                try:
                    page.compress_content_streams()
                except:
                    pass
                writer.add_page(page)
                total_pages += 1

        except Exception as e:
            print(f"Error processing {pdf_path}: {str(e)}")
            continue

    if total_pages == 0:
        raise Exception("Tidak ada halaman yang bisa digabungkan")

    if remove_metadata:
        try:
            writer.add_metadata({})
        except:
            pass

    with open(output_path, 'wb') as output_file:
        writer.write(output_file)

    return {
        'total_pages': total_pages,
        'output_size': os.path.getsize(output_path)
    }

def merge_images_to_pdf(
    image_paths: list,
    output_path: str,
    page_size: str = 'A4',
    fit: str = 'inside',
    rotation_angles: list = None
):
    """Gabungkan multiple gambar menjadi satu PDF"""
    if rotation_angles is None:
        rotation_angles = [0] * len(image_paths)

    return convert_images_to_pdf_with_options(
        image_paths,
        output_path,
        rotation_angles,
        page_size,
        fit
    )

# =============================================
# REMOVE WATERMARK FUNCTIONS
# =============================================

def remove_watermark_auto(
    input_path: str,
    output_path: str,
    mode: str = 'bottom',
    threshold: float = 0.5
):
    """Hapus watermark secara otomatis dari gambar/PDF"""
    if input_path.lower().endswith('.pdf'):
        return _remove_watermark_from_pdf(input_path, output_path, mode, threshold)
    else:
        return _remove_watermark_from_image(input_path, output_path, mode, threshold)

def _remove_watermark_from_image(image_path, output_path, mode='bottom', threshold=0.5):
    """Hapus watermark dari satu gambar"""
    img = cv2.imread(image_path)
    if img is None:
        raise Exception("Gagal membaca gambar")

    h, w = img.shape[:2]
    result = img.copy()

    if mode == 'bottom':
        crop_height = int(h * 0.15)
        result[h-crop_height:, :] = (255, 255, 255)
    elif mode == 'top':
        crop_height = int(h * 0.15)
        result[:crop_height, :] = (255, 255, 255)
    elif mode == 'all_corners':
        corner_h = int(h * 0.1)
        corner_w = int(w * 0.1)
        result[:corner_h, :corner_w] = (255, 255, 255)
        result[:corner_h, w-corner_w:] = (255, 255, 255)
        result[h-corner_h:, :corner_w] = (255, 255, 255)
        result[h-corner_h:, w-corner_w:] = (255, 255, 255)
    elif mode == 'border':
        border_size = int(min(h, w) * 0.03)
        result[:border_size, :] = (255, 255, 255)
        result[h-border_size:, :] = (255, 255, 255)
        result[:, :border_size] = (255, 255, 255)
        result[:, w-border_size:] = (255, 255, 255)

    cv2.imwrite(output_path, result, [cv2.IMWRITE_JPEG_QUALITY, 85])
    return output_path

def _remove_watermark_from_pdf(pdf_path, output_path, mode='bottom', threshold=0.5):
    """
    Hapus watermark dari semua halaman PDF TANPA rasterisasi - watermark
    ditutup dengan kotak putih langsung di halaman vektor asli, sehingga
    ukuran file tidak membengkak.
    """
    import fitz  

    doc = fitz.open(pdf_path)
    white = (1, 1, 1)

    for page in doc:
        rect = page.rect
        w, h = rect.width, rect.height

        if mode == 'bottom':
            crop_h = h * 0.12
            page.draw_rect(fitz.Rect(0, h - crop_h, w, h), color=None, fill=white, overlay=True)
        elif mode == 'top':
            crop_h = h * 0.12
            page.draw_rect(fitz.Rect(0, 0, w, crop_h), color=None, fill=white, overlay=True)
        elif mode == 'all_corners':
            ch, cw = h * 0.08, w * 0.08
            page.draw_rect(fitz.Rect(0, 0, cw, ch), color=None, fill=white, overlay=True)
            page.draw_rect(fitz.Rect(w - cw, 0, w, ch), color=None, fill=white, overlay=True)
            page.draw_rect(fitz.Rect(0, h - ch, cw, h), color=None, fill=white, overlay=True)
            page.draw_rect(fitz.Rect(w - cw, h - ch, w, h), color=None, fill=white, overlay=True)
        elif mode == 'border':
            bs = min(h, w) * 0.02
            page.draw_rect(fitz.Rect(0, 0, w, bs), color=None, fill=white, overlay=True)
            page.draw_rect(fitz.Rect(0, h - bs, w, h), color=None, fill=white, overlay=True)
            page.draw_rect(fitz.Rect(0, 0, bs, h), color=None, fill=white, overlay=True)
            page.draw_rect(fitz.Rect(w - bs, 0, w, h), color=None, fill=white, overlay=True)

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    return output_path

def remove_watermark_inpaint(
    input_path: str,
    output_path: str,
    mask_path: str = None,
    inpaint_radius: int = 5
):
    """Hapus watermark menggunakan OpenCV inpainting. Mendukung PDF dan gambar"""
    if input_path.lower().endswith('.pdf'):
        return _remove_watermark_inpaint_from_pdf(input_path, output_path, inpaint_radius)
    else:
        return _remove_watermark_inpaint_from_image(input_path, output_path, mask_path, inpaint_radius)

def _remove_watermark_inpaint_from_image(image_path, output_path, mask_path=None, inpaint_radius=5):
    """Inpainting untuk satu gambar"""
    img = cv2.imread(image_path)
    if img is None:
        raise Exception("Gagal membaca gambar. Pastikan file adalah JPG/PNG yang valid.")

    if mask_path:
        mask = cv2.imread(mask_path, cv2.IMREAD_GRAYSCALE)
    else:
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, mask = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)
        mask = cv2.bitwise_not(mask)

    if mask.shape != img.shape[:2]:
        mask = cv2.resize(mask, (img.shape[1], img.shape[0]))

    result = cv2.inpaint(img, mask, inpaint_radius, cv2.INPAINT_TELEA)
    cv2.imwrite(output_path, result, [cv2.IMWRITE_JPEG_QUALITY, 85])
    return output_path

def _remove_watermark_inpaint_from_pdf(pdf_path, output_path, inpaint_radius=5):
    """
    Inpainting untuk semua halaman PDF, dengan kontrol ukuran output.
    File gambar sementara ditulis ke folder temp OS (tempfile.TemporaryDirectory)
    yang otomatis dihapus Python sendiri begitu proses selesai - baik berhasil
    maupun gagal - bukan ke working directory yang tidak ikut tersapu proses
    pembersihan file lain.
    """
    try:
        import fitz  
    except ImportError:
        raise Exception("PyMuPDF tidak terinstall. Jalankan: pip install PyMuPDF")

    doc = fitz.open(pdf_path)
    output_doc = fitz.open()

    with tempfile.TemporaryDirectory() as tmp_dir:
        for page_num in range(len(doc)):
            page = doc[page_num]

            pix = page.get_pixmap(dpi=150)
            img_data = pix.tobytes("png")
            nparr = np.frombuffer(img_data, np.uint8)
            img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
            if img is None:
                continue

            gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
            _, mask = cv2.threshold(gray, 200, 255, cv2.THRESH_BINARY)
            mask = cv2.bitwise_not(mask)

            result = cv2.inpaint(img, mask, inpaint_radius, cv2.INPAINT_TELEA)

            temp_path = os.path.join(tmp_dir, f"page_{page_num}.jpg")
            cv2.imwrite(temp_path, result, [cv2.IMWRITE_JPEG_QUALITY, 75])

            new_page = output_doc.new_page(width=page.rect.width, height=page.rect.height)
            new_page.insert_image(new_page.rect, filename=temp_path)

    output_doc.save(output_path, garbage=4, deflate=True)
    output_doc.close()
    doc.close()

    return output_path

def remove_camscanner_watermark(
    input_path: str,
    output_path: str
):
    """
    Khusus untuk dokumen CamScanner:
    1. Hapus border/frame hitam
    2. Hapus watermark "Scanned by CamScanner" di bagian bawah
    3. Bersihkan background
    """
    if input_path.lower().endswith('.pdf'):
        return _remove_camscanner_from_pdf(input_path, output_path)
    else:
        return _remove_camscanner_from_image(input_path, output_path)

def _remove_camscanner_from_image(image_path, output_path):
    """Hapus watermark CamScanner dari gambar"""
    img = cv2.imread(image_path)
    if img is None:
        raise Exception("Gagal membaca gambar")

    h, w = img.shape[:2]

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 50, 255, cv2.THRESH_BINARY)
    contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    if contours:
        largest = max(contours, key=cv2.contourArea)
        x, y, bw, bh = cv2.boundingRect(largest)
        img = img[y:y+bh, x:x+bw]
        h, w = img.shape[:2]

    crop_height = int(h * 0.08)
    img[h-crop_height:, :] = (255, 255, 255)

    img = cv2.convertScaleAbs(img, alpha=1.1, beta=10)

    cv2.imwrite(output_path, img, [cv2.IMWRITE_JPEG_QUALITY, 85])
    return output_path

def _remove_camscanner_from_pdf(pdf_path, output_path):
    """
    Hapus border + watermark CamScanner TANPA rasterisasi konten asli.
    Border dideteksi lewat render resolusi rendah (hanya untuk kalkulasi
    posisi), lalu diterapkan sebagai crop pada halaman vektor asli.
    """
    import fitz

    doc = fitz.open(pdf_path)
    white = (1, 1, 1)

    for page in doc:
        rect = page.rect
        w_pt, h_pt = rect.width, rect.height

        pix = page.get_pixmap(dpi=100)
        img_data = pix.tobytes("png")
        nparr = np.frombuffer(img_data, np.uint8)
        img = cv2.imdecode(nparr, cv2.IMREAD_COLOR)
        h_px, w_px = img.shape[:2]

        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 50, 255, cv2.THRESH_BINARY)
        contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

        crop_rect = fitz.Rect(0, 0, w_pt, h_pt)
        if contours:
            largest = max(contours, key=cv2.contourArea)
            x, y, bw, bh = cv2.boundingRect(largest)
            if bw > w_px * 0.5 and bh > h_px * 0.5:
                scale_x = w_pt / w_px
                scale_y = h_pt / h_px
                crop_rect = fitz.Rect(
                    x * scale_x, y * scale_y,
                    (x + bw) * scale_x, (y + bh) * scale_y
                )

        page.set_cropbox(crop_rect)

        crop_h_pt = crop_rect.height
        wm_h = crop_h_pt * 0.06
        page.draw_rect(
            fitz.Rect(crop_rect.x0, crop_rect.y1 - wm_h, crop_rect.x1, crop_rect.y1),
            color=None, fill=white, overlay=True
        )

    doc.save(output_path, garbage=4, deflate=True)
    doc.close()
    return output_path

# =============================================
# SCAN & CROP FUNCTIONS
# =============================================

def scan_document(
    input_path: str,
    output_path: str,
    crop_x: int = 0,
    crop_y: int = 0,
    crop_w: int = 0,
    crop_h: int = 0
):
    """Scan dokumen: crop area konten dari gambar"""
    img = cv2.imread(input_path)
    if img is None:
        raise Exception("Gagal membaca gambar")

    h, w = img.shape[:2]

    if crop_w > 0 and crop_h > 0:
        crop_x = max(0, min(crop_x, w - 10))
        crop_y = max(0, min(crop_y, h - 10))
        crop_w = min(crop_w, w - crop_x)
        crop_h = min(crop_h, h - crop_y)

        result = img[crop_y:crop_y+crop_h, crop_x:crop_x+crop_w]
    else:
        result = _auto_crop_document(img)

    cv2.imwrite(output_path, result)
    return output_path

def _auto_crop_document(img):
    """Deteksi dokumen seperti CamScanner: cari 4 sudut lalu perspective transform"""
    h, w = img.shape[:2]
    original = img.copy()

    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
    blurred = cv2.GaussianBlur(gray, (5, 5), 0)
    thresh = cv2.adaptiveThreshold(
        blurred, 255,
        cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY_INV,
        11, 2
    )

    kernel = np.ones((5, 5), np.uint8)
    dilated = cv2.dilate(thresh, kernel, iterations=3)

    contours, _ = cv2.findContours(dilated, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    if contours:
        contours = sorted(contours, key=cv2.contourArea, reverse=True)

        for contour in contours[:5]:
            area = cv2.contourArea(contour)
            if area < (w * h * 0.2):
                continue

            peri = cv2.arcLength(contour, True)
            approx = cv2.approxPolyDP(contour, 0.02 * peri, True)

            if len(approx) == 4:
                return _perspective_transform(original, approx)
            elif len(approx) > 4:
                approx = cv2.approxPolyDP(contour, 0.05 * peri, True)
                if len(approx) == 4:
                    return _perspective_transform(original, approx)

    margin_x = int(w * 0.03)
    margin_y = int(h * 0.03)
    return original[margin_y:h-margin_y, margin_x:w-margin_x]

def _perspective_transform(img, corners):
    """Transform perspektif: luruskan dokumen dari 4 sudut"""
    corners = corners.reshape(4, 2)
    rect = _order_corners(corners)
    (tl, tr, br, bl) = rect

    width_a = np.sqrt(((br[0] - bl[0]) ** 2) + ((br[1] - bl[1]) ** 2))
    width_b = np.sqrt(((tr[0] - tl[0]) ** 2) + ((tr[1] - tl[1]) ** 2))
    max_width = max(int(width_a), int(width_b))

    height_a = np.sqrt(((tr[0] - br[0]) ** 2) + ((tr[1] - br[1]) ** 2))
    height_b = np.sqrt(((tl[0] - bl[0]) ** 2) + ((tl[1] - bl[1]) ** 2))
    max_height = max(int(height_a), int(height_b))

    dst = np.array([
        [0, 0],
        [max_width - 1, 0],
        [max_width - 1, max_height - 1],
        [0, max_height - 1]
    ], dtype="float32")

    matrix = cv2.getPerspectiveTransform(rect, dst)
    warped = cv2.warpPerspective(img, matrix, (max_width, max_height))

    return warped

def _order_corners(pts):
    """Urutkan 4 sudut: top-left, top-right, bottom-right, bottom-left"""
    rect = np.zeros((4, 2), dtype="float32")

    s = pts.sum(axis=1)
    rect[0] = pts[np.argmin(s)]
    rect[2] = pts[np.argmax(s)]

    diff = np.diff(pts, axis=1)
    rect[1] = pts[np.argmin(diff)]
    rect[3] = pts[np.argmax(diff)]

    return rect

def get_document_boundary(input_path: str):
    """Deteksi boundary super agresif untuk objek dengan background warna senada"""
    img = cv2.imread(input_path)
    if img is None:
        raise Exception("Gagal membaca gambar")

    h, w = img.shape[:2]

    lab = cv2.cvtColor(img, cv2.COLOR_BGR2LAB)
    l, a, b = cv2.split(lab)
    clahe = cv2.createCLAHE(clipLimit=3.0, tileGridSize=(8,8))
    cl = clahe.apply(l)
    limg = cv2.merge((cl, a, b))
    enhanced = cv2.cvtColor(limg, cv2.COLOR_LAB2BGR)

    kernel_grad = cv2.getStructuringElement(cv2.MORPH_RECT, (5, 5))
    gradient = cv2.morphologyEx(enhanced, cv2.MORPH_GRADIENT, kernel_grad)

    gray = cv2.cvtColor(gradient, cv2.COLOR_BGR2GRAY)
    _, thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)

    kernel_close = cv2.getStructuringElement(cv2.MORPH_RECT, (15, 15))
    closed = cv2.morphologyEx(thresh, cv2.MORPH_CLOSE, kernel_close)

    contours, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)

    x, y, bw, bh = int(w * 0.05), int(h * 0.05), int(w * 0.9), int(h * 0.9)

    if contours:
        contours = sorted(contours, key=cv2.contourArea, reverse=True)

        for contour in contours[:10]:
            area = cv2.contourArea(contour)
            if area < (w * h * 0.1):
                continue

            cx, cy, cbw, cbh = cv2.boundingRect(contour)

            if cbw > w * 0.95 and cbh > h * 0.95:
                continue

            x, y, bw, bh = cx, cy, cbw, cbh
            break

    padding = 10
    x = max(0, x - padding)
    y = max(0, y - padding)
    bw = min(w - x, bw + padding * 2)
    bh = min(h - y, bh + padding * 2)

    return {
        'x': x,
        'y': y,
        'width': bw,
        'height': bh,
        'image_width': w,
        'image_height': h
    }