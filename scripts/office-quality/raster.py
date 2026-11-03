import ctypes
import math
from contextlib import contextmanager
from importlib.metadata import version

import pypdfium2 as pdfium
import pypdfium2.raw as pdfium_c
from PIL import Image

RASTERIZER = f"pypdfium2 {version('pypdfium2')}"


def load_pdf(path):
    try:
        return pdfium.PdfDocument(str(path))
    except pdfium.PdfiumError as error:
        raise ValueError(f'{path}: unreadable PDF ({error})') from error


@contextmanager
def open_pdf(path):
    document = load_pdf(path)
    try:
        yield document
    finally:
        document.close()


def page_size_pt(document, index):
    page = document[index]
    try:
        return page.get_width(), page.get_height()
    finally:
        page.close()


def render_page(document, index, dpi=150):
    page = document[index]
    try:
        bitmap = page.render(scale=dpi / 72, fill_color=(255, 255, 255, 255), may_draw_forms=True)
        try:
            image = bitmap.to_pil().convert('RGB')
        finally:
            bitmap.close()
        target = tuple(math.ceil(points * dpi / 72 - 0.001) for points in (page.get_width(), page.get_height()))
        if image.size == target or any(abs(a - b) > 1 for a, b in zip(image.size, target)):
            return image
        snapped = Image.new('RGB', target, 'white')
        snapped.paste(image, (0, 0))
        return snapped
    finally:
        page.close()



def pdf_metadata(document):
    info = document.get_metadata_dict(skip_empty=False)
    version = document.get_version()
    return dict(format=f'PDF {version // 10}.{version % 10}' if version else '', title=info.get('Title', ''),
                author=info.get('Author', ''), subject=info.get('Subject', ''), keywords=info.get('Keywords', ''),
                creator=info.get('Creator', ''), producer=info.get('Producer', ''),
                creationDate=info.get('CreationDate', ''), modDate=info.get('ModDate', ''),
                trapped=document.get_metadata_value('Trapped'), encryption=None)


def page_fonts(document, index):
    page = document[index]
    names = set()
    try:
        for item in page.get_objects(filter=(pdfium_c.FPDF_PAGEOBJ_TEXT,)):
            font = pdfium_c.FPDFTextObj_GetFont(item.raw)
            if not font:
                continue
            size = pdfium_c.FPDFFont_GetBaseFontName(font, None, 0)
            buffer = ctypes.create_string_buffer(size)
            pdfium_c.FPDFFont_GetBaseFontName(font, buffer, size)
            if buffer.value:
                names.add(buffer.value.decode('utf-8', 'replace'))
    finally:
        page.close()
    return names


def merge_pdfs(paths, output):
    combined = pdfium.PdfDocument.new()
    try:
        for path in paths:
            with open_pdf(path) as document:
                combined.import_pages(document)
        combined.save(str(output))
    finally:
        combined.close()
