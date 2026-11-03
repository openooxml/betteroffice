from contextlib import contextmanager
from importlib.metadata import version

import pypdfium2 as pdfium
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
        target = (round(page.get_width() * dpi / 72), round(page.get_height() * dpi / 72))
        if image.size == target or any(abs(a - b) > 1 for a, b in zip(image.size, target)):
            return image
        snapped = Image.new('RGB', target, 'white')
        snapped.paste(image, (0, 0))
        return snapped
    finally:
        page.close()

