import io
import json
import math
import os
import platform
import re
import zipfile
from pathlib import Path
from xml.dom import minidom

import fitz
from PIL import Image

from compare import Pages, compare
from docx_benchmark import edge_adjust, execute, font_environment, lo_command, sha256, write_json
from xlsx_calc import NS as XLSX_NS, address, package, sheets

NS = XLSX_NS['s']

WORKBOOK_ORDER = 'fileVersion fileSharing workbookPr workbookProtection bookViews sheets functionGroups externalReferences definedNames calcPr oleSize customWorkbookViews pivotCaches smartTagPr smartTagTypes webPublishing fileRecoveryPr webPublishObjects extLst'.split()
WORKSHEET_ORDER = 'sheetPr dimension sheetViews sheetFormatPr cols sheetData sheetCalcPr sheetProtection protectedRanges scenarios autoFilter sortState dataConsolidate customSheetViews mergeCells phoneticPr conditionalFormatting dataValidations hyperlinks printOptions pageMargins pageSetup headerFooter rowBreaks colBreaks customProperties cellWatches ignoredErrors smartTags drawing legacyDrawing legacyDrawingHF picture oleObjects controls webPublishItems tableParts extLst'.split()
PAPERS = [(1, 792, 612), (5, 1008, 612), (8, 1190.55, 841.89), (9, 841.89, 595.28)]


def children(parent, name):
    return [node for node in parent.childNodes if node.nodeType == node.ELEMENT_NODE
            and node.namespaceURI == NS and node.localName == name]


def element(parent, name, order):
    found = children(parent, name)
    if found:
        return found[0]
    prefix = parent.prefix + ':' if parent.prefix else ''
    node = parent.ownerDocument.createElementNS(NS, prefix + name)
    for sibling in parent.childNodes:
        if (sibling.nodeType == sibling.ELEMENT_NODE and sibling.namespaceURI == NS
                and sibling.localName in order and order.index(sibling.localName) > order.index(name)):
            parent.insertBefore(node, sibling)
            return node
    parent.appendChild(node)
    return node


def set_attributes(node, **attrs):
    for key, value in attrs.items():
        node.setAttribute(key, str(value))


def range_workbook(data, profile, page):
    scale, margin = profile.get('scale_percent'), profile.get('margin_pt')
    if type(scale) not in (int, float) or not math.isfinite(scale) or not 10 <= scale <= 100 or int(scale) != scale:
        raise ValueError('Invalid recorded print scale')
    if type(margin) not in (int, float) or not math.isfinite(margin) or not 0 <= margin <= 72:
        raise ValueError('Invalid recorded margins')
    match = re.fullmatch(r'([A-Z]{1,3})([1-9][0-9]*):([A-Z]{1,3})([1-9][0-9]*)', page.get('range', ''))
    if not match or type(page.get('sheet')) is not int or not 0 <= page['sheet'] < 100:
        raise ValueError('Invalid recorded print range')
    start, end = (address(cell) for cell in page['range'].split(':'))
    if start[0] > end[0] or start[1] > end[1]:
        raise ValueError('Inverted recorded print range')
    paper = next((code for code, width, height in PAPERS
                  if abs(page.get('width_pt', 0) - width) < 0.1 and abs(page.get('height_pt', 0) - height) < 0.1), None)
    if paper is None:
        raise ValueError('Unsupported recorded paper dimensions')
    parts = package(data)
    worksheets = sheets(parts)
    if page['sheet'] >= len(worksheets):
        raise ValueError('Recorded worksheet is missing')
    name, path, _ = worksheets[page['sheet']]
    book = minidom.parseString(parts['xl/workbook.xml'])
    root = book.documentElement
    tabs = children(children(root, 'sheets')[0], 'sheet')
    selected = next(index for index, tab in enumerate(tabs) if tab.getAttribute('name') == name)
    for index, tab in enumerate(tabs):
        tab.setAttribute('state', 'visible' if index == selected else 'hidden')
    for views in children(root, 'bookViews'):
        for view in children(views, 'workbookView'):
            set_attributes(view, activeTab=selected, firstSheet=selected)
    names = element(root, 'definedNames', WORKBOOK_ORDER)
    for node in children(names, 'definedName'):
        if node.getAttribute('name') in ('_xlnm.Print_Area', '_xlnm.Print_Titles'):
            names.removeChild(node)
    area = book.createElementNS(NS, (root.prefix + ':' if root.prefix else '') + 'definedName')
    set_attributes(area, name='_xlnm.Print_Area', localSheetId=selected)
    c1, r1, c2, r2 = match.groups()
    area.appendChild(book.createTextNode("'" + name.replace("'", "''") + f"'!${c1}${r1}:${c2}${r2}"))
    names.appendChild(area)
    sheet = minidom.parseString(parts[path])
    root = sheet.documentElement
    properties = element(root, 'sheetPr', WORKSHEET_ORDER)
    set_attributes(element(properties, 'pageSetUpPr', ['tabColor', 'outlinePr', 'pageSetUpPr']), fitToPage=0)
    set_attributes(element(root, 'printOptions', WORKSHEET_ORDER), gridLines=1, gridLinesSet=1,
                   headings=0, horizontalCentered=0, verticalCentered=0)
    set_attributes(element(root, 'pageMargins', WORKSHEET_ORDER), left=margin / 72, right=margin / 72,
                   top=margin / 72, bottom=margin / 72, header=0, footer=0)
    setup = element(root, 'pageSetup', WORKSHEET_ORDER)
    for attr in list(setup.attributes.values()):
        if attr.localName in ('fitToWidth', 'fitToHeight', 'paperWidth', 'paperHeight', 'paperUnits', 'id'):
            setup.removeAttributeNode(attr)
    set_attributes(setup, orientation='landscape', scale=int(scale), paperSize=paper, usePrinterDefaults=0)
    for header in children(root, 'headerFooter'):
        root.removeChild(header)
    parts['xl/workbook.xml'] = book.toxml(encoding='utf-8')
    parts[path] = sheet.toxml(encoding='utf-8')
    output = io.BytesIO()
    with zipfile.ZipFile(output, 'w', zipfile.ZIP_DEFLATED) as archive:
        for filename, content in parts.items():
            archive.writestr(filename, content)
    return output.getvalue()


def libreoffice_xlsx_fidelity(sample, root, soffice, version):
    actual = root / 'libreoffice'
    actual.mkdir()
    result = dict(channel='libreoffice', version=version)
    stage = 'capture'
    try:
        reference = sample['metadata']['reference']
        profile = reference['capture_profile']
        pages = profile['pages']
        if not 1 <= len(pages) <= 100 or len(pages) != reference['pages']:
            raise ValueError('Recorded ranges do not match the reference pages')
        source = (root / 'source.xlsx').read_bytes()
        adjustments, exports = [], []
        for index, page in enumerate(pages, 1):
            print(f"{sample['id']}: LibreOffice range {index}/{len(pages)} {page['range']}", flush=True)
            directory = actual / f'range-{index:04d}'
            directory.mkdir()
            prepared = directory / 'input.xlsx'
            prepared.write_bytes(range_workbook(source, profile, page))
            execute(lo_command(soffice, root / 'profile-fidelity', directory, prepared,
                               'pdf:calc_pdf_Export:{"SinglePageSheets":{"type":"boolean","value":false}}'),
                    timeout=180, log=directory / 'convert.log')
            with fitz.open(directory / 'input.pdf') as pdf:
                if len(pdf) != 1:
                    raise ValueError(f"Range {page['range']} produced {len(pdf)} pages; expected one at the recorded scale")
                if abs(pdf[0].rect.width - page['width_pt']) > 0.1 or abs(pdf[0].rect.height - page['height_pt']) > 0.1:
                    raise ValueError('LibreOffice paper dimensions differ from the recorded reference')
                pixmap = pdf[0].get_pixmap(dpi=150, alpha=False)
                image = Image.frombytes('RGB', (pixmap.width, pixmap.height), pixmap.samples)
                with Image.open(root / 'reference' / f'page_{index:04d}.png') as expected:
                    if expected.size != (page['width_px'], page['height_px']):
                        raise ValueError('Reference PNG dimensions differ from the capture profile')
                    image, adjustment = edge_adjust(image, expected.size)
                    if image.size != expected.size:
                        raise ValueError('LibreOffice raster dimensions differ from the reference')
                if adjustment:
                    adjustments.append(dict(page=index, **adjustment))
                image.save(actual / f'page_{index:04d}.png')
                exports.append(dict(page=index, sheet=page['sheet'], range=page['range'],
                                    prepared_sha256=sha256(prepared), pdf_sha256=sha256(directory / 'input.pdf')))
        write_json(actual / 'result.json', dict(status='ok', sha256=sample['metadata']['source']['sha256'],
                   dpi=150, pages=len(pages), engine='LibreOffice ' + version, rasterizer='PyMuPDF ' + fitz.VersionBind,
                   capture_profile=profile, exports=exports, edge_adjustments=adjustments))
        stage = 'compare'
        result.update(compare(Pages(root / 'reference', 150), Pages(actual, 150), root / 'lo-comparison', gallery=False))
        result['status'] = 'ok'
    except Exception as error:
        result.update(status='failed', stage=stage, error=str(error))
    return result


def main():
    root = Path(os.environ['QUALITY_OUTPUT']).resolve()
    job = json.loads((root / 'job.json').read_text())
    soffice = os.environ.get('SOFFICE', 'libreoffice')
    os.environ['SAL_USE_VCLPLUGIN'] = 'svp'
    fonts_hash = font_environment(Path(os.environ['QUALITY_FONTS']).resolve(), root)
    _, version_text = execute([soffice, '--version'])
    version = re.search(r'LibreOffice (\d+(?:\.\d+){2,3})\b', version_text)[1]
    report = dict(schema_version=1, plan_sha256=job['plan_sha256'], shard=job['shard'], samples=[], benchmark=dict(
        method=job['method'], dpi=150, source_sha=job['source_sha'], libreoffice_version=version,
        libreoffice_build=version_text.strip(), fonts_sha256=fonts_hash, rasterizer='PyMuPDF ' + fitz.VersionBind,
        font_policy='bundled-fontconfig' if platform.system() == 'Linux' else 'system-fonts'))
    for index, sample in enumerate(job['samples'], 1):
        directory = root / sample['id']
        if sha256(directory / 'source.xlsx') != sample['metadata']['source']['sha256']:
            raise ValueError('Source hash mismatch')
        for page, asset in enumerate(sample['metadata']['reference_pages'], 1):
            if sha256(directory / 'reference' / f'page_{page:04d}.png') != asset['sha256']:
                raise ValueError('Reference hash mismatch')
        print(f"XLSX LibreOffice: {index}/{len(job['samples'])} {sample['id']}", flush=True)
        result = libreoffice_xlsx_fidelity(sample, directory, soffice, version)
        report['samples'].append(dict(id=sample['id'], source_sha256=sample['metadata']['source']['sha256'], libreoffice=result))
        write_json(root / 'report.json', report)
        print(f"{sample['id']}: {result['status']}, SSIM {result.get('penalized_ssim', '—')}", flush=True)


if __name__ == '__main__':
    main()
