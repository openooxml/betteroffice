import re
import subprocess

import fitz


def export_profile(source, pdf, profile, timeout, dpi):
    pages = profile.get('pages', [])
    scale = profile.get('scale_percent')
    margin = profile.get('margin_pt')
    if not isinstance(scale, (int, float)) or not 10 <= scale <= 100:
        raise ValueError('XLSX scale_percent must be between 10 and 100')
    if not isinstance(margin, (int, float)) or not 0 <= margin <= 72:
        raise ValueError('XLSX margin_pt must be between 0 and 72')
    if not 1 <= len(pages) <= 100:
        raise ValueError('XLSX profile must contain 1–100 page ranges')
    for page in pages:
        if not isinstance(page.get('sheet'), int) or not 0 <= page['sheet'] < 100:
            raise ValueError('Invalid worksheet index')
        if not re.fullmatch(r'[A-Z]{1,3}[1-9][0-9]*:[A-Z]{1,3}[1-9][0-9]*', page.get('range', '')):
            raise ValueError('Invalid rectangular print range')
    sheets = ', '.join(str(page['sheet'] + 1) for page in pages)
    ranges = ', '.join('"' + page['range'] + '"' for page in pages)
    script = f'''on run argv
    set inputPath to item 1 of argv
    set outputDirectory to item 2 of argv
    set sheetIndices to {{{sheets}}}
    set printRanges to {{{ranges}}}
    with timeout of {timeout} seconds
        tell application "Microsoft Excel"
            set wb to open workbook workbook file name inputPath update links do not update links read only true add to mru false
            try
                repeat with pageIndex from 1 to count of sheetIndices
                    set sheetIndex to item pageIndex of sheetIndices
                    set sh to worksheet sheetIndex of wb
                    set visible of sh to sheet visible
                    repeat with otherIndex from 1 to count of worksheets of wb
                        if otherIndex is not sheetIndex then set visible of worksheet otherIndex of wb to sheet hidden
                    end repeat
                    tell page setup object of sh
                        set print area to item pageIndex of printRanges
                        set page orientation to landscape
                        set zoom to {scale}
                        set left margin to {margin}
                        set right margin to {margin}
                        set top margin to {margin}
                        set bottom margin to {margin}
                        set print gridlines to true
                        set print headings to false
                        set center horizontally to false
                        set center vertically to false
                        set left header to ""
                        set center header to ""
                        set right header to ""
                        set left footer to ""
                        set center footer to ""
                        set right footer to ""
                        set print title rows to ""
                        set print title columns to ""
                    end tell
                    save workbook as wb filename (outputDirectory & "/sheet-" & pageIndex & ".pdf") file format PDF file format
                end repeat
                close wb saving no
            on error messageText number errorNumber
                close wb saving no
                error messageText number errorNumber
            end try
        end tell
    end timeout
end run
'''
    subprocess.run(['osascript', '-', str(source), str(pdf.parent)], input=script,
                   text=True, capture_output=True, check=True, timeout=timeout + 5)
    capture = dict(profile, paper='landscape; Office paper size recorded per page', gridlines=True, headings=False,
                   pagination='one recorded range per page; fixed print scale', pages=[])
    with fitz.open() as combined:
        exports = []
        for index, page in enumerate(pages):
            with fitz.open(pdf.parent / f'sheet-{index + 1}.pdf') as document:
                if len(document) != 1:
                    raise ValueError(f'Print range {page["range"]} produced {len(document)} pages; choose a smaller range or scale')
                pixmap = document[0].get_pixmap(dpi=dpi, alpha=False)
                capture['pages'].append(dict(page, width_px=pixmap.width, height_px=pixmap.height,
                                            width_pt=document[0].rect.width, height_pt=document[0].rect.height))
                exports.append(document.metadata)
                combined.insert_pdf(document)
        combined.set_metadata(exports[0])
        combined.save(pdf)
    capture['office_pdf_metadata'] = exports
    return capture
