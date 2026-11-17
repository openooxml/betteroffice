def pdf_bytes(pages, info=None):
    objects = {}
    kids = []
    number = 3
    for width, height, text in pages:
        stream = f'BT /F1 24 Tf 30 30 Td ({text}) Tj ET'.encode() if text else b''
        objects[number + 1] = b'<< /Length %d >>\nstream\n' % len(stream) + stream + b'\nendstream'
        objects[number + 2] = (f'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {width} {height}] '
                               f'/Resources << /Font << /F1 3 0 R >> >> /Contents {number + 1} 0 R >>').encode()
        kids.append(number + 2)
        number += 2
    objects[1] = b'<< /Type /Catalog /Pages 2 0 R >>'
    objects[2] = f'<< /Type /Pages /Kids [{" ".join(f"{kid} 0 R" for kid in kids)}] /Count {len(kids)} >>'.encode()
    objects[3] = b'<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>'
    info_reference = ''
    if info:
        number += 1
        objects[number] = ('<< ' + ' '.join(f'/{key} ({value})' for key, value in info.items()) + ' >>').encode()
        info_reference = f' /Info {number} 0 R'
    output = bytearray(b'%PDF-1.7\n')
    offsets = {}
    for index in sorted(objects):
        offsets[index] = len(output)
        output += b'%d 0 obj\n' % index + objects[index] + b'\nendobj\n'
    xref = len(output)
    output += b'xref\n0 %d\n0000000000 65535 f \n' % (number + 1)
    for index in range(1, number + 1):
        output += b'%010d 00000 n \n' % offsets[index]
    output += f'trailer\n<< /Size {number + 1} /Root 1 0 R{info_reference} >>\nstartxref\n{xref}\n%%EOF\n'.encode()
    return bytes(output)


def form_pdf_bytes():
    appearance = b'0 g 0 0 200 100 re f'
    objects = {
        1: b'<< /Type /Catalog /Pages 2 0 R /AcroForm << /Fields [4 0 R] >> >>',
        2: b'<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
        3: b'<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Annots [4 0 R] >>',
        4: b'<< /Type /Annot /Subtype /Widget /FT /Tx /T (field) /F 4 /Rect [100 100 300 200] /P 3 0 R /AP << /N 5 0 R >> >>',
        5: b'<< /Type /XObject /Subtype /Form /BBox [0 0 200 100] /Length %d >>\nstream\n' % len(appearance) + appearance + b'\nendstream',
    }
    output = bytearray(b'%PDF-1.7\n')
    offsets = {}
    for index in sorted(objects):
        offsets[index] = len(output)
        output += b'%d 0 obj\n' % index + objects[index] + b'\nendobj\n'
    xref = len(output)
    output += b'xref\n0 6\n0000000000 65535 f \n' + b''.join(b'%010d 00000 n \n' % offsets[index] for index in range(1, 6))
    output += f'trailer\n<< /Size 6 /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n'.encode()
    return bytes(output)
