# A real-world diagram with a Word reference

`libreoffice-one-line.svg` is a synthetic single-line electrical diagram: a
bus bar, a transformer, converter blocks joined by a DC link, a motor, a
legend, labels in Arial and italic Times New Roman, and a small embedded PNG.
LibreOffice 26.2 exported it as `libreoffice-one-line.emf`, a dual EMF+ file
whose EMF+ records only set rendering hints and hand over with `GetDC`, so its
GDI records draw it: 1,277 object selections, 181 Bézier and 223 polygon
records, 14 `EXTTEXTOUTW` runs and one `STRETCHDIBITS`. LibreOffice's export
drops the SVG's dashed outer frame, and the EMF frame is the remaining ink's
bounding box.

`libreoffice-one-line.word.png` is Microsoft Word 16.113's rendering of that
EMF: a one-page A4 document anchors the picture at its natural size
(5,702,040 x 3,416,040 EMU) one inch from the page's top left corner,
`scripts/office-quality/reference.py` exports it to PDF at 192 DPI, and the
PNG is the picture box cropped from that page, scale 2 against CSS pixels.
