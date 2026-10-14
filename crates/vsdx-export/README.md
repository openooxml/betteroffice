# VSDX office export

Headless export of Visio diagrams to PowerPoint (`.pptx`) and Word (`.docx`).

One slide per page carries the diagram as native DrawingML shapes translated
from the `vsdx-render` display list, so output stays editable. The Word
document carries the same shapes plus a table of shape data per page.

The slide canvas is sized from the first page; every later page is uniformly
fitted inside it and centred. Shape data is the resolved `Property` section;
rows without a cached value or formula are omitted. Word diagrams fit within
a 6.5 by 8 inch frame, leaving room for the page heading. Gradients preserve
their linear angle, text re-flows in the host, and placeholders render as
dashed boxes labelled with their reason.

DrawingML fragments (`custGeom`, fills, outlines, paragraphs, placement) are
emitted through `ooxml-drawingml`; this crate only owns the package
scaffolding. An export never drops content silently: skewed transforms fall
back to the upright bounding box, unusable geometry falls back to a rectangle,
and unsupported images become labelled placeholders. Use
`export_pptx_with_report` / `export_docx_with_report` for the per-shape
breakdown alongside the bytes.
