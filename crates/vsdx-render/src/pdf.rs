use std::collections::BTreeSet;

use ooxml_drawingml::GeometryPathCommand;
use vsdx_parse::VsdxPackage;

use crate::display_list::{Affine, Paint, Primitive, TextRun, VsdxDisplayList};
use crate::vector::{collect_ordered, num, rgb, z_order};
use crate::{RenderError, Renderer};

const POINTS_PER_INCH: f32 = 72.0;

#[derive(Clone, Copy)]
struct FontKey {
    class: u8,
    bold: bool,
    italic: bool,
}

impl FontKey {
    fn of(run: &TextRun) -> Self {
        let family = run.family.to_lowercase();
        let class = if ["times", "serif", "georgia", "garamond", "palatino"]
            .iter()
            .any(|marker| family.contains(marker))
        {
            1
        } else if ["courier", "mono", "consol", "menlo"]
            .iter()
            .any(|marker| family.contains(marker))
        {
            2
        } else {
            0
        };
        Self {
            class,
            bold: run.bold,
            italic: run.italic,
        }
    }

    fn base_font(&self) -> &'static str {
        match (self.class, self.bold, self.italic) {
            (1, false, false) => "Times-Roman",
            (1, true, false) => "Times-Bold",
            (1, false, true) => "Times-Italic",
            (1, true, true) => "Times-BoldItalic",
            (2, false, false) => "Courier",
            (2, true, false) => "Courier-Bold",
            (2, false, true) => "Courier-Oblique",
            (2, true, true) => "Courier-BoldOblique",
            (_, false, false) => "Helvetica",
            (_, true, false) => "Helvetica-Bold",
            (_, false, true) => "Helvetica-Oblique",
            (_, true, true) => "Helvetica-BoldOblique",
        }
    }
}

impl Ord for FontKey {
    fn cmp(&self, other: &Self) -> std::cmp::Ordering {
        (self.class, self.bold, self.italic).cmp(&(other.class, other.bold, other.italic))
    }
}

impl PartialOrd for FontKey {
    fn partial_cmp(&self, other: &Self) -> Option<std::cmp::Ordering> {
        Some(self.cmp(other))
    }
}

impl PartialEq for FontKey {
    fn eq(&self, other: &Self) -> bool {
        (self.class, self.bold, self.italic) == (other.class, other.bold, other.italic)
    }
}

impl Eq for FontKey {}

struct PageContent {
    width_pt: f32,
    height_pt: f32,
    stream: Vec<u8>,
    fonts: BTreeSet<FontKey>,
    images: Vec<PlacedImage>,
}

type FontRefs = Vec<(String, u32)>;
type PlacedImageRefs = Vec<(PlacedImage, u32)>;
type PageNumbers = Vec<(u32, u32, FontRefs, PlacedImageRefs)>;

struct PlacedImage {
    name: String,
    bytes: Vec<u8>,
    width: u32,
    height: u32,
    gray: bool,
    matrix: Affine,
}

struct Emitter<'a> {
    package: &'a VsdxPackage,
    content: PageContent,
    image_count: usize,
}

fn solid(paint: &Option<Paint>) -> Option<(f64, f64, f64)> {
    match paint {
        Some(Paint::Solid { color }) => rgb(color),
        Some(Paint::Gradient { stops, .. }) => stops.first().and_then(|stop| rgb(&stop.color)),
        None => None,
    }
}

fn pdf_text(value: &str) -> String {
    let filtered: String = value
        .chars()
        .filter_map(|ch| match ch {
            '\t' | '\n' | '\r' => None,
            ' '..='~' => Some(ch),
            _ => Some('?'),
        })
        .collect();
    let mut out = String::with_capacity(filtered.len() + 2);
    out.push('(');
    for ch in filtered.chars() {
        match ch {
            '(' | ')' | '\\' => {
                out.push('\\');
                out.push(ch);
            }
            _ => out.push(ch),
        }
    }
    out.push(')');
    out
}

fn jpeg_dimensions(bytes: &[u8]) -> Option<(u32, u32, bool)> {
    if bytes.len() < 4 || bytes[0] != 0xFF || bytes[1] != 0xD8 {
        return None;
    }
    let mut index = 2;
    while index + 4 <= bytes.len() {
        if bytes[index] != 0xFF {
            return None;
        }
        let marker = bytes[index + 1];
        index += 2;
        if marker == 0xD9 {
            return None;
        }
        if marker == 0x01 || (0xD0..=0xD7).contains(&marker) {
            continue;
        }
        if index + 2 > bytes.len() {
            return None;
        }
        let length = u16::from_be_bytes([bytes[index], bytes[index + 1]]) as usize;
        if length < 2 || index + length > bytes.len() {
            return None;
        }
        if matches!(marker, 0xC0..=0xC2) && length >= 8 {
            let height = u16::from_be_bytes([bytes[index + 3], bytes[index + 4]]) as u32;
            let width = u16::from_be_bytes([bytes[index + 5], bytes[index + 6]]) as u32;
            let components = bytes[index + 7];
            if width == 0 || height == 0 {
                return None;
            }
            return match components {
                1 => Some((width, height, true)),
                3 => Some((width, height, false)),
                _ => None,
            };
        }
        index += length;
    }
    None
}

fn push_matrix(out: &mut Vec<u8>, matrix: Affine) {
    out.extend_from_slice(
        format!(
            "{} {} {} {} {} {} cm\n",
            num(f64::from(matrix.a)),
            num(f64::from(matrix.b)),
            num(f64::from(matrix.c)),
            num(f64::from(matrix.d)),
            num(f64::from(matrix.e)),
            num(f64::from(matrix.f))
        )
        .as_bytes(),
    );
}

fn paint_path(
    out: &mut Vec<u8>,
    path: &[GeometryPathCommand],
    fill: &Option<Paint>,
    stroke: &Option<crate::display_list::Stroke>,
) {
    let mut current = None;
    let mut subpath_start = None;
    for command in path {
        let line = match command {
            GeometryPathCommand::Move { x, y } => {
                current = Some((*x, *y));
                subpath_start = current;
                format!("{} {} m\n", num(*x), num(*y))
            }
            GeometryPathCommand::Line { x, y } => {
                current = Some((*x, *y));
                format!("{} {} l\n", num(*x), num(*y))
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                let (current_x, current_y) = current.unwrap_or((*x, *y));
                let (c1x, c1y) = (
                    current_x + 2.0 * (*cpx - current_x) / 3.0,
                    current_y + 2.0 * (*cpy - current_y) / 3.0,
                );
                let (c2x, c2y) = (*x + 2.0 * (*cpx - *x) / 3.0, *y + 2.0 * (*cpy - *y) / 3.0);
                current = Some((*x, *y));
                format!(
                    "{} {} {} {} {} {} c\n",
                    num(c1x),
                    num(c1y),
                    num(c2x),
                    num(c2y),
                    num(*x),
                    num(*y)
                )
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                current = Some((*x, *y));
                format!(
                    "{} {} {} {} {} {} c\n",
                    num(*cp1x),
                    num(*cp1y),
                    num(*cp2x),
                    num(*cp2y),
                    num(*x),
                    num(*y)
                )
            }
            GeometryPathCommand::Close => {
                current = subpath_start;
                "h\n".to_owned()
            }
        };
        out.extend_from_slice(line.as_bytes());
    }
    let fill_rgb = solid(fill);
    let stroke_paint = stroke.as_ref().and_then(|stroke| {
        rgb(&stroke.color).map(|color| (color, f64::from(stroke.width), stroke.dashed))
    });
    match (fill_rgb, stroke_paint) {
        (Some((r, g, b)), Some(((sr, sg, sb), width, dashed))) => {
            out.extend_from_slice(format!("{} {} {} rg\n", num(r), num(g), num(b)).as_bytes());
            out.extend_from_slice(format!("{} {} {} RG\n", num(sr), num(sg), num(sb)).as_bytes());
            out.extend_from_slice(format!("{} w\n", num(width.max(0.0))).as_bytes());
            if dashed {
                let step = num(width.max(0.0) * 2.0);
                out.extend_from_slice(format!("[{step} {step}] 0 d\n").as_bytes());
            }
            out.extend_from_slice(b"B\n");
        }
        (Some((r, g, b)), None) => {
            out.extend_from_slice(format!("{} {} {} rg\n", num(r), num(g), num(b)).as_bytes());
            out.extend_from_slice(b"f\n");
        }
        (None, Some(((sr, sg, sb), width, dashed))) => {
            out.extend_from_slice(format!("{} {} {} RG\n", num(sr), num(sg), num(sb)).as_bytes());
            out.extend_from_slice(format!("{} w\n", num(width.max(0.0))).as_bytes());
            if dashed {
                let step = num(width.max(0.0) * 2.0);
                out.extend_from_slice(format!("[{step} {step}] 0 d\n").as_bytes());
            }
            out.extend_from_slice(b"S\n");
        }
        (None, None) => {}
    }
}

fn stop_x(line: &crate::display_list::PositionedLine, position: u32) -> Option<f32> {
    line.caret_stops
        .iter()
        .find(|stop| stop.position == position)
        .map(|stop| stop.x)
}

impl<'a> Emitter<'a> {
    fn primitive(&mut self, primitive: &Primitive, outer: Affine, fonts: &mut BTreeSet<FontKey>) {
        let mut ordered = Vec::new();
        collect_ordered(primitive, &mut ordered);
        for (item, matrix) in ordered {
            let matrix = outer.compose(matrix);
            match item {
                Primitive::Shape {
                    path, fill, stroke, ..
                } => {
                    let ctm = Affine {
                        a: POINTS_PER_INCH,
                        b: 0.0,
                        c: 0.0,
                        d: POINTS_PER_INCH,
                        e: 0.0,
                        f: 0.0,
                    }
                    .compose(matrix);
                    self.content.stream.extend_from_slice(b"q\n");
                    push_matrix(&mut self.content.stream, ctm);
                    paint_path(&mut self.content.stream, path, fill, stroke);
                    self.content.stream.extend_from_slice(b"Q\n");
                }
                Primitive::TextBox {
                    paragraphs, lines, ..
                } => self.textbox(paragraphs, lines, matrix, fonts),
                Primitive::Image {
                    asset_id,
                    x,
                    y,
                    width,
                    height,
                    ..
                } => self.image(asset_id, *x, *y, *width, *height, matrix),
                Primitive::Placeholder {
                    x,
                    y,
                    width,
                    height,
                    reason,
                    ..
                } => {
                    self.placeholder(*x, *y, *width, *height, matrix, reason, fonts);
                }
                Primitive::Group { .. } => {}
            }
        }
    }

    fn textbox(
        &mut self,
        paragraphs: &[crate::display_list::TextParagraph],
        lines: &[crate::display_list::PositionedLine],
        matrix: Affine,
        fonts: &mut BTreeSet<FontKey>,
    ) {
        let ctm = Affine {
            a: POINTS_PER_INCH,
            b: 0.0,
            c: 0.0,
            d: POINTS_PER_INCH,
            e: 0.0,
            f: 0.0,
        }
        .compose(matrix);
        let mut offset = 0u32;
        let runs: Vec<(&TextRun, u32, u32)> = paragraphs
            .iter()
            .flat_map(|paragraph| paragraph.runs.iter())
            .map(|run| {
                let start = offset;
                offset += run.text.len() as u32;
                (run, start, offset)
            })
            .collect();
        self.content.stream.extend_from_slice(b"q\n");
        push_matrix(&mut self.content.stream, ctm);
        for line in lines {
            for (run, run_start, run_end) in &runs {
                let start = (*run_start).max(line.start);
                let end = (*run_end).min(line.end);
                if start >= end {
                    continue;
                }
                let fragment = &run.text[(start - run_start) as usize..(end - run_start) as usize];
                let visible: String = fragment
                    .chars()
                    .filter(|ch| !matches!(ch, '\t' | '\n' | '\r'))
                    .collect();
                if visible.is_empty() || !run.size_in.is_finite() || run.size_in <= 0.0 {
                    continue;
                }
                let key = FontKey::of(run);
                fonts.insert(key);
                let index = font_index(&key);
                let x = stop_x(line, start).unwrap_or(line.x);
                let baseline = f64::from(line.y) - f64::from(run.size_in) * 0.8;
                let (r, g, b) = rgb(&run.color).unwrap_or((0.0, 0.0, 0.0));
                self.content.stream.extend_from_slice(
                    format!(
                        "BT /F{} {} Tf {} {} {} rg 1 0 0 1 {} {} Tm {} Tj ET\n",
                        index,
                        num(f64::from(run.size_in)),
                        num(r),
                        num(g),
                        num(b),
                        num(f64::from(x)),
                        num(baseline),
                        pdf_text(&visible)
                    )
                    .as_bytes(),
                );
                if run.underline {
                    let x_end = stop_x(line, end).unwrap_or(line.x + line.width);
                    let underline_y = baseline - f64::from(run.size_in) * 0.08;
                    self.content.stream.extend_from_slice(
                        format!(
                            "{} w {} {} m {} {} l S\n",
                            num(f64::from(run.size_in) * 0.04),
                            num(f64::from(x)),
                            num(underline_y),
                            num(f64::from(x_end)),
                            num(underline_y)
                        )
                        .as_bytes(),
                    );
                }
            }
        }
        self.content.stream.extend_from_slice(b"Q\n");
    }

    fn image(&mut self, asset_id: &str, x: f32, y: f32, width: f32, height: f32, matrix: Affine) {
        let bytes = self.package.part_bytes(asset_id).and_then(|bytes| {
            jpeg_dimensions(bytes).map(|(w, h, gray)| (bytes.to_vec(), w, h, gray))
        });
        let Some((bytes, w, h, gray)) = bytes else {
            self.placeholder_outline(x, y, width, height, matrix);
            return;
        };
        self.image_count += 1;
        let name = format!("Im{}", self.image_count);
        let rect = Affine {
            a: width,
            b: 0.0,
            c: 0.0,
            d: height,
            e: x,
            f: y,
        };
        let ctm = Affine {
            a: POINTS_PER_INCH,
            b: 0.0,
            c: 0.0,
            d: POINTS_PER_INCH,
            e: 0.0,
            f: 0.0,
        }
        .compose(matrix.compose(rect));
        self.content.stream.extend_from_slice(b"q\n");
        push_matrix(&mut self.content.stream, ctm);
        self.content
            .stream
            .extend_from_slice(format!("/{name} Do\n").as_bytes());
        self.content.stream.extend_from_slice(b"Q\n");
        self.content.images.push(PlacedImage {
            name,
            bytes,
            width: w,
            height: h,
            gray,
            matrix: ctm,
        });
    }

    #[allow(clippy::too_many_arguments)]
    fn placeholder(
        &mut self,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        matrix: Affine,
        reason: &str,
        fonts: &mut BTreeSet<FontKey>,
    ) {
        self.placeholder_outline(x, y, width, height, matrix);
        if reason.is_empty() {
            return;
        }
        fonts.insert(FontKey {
            class: 0,
            bold: false,
            italic: false,
        });
        let ctm = Affine {
            a: POINTS_PER_INCH,
            b: 0.0,
            c: 0.0,
            d: POINTS_PER_INCH,
            e: 0.0,
            f: 0.0,
        }
        .compose(matrix);
        self.content.stream.extend_from_slice(b"q\n");
        push_matrix(&mut self.content.stream, ctm);
        self.content.stream.extend_from_slice(
            format!(
                "BT /F{} {} Tf 0.35 0.38 0.45 rg 1 0 0 1 {} {} Tm {} Tj ET\n",
                font_index(&FontKey {
                    class: 0,
                    bold: false,
                    italic: false,
                }),
                num(10.0 / 72.0),
                num(f64::from(x)),
                num(f64::from(y)),
                pdf_text(reason)
            )
            .as_bytes(),
        );
        self.content.stream.extend_from_slice(b"Q\n");
    }

    fn placeholder_outline(&mut self, x: f32, y: f32, width: f32, height: f32, matrix: Affine) {
        let ctm = Affine {
            a: POINTS_PER_INCH,
            b: 0.0,
            c: 0.0,
            d: POINTS_PER_INCH,
            e: 0.0,
            f: 0.0,
        }
        .compose(matrix);
        self.content.stream.extend_from_slice(b"q\n");
        push_matrix(&mut self.content.stream, ctm);
        self.content.stream.extend_from_slice(
            format!(
                "0.54 0.58 0.65 RG {} w [{} {}] 0 d {} {} {} {} re S\n",
                num(1.0 / 72.0),
                num(5.0 / 72.0),
                num(4.0 / 72.0),
                num(f64::from(x)),
                num(f64::from(y)),
                num(f64::from(width).max(0.0)),
                num(f64::from(height).max(0.0))
            )
            .as_bytes(),
        );
        self.content.stream.extend_from_slice(b"Q\n");
    }
}

fn font_index(key: &FontKey) -> usize {
    (key.class as usize) * 4 + ((key.bold as usize) << 1) + (key.italic as usize) + 1
}

fn base_font(index: usize) -> &'static str {
    FontKey {
        class: (index / 4) as u8,
        bold: (index % 4) >= 2,
        italic: (index % 2) == 1,
    }
    .base_font()
}

fn emit_page(list: &VsdxDisplayList, package: &VsdxPackage) -> PageContent {
    let mut primitives: Vec<&Primitive> = list.primitives.iter().collect();
    primitives.sort_by_key(|primitive| z_order(primitive));
    let mut fonts = BTreeSet::new();
    let mut emitter = Emitter {
        package,
        content: PageContent {
            width_pt: list.width / 96.0 * POINTS_PER_INCH,
            height_pt: list.height / 96.0 * POINTS_PER_INCH,
            stream: Vec::new(),
            fonts: BTreeSet::new(),
            images: Vec::new(),
        },
        image_count: 0,
    };
    for primitive in primitives {
        emitter.primitive(primitive, Affine::identity(), &mut fonts);
    }
    emitter.content.fonts = fonts;
    emitter.content
}

fn emit_document(pages: Vec<PageContent>) -> Vec<u8> {
    let mut out = Vec::new();
    out.extend_from_slice(b"%PDF-1.4\n%\xE2\xE3\xCF\xD3\n");
    let mut offsets: Vec<usize> = Vec::new();
    let object = |out: &mut Vec<u8>, offsets: &mut Vec<usize>, body: &[u8]| -> u32 {
        offsets.push(out.len());
        let number = offsets.len() as u32;
        out.extend_from_slice(format!("{number} 0 obj\n").as_bytes());
        out.extend_from_slice(body);
        out.extend_from_slice(b"\nendobj\n");
        number
    };
    let mut next_number = 3u32;
    let mut page_numbers: PageNumbers = Vec::new();
    for page in &pages {
        let page_number = next_number;
        next_number += 1;
        let content_number = next_number;
        next_number += 1;
        let mut font_refs: Vec<(FontKey, u32)> = Vec::new();
        for key in &page.fonts {
            font_refs.push((
                FontKey {
                    class: key.class,
                    bold: key.bold,
                    italic: key.italic,
                },
                next_number,
            ));
            next_number += 1;
        }
        let mut image_refs: Vec<(usize, u32)> = Vec::new();
        for (index, _) in page.images.iter().enumerate() {
            image_refs.push((index, next_number));
            next_number += 1;
        }
        page_numbers.push((
            page_number,
            content_number,
            font_refs
                .iter()
                .map(|(key, number)| (format!("F{}", font_index(key)), *number))
                .collect(),
            page.images
                .iter()
                .zip(image_refs.iter().map(|(_, number)| *number))
                .map(|(image, number)| {
                    (
                        PlacedImage {
                            name: image.name.clone(),
                            bytes: image.bytes.clone(),
                            width: image.width,
                            height: image.height,
                            gray: image.gray,
                            matrix: image.matrix,
                        },
                        number,
                    )
                })
                .collect(),
        ));
    }
    let kids: String = page_numbers
        .iter()
        .map(|(page, _, _, _)| format!("{page} 0 R"))
        .collect::<Vec<_>>()
        .join(" ");
    let catalog = "<< /Type /Catalog /Pages 2 0 R >>".to_string();
    let pages_obj = format!("<< /Type /Pages /Kids [{kids}] /Count {} >>", pages.len());
    let catalog_number = object(&mut out, &mut offsets, catalog.as_bytes());
    debug_assert_eq!(catalog_number, 1);
    let pages_number = object(&mut out, &mut offsets, pages_obj.as_bytes());
    debug_assert_eq!(pages_number, 2);
    for (index, page) in pages.iter().enumerate() {
        let (page_number, content_number, font_refs, image_refs) = &page_numbers[index];
        let mut resources = String::from("<< /Font << ");
        for (name, number) in font_refs {
            resources.push_str(&format!("/{name} {number} 0 R "));
        }
        resources.push_str(">>");
        if !image_refs.is_empty() {
            resources.push_str(" /XObject << ");
            for (image, number) in image_refs {
                resources.push_str(&format!("/{} {number} 0 R ", image.name));
            }
            resources.push_str(">>");
        }
        resources.push_str(" >>");
        let page_obj = format!(
            "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 {} {}] /Contents {content_number} 0 R /Resources {resources} >>",
            num(f64::from(page.width_pt)),
            num(f64::from(page.height_pt))
        );
        let written_page = object(&mut out, &mut offsets, page_obj.as_bytes());
        debug_assert_eq!(written_page, *page_number);
        let mut stream_obj = format!("<< /Length {} >>\nstream\n", page.stream.len()).into_bytes();
        stream_obj.extend_from_slice(&page.stream);
        stream_obj.extend_from_slice(b"endstream");
        let written_content = object(&mut out, &mut offsets, &stream_obj);
        debug_assert_eq!(written_content, *content_number);
        for (name, number) in font_refs {
            let key_index = name.trim_start_matches('F').parse::<usize>().unwrap_or(1);
            let font_obj = format!(
                "<< /Type /Font /Subtype /Type1 /BaseFont /{} >>",
                base_font(key_index.saturating_sub(1))
            );
            let written_font = object(&mut out, &mut offsets, font_obj.as_bytes());
            debug_assert_eq!(written_font, *number);
        }
        for (image, number) in image_refs {
            let colorspace = if image.gray {
                "DeviceGray"
            } else {
                "DeviceRGB"
            };
            let image_obj = format!(
                "<< /Type /XObject /Subtype /Image /Width {} /Height {} /ColorSpace /{} /BitsPerComponent 8 /Filter /DCTDecode /Length {} >>\nstream\n",
                image.width,
                image.height,
                colorspace,
                image.bytes.len()
            )
            .into_bytes();
            let mut body = image_obj;
            body.extend_from_slice(&image.bytes);
            body.extend_from_slice(b"\nendstream");
            let written_image = object(&mut out, &mut offsets, &body);
            debug_assert_eq!(written_image, *number);
        }
    }
    let xref_start = out.len();
    out.extend_from_slice(format!("xref\n0 {}\n", offsets.len() + 1).as_bytes());
    out.extend_from_slice(b"0000000000 65535 f \n");
    for offset in &offsets {
        out.extend_from_slice(format!("{offset:010} 00000 n \n").as_bytes());
    }
    out.extend_from_slice(
        format!(
            "trailer\n<< /Size {} /Root 1 0 R >>\nstartxref\n{xref_start}\n%%EOF",
            offsets.len() + 1
        )
        .as_bytes(),
    );
    out
}

impl Renderer {
    /// Renders every diagram page to a vector PDF with selectable text.
    pub fn export_pdf(&self, package: &VsdxPackage) -> Result<Vec<u8>, RenderError> {
        let mut pages = Vec::with_capacity(package.page_part_paths.len().max(1));
        for part in &package.page_part_paths.clone() {
            pages.push(emit_page(&self.layout_page(package, part)?, package));
        }
        Ok(emit_document(pages))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn pages_of(pdf: &[u8]) -> usize {
        let text = String::from_utf8_lossy(pdf);
        text.match_indices("/Type /Page ").count()
    }

    #[test]
    fn truncated_sof_jpeg_returns_no_dimensions() {
        let jpeg = [
            0xFF, 0xD8, 0xFF, 0xC0, 0x00, 0x07, 0x08, 0x00, 0x01, 0x00, 0x01,
        ];
        assert_eq!(jpeg_dimensions(&jpeg), None);
    }

    #[test]
    fn quadratic_paths_match_their_cubic_equivalent() {
        let quadratic = [
            GeometryPathCommand::Move { x: 1.0, y: 2.0 },
            GeometryPathCommand::Quad {
                cpx: 4.0,
                cpy: 8.0,
                x: 10.0,
                y: 5.0,
            },
        ];
        let cubic = [
            GeometryPathCommand::Move { x: 1.0, y: 2.0 },
            GeometryPathCommand::Cubic {
                cp1x: 3.0,
                cp1y: 6.0,
                cp2x: 6.0,
                cp2y: 7.0,
                x: 10.0,
                y: 5.0,
            },
        ];
        let mut quadratic_pdf = Vec::new();
        let mut cubic_pdf = Vec::new();
        paint_path(&mut quadratic_pdf, &quadratic, &None, &None);
        paint_path(&mut cubic_pdf, &cubic, &None, &None);
        assert_eq!(quadratic_pdf, cubic_pdf);
    }

    #[test]
    fn limits_text_to_ascii_supported_by_standard_fonts() {
        assert_eq!(pdf_text("Caf\u{e9} \u{03a9} \u{1f642}"), "(Caf? ? ?)");
    }

    #[test]
    fn exports_a_header_and_one_page_per_diagram_page() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let pdf = Renderer::default().export_pdf(&package).unwrap();
        assert!(pdf.starts_with(b"%PDF-1.4"));
        assert!(pdf.ends_with(b"%%EOF"));
        assert_eq!(pages_of(&pdf), package.page_part_paths.len());
        assert!(!package.page_part_paths.is_empty());
    }

    #[test]
    fn sizes_pages_from_the_display_list_rather_than_a_hardcoded_format() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let list = Renderer::default()
            .layout_page(&package, &package.page_part_paths[0])
            .unwrap();
        let expected = format!(
            "/MediaBox [0 0 {} {}]",
            num(f64::from(list.width) / 96.0 * 72.0),
            num(f64::from(list.height) / 96.0 * 72.0)
        );
        let pdf = Renderer::default().export_pdf(&package).unwrap();
        let text = String::from_utf8_lossy(&pdf);
        assert!(text.contains(&expected), "missing {expected}");
        assert!(!text.contains("/MediaBox [0 0 595") && !text.contains("/MediaBox [0 0 842"));
    }

    #[test]
    fn keeps_text_selectable_as_text_operators() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let list = Renderer::default()
            .layout_page(&package, &package.page_part_paths[0])
            .unwrap();
        let run = list
            .primitives
            .iter()
            .filter_map(|primitive| match primitive {
                Primitive::TextBox { paragraphs, .. } => Some(paragraphs),
                _ => None,
            })
            .flat_map(|paragraphs| paragraphs.iter().flat_map(|p| p.runs.iter()))
            .find(|run| !run.text.trim().is_empty())
            .expect("fixture carries text");
        let pdf = Renderer::default().export_pdf(&package).unwrap();
        let text = String::from_utf8_lossy(&pdf);
        assert!(text.contains("BT") && text.contains("Tj") && text.contains("Tf"));
        let visible: String = run
            .text
            .chars()
            .filter(|ch| !matches!(ch, '\t' | '\n' | '\r'))
            .collect();
        assert!(text.contains(&pdf_text(&visible)));
    }

    #[test]
    fn emits_shape_geometry_as_path_operators() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/indexed-geometry.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let pdf = Renderer::default().export_pdf(&package).unwrap();
        let text = String::from_utf8_lossy(&pdf);
        assert!(text.contains(" m\n") && text.contains(" l\n"));
    }

    #[test]
    fn references_only_standard_fonts_without_embedding() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let pdf = Renderer::default().export_pdf(&package).unwrap();
        let text = String::from_utf8_lossy(&pdf);
        assert!(text.contains("/Subtype /Type1"));
        assert!(!text.contains("/FontFile") && !text.contains("/FontFile2"));
    }

    #[test]
    fn xref_offsets_point_at_objects() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let pdf = Renderer::default().export_pdf(&package).unwrap();
        let find = |needle: &[u8]| {
            pdf.windows(needle.len())
                .position(|window| window == needle)
                .expect("pdf section")
        };
        let xref = find(b"xref\n");
        let trailer = find(b"trailer\n");
        let section = std::str::from_utf8(&pdf[xref..trailer]).expect("xref is ascii");
        let mut entries = 0;
        for line in section.lines().skip(2) {
            if line.ends_with("n ") {
                entries += 1;
                let offset: usize = line[..10].parse().expect("offset entry");
                let end = pdf[offset..]
                    .iter()
                    .position(|byte| *byte == b'\n')
                    .map(|pos| offset + pos)
                    .expect("object header line");
                let header =
                    std::str::from_utf8(&pdf[offset..end]).expect("object header is ascii");
                assert!(
                    header.ends_with(" 0 obj"),
                    "offset {offset} misses its object"
                );
            }
        }
        assert!(entries >= 4);
    }

    #[test]
    fn graphics_blocks_balance_and_font_references_resolve() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let pdf = Renderer::default().export_pdf(&package).unwrap();
        let text = String::from_utf8_lossy(&pdf);
        let count = |needle: &str| text.match_indices(needle).count();
        assert_eq!(count("\nq\n"), count("\nQ\n"));
        assert!(count("\nq\n") > 0);
        assert_eq!(count("BT "), count(" ET\n"));
        let mut defined = std::collections::BTreeSet::new();
        let mut rest: &str = &text;
        while let Some(at) = rest.find("/F") {
            let candidate = &rest[at + 2..];
            let digits: String = candidate
                .chars()
                .take_while(|ch| ch.is_ascii_digit())
                .collect();
            let tokens: Vec<&str> = candidate[digits.len()..]
                .split_whitespace()
                .take(3)
                .collect();
            if !digits.is_empty() && tokens.len() == 3 && tokens[1] == "0" && tokens[2] == "R" {
                defined.insert(format!("/F{digits}"));
            }
            rest = &candidate[digits.len().max(1)..];
        }
        for caps in text.match_indices("BT /F") {
            let name: String = text[caps.0 + 3..]
                .chars()
                .take_while(|ch| *ch != ' ')
                .collect();
            assert!(defined.contains(name.as_str()), "undefined font {name}");
        }
    }
}
