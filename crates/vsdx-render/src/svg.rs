use base64::Engine as _;
use ooxml_drawingml::GeometryPathCommand;
use vsdx_parse::VsdxPackage;

use crate::display_list::{Affine, Primitive, VsdxDisplayList};
use crate::vector::{
    collect_ordered, escape, fragments, jpeg_dimensions, linear_gradient, matrix, num,
    png_dimensions, solid_color, z_order,
};
use crate::{RenderError, Renderer};

struct Emitter<'a> {
    package: &'a VsdxPackage,
    out: String,
    defs: String,
    gradients: usize,
    shadows: usize,
}

/// Splits `#RRGGBBAA` into the flood colour and its opacity.
fn flood(color: &str) -> (String, f64) {
    let digits = color.strip_prefix('#').unwrap_or(color);
    if digits.len() == 8
        && digits.bytes().all(|byte| byte.is_ascii_hexdigit())
        && let Ok(alpha) = u8::from_str_radix(&digits[6..8], 16)
    {
        return (format!("#{}", &digits[..6]), f64::from(alpha) / 255.0);
    }
    (color.to_owned(), 1.0)
}

fn path_data(path: &[GeometryPathCommand]) -> String {
    let mut data = String::new();
    for command in path {
        match command {
            GeometryPathCommand::Move { x, y } => {
                data.push_str(&format!("M{} {} ", num(*x), num(*y)));
            }
            GeometryPathCommand::Line { x, y } => {
                data.push_str(&format!("L{} {} ", num(*x), num(*y)));
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                data.push_str(&format!(
                    "Q{} {} {} {} ",
                    num(*cpx),
                    num(*cpy),
                    num(*x),
                    num(*y)
                ));
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                data.push_str(&format!(
                    "C{} {} {} {} {} {} ",
                    num(*cp1x),
                    num(*cp1y),
                    num(*cp2x),
                    num(*cp2y),
                    num(*x),
                    num(*y)
                ));
            }
            GeometryPathCommand::Close => data.push_str("Z "),
        }
    }
    data.trim_end().to_owned()
}

fn stroke_attributes(stroke: &Option<crate::display_list::Stroke>) -> String {
    let mut attributes = String::new();
    if let Some(stroke) = stroke
        && crate::vector::rgb(&stroke.color).is_some()
    {
        attributes.push_str(&format!(
            " stroke=\"{}\" stroke-width=\"{}",
            escape(&stroke.color),
            num(f64::from(stroke.width).max(0.0))
        ));
        if stroke.dashed {
            let step = num(f64::from(stroke.width).max(0.0) * 2.0);
            attributes.push_str(&format!("\" stroke-dasharray=\"{step} {step}"));
        }
        attributes.push('"');
    }
    attributes
}

fn translate(x: f32, y: f32) -> Affine {
    Affine {
        a: 1.0,
        b: 0.0,
        c: 0.0,
        d: 1.0,
        e: x,
        f: y,
    }
}

const FLIP_Y: Affine = Affine {
    a: 1.0,
    b: 0.0,
    c: 0.0,
    d: -1.0,
    e: 0.0,
    f: 0.0,
};

impl<'a> Emitter<'a> {
    /// Multi-stop fills become a `<defs>` gradient; everything else stays flat.
    fn fill_attribute(
        &mut self,
        fill: &Option<crate::display_list::Paint>,
        path: &[GeometryPathCommand],
    ) -> String {
        if let Some(paint) = fill
            && let Some(gradient) = linear_gradient(paint, path)
        {
            let id = format!("vsdxGradient{}", self.gradients);
            self.gradients += 1;
            self.defs.push_str(&format!(
                "<linearGradient id=\"{id}\" gradientUnits=\"userSpaceOnUse\" x1=\"{}\" y1=\"{}\" x2=\"{}\" y2=\"{}\">",
                num(f64::from(gradient.start.0)),
                num(f64::from(gradient.start.1)),
                num(f64::from(gradient.end.0)),
                num(f64::from(gradient.end.1))
            ));
            for (position, color) in &gradient.stops {
                self.defs.push_str(&format!(
                    "<stop offset=\"{}\" stop-color=\"{}\"/>",
                    num(f64::from(*position)),
                    escape(color)
                ));
            }
            self.defs.push_str("</linearGradient>");
            return format!(" fill=\"url(#{id})\"");
        }
        match solid_color(fill) {
            Some(color) => format!(" fill=\"{}\"", escape(color)),
            None => " fill=\"none\"".to_owned(),
        }
    }

    /// Offsets stay in the element's own inch space, so the page flip carries them.
    fn shadow_attribute(&mut self, shadow: &Option<crate::display_list::Shadow>) -> String {
        let Some(shadow) = shadow else {
            return String::new();
        };
        let (color, opacity) = flood(&shadow.color);
        if crate::vector::rgb(&color).is_none() {
            return String::new();
        }
        let id = format!("vsdxShadow{}", self.shadows);
        self.shadows += 1;
        self.defs.push_str(&format!(
            "<filter id=\"{id}\" x=\"-50%\" y=\"-50%\" width=\"200%\" height=\"200%\"><feDropShadow dx=\"{}\" dy=\"{}\" stdDeviation=\"{}\" flood-color=\"{}\" flood-opacity=\"{}\"/></filter>",
            num(f64::from(shadow.offset_x_in)),
            num(f64::from(shadow.offset_y_in)),
            num(f64::from(shadow.blur_in).max(0.0) / 2.0),
            escape(&color),
            num(opacity)
        ));
        format!(" filter=\"url(#{id})\"")
    }

    fn primitive(&mut self, primitive: &Primitive, outer: Affine) {
        let mut ordered = Vec::new();
        collect_ordered(primitive, &mut ordered);
        for (item, transform) in ordered {
            let composed = outer.compose(transform);
            match item {
                Primitive::Shape {
                    path,
                    fill,
                    stroke,
                    shadow,
                    ..
                } => {
                    let data = path_data(path);
                    if data.is_empty() {
                        continue;
                    }
                    let mut element = String::from("<path d=\"");
                    element.push_str(&escape(&data));
                    element.push('"');
                    element.push_str(&self.fill_attribute(fill, path));
                    element.push_str(&stroke_attributes(stroke));
                    element.push_str(&self.shadow_attribute(shadow));
                    element.push_str("/>");
                    self.wrapped(&element, composed);
                }
                Primitive::TextBox {
                    y,
                    height,
                    paragraphs,
                    lines,
                    ..
                } => {
                    let frame = composed
                        .compose(translate(0.0, y * 2.0 + height))
                        .compose(FLIP_Y);
                    for run in fragments(paragraphs, lines) {
                        let anchor = frame.compose(translate(run.x, run.line_y));
                        self.out.push_str(&format!(
                            "<text transform=\"{}\" x=\"0\" y=\"0\" dominant-baseline=\"text-before-edge\" font-family=\"'{}', {}\" font-size=\"{}\" fill=\"{}\"",
                            matrix(anchor),
                            escape(&run.family),
                            run.generic,
                            num(f64::from(run.size_in)),
                            escape(&run.color)
                        ));
                        if run.bold {
                            self.out.push_str(" font-weight=\"bold\"");
                        }
                        if run.italic {
                            self.out.push_str(" font-style=\"italic\"");
                        }
                        if run.underline {
                            self.out.push_str(" text-decoration=\"underline\"");
                        }
                        if run.letter_spacing != 0.0 && run.letter_spacing.is_finite() {
                            self.out.push_str(&format!(
                                " letter-spacing=\"{}\"",
                                num(f64::from(run.letter_spacing))
                            ));
                        }
                        self.out.push('>');
                        self.out.push_str(&escape(&run.text));
                        self.out.push_str("</text>");
                    }
                }
                Primitive::Image {
                    asset_id,
                    x,
                    y,
                    width,
                    height,
                    ..
                } => self.image(asset_id, *x, *y, *width, *height, composed),
                Primitive::Placeholder {
                    x,
                    y,
                    width,
                    height,
                    reason,
                    ..
                } => self.placeholder(*x, *y, *width, *height, reason, composed),
                Primitive::Group { .. } => {}
            }
        }
    }

    fn wrapped(&mut self, element: &str, transform: Affine) {
        if transform.is_identity() {
            self.out.push_str(element);
            return;
        }
        self.out.push_str(&format!(
            "<g transform=\"{}\">{element}</g>",
            matrix(transform)
        ));
    }

    fn image(
        &mut self,
        asset_id: &str,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        transform: Affine,
    ) {
        let uri = self.package.part_bytes(asset_id).and_then(|bytes| {
            if jpeg_dimensions(bytes).is_some() {
                Some((bytes, "image/jpeg"))
            } else if png_dimensions(bytes).is_some() {
                Some((bytes, "image/png"))
            } else {
                None
            }
        });
        let Some((bytes, mime)) = uri else {
            self.placeholder(x, y, width, height, "", transform);
            return;
        };
        let encoded = base64::engine::general_purpose::STANDARD.encode(bytes);
        let anchor = transform
            .compose(translate(0.0, y * 2.0 + height))
            .compose(FLIP_Y);
        let element = format!(
            "<image transform=\"{}\" x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" preserveAspectRatio=\"none\" href=\"data:{mime};base64,{encoded}\"/>",
            matrix(anchor),
            num(f64::from(x)),
            num(f64::from(y)),
            num(f64::from(width).max(0.0)),
            num(f64::from(height).max(0.0))
        );
        self.out.push_str(&element);
    }

    fn placeholder(
        &mut self,
        x: f32,
        y: f32,
        width: f32,
        height: f32,
        reason: &str,
        transform: Affine,
    ) {
        let mut element = format!(
            "<rect x=\"{}\" y=\"{}\" width=\"{}\" height=\"{}\" fill=\"none\" stroke=\"#8a94a6\" stroke-width=\"{}\" stroke-dasharray=\"{} {}\"/>",
            num(f64::from(x)),
            num(f64::from(y)),
            num(f64::from(width).max(0.0)),
            num(f64::from(height).max(0.0)),
            num(1.0 / 96.0),
            num(5.0 / 96.0),
            num(4.0 / 96.0)
        );
        if !reason.is_empty() {
            let anchor = transform.compose(translate(x, y)).compose(FLIP_Y);
            element.push_str(&format!(
                "<text transform=\"{}\" x=\"0\" y=\"0\" dominant-baseline=\"text-before-edge\" font-family=\"'sans-serif'\" font-size=\"{}\" fill=\"#5d6675\">{}</text>",
                matrix(anchor),
                num(10.0 / 72.0),
                escape(reason)
            ));
        }
        self.wrapped(&element, transform);
    }
}

fn emit_page(list: &VsdxDisplayList, package: &VsdxPackage) -> String {
    let mut primitives: Vec<&Primitive> = list.primitives.iter().collect();
    primitives.sort_by_key(|primitive| z_order(primitive));
    let mut emitter = Emitter {
        package,
        out: String::new(),
        defs: String::new(),
        gradients: 0,
        shadows: 0,
    };
    for primitive in primitives {
        emitter.primitive(primitive, Affine::identity());
    }
    let defs = if emitter.defs.is_empty() {
        String::new()
    } else {
        format!("<defs>{}</defs>", emitter.defs)
    };
    let paint = list.paint_transform;
    format!(
        "<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"{width}\" height=\"{height}\" viewBox=\"0 0 {width} {height}\">{defs}<g transform=\"{transform}\">{body}</g></svg>",
        width = num(f64::from(list.width)),
        height = num(f64::from(list.height)),
        transform = matrix(Affine {
            a: paint.a,
            b: paint.b,
            c: paint.c,
            d: paint.d,
            e: paint.e,
            f: paint.f,
        }),
        body = emitter.out
    )
}

impl Renderer {
    /// Renders every diagram page to one SVG string per page, sized from the PageSheet.
    pub fn export_svg(&self, package: &VsdxPackage) -> Result<Vec<String>, RenderError> {
        let mut pages = Vec::with_capacity(package.page_part_paths.len().max(1));
        for part in &package.page_part_paths.clone() {
            pages.push(emit_page(&self.layout_page(package, part)?, package));
        }
        Ok(pages)
    }

    /// Renders one diagram page to an SVG string sized from the PageSheet.
    pub fn export_svg_page(
        &self,
        package: &VsdxPackage,
        page_index: usize,
    ) -> Result<String, RenderError> {
        let part = package
            .page_part_paths
            .get(page_index)
            .ok_or_else(|| RenderError::MissingPage(format!("page index {page_index}")))?;
        Ok(emit_page(&self.layout_page(package, part)?, package))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn package() -> VsdxPackage {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
        vsdx_parse::parse_vsdx(source).unwrap()
    }

    #[test]
    fn exports_one_svg_per_diagram_page() {
        let package = package();
        let pages = Renderer::default().export_svg(&package).unwrap();
        assert_eq!(pages.len(), package.page_part_paths.len());
        assert!(!pages.is_empty());
        for page in &pages {
            assert!(page.starts_with("<svg xmlns=\"http://www.w3.org/2000/svg\""));
            assert!(page.ends_with("</svg>"));
        }
    }

    #[test]
    fn sizes_pages_from_the_display_list_rather_than_a_hardcoded_format() {
        let package = package();
        let list = Renderer::default()
            .layout_page(&package, &package.page_part_paths[0])
            .unwrap();
        let expected = format!(
            "width=\"{}\" height=\"{}\" viewBox=\"0 0 {0} {1}\"",
            num(f64::from(list.width)),
            num(f64::from(list.height))
        );
        let pages = Renderer::default().export_svg(&package).unwrap();
        assert!(pages[0].contains(&expected), "missing {expected}");
        assert!(!pages[0].contains("viewBox=\"0 0 595") && !pages[0].contains("viewBox=\"0 0 842"));
    }

    #[test]
    fn emits_shape_geometry_as_path_data() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/indexed-geometry.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let pages = Renderer::default().export_svg(&package).unwrap();
        let body = pages.join("");
        assert!(body.contains("<path d=\"M"));
    }

    #[test]
    fn references_fonts_by_name_without_embedding() {
        let source = include_bytes!("../../vsdx-parse/tests/fixtures/text-accounting.vsdx");
        let package = vsdx_parse::parse_vsdx(source).unwrap();
        let pages = Renderer::default().export_svg(&package).unwrap();
        let body = pages.join("");
        assert!(body.contains("font-family="));
        assert!(!body.contains("@font-face") && !body.contains("data:font"));
    }

    #[test]
    fn rejects_an_unknown_page_index() {
        let package = package();
        assert!(Renderer::default().export_svg_page(&package, 99).is_err());
    }

    fn list_with(primitive: Primitive) -> VsdxDisplayList {
        VsdxDisplayList {
            contract_version: crate::CONTRACT_VERSION,
            width: 384.0,
            height: 384.0,
            print_width: 384.0,
            print_height: 384.0,
            paint_transform: crate::PaintTransform {
                a: 96.0,
                b: 0.0,
                c: 0.0,
                d: -96.0,
                e: 0.0,
                f: 384.0,
            },
            primitives: vec![primitive],
            connectors: Vec::new(),
        }
    }

    fn unit_rect() -> Vec<GeometryPathCommand> {
        vec![
            GeometryPathCommand::Move { x: 0.0, y: 0.0 },
            GeometryPathCommand::Line { x: 2.0, y: 0.0 },
            GeometryPathCommand::Line { x: 2.0, y: 1.0 },
            GeometryPathCommand::Line { x: 0.0, y: 1.0 },
            GeometryPathCommand::Close,
        ]
    }

    fn filled(stops: Vec<crate::display_list::GradientStop>) -> Primitive {
        Primitive::Shape {
            id: "rect".into(),
            z_order: 0,
            path: unit_rect(),
            fill: Some(crate::display_list::Paint::Gradient {
                angle_deg: Some(0.0),
                stops,
            }),
            stroke: None,
            shadow: None,
            transform: Affine::identity(),
            diagnostics: Vec::new(),
        }
    }

    fn stop(position: f32, color: &str) -> crate::display_list::GradientStop {
        crate::display_list::GradientStop {
            position,
            color: color.into(),
        }
    }

    #[test]
    fn shape_shadows_export_as_a_drop_shadow_filter() {
        let mut shape = filled(vec![stop(0.0, "#102030"), stop(1.0, "#405060")]);
        if let Primitive::Shape { shadow, .. } = &mut shape {
            *shadow = Some(crate::display_list::Shadow {
                color: "#11223380".into(),
                blur_in: 0.5,
                offset_x_in: 0.125,
                offset_y_in: -0.125,
            });
        }
        let svg = emit_page(&list_with(shape), &package());
        assert!(
            svg.contains("<feDropShadow dx=\"0.125\" dy=\"-0.125\" stdDeviation=\"0.25\" flood-color=\"#112233\" flood-opacity=\"0.502\"/>"),
            "missing drop shadow: {svg}"
        );
        assert!(svg.contains(" filter=\"url(#vsdxShadow0)\""), "{svg}");
    }

    #[test]
    fn multi_stop_fills_export_as_gradients_not_flat_colour() {
        let svg = emit_page(
            &list_with(filled(vec![stop(0.0, "#102030"), stop(1.0, "#405060")])),
            &package(),
        );
        assert!(
            svg.contains(
                "<defs><linearGradient id=\"vsdxGradient0\" gradientUnits=\"userSpaceOnUse\""
            ),
            "missing gradient definition: {svg}"
        );
        assert!(svg.contains(" fill=\"url(#vsdxGradient0)\""), "{svg}");
        assert!(svg.contains("stop-color=\"#102030\"") && svg.contains("stop-color=\"#405060\""));
        assert!(
            svg.contains(" y1=\"0.5\"") && svg.contains(" y2=\"0.5\""),
            "{svg}"
        );
    }

    #[test]
    fn single_stop_fills_stay_flat() {
        let svg = emit_page(&list_with(filled(vec![stop(0.0, "#102030")])), &package());
        assert!(svg.contains(" fill=\"#102030\""), "{svg}");
        assert!(!svg.contains("linearGradient"));
    }

    #[test]
    fn text_anchors_match_the_canvas_flip() {
        let run = crate::display_list::TextRun {
            text: "Hi".into(),
            family: "Arial".into(),
            size_in: 0.2,
            bold: false,
            italic: false,
            underline: false,
            small_caps: false,
            superscript: false,
            subscript: false,
            letter_spacing: 0.0,
            case: 0,
            color: "#102030".into(),
            diagnostics: Vec::new(),
            tab: None,
            diagnosed_face: None,
        };
        let svg = emit_page(
            &list_with(Primitive::TextBox {
                id: "t".into(),
                z_order: 0,
                x: 2.0,
                y: 1.0,
                width: 1.0,
                height: 1.0,
                paragraphs: vec![crate::display_list::TextParagraph { runs: vec![run] }],
                lines: vec![crate::display_list::PositionedLine {
                    x: 2.1,
                    y: 1.4,
                    width: 0.5,
                    height: 0.2,
                    start: 0,
                    end: 2,
                    caret_stops: Vec::new(),
                }],
                transform: Affine::identity(),
            }),
            &package(),
        );
        assert!(
            svg.contains("<text transform=\"matrix(1 0 0 -1 2.1 1.6)\""),
            "text must sit where the canvas flip puts it: {svg}"
        );
    }
}
