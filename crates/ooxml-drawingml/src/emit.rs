//! From-scratch DrawingML emission over neutral inputs.

use crate::GeometryPathCommand;

pub const EMU_PER_INCH: f64 = 914400.0;
const EMU_PER_DEGREE: f64 = 60000.0;

/// Converts inches to English Metric Units.
pub fn emu(inches: f64) -> i64 {
    (inches * EMU_PER_INCH).round() as i64
}

/// Escapes text for embedding in XML element content or attributes.
pub fn escape_xml(value: &str) -> String {
    let mut out = String::with_capacity(value.len());
    for ch in value.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&apos;"),
            _ => out.push(ch),
        }
    }
    out
}

/// Returns the hex body of a `#rrggbb` colour, or `None` when malformed.
pub fn srgb_hex(color: &str) -> Option<&str> {
    let hex = color.strip_prefix('#').unwrap_or(color);
    (hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit())).then_some(hex)
}

/// Row-major 2D affine map taking local inches to page inches, Y-up.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Matrix {
    pub a: f32,
    pub b: f32,
    pub c: f32,
    pub d: f32,
    pub e: f32,
    pub f: f32,
}

impl Matrix {
    /// Identity map.
    pub const fn identity() -> Self {
        Self {
            a: 1.0,
            b: 0.0,
            c: 0.0,
            d: 1.0,
            e: 0.0,
            f: 0.0,
        }
    }

    fn apply(self, x: f32, y: f32) -> (f32, f32) {
        (
            self.a * x + self.c * y + self.e,
            self.b * x + self.d * y + self.f,
        )
    }
}

/// Placed DrawingML frame in EMUs, Y-down from the page top.
#[derive(Clone, Copy, Debug, PartialEq)]
pub struct Placed {
    pub x: i64,
    pub y: i64,
    pub w: i64,
    pub h: i64,
    pub rot: i64,
    pub flip_h: bool,
    pub flip_v: bool,
    pub degraded: bool,
}

/// Places a local rect as a pre-rotation frame; skew falls back to the bbox.
pub fn place_rect(
    x: f32,
    y: f32,
    width: f32,
    height: f32,
    matrix: Matrix,
    page_height_in: f64,
) -> Option<Placed> {
    let corners = [
        matrix.apply(x, y),
        matrix.apply(x + width, y),
        matrix.apply(x, y + height),
        matrix.apply(x + width, y + height),
    ];
    if corners
        .iter()
        .any(|(px, py)| !px.is_finite() || !py.is_finite())
    {
        return None;
    }
    if width < 0.0
        || height < 0.0
        || !width.is_finite()
        || !height.is_finite()
        || !page_height_in.is_finite()
    {
        return None;
    }
    if let Some((rotation, flip_h, flip_v)) = affine_placement(matrix) {
        let (cx, cy) = matrix.apply(x + width / 2.0, y + height / 2.0);
        let ext_w = f64::from(width)
            * (f64::from(matrix.a) * f64::from(matrix.a)
                + f64::from(matrix.b) * f64::from(matrix.b))
            .sqrt();
        let ext_h = f64::from(height)
            * (f64::from(matrix.c) * f64::from(matrix.c)
                + f64::from(matrix.d) * f64::from(matrix.d))
            .sqrt();
        if !ext_w.is_finite() || !ext_h.is_finite() {
            return None;
        }
        return Some(Placed {
            x: emu(f64::from(cx) - ext_w / 2.0),
            y: emu(page_height_in - (f64::from(cy) + ext_h / 2.0)),
            w: emu(ext_w).max(1),
            h: emu(ext_h).max(1),
            rot: rotation,
            flip_h,
            flip_v,
            degraded: false,
        });
    }
    let min_x = corners
        .iter()
        .map(|corner| corner.0)
        .fold(f32::INFINITY, f32::min);
    let max_x = corners
        .iter()
        .map(|corner| corner.0)
        .fold(f32::NEG_INFINITY, f32::max);
    let min_y = corners
        .iter()
        .map(|corner| corner.1)
        .fold(f32::INFINITY, f32::min);
    let max_y = corners
        .iter()
        .map(|corner| corner.1)
        .fold(f32::NEG_INFINITY, f32::max);
    Some(Placed {
        x: emu(f64::from(min_x)),
        y: emu(page_height_in - f64::from(max_y)),
        w: emu(f64::from(max_x - min_x)).max(1),
        h: emu(f64::from(max_y - min_y)).max(1),
        rot: 0,
        flip_h: false,
        flip_v: false,
        degraded: true,
    })
}

fn affine_placement(matrix: Matrix) -> Option<(i64, bool, bool)> {
    const EPSILON: f32 = 1e-6;
    if matrix.b.abs() <= EPSILON && matrix.c.abs() <= EPSILON {
        return Some((0, matrix.a < -EPSILON, matrix.d < -EPSILON));
    }
    let x_scale = f64::from(matrix.a).hypot(f64::from(matrix.b));
    let y_scale = f64::from(matrix.c).hypot(f64::from(matrix.d));
    if x_scale <= f64::from(EPSILON) || y_scale <= f64::from(EPSILON) {
        return None;
    }
    let cos = f64::from(matrix.a) / x_scale;
    let sin = f64::from(matrix.b) / x_scale;
    let y_x = f64::from(matrix.c) / y_scale;
    let y_y = f64::from(matrix.d) / y_scale;
    if (cos * y_x + sin * y_y).abs() > f64::from(EPSILON) {
        return None;
    }
    Some((
        (-sin.atan2(cos).to_degrees() * EMU_PER_DEGREE).round() as i64,
        false,
        cos * y_y - sin * y_x < 0.0,
    ))
}

/// Renders an `a:xfrm` element for a placed frame.
pub fn xfrm_xml(placed: &Placed) -> String {
    let flip_h = if placed.flip_h { " flipH=\"1\"" } else { "" };
    let flip_v = if placed.flip_v { " flipV=\"1\"" } else { "" };
    if placed.rot == 0 {
        format!(
            "<a:xfrm{flip_h}{flip_v}><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm>",
            placed.x, placed.y, placed.w, placed.h
        )
    } else {
        format!(
            "<a:xfrm rot=\"{}\"{flip_h}{flip_v}><a:off x=\"{}\" y=\"{}\"/><a:ext cx=\"{}\" cy=\"{}\"/></a:xfrm>",
            placed.rot, placed.x, placed.y, placed.w, placed.h
        )
    }
}

/// Custom geometry frame with its path body, in EMUs Y-down.
#[derive(Clone, Debug, PartialEq)]
pub struct CustGeom {
    pub x: i64,
    pub y: i64,
    pub w: i64,
    pub h: i64,
    pub body: String,
}

/// Converts a Y-up inch path to a `custGeom` frame, or `None` when empty.
pub fn cust_geom(path: &[GeometryPathCommand], page_height_in: f64) -> Option<CustGeom> {
    if path.is_empty() {
        return None;
    }
    let mut points: Vec<(f64, f64)> = Vec::new();
    for command in path {
        match command {
            GeometryPathCommand::Move { x, y } | GeometryPathCommand::Line { x, y } => {
                points.push((*x, page_height_in - *y));
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                points.push((*cpx, page_height_in - *cpy));
                points.push((*x, page_height_in - *y));
            }
            GeometryPathCommand::Cubic {
                cp1x,
                cp1y,
                cp2x,
                cp2y,
                x,
                y,
            } => {
                points.push((*cp1x, page_height_in - *cp1y));
                points.push((*cp2x, page_height_in - *cp2y));
                points.push((*x, page_height_in - *y));
            }
            GeometryPathCommand::Close => {}
        }
    }
    if points.iter().any(|(x, y)| !x.is_finite() || !y.is_finite()) {
        return None;
    }
    let min_x = points
        .iter()
        .map(|point| point.0)
        .fold(f64::INFINITY, f64::min);
    let max_x = points
        .iter()
        .map(|point| point.0)
        .fold(f64::NEG_INFINITY, f64::max);
    let min_y = points
        .iter()
        .map(|point| point.1)
        .fold(f64::INFINITY, f64::min);
    let max_y = points
        .iter()
        .map(|point| point.1)
        .fold(f64::NEG_INFINITY, f64::max);
    let x = emu(min_x);
    let y = emu(min_y);
    let w = emu(max_x - min_x).max(1);
    let h = emu(max_y - min_y).max(1);
    let pt = |x: f64, y: f64| format!("<a:pt x=\"{}\" y=\"{}\"/>", emu(x - min_x), emu(y - min_y));
    let mut body = String::new();
    for command in path {
        match command {
            GeometryPathCommand::Move { x, y } => {
                body.push_str(&format!(
                    "<a:moveTo>{}</a:moveTo>",
                    pt(*x, page_height_in - *y)
                ));
            }
            GeometryPathCommand::Line { x, y } => {
                body.push_str(&format!("<a:lnTo>{}</a:lnTo>", pt(*x, page_height_in - *y)));
            }
            GeometryPathCommand::Quad { cpx, cpy, x, y } => {
                body.push_str(&format!(
                    "<a:quadBezTo>{}{}</a:quadBezTo>",
                    pt(*cpx, page_height_in - *cpy),
                    pt(*x, page_height_in - *y)
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
                body.push_str(&format!(
                    "<a:cubicBezTo>{}{}{}</a:cubicBezTo>",
                    pt(*cp1x, page_height_in - *cp1y),
                    pt(*cp2x, page_height_in - *cp2y),
                    pt(*x, page_height_in - *y)
                ));
            }
            GeometryPathCommand::Close => body.push_str("<a:close/>"),
        }
    }
    Some(CustGeom { x, y, w, h, body })
}

/// One gradient stop with its colour as `#rrggbb`.
pub struct EmitGradientStop<'a> {
    pub position: f32,
    pub color: &'a str,
}

/// Fill over neutral inputs; unparseable colours degrade to no fill.
pub enum Fill<'a> {
    NoFill,
    Solid {
        hex: &'a str,
    },
    Gradient {
        stops: &'a [EmitGradientStop<'a>],
        /// Degrees clockwise from the positive x-axis; 90 when absent.
        angle_deg: Option<f32>,
    },
}

/// DrawingML `ang`, in 60000ths of a degree clockwise from the positive x-axis.
fn gradient_angle(angle_deg: Option<f32>) -> i64 {
    let degrees = angle_deg.filter(|value| value.is_finite()).unwrap_or(90.0);
    let units = (f64::from(degrees).rem_euclid(360.0) * 60_000.0).round() as i64;
    units.rem_euclid(21_600_000)
}

/// Renders a fill element (`a:solidFill`, `a:gradFill` or `a:noFill`).
pub fn fill_xml(fill: &Fill) -> String {
    match fill {
        Fill::NoFill => "<a:noFill/>".to_owned(),
        Fill::Solid { hex } => match srgb_hex(hex) {
            Some(hex) => format!("<a:solidFill><a:srgbClr val=\"{hex}\"/></a:solidFill>"),
            None => "<a:noFill/>".to_owned(),
        },
        Fill::Gradient { stops, angle_deg } => {
            let mut list = String::new();
            for stop in stops.iter().filter(|stop| stop.position.is_finite()) {
                let Some(hex) = srgb_hex(stop.color) else {
                    continue;
                };
                let position = (f64::from(stop.position) * 1000.0)
                    .round()
                    .clamp(0.0, 100000.0);
                list.push_str(&format!(
                    "<a:gs pos=\"{position}\"><a:srgbClr val=\"{hex}\"/></a:gs>"
                ));
            }
            if list.is_empty() {
                return "<a:noFill/>".to_owned();
            }
            format!(
                "<a:gradFill><a:gsLst>{list}</a:gsLst><a:lin ang=\"{}\" scaled=\"0\"/></a:gradFill>",
                gradient_angle(*angle_deg)
            )
        }
    }
}

/// Outline over neutral inputs with width in inches.
pub struct Line<'a> {
    pub color: &'a str,
    pub width_in: f64,
    pub dashed: bool,
}

/// Renders an `a:ln` element; absent for no line, `noFill` for bad colour.
pub fn line_xml(line: Option<&Line>) -> String {
    let Some(line) = line else {
        return String::new();
    };
    let width = emu(line.width_in).max(1);
    let dash = if line.dashed {
        "<a:prstDash val=\"dash\"/>"
    } else {
        ""
    };
    match srgb_hex(line.color) {
        Some(color) => format!(
            "<a:ln w=\"{width}\"><a:solidFill><a:srgbClr val=\"{color}\"/></a:solidFill>{dash}</a:ln>"
        ),
        None => format!("<a:ln w=\"{width}\"><a:noFill/>{dash}</a:ln>"),
    }
}

/// One DrawingML text run over neutral inputs.
pub struct EmitRun<'a> {
    pub text: &'a str,
    pub family: &'a str,
    pub size_in: f32,
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub small_caps: bool,
    pub superscript: bool,
    pub subscript: bool,
    pub color: &'a str,
}

/// One DrawingML paragraph of runs.
pub struct EmitParagraph<'a> {
    pub runs: &'a [EmitRun<'a>],
}

/// Reports whether any paragraph carries text.
pub fn has_emit_text(paragraphs: &[EmitParagraph]) -> bool {
    paragraphs
        .iter()
        .any(|paragraph| paragraph.runs.iter().any(|run| !run.text.is_empty()))
}

fn font_points(size_in: f32) -> i64 {
    (f64::from(size_in) * 72.0 * 100.0).round() as i64
}

fn run_props(run: &EmitRun) -> String {
    let mut props = format!(" sz=\"{}\"", font_points(run.size_in).max(100));
    if run.bold {
        props.push_str(" b=\"1\"");
    }
    if run.italic {
        props.push_str(" i=\"1\"");
    }
    if run.underline {
        props.push_str(" u=\"sng\"");
    }
    if run.superscript {
        props.push_str(" baseline=\"30000\"");
    } else if run.subscript {
        props.push_str(" baseline=\"-25000\"");
    }
    if run.small_caps {
        props.push_str(" cap=\"smCaps\"");
    }
    props
}

/// Renders DrawingML paragraphs (`a:p`/`a:r`/`a:rPr`) for the given runs.
pub fn dml_paragraphs(paragraphs: &[EmitParagraph]) -> String {
    let mut out = String::new();
    for paragraph in paragraphs {
        out.push_str("<a:p>");
        for run in paragraph.runs {
            let text = escape_xml(run.text);
            let props = run_props(run);
            let color = match srgb_hex(run.color) {
                Some(hex) => format!("<a:solidFill><a:srgbClr val=\"{hex}\"/></a:solidFill>"),
                None => String::new(),
            };
            let family = escape_xml(run.family);
            let face = format!("<a:latin typeface=\"{family}\"/>");
            out.push_str(&format!(
                "<a:r><a:rPr{props}>{color}{face}</a:rPr><a:t>{text}</a:t></a:r>"
            ));
        }
        out.push_str("<a:endParaRPr/>");
        out.push_str("</a:p>");
    }
    out
}

#[cfg(test)]
mod tests {
    use super::*;
    use std::f32::consts::FRAC_PI_4;

    #[test]
    fn translation_keeps_size_and_offset() {
        let placed = place_rect(
            1.0,
            2.0,
            3.0,
            4.0,
            Matrix {
                e: 5.0,
                f: 1.0,
                ..Matrix::identity()
            },
            12.0,
        )
        .unwrap();
        assert!(!placed.degraded);
        assert_eq!((placed.x, placed.w), (emu(6.0), emu(3.0)));
        assert_eq!((placed.y, placed.h), (emu(12.0 - 7.0), emu(4.0)));
        assert_eq!(placed.rot, 0);
    }

    #[test]
    fn rotation_uses_the_pre_rotation_frame() {
        let (sin, cos) = FRAC_PI_4.sin_cos();
        let placed = place_rect(
            0.0,
            0.0,
            1.0,
            1.0,
            Matrix {
                a: cos,
                b: sin,
                c: -sin,
                d: cos,
                e: 0.0,
                f: 0.0,
            },
            4.0,
        )
        .unwrap();
        assert!(!placed.degraded);
        assert_eq!((placed.w, placed.h), (emu(1.0), emu(1.0)));
        assert_eq!(placed.rot, -2700000);
        let root_half = std::f64::consts::SQRT_2 / 2.0;
        let expected = (emu(-0.5), emu(4.0 - (root_half + 0.5)));
        assert!((placed.x - expected.0).abs() <= 2, "{placed:?}");
        assert!((placed.y - expected.1).abs() <= 2, "{placed:?}");
    }

    #[test]
    fn quarter_turn_keeps_aspect_with_rotation() {
        let placed = place_rect(
            0.0,
            0.0,
            2.0,
            1.0,
            Matrix {
                a: 0.0,
                b: 1.0,
                c: -1.0,
                d: 0.0,
                e: 2.0,
                f: 0.0,
            },
            4.0,
        )
        .unwrap();
        assert!(!placed.degraded);
        assert_eq!((placed.w, placed.h), (emu(2.0), emu(1.0)));
        assert_ne!(placed.rot, 0);
    }

    #[test]
    fn skew_falls_back_to_the_upright_box() {
        let placed = place_rect(
            0.0,
            0.0,
            1.0,
            1.0,
            Matrix {
                a: 1.0,
                b: 0.5,
                c: 0.0,
                d: 1.0,
                e: 0.0,
                f: 0.0,
            },
            2.0,
        )
        .unwrap();
        assert!(placed.degraded);
        assert_eq!(placed.rot, 0);
        assert_eq!((placed.w, placed.h), (emu(1.0), emu(1.5)));
    }

    #[test]
    fn rotated_nonuniform_scale_and_reflection_stay_native() {
        for (c, reflected) in [(-3.0, false), (3.0, true)] {
            let placed = place_rect(
                0.0,
                0.0,
                2.0,
                1.0,
                Matrix {
                    a: 0.0,
                    b: 2.0,
                    c,
                    d: 0.0,
                    e: 0.0,
                    f: 0.0,
                },
                8.0,
            )
            .unwrap();
            assert!(!placed.degraded);
            assert_eq!((placed.w, placed.h), (emu(4.0), emu(3.0)));
            assert_eq!(placed.rot, -5_400_000);
            assert!(!placed.flip_h);
            assert_eq!(placed.flip_v, reflected);
        }
    }

    #[test]
    fn invalid_gradient_positions_do_not_emit_invalid_xml() {
        let stops = [EmitGradientStop {
            position: f32::NAN,
            color: "#FF0000",
        }];
        assert_eq!(
            fill_xml(&Fill::Gradient {
                stops: &stops,
                angle_deg: None
            }),
            "<a:noFill/>"
        );
        assert!((0..21_600_000).contains(&gradient_angle(Some(f32::MAX))));
    }

    #[test]
    fn reflected_rect_emits_flip_flag() {
        let placed = place_rect(
            0.0,
            0.0,
            1.0,
            1.0,
            Matrix {
                a: -1.0,
                e: 1.0,
                ..Matrix::identity()
            },
            1.0,
        )
        .unwrap();
        assert!(placed.flip_h);
        assert!(!placed.flip_v);
    }

    #[test]
    fn bad_line_colour_degrades_instead_of_vanishing() {
        assert_eq!(line_xml(None), "");
        let bad = line_xml(Some(&Line {
            color: "not-a-colour",
            width_in: 0.02,
            dashed: false,
        }));
        assert!(bad.contains("<a:noFill/>"), "{bad}");
        assert_eq!(
            fill_xml(&Fill::Solid {
                hex: "not-a-colour"
            }),
            "<a:noFill/>"
        );
    }

    #[test]
    fn gradient_emits_stops() {
        let stops = [
            EmitGradientStop {
                position: 0.0,
                color: "#FF0000",
            },
            EmitGradientStop {
                position: 100.0,
                color: "#0000FF",
            },
        ];
        let xml = fill_xml(&Fill::Gradient {
            stops: &stops,
            angle_deg: None,
        });
        assert!(xml.contains("<a:gradFill>"), "{xml}");
        assert!(xml.contains("val=\"FF0000\""), "{xml}");
        assert!(xml.contains("ang=\"5400000\""), "{xml}");
    }

    #[test]
    fn gradient_carries_its_angle() {
        let stops = [EmitGradientStop {
            position: 0.0,
            color: "#FF0000",
        }];
        for (degrees, units) in [(0.0, 0), (45.0, 2_700_000), (270.0, 16_200_000)] {
            let xml = fill_xml(&Fill::Gradient {
                stops: &stops,
                angle_deg: Some(degrees),
            });
            assert!(
                xml.contains(&format!("ang=\"{units}\"")),
                "{degrees}: {xml}"
            );
        }
        let wrapped = fill_xml(&Fill::Gradient {
            stops: &stops,
            angle_deg: Some(-90.0),
        });
        assert!(wrapped.contains("ang=\"16200000\""), "{wrapped}");
        let invalid = fill_xml(&Fill::Gradient {
            stops: &stops,
            angle_deg: Some(f32::NAN),
        });
        assert!(invalid.contains("ang=\"5400000\""), "{invalid}");
    }
}
