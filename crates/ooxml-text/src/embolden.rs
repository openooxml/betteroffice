//! Synthetic bold — real contour widening for faces without a bold cut.
//!
//! The outline is flattened to polylines and offset with `i_overlay`:
//! counter-clockwise (ink) contours grow outward, clockwise counters grow
//! inward, and the per-contour offsets are unioned, so overlapping glyph
//! components merge instead of showing seams.
//!
//! Output coordinates stay in the input's font-unit space (including any
//! advance widening) and are lines only — offsetting is a polygon operation,
//! so the original curve commands cannot survive it.

use i_overlay::mesh::float::outline::offset::OutlineOffset;
use i_overlay::mesh::float::style::{LineJoin, OutlineStyle};

use crate::outline::{GlyphOutline, PathCmd};

/// Ink growth on each side of every contour, as a fraction of the em.
/// A stem gains `2 / 64` em, the weight jump synthetic bold is expected to
/// show, and counters close at the same absolute rate.
pub const EMBOLDEN_EM_PER_SIDE: f32 = 1.0 / 64.0;

/// Curve flattening tolerance in em units, ~1 font unit at a 2048 upem.
const FLATTEN_EM: f32 = 1.0 / 2048.0;

/// Largest angle (radians) between segments of the rounded corner joins.
const JOIN_STEP_RADIANS: f32 = 0.25;

const MAX_FLATTEN_DEPTH: u8 = 10;

/// Guard against hostile outlines fanning out into a huge polygon set.
const MAX_FLATTENED_POINTS: usize = 16_384;

/// A widened copy of `outline`'s glyph, or `None` when no widened geometry is
/// available (blank glyph, non-finite coordinates, empty offset result).
///
/// The result is in the same design-unit space as the input and fills solid
/// under the non-zero rule; callers paint it instead of the crisp outline when
/// they synthesize bold.
pub fn embolden_glyph_outline(outline: &GlyphOutline) -> Option<Vec<PathCmd>> {
    if outline.cmds.is_empty() || outline.upem == 0 {
        return None;
    }

    let upem = f32::from(outline.upem);
    let mut contours = flatten(outline, upem * FLATTEN_EM)?;
    if contours.is_empty() {
        return None;
    }
    normalize_winding(&mut contours);

    let style = OutlineStyle::new(upem * EMBOLDEN_EM_PER_SIDE)
        .line_join(LineJoin::Round(JOIN_STEP_RADIANS));
    let shapes = contours.outline(&style);

    let mut cmds = Vec::new();
    for shape in shapes {
        for contour in shape {
            emit_contour(&contour, &mut cmds);
        }
    }

    (!cmds.is_empty()).then_some(cmds)
}

/// Flattens the path into closed contours; curves are subdivided until they
/// deviate from a chord by less than `tolerance`.
fn flatten(outline: &GlyphOutline, tolerance: f32) -> Option<Vec<Vec<[f32; 2]>>> {
    let mut contours = Vec::new();
    let mut current: Vec<[f32; 2]> = Vec::new();

    for cmd in &outline.cmds {
        match *cmd {
            PathCmd::MoveTo { x, y } => {
                take_contour(&mut contours, &mut current);
                current.push(finite(x, y)?);
            }
            PathCmd::LineTo { x, y } => {
                if current.is_empty() {
                    return None;
                }
                current.push(finite(x, y)?);
            }
            PathCmd::QuadTo { cx, cy, x, y } => {
                let start = *current.last()?;
                flatten_quad(
                    start,
                    finite(cx, cy)?,
                    finite(x, y)?,
                    tolerance,
                    0,
                    &mut current,
                );
                if current.len() > MAX_FLATTENED_POINTS {
                    return None;
                }
            }
            PathCmd::CubicTo {
                c1x,
                c1y,
                c2x,
                c2y,
                x,
                y,
            } => {
                let start = *current.last()?;
                flatten_cubic(
                    start,
                    finite(c1x, c1y)?,
                    finite(c2x, c2y)?,
                    finite(x, y)?,
                    tolerance,
                    0,
                    &mut current,
                );
                if current.len() > MAX_FLATTENED_POINTS {
                    return None;
                }
            }
            PathCmd::Close => take_contour(&mut contours, &mut current),
        }
    }
    take_contour(&mut contours, &mut current);

    Some(contours)
}

fn flatten_quad(
    p0: [f32; 2],
    control: [f32; 2],
    p1: [f32; 2],
    tolerance: f32,
    depth: u8,
    out: &mut Vec<[f32; 2]>,
) {
    if depth >= MAX_FLATTEN_DEPTH || 0.25 * norm(sub(add(p0, p1), scale(control, 2.0))) <= tolerance
    {
        out.push(p1);
        return;
    }
    let a = midpoint(p0, control);
    let b = midpoint(control, p1);
    let middle = midpoint(a, b);
    flatten_quad(p0, a, middle, tolerance, depth + 1, out);
    flatten_quad(middle, b, p1, tolerance, depth + 1, out);
}

fn flatten_cubic(
    p0: [f32; 2],
    c1: [f32; 2],
    c2: [f32; 2],
    p1: [f32; 2],
    tolerance: f32,
    depth: u8,
    out: &mut Vec<[f32; 2]>,
) {
    let deviation =
        0.75 * norm(sub(add(p0, c2), scale(c1, 2.0))).max(norm(sub(add(c1, p1), scale(c2, 2.0))));
    if depth >= MAX_FLATTEN_DEPTH || deviation <= tolerance {
        out.push(p1);
        return;
    }
    let a = midpoint(p0, c1);
    let b = midpoint(c1, c2);
    let c = midpoint(c2, p1);
    let ab = midpoint(a, b);
    let bc = midpoint(b, c);
    let middle = midpoint(ab, bc);
    flatten_cubic(p0, a, ab, middle, tolerance, depth + 1, out);
    flatten_cubic(middle, bc, c, p1, tolerance, depth + 1, out);
}

/// Outer contours must wind counter-clockwise for the offset to grow ink and
/// shrink counters. TrueType (`glyf`) glyphs wind outers clockwise, CFF ones
/// counter-clockwise; the largest contour is always an outer, so its sign
/// identifies the convention for the whole glyph.
fn normalize_winding(contours: &mut [Vec<[f32; 2]>]) {
    let mut largest = 0.0_f64;
    let mut largest_area = 0.0_f64;
    for contour in contours.iter() {
        let area = signed_area(contour);
        if area.abs() > largest {
            largest = area.abs();
            largest_area = area;
        }
    }
    if largest_area < 0.0 {
        for contour in contours.iter_mut() {
            contour.reverse();
        }
    }
}

fn signed_area(contour: &[[f32; 2]]) -> f64 {
    if contour.len() < 3 {
        return 0.0;
    }
    let mut area = 0.0_f64;
    for (index, a) in contour.iter().enumerate() {
        let b = contour[(index + 1) % contour.len()];
        area += f64::from(a[0]) * f64::from(b[1]) - f64::from(b[0]) * f64::from(a[1]);
    }
    0.5 * area
}

fn emit_contour(contour: &[[f32; 2]], cmds: &mut Vec<PathCmd>) {
    let contour = match contour.split_last() {
        Some((last, rest)) if !rest.is_empty() && *last == rest[0] => rest,
        _ => contour,
    };
    let Some((first, rest)) = contour.split_first() else {
        return;
    };
    if rest.len() < 2 {
        return;
    }
    cmds.push(PathCmd::MoveTo {
        x: first[0],
        y: first[1],
    });
    for point in rest {
        cmds.push(PathCmd::LineTo {
            x: point[0],
            y: point[1],
        });
    }
    cmds.push(PathCmd::Close);
}

fn take_contour(contours: &mut Vec<Vec<[f32; 2]>>, current: &mut Vec<[f32; 2]>) {
    if current.len() >= 3 {
        contours.push(std::mem::take(current));
    } else {
        current.clear();
    }
}

fn finite(x: f32, y: f32) -> Option<[f32; 2]> {
    (x.is_finite() && y.is_finite()).then_some([x, y])
}

fn add(a: [f32; 2], b: [f32; 2]) -> [f32; 2] {
    [a[0] + b[0], a[1] + b[1]]
}

fn sub(a: [f32; 2], b: [f32; 2]) -> [f32; 2] {
    [a[0] - b[0], a[1] - b[1]]
}

fn scale(point: [f32; 2], factor: f32) -> [f32; 2] {
    [point[0] * factor, point[1] * factor]
}

fn midpoint(a: [f32; 2], b: [f32; 2]) -> [f32; 2] {
    [(a[0] + b[0]) * 0.5, (a[1] + b[1]) * 0.5]
}

fn norm(point: [f32; 2]) -> f32 {
    point[0].hypot(point[1])
}

#[cfg(test)]
mod tests {
    use super::*;

    const LIBERATION_SANS: &[u8] = include_bytes!("../tests/fonts/LiberationSans-Regular.ttf");

    #[test]
    fn widens_a_glyph_and_keeps_its_counter() {
        let mut store = crate::FontStore::new();
        let id = store.register(LIBERATION_SANS.to_vec()).expect("registers");
        let glyph = store.glyph_id(id, 'o').expect("known font").expect("cmap");
        let crisp = store.outline_glyph(id, glyph).expect("outline 'o'");

        let widened = embolden_glyph_outline(&crisp).expect("'o' widens");

        let d = f64::from(f32::from(crisp.upem) * EMBOLDEN_EM_PER_SIDE);
        let crisp = sampled(&crisp.cmds);
        let wide = sampled(&widened);
        assert_eq!(crisp.len(), 2, "'o' has an outer and one counter");
        assert_eq!(wide.len(), 2, "widening keeps the counter");

        let (outer, counter) = widest(&crisp);
        let (wide_outer, wide_counter) = widest(&wide);
        // LiberationSans is glyf: its outers wind clockwise, and only the
        // widened output is normalized to counter-clockwise ink.
        assert!(signed_area(&crisp[outer]) < 0.0);
        assert!(signed_area(&wide[wide_outer]) > 0.0);
        assert!(signed_area(&wide[wide_counter]) < 0.0);

        assert_close(
            "outer grows",
            signed_area(&wide[wide_outer]),
            grown_out(
                signed_area(&crisp[outer]).abs(),
                perimeter(&crisp[outer]),
                d,
            ),
        );
        assert_close(
            "counter shrinks",
            signed_area(&wide[wide_counter]).abs(),
            grown_in(
                signed_area(&crisp[counter]).abs(),
                perimeter(&crisp[counter]),
                d,
            ),
        );
    }

    #[test]
    fn merges_overlapping_contours_without_a_seam() {
        let outline = GlyphOutline {
            upem: 1000,
            cmds: [square(0.0, 0.0, 600.0), square(400.0, 0.0, 600.0)].concat(),
        };

        let d = f64::from(1000.0_f32 * EMBOLDEN_EM_PER_SIDE);
        let widened = embolden_glyph_outline(&outline).expect("overlap widens");
        let contours = sampled(&widened);

        assert_eq!(
            contours.len(),
            1,
            "the two strokes union into one outer contour"
        );
        let area = signed_area(&contours[0]);
        assert!(area > 0.0, "widened ink winds counter-clockwise");
        // The two 600x600 squares union into one 1000x600 rectangle, so a
        // doubled offset would land far outside this analytic answer.
        assert_close("merged union", area, grown_out(1_000.0 * 600.0, 3_200.0, d));
    }

    #[test]
    fn widens_a_cff_glyph_and_keeps_its_counter() {
        let frame = std::fs::read(concat!(
            env!("CARGO_MANIFEST_DIR"),
            "/../../packages/fonts-cjk/assets/NotoSansSC-Regular.otf.zst"
        ))
        .expect("CJK asset is present");
        let font = betteroffice_fonts::decode_face(
            &frame,
            None,
            betteroffice_fonts::Codec::Zstd,
            64 << 20,
        )
        .expect("asset decodes")
        .font;
        let mut store = crate::FontStore::new();
        let id = store.register(font).expect("registers");
        let glyph = store.glyph_id(id, 'e').expect("known font").expect("cmap");
        let crisp = store.outline_glyph(id, glyph).expect("outline 'e'");
        assert!(
            crisp
                .cmds
                .iter()
                .any(|c| matches!(c, PathCmd::CubicTo { .. })),
            "CFF glyphs draw cubics"
        );

        let widened = embolden_glyph_outline(&crisp).expect("'e' widens");
        let d = f64::from(f32::from(crisp.upem) * EMBOLDEN_EM_PER_SIDE);
        let crisp = sampled(&crisp.cmds);
        let wide = sampled(&widened);
        assert_eq!(wide.len(), 2, "widening keeps the counter");

        let (outer, counter) = widest(&crisp);
        let (wide_outer, wide_counter) = widest(&wide);
        // CFF outers already wind counter-clockwise.
        assert!(signed_area(&crisp[outer]) > 0.0);
        assert!(signed_area(&wide[wide_outer]) > 0.0);
        assert!(signed_area(&wide[wide_counter]) < 0.0);

        assert_close(
            "outer grows",
            signed_area(&wide[wide_outer]),
            grown_out(
                signed_area(&crisp[outer]).abs(),
                perimeter(&crisp[outer]),
                d,
            ),
        );
        assert_close(
            "counter shrinks",
            signed_area(&wide[wide_counter]).abs(),
            grown_in(
                signed_area(&crisp[counter]).abs(),
                perimeter(&crisp[counter]),
                d,
            ),
        );
    }

    #[test]
    fn blank_glyph_has_no_widened_geometry() {
        let mut store = crate::FontStore::new();
        let id = store.register(LIBERATION_SANS.to_vec()).expect("registers");
        let space = store.glyph_id(id, ' ').expect("known font").expect("cmap");
        let outline = store.outline_glyph(id, space).expect("outline space");
        assert_eq!(embolden_glyph_outline(&outline), None);
    }

    fn square(x: f32, y: f32, size: f32) -> Vec<PathCmd> {
        vec![
            PathCmd::MoveTo { x, y },
            PathCmd::LineTo { x: x + size, y },
            PathCmd::LineTo {
                x: x + size,
                y: y + size,
            },
            PathCmd::LineTo { x, y: y + size },
            PathCmd::Close,
        ]
    }

    /// Areas and perimeters measure the sampled outline, not the control
    /// polygon, so offsets can be checked against analytic geometry.
    fn sampled(cmds: &[PathCmd]) -> Vec<Vec<[f32; 2]>> {
        const STEPS: usize = 64;
        let mut contours = Vec::new();
        let mut contour: Vec<[f32; 2]> = Vec::new();
        let mut cursor = [0.0_f32; 2];
        for cmd in cmds {
            match *cmd {
                PathCmd::MoveTo { x, y } => {
                    contour.clear();
                    cursor = [x, y];
                    contour.push(cursor);
                }
                PathCmd::LineTo { x, y } => {
                    cursor = [x, y];
                    contour.push(cursor);
                }
                PathCmd::QuadTo { cx, cy, x, y } => {
                    let control = [cx, cy];
                    let end = [x, y];
                    for step in 1..=STEPS {
                        let t = step as f32 / STEPS as f32;
                        let u = 1.0 - t;
                        contour.push([
                            u * u * cursor[0] + 2.0 * u * t * control[0] + t * t * end[0],
                            u * u * cursor[1] + 2.0 * u * t * control[1] + t * t * end[1],
                        ]);
                    }
                    cursor = end;
                }
                PathCmd::CubicTo {
                    c1x,
                    c1y,
                    c2x,
                    c2y,
                    x,
                    y,
                } => {
                    let c1 = [c1x, c1y];
                    let c2 = [c2x, c2y];
                    let end = [x, y];
                    for step in 1..=STEPS {
                        let t = step as f32 / STEPS as f32;
                        let u = 1.0 - t;
                        contour.push([
                            u * u * u * cursor[0]
                                + 3.0 * u * u * t * c1[0]
                                + 3.0 * u * t * t * c2[0]
                                + t * t * t * end[0],
                            u * u * u * cursor[1]
                                + 3.0 * u * u * t * c1[1]
                                + 3.0 * u * t * t * c2[1]
                                + t * t * t * end[1],
                        ]);
                    }
                    cursor = end;
                }
                PathCmd::Close => {
                    if contour.len() >= 3 {
                        contours.push(std::mem::take(&mut contour));
                    }
                    contour.clear();
                }
            }
        }
        contours
    }

    fn perimeter(contour: &[[f32; 2]]) -> f64 {
        (0..contour.len())
            .map(|index| {
                let a = contour[index];
                let b = contour[(index + 1) % contour.len()];
                (f64::from(b[0] - a[0]).powi(2) + f64::from(b[1] - a[1]).powi(2)).sqrt()
            })
            .sum()
    }

    fn grown_out(area: f64, perimeter: f64, d: f64) -> f64 {
        area + perimeter * d + std::f64::consts::PI * d * d
    }

    fn grown_in(area: f64, perimeter: f64, d: f64) -> f64 {
        area - perimeter * d + std::f64::consts::PI * d * d
    }

    /// Indices of the outer contour (largest area) and the next contour.
    fn widest(contours: &[Vec<[f32; 2]>]) -> (usize, usize) {
        let mut order: Vec<usize> = (0..contours.len()).collect();
        order.sort_by(|&a, &b| {
            signed_area(&contours[b])
                .abs()
                .total_cmp(&signed_area(&contours[a]).abs())
        });
        (order[0], order[1])
    }

    fn assert_close(label: &str, got: f64, expected: f64) {
        let error = (got - expected).abs() / expected;
        assert!(
            error < 0.01,
            "{label}: {got:.0} vs expected {expected:.0} ({:.2}% off)",
            error * 100.0
        );
    }
}
