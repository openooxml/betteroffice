use crate::display_list::{Diagnostic, GradientStop, Paint, Stroke};
use vsdx_eval::{Evaluation, PageShapeReferences, Value, evaluate_cell_with_shape_package_theme};
use vsdx_parse::{ParseLimits, VsdxPackage};
use vsdx_resolve::{Lookup, ResolvedShape};

/// Paint for the channels an emitting section uses; a channel that fails keeps the Visio default.
pub struct PaintOutcome {
    pub fill: Option<Paint>,
    pub stroke: Option<Stroke>,
    pub diagnostics: Vec<Diagnostic>,
}

const DEFAULT_STROKE_WIDTH: f64 = 0.01;

pub fn paint(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    needs_fill: bool,
    needs_stroke: bool,
) -> PaintOutcome {
    let mut diagnostics = Vec::new();
    let fill = if needs_fill && value(shape, "FillPattern").is_some_and(|v| v != "0") {
        match gradient(package, references, shape, shape_id) {
            GradientOutcome::Painted(paint, lossy) => {
                if let Some(reason) = lossy {
                    diagnostics.push(Diagnostic::for_code(
                        "lossy-fill-gradient",
                        format!("lossy fill gradient: {reason}"),
                    ));
                }
                Some(paint)
            }
            GradientOutcome::Unsupported => {
                solid(package, references, shape, shape_id, &mut diagnostics)
            }
            GradientOutcome::Broken(reason) => {
                diagnostics.push(Diagnostic::for_code(
                    "unresolvable-fill-gradient",
                    format!("unresolvable fill gradient: {reason}"),
                ));
                solid(package, references, shape, shape_id, &mut diagnostics)
            }
        }
    } else {
        None
    };
    let stroke = if needs_stroke && value(shape, "LinePattern").is_some_and(|v| v != "0") {
        let width = number(shape, "LineWeight")
            .filter(|candidate| *candidate > 0.0)
            .unwrap_or(DEFAULT_STROKE_WIDTH) as f32;
        let width = if width.is_finite() {
            width
        } else {
            diagnostics.push(Diagnostic::for_code(
                "unresolvable-stroke-width",
                "unresolvable stroke width: non-finite LineWeight",
            ));
            DEFAULT_STROKE_WIDTH as f32
        };
        let color = match colour(package, references, shape, shape_id, "LineColor") {
            Ok(color) => color,
            Err(reason) => {
                diagnostics.push(Diagnostic::for_code(
                    "unresolvable-stroke-colour",
                    format!("unresolvable stroke colour: {reason}"),
                ));
                default_colour(package)
            }
        };
        Some(Stroke {
            color,
            width,
            dashed: value(shape, "LinePattern").is_some_and(|v| v != "1"),
        })
    } else {
        None
    };
    PaintOutcome {
        fill,
        stroke,
        diagnostics,
    }
}
pub fn number(shape: &ResolvedShape, name: &str) -> Option<f64> {
    value(shape, name)?
        .parse()
        .ok()
        .filter(|n: &f64| n.is_finite())
}
fn solid(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    diagnostics: &mut Vec<Diagnostic>,
) -> Option<Paint> {
    let color = match colour(package, references, shape, shape_id, "FillForegnd") {
        Ok(color) => color,
        Err(reason) => {
            diagnostics.push(Diagnostic::for_code(
                "unresolvable-fill-colour",
                format!("unresolvable fill colour: {reason}"),
            ));
            default_colour(package)
        }
    };
    Some(Paint::Solid { color })
}
enum GradientOutcome {
    /// The gradient, plus the reason it is a lossy rendering of the ShapeSheet.
    Painted(Paint, Option<String>),
    Unsupported,
    Broken(String),
}
/// `FillGradientDir` 0 is Visio's linear gradient; every other direction is radial, rectangular or
/// path, which the display list cannot express, so it degrades to the solid fill.
const LINEAR_GRADIENT_DIR: f64 = 0.0;
fn gradient(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
) -> GradientOutcome {
    if !resolved_number(package, references, shape, shape_id, "FillGradientEnabled")
        .is_some_and(|enabled| enabled != 0.0)
    {
        return GradientOutcome::Unsupported;
    }
    if let Some(dir) = resolved_number(package, references, shape, shape_id, "FillGradientDir")
        && dir != LINEAR_GRADIENT_DIR
    {
        return GradientOutcome::Broken(format!("unsupported FillGradientDir {dir}"));
    }
    let Some(radians) = resolved_number(package, references, shape, shape_id, "FillGradientAngle")
    else {
        return GradientOutcome::Unsupported;
    };
    if !radians.is_finite() {
        return GradientOutcome::Broken("non-finite FillGradientAngle".into());
    }
    let Some(section) = shape.sections.get("FillGradient") else {
        return GradientOutcome::Broken("missing FillGradient section".into());
    };
    let mut lossy = None;
    let mut stops = Vec::new();
    for row in section
        .row_order
        .iter()
        .filter_map(|key| section.rows.get(key))
        .filter(|row| !row.deleted)
    {
        match stop(package, references, shape, shape_id, row) {
            StopOutcome::Resolved(resolved) => stops.push(resolved),
            StopOutcome::Opaqued(resolved, reason) => {
                if lossy.is_none() {
                    lossy = Some(reason);
                }
                stops.push(resolved);
            }
            StopOutcome::Broken(reason) => return GradientOutcome::Broken(reason),
        }
    }
    stops.sort_by(|left, right| left.position.total_cmp(&right.position));
    if stops.len() < 2 {
        return GradientOutcome::Broken("fewer than two resolvable gradient stops".into());
    }
    GradientOutcome::Painted(
        Paint::Gradient {
            angle_deg: Some((radians * 180.0 / std::f64::consts::PI) as f32),
            stops,
        },
        lossy,
    )
}
enum StopOutcome {
    Resolved(GradientStop),
    /// Visio paints this stop with transparency; the display list carries opaque colours only.
    Opaqued(GradientStop, String),
    Broken(String),
}
fn stop(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    row: &vsdx_resolve::ResolvedRow,
) -> StopOutcome {
    let Some(color) = row_cell(row, "GradientStopColor")
        .and_then(|cell| stop_colour(package, references, shape, shape_id, cell))
    else {
        return StopOutcome::Broken("unresolvable GradientStopColor".into());
    };
    let position = row_cell(row, "GradientStopPosition").and_then(|cell| {
        stop_number(
            package,
            references,
            shape,
            shape_id,
            "GradientStopPosition",
            cell,
        )
    });
    let Some(position) = position.filter(|position| position.is_finite()) else {
        return StopOutcome::Broken("unresolvable GradientStopPosition".into());
    };
    let resolved = GradientStop {
        position: position.clamp(0.0, 1.0) as f32,
        color,
    };
    let Some(cell) = row_cell(row, "GradientStopColorTrans")
        .filter(|cell| cell.formula.is_some() || cell.value.is_some())
    else {
        return StopOutcome::Resolved(resolved);
    };
    match stop_number(
        package,
        references,
        shape,
        shape_id,
        "GradientStopColorTrans",
        cell,
    ) {
        Some(0.0) => StopOutcome::Resolved(resolved),
        Some(transparency) => StopOutcome::Opaqued(
            resolved,
            format!("GradientStopColorTrans {transparency} painted opaque"),
        ),
        None => StopOutcome::Broken("unresolvable GradientStopColorTrans".into()),
    }
}
fn row_cell<'a>(row: &'a vsdx_resolve::ResolvedRow, name: &str) -> Option<&'a vsdx_parse::Cell> {
    match row.cells.get(name)? {
        Lookup::Found(cell) => Some(&cell.cell),
        Lookup::Deleted | Lookup::Absent => None,
    }
}
fn stop_colour(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    cell: &vsdx_parse::Cell,
) -> Option<String> {
    if let Some(formula) = cell.formula.as_deref().or(cell.value.as_deref())
        && let Some(refs) = references
    {
        match evaluate_cell_with_shape_package_theme(
            "GradientStopColor",
            formula,
            &refs.for_shape(shape_id),
            &ParseLimits::default(),
            shape,
            package,
        ) {
            Evaluation::Evaluated(result) => match result.value {
                Value::Color(color) => {
                    return Some(format!(
                        "#{:02X}{:02X}{:02X}",
                        color.red, color.green, color.blue
                    ));
                }
                Value::Number(value) => {
                    if let Some(color) = crate::palette_colour(package, value.number) {
                        return Some(color);
                    }
                }
            },
            Evaluation::Unsupported(_) | Evaluation::Error(_) => {}
        }
    }
    cached_colour(package, cell.value.as_deref()?)
}
fn cached_colour(package: &VsdxPackage, value: &str) -> Option<String> {
    let hex = value.strip_prefix('#').unwrap_or(value);
    if hex.len() == 6 && hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Some(format!("#{}", hex.to_ascii_uppercase()));
    }
    let index: f64 = value.parse().ok()?;
    crate::palette_colour(package, index)
}
fn stop_number(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    name: &str,
    cell: &vsdx_parse::Cell,
) -> Option<f64> {
    if let Some(formula) = cell.formula.as_deref()
        && let Some(refs) = references
    {
        match evaluate_cell_with_shape_package_theme(
            name,
            formula,
            &refs.for_shape(shape_id),
            &ParseLimits::default(),
            shape,
            package,
        ) {
            Evaluation::Evaluated(result) => match result.value {
                Value::Number(value) => return Some(value.number),
                Value::Color(_) => return None,
            },
            Evaluation::Unsupported(_) | Evaluation::Error(_) => {}
        }
    }
    cell.value.as_deref()?.parse().ok()
}
pub(crate) fn resolved_number(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    name: &str,
) -> Option<f64> {
    if let Some(cached) = value(shape, name)
        && let Ok(number) = cached.parse::<f64>()
        && number.is_finite()
    {
        return Some(number);
    }
    let cell = match shape.cell(name)? {
        Lookup::Found(cell) => &cell.cell,
        Lookup::Deleted | Lookup::Absent => return None,
    };
    let formula = cell.formula.as_deref().or(cell.value.as_deref())?;
    let refs = references?;
    match evaluate_cell_with_shape_package_theme(
        name,
        formula,
        &refs.for_shape(shape_id),
        &ParseLimits::default(),
        shape,
        package,
    ) {
        Evaluation::Evaluated(result) => match result.value {
            Value::Number(value) => Some(value.number),
            Value::Color(_) => None,
        },
        Evaluation::Unsupported(_) | Evaluation::Error(_) => None,
    }
}
fn value<'a>(shape: &'a ResolvedShape, name: &str) -> Option<&'a str> {
    match shape.cell(name)? {
        Lookup::Found(cell) => cell.cell.value.as_deref(),
        Lookup::Deleted | Lookup::Absent => None,
    }
}
/// Visio's unformatted foreground: palette index 0, black when the file has no palette.
fn default_colour(package: &VsdxPackage) -> String {
    crate::palette_colour(package, 0.0).unwrap_or_else(|| "#000000".into())
}
pub(crate) fn colour(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
    name: &str,
) -> Result<String, String> {
    let cell = match shape.cell(name) {
        Some(Lookup::Found(cell)) => &cell.cell,
        _ => return Err(format!("missing colour cell {name}")),
    };
    let formula = cell
        .formula
        .as_deref()
        .or(cell.value.as_deref())
        .ok_or_else(|| format!("missing colour value {name}"))?;
    let refs = references.ok_or_else(|| format!("unavailable colour references for {name}"))?;
    match evaluate_cell_with_shape_package_theme(
        name,
        formula,
        &refs.for_shape(shape_id),
        &ParseLimits::default(),
        shape,
        package,
    ) {
        Evaluation::Evaluated(result) => match result.value {
            Value::Color(color) => Ok(format!(
                "#{:02X}{:02X}{:02X}",
                color.red, color.green, color.blue
            )),
            Value::Number(value) => crate::palette_colour(package, value.number)
                .ok_or_else(|| format!("colour cell {name} has an unknown palette index")),
        },
        Evaluation::Unsupported(reason) => Err(reason),
        Evaluation::Error(error) => Err(error.message),
    }
}
