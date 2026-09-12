use crate::display_list::{Paint, Stroke};
use vsdx_eval::{Evaluation, PageShapeReferences, Value, evaluate_cell_with_shape_package_theme};
use vsdx_parse::{ParseLimits, VsdxPackage};
use vsdx_resolve::{Lookup, ResolvedShape};

pub fn paint(
    package: &VsdxPackage,
    references: Option<&PageShapeReferences>,
    shape: &ResolvedShape,
    shape_id: u32,
) -> Result<(Option<Paint>, Option<Stroke>), String> {
    let fill = value(shape, "FillPattern")
        .filter(|v| *v != "0")
        .map(|_| colour(package, references, shape, shape_id, "FillForegnd"))
        .transpose()?
        .map(|color| Paint::Solid { color });
    let stroke = value(shape, "LinePattern")
        .filter(|v| *v != "0")
        .map(|_| colour(package, references, shape, shape_id, "LineColor"))
        .transpose()?
        .map(|color| {
            let width = number(shape, "LineWeight").unwrap_or(0.01) as f32;
            if !width.is_finite() {
                return Err::<Stroke, String>("non-finite stroke width".into());
            }
            Ok(Stroke {
                color,
                width,
                dashed: value(shape, "LinePattern").is_some_and(|v| v != "1"),
            })
        })
        .transpose()?;
    Ok((fill, stroke))
}
pub fn number(shape: &ResolvedShape, name: &str) -> Option<f64> {
    value(shape, name)?
        .parse()
        .ok()
        .filter(|n: &f64| n.is_finite())
}
fn value<'a>(shape: &'a ResolvedShape, name: &str) -> Option<&'a str> {
    match shape.cell(name)? {
        Lookup::Found(cell) => cell.cell.value.as_deref(),
        Lookup::Deleted | Lookup::Absent => None,
    }
}
fn colour(
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
            _ => Err(format!("colour cell {name} evaluated to a number")),
        },
        Evaluation::Unsupported(reason) => Err(reason),
        Evaluation::Error(error) => Err(error.message),
    }
}
