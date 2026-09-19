use vsdx_parse::VsdxPackage;
use vsdx_resolve::{Lookup, Resolver};

use crate::ExportError;

pub struct ShapeDatum {
    pub page: String,
    pub shape_id: u32,
    pub shape: String,
    pub label: String,
    pub value: String,
}

pub fn shape_data(package: &VsdxPackage) -> Result<Vec<ShapeDatum>, ExportError> {
    let resolver = Resolver::new(package);
    let mut data = Vec::new();
    for part in &package.page_part_paths {
        let name = page_name_for(package, part);
        let shapes = resolver
            .resolve_page_shapes(part)
            .map_err(|error| ExportError::Resolve(error.to_string()))?;
        let mut ids: Vec<u32> = shapes.keys().copied().collect();
        ids.sort_unstable();
        for id in ids {
            let resolved = &shapes[&id];
            let Some(section) = resolved.sections.get("Property") else {
                continue;
            };
            for key in &section.row_order {
                let Some(row) = section.rows.get(key) else {
                    continue;
                };
                if row.deleted {
                    continue;
                }
                let fallback = key.strip_prefix("N:").unwrap_or(key);
                let label = cell_text(row, "Label");
                let value = cell_text(row, "Value");
                let label = label.or_else(|| {
                    (!fallback.is_empty() && fallback != "row").then(|| fallback.to_owned())
                });
                match (label, value) {
                    (Some(label), Some(value)) if !(label.is_empty() && value.is_empty()) => {
                        data.push(ShapeDatum {
                            page: name.clone(),
                            shape_id: id,
                            shape: shape_name(package, part, id),
                            label,
                            value,
                        });
                    }
                    _ => {}
                }
            }
        }
    }
    Ok(data)
}

fn cell_text(row: &vsdx_resolve::ResolvedRow, name: &str) -> Option<String> {
    match row.cells.get(name)? {
        Lookup::Found(cell) => cell
            .cell
            .value
            .clone()
            .or_else(|| cell.cell.formula.clone()),
        Lookup::Deleted | Lookup::Absent => None,
    }
}

pub(crate) fn page_name_for(package: &VsdxPackage, part: &str) -> String {
    package
        .page_part_ids
        .get(part)
        .and_then(|id| package.page_names.get(id))
        .cloned()
        .unwrap_or_else(|| part.to_owned())
}

fn shape_name(package: &VsdxPackage, part: &str, id: u32) -> String {
    package
        .page_contents
        .get(part)
        .and_then(|sheet| find_shape(sheet.shapes(), id))
        .and_then(|shape| shape.name.clone().or_else(|| shape.name_u.clone()))
        .unwrap_or_else(|| format!("Shape.{id}"))
}

fn find_shape<'a>(
    shapes: impl Iterator<Item = &'a vsdx_parse::Shape>,
    id: u32,
) -> Option<&'a vsdx_parse::Shape> {
    for shape in shapes {
        if shape.id == id {
            return Some(shape);
        }
        if let Some(found) = find_shape(shape.shapes(), id) {
            return Some(found);
        }
    }
    None
}
