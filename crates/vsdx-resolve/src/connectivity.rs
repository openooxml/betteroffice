//! Resolved glue records and connection-point positions.

use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use vsdx_formula::{Limits as FormulaLimits, evaluate_number};
use vsdx_parse::{Connect, Shape};

use crate::{Lookup, ResolvedShape, Resolver};

/// A scene-space point in Visio inches (with Visio's Y-up convention).
#[derive(Clone, Copy, Debug, PartialEq, Serialize, Deserialize)]
pub struct ScenePoint {
    pub x: f64,
    pub y: f64,
}

/// An affine transform in Visio scene coordinates.
#[derive(Clone, Copy, Debug, PartialEq, Serialize, Deserialize)]
pub struct SceneAffine {
    pub a: f64,
    pub b: f64,
    pub c: f64,
    pub d: f64,
    pub e: f64,
    pub f: f64,
}

impl SceneAffine {
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
    pub fn compose(self, other: Self) -> Self {
        Self {
            a: self.a * other.a + self.c * other.b,
            b: self.b * other.a + self.d * other.b,
            c: self.a * other.c + self.c * other.d,
            d: self.b * other.c + self.d * other.d,
            e: self.a * other.e + self.c * other.f + self.e,
            f: self.b * other.e + self.d * other.f + self.f,
        }
    }
    pub fn apply_point(self, x: f64, y: f64) -> ScenePoint {
        ScenePoint {
            x: self.a * x + self.c * y + self.e,
            y: self.b * x + self.d * y + self.f,
        }
    }
}

#[derive(Clone, Copy, Debug)]
pub struct ShapeBounds {
    pub x: f64,
    pub y: f64,
    pub width: f64,
    pub height: f64,
    pub loc_pin_x: f64,
    pub loc_pin_y: f64,
    pub angle: f64,
    pub flip_x: bool,
    pub flip_y: bool,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct SceneTransform {
    pub local: SceneAffine,
    pub scene: SceneAffine,
}

pub fn bounds_affine(
    bounds: ShapeBounds,
    child_extent: Option<(f64, f64, f64, f64)>,
) -> SceneAffine {
    let (origin_x, origin_y, scale_x, scale_y) = child_extent
        .map(|(x, y, width, height)| (x, y, bounds.width / width, bounds.height / height))
        .unwrap_or((0.0, 0.0, 1.0, 1.0));
    let (sin, cos) = bounds.angle.sin_cos();
    let sx = scale_x * if bounds.flip_x { -1.0 } else { 1.0 };
    let sy = scale_y * if bounds.flip_y { -1.0 } else { 1.0 };
    let pin_x = bounds.x + bounds.loc_pin_x;
    let pin_y = bounds.y + bounds.loc_pin_y;
    SceneAffine {
        a: cos * sx,
        b: sin * sx,
        c: -sin * sy,
        d: cos * sy,
        e: pin_x - cos * sx * (bounds.loc_pin_x + origin_x)
            + sin * sy * (bounds.loc_pin_y + origin_y),
        f: pin_y
            - sin * sx * (bounds.loc_pin_x + origin_x)
            - cos * sy * (bounds.loc_pin_y + origin_y),
    }
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ConnectionPoint {
    pub row: u32,
    pub position: ScenePoint,
    pub x_provenance: NumericProvenance,
    pub y_provenance: NumericProvenance,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum NumericProvenance {
    Formula,
    CachedValue,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum ConnectivityDiagnostic {
    MissingFromShape { shape_id: u32 },
    MissingToShape { shape_id: u32 },
    MissingConnectionPoint { shape_id: u32, row: u32 },
    UnsupportedFromCell { shape_id: u32, cell: String },
    UnsupportedToCell { shape_id: u32, cell: String },
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct GluedEnd {
    pub shape_id: u32,
    pub cell: Option<String>,
    pub part: Option<i32>,
    pub connection_point: Option<ConnectionPoint>,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ResolvedGlue {
    pub connector_id: u32,
    pub endpoint: ConnectorEndpoint,
    pub from_part: Option<i32>,
    pub to: Option<GluedEnd>,
    pub diagnostics: Vec<ConnectivityDiagnostic>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum ConnectorEndpoint {
    Begin,
    End,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ResolvedConnector {
    pub shape_id: u32,
    pub is_1d: bool,
    pub begin: Option<ScenePoint>,
    pub end: Option<ScenePoint>,
    pub glue: Vec<ResolvedGlue>,
}

#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
pub struct PageConnectivity {
    pub connectors: BTreeMap<u32, ResolvedConnector>,
    pub diagnostics: Vec<ConnectivityDiagnostic>,
}

impl<'a> Resolver<'a> {
    /// Resolves page glue. A shape is 1D when its effective `OneD` ShapeSheet cell is nonzero,
    /// as defined by MS-VSDX's Shape element/OneD cell semantics.
    pub fn resolve_page_connectivity(
        &self,
        page_part: &str,
    ) -> Result<PageConnectivity, crate::ResolveError> {
        let page = self
            .package()
            .page_contents
            .get(page_part)
            .ok_or_else(|| crate::ResolveError::MissingPage(page_part.into()))?;
        let shapes = self.resolve_page_shapes(page_part)?;
        let mut out = PageConnectivity::default();
        for (id, shape) in &shapes {
            if is_one_d(shape) {
                out.connectors.insert(
                    *id,
                    ResolvedConnector {
                        shape_id: *id,
                        is_1d: true,
                        begin: endpoint(shape, "BeginX", "BeginY"),
                        end: endpoint(shape, "EndX", "EndY"),
                        glue: Vec::new(),
                    },
                );
            }
        }
        let transforms = scene_transforms(page, &shapes, |_, shape, name| number(shape, name));
        for connect in page.connects() {
            self.add_connectivity_record(connect, &shapes, &transforms, &mut out);
        }
        Ok(out)
    }

    fn add_connectivity_record(
        &self,
        connect: &Connect,
        shapes: &BTreeMap<u32, ResolvedShape>,
        transforms: &BTreeMap<u32, SceneTransform>,
        out: &mut PageConnectivity,
    ) {
        let mut diagnostics = Vec::new();
        let Some(source) = shapes.get(&connect.from_sheet) else {
            diagnostics.push(ConnectivityDiagnostic::MissingFromShape {
                shape_id: connect.from_sheet,
            });
            out.diagnostics.extend(diagnostics);
            return;
        };
        let Some(glue_endpoint) = connect.from_cell.as_deref().and_then(endpoint_name) else {
            diagnostics.push(ConnectivityDiagnostic::UnsupportedFromCell {
                shape_id: connect.from_sheet,
                cell: connect.from_cell.clone().unwrap_or_default(),
            });
            out.diagnostics.extend(diagnostics);
            return;
        };
        if !is_one_d(source) {
            diagnostics.push(ConnectivityDiagnostic::UnsupportedFromCell {
                shape_id: connect.from_sheet,
                cell: connect.from_cell.clone().unwrap_or_default(),
            });
        }
        out.connectors
            .entry(connect.from_sheet)
            .or_insert_with(|| ResolvedConnector {
                shape_id: connect.from_sheet,
                is_1d: is_one_d(source),
                begin: endpoint(source, "BeginX", "BeginY"),
                end: endpoint(source, "EndX", "EndY"),
                glue: Vec::new(),
            });
        let to = match shapes.get(&connect.to_sheet) {
            None => {
                diagnostics.push(ConnectivityDiagnostic::MissingToShape {
                    shape_id: connect.to_sheet,
                });
                None
            }
            Some(target) => {
                let point = match connect.to_cell.as_deref().and_then(connection_row) {
                    Some(row) => {
                        let point = connection_point(
                            target,
                            row,
                            transforms
                                .get(&connect.to_sheet)
                                .map(|transform| transform.scene),
                        );
                        point.or_else(|| {
                            diagnostics.push(ConnectivityDiagnostic::MissingConnectionPoint {
                                shape_id: connect.to_sheet,
                                row,
                            });
                            None
                        })
                    }
                    None if matches!(connect.to_cell.as_deref(), Some("PinX") | Some("PinY")) => {
                        shape_pin(
                            target,
                            transforms
                                .get(&connect.to_sheet)
                                .map(|transform| transform.scene),
                        )
                        .map(|(position, x_provenance, y_provenance)| {
                            ConnectionPoint {
                                row: 0,
                                position,
                                x_provenance,
                                y_provenance,
                            }
                        })
                    }
                    None => {
                        diagnostics.push(ConnectivityDiagnostic::UnsupportedToCell {
                            shape_id: connect.to_sheet,
                            cell: connect.to_cell.clone().unwrap_or_default(),
                        });
                        None
                    }
                };
                Some(GluedEnd {
                    shape_id: connect.to_sheet,
                    cell: connect.to_cell.clone(),
                    part: connect.to_part,
                    connection_point: point,
                })
            }
        };
        let glue = ResolvedGlue {
            connector_id: connect.from_sheet,
            endpoint: glue_endpoint,
            from_part: connect.from_part,
            to,
            diagnostics: diagnostics.clone(),
        };
        if let Some(connector) = out.connectors.get_mut(&connect.from_sheet) {
            connector.glue.push(glue);
        }
        out.diagnostics.extend(diagnostics);
    }
}

fn is_one_d(shape: &ResolvedShape) -> bool {
    number(shape, "OneD").is_some_and(|value| value != 0.0)
}
fn endpoint(shape: &ResolvedShape, x: &str, y: &str) -> Option<ScenePoint> {
    Some(ScenePoint {
        x: number(shape, x)?,
        y: number(shape, y)?,
    })
}
fn number(shape: &ResolvedShape, name: &str) -> Option<f64> {
    number_with_provenance(shape, name).map(|(value, _)| value)
}

fn number_with_provenance(shape: &ResolvedShape, name: &str) -> Option<(f64, NumericProvenance)> {
    let Lookup::Found(value) = shape.cell(name)? else {
        return None;
    };
    number_from_cell(shape, &value.cell)
}

fn number_from_cell(
    shape: &ResolvedShape,
    cell: &vsdx_parse::Cell,
) -> Option<(f64, NumericProvenance)> {
    if let Some(number) = cell
        .formula
        .as_deref()
        .and_then(|formula| formula_number(shape, formula))
    {
        return Some((number, NumericProvenance::Formula));
    }
    cell.value
        .as_deref()?
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
        .map(|value| (value, NumericProvenance::CachedValue))
}

/// Evaluates only finite numeric literals, one binary `+`, `-`, `*`, or `/` expression (with
/// recursive operands), and local cell-name references. It intentionally does not implement
/// ShapeSheet functions, units, parentheses, cross-sheet references, or general ShapeSheet
/// syntax because `vsdx-eval` depends on this crate. If a formula is unsupported, callers fall
/// back to a finite cached `V`; a supported formula always wins over a conflicting cache.
fn formula_number(shape: &ResolvedShape, formula: &str) -> Option<f64> {
    evaluate_number(
        formula,
        FormulaLimits {
            max_depth: 64,
            max_nodes: 1_024,
            max_tokens: 1_024,
        },
        &mut |name| {
            let Lookup::Found(value) = shape.cell(name.trim())? else {
                return None;
            };
            value
                .cell
                .formula
                .clone()
                .or_else(|| value.cell.value.clone())
        },
    )
}

fn endpoint_name(name: &str) -> Option<ConnectorEndpoint> {
    match name {
        "BeginX" | "BeginY" => Some(ConnectorEndpoint::Begin),
        "EndX" | "EndY" => Some(ConnectorEndpoint::End),
        _ => None,
    }
}
fn connection_row(name: &str) -> Option<u32> {
    name.strip_prefix("Connections.X")?.parse().ok()
}
fn connection_point(
    shape: &ResolvedShape,
    row: u32,
    transform: Option<SceneAffine>,
) -> Option<ConnectionPoint> {
    let section = shape.sections.get("Connection")?;
    let resolved_row = section.rows.get(&format!("IX:{row}"))?;
    if resolved_row.deleted {
        return None;
    }
    let value = |name: &str| match resolved_row.cells.get(name)? {
        Lookup::Found(cell) => number_from_cell(shape, &cell.cell),
        _ => None,
    };
    let (x, x_provenance) = value("X")?;
    let (y, y_provenance) = value("Y")?;
    Some(ConnectionPoint {
        row,
        position: transform?.apply_point(x, y),
        x_provenance,
        y_provenance,
    })
}

pub fn scene_transforms(
    page: &vsdx_parse::Sheet,
    shapes: &BTreeMap<u32, ResolvedShape>,
    value: impl Fn(u32, &ResolvedShape, &str) -> Option<f64> + Copy,
) -> BTreeMap<u32, SceneTransform> {
    let mut transforms = BTreeMap::new();
    for shape in page.shapes() {
        add_scene_transforms(
            shape,
            SceneAffine::identity(),
            shapes,
            value,
            &mut transforms,
        );
    }
    transforms
}

fn add_scene_transforms(
    shape: &Shape,
    parent: SceneAffine,
    shapes: &BTreeMap<u32, ResolvedShape>,
    value: impl Fn(u32, &ResolvedShape, &str) -> Option<f64> + Copy,
    transforms: &mut BTreeMap<u32, SceneTransform>,
) {
    let Some(resolved) = shapes.get(&shape.id) else {
        return;
    };
    let Some(bounds) = shape_bounds(shape.id, resolved, value) else {
        return;
    };
    let children = shape.shapes().collect::<Vec<_>>();
    let extent = child_extent(&children, shapes, value);
    let local = bounds_affine(bounds, extent);
    let scene = parent.compose(local);
    transforms.insert(shape.id, SceneTransform { local, scene });
    for child in children {
        add_scene_transforms(child, scene, shapes, value, transforms);
    }
}

fn child_extent(
    children: &[&Shape],
    shapes: &BTreeMap<u32, ResolvedShape>,
    value: impl Fn(u32, &ResolvedShape, &str) -> Option<f64> + Copy,
) -> Option<(f64, f64, f64, f64)> {
    let bounds = children
        .iter()
        .filter_map(|source| {
            shapes
                .get(&source.id)
                .and_then(|shape| shape_bounds(source.id, shape, value))
        })
        .collect::<Vec<_>>();
    let min_x = bounds.iter().map(|bounds| bounds.x).reduce(f64::min)?;
    let min_y = bounds.iter().map(|bounds| bounds.y).reduce(f64::min)?;
    let max_x = bounds
        .iter()
        .map(|bounds| bounds.x + bounds.width)
        .reduce(f64::max)?;
    let max_y = bounds
        .iter()
        .map(|bounds| bounds.y + bounds.height)
        .reduce(f64::max)?;
    let width = max_x - min_x;
    let height = max_y - min_y;
    (width > 0.0 && height > 0.0).then_some((min_x, min_y, width, height))
}

fn shape_bounds(
    id: u32,
    shape: &ResolvedShape,
    value: impl Fn(u32, &ResolvedShape, &str) -> Option<f64>,
) -> Option<ShapeBounds> {
    let width = value(id, shape, "Width")?;
    let height = value(id, shape, "Height")?;
    let loc_pin_x = value(id, shape, "LocPinX").unwrap_or(width / 2.0);
    let loc_pin_y = value(id, shape, "LocPinY").unwrap_or(height / 2.0);
    Some(ShapeBounds {
        x: value(id, shape, "PinX")? - loc_pin_x,
        y: value(id, shape, "PinY")? - loc_pin_y,
        width,
        height,
        loc_pin_x,
        loc_pin_y,
        angle: value(id, shape, "Angle").unwrap_or(0.0),
        flip_x: value(id, shape, "FlipX").unwrap_or(0.0) != 0.0,
        flip_y: value(id, shape, "FlipY").unwrap_or(0.0) != 0.0,
    })
}

fn shape_pin(
    shape: &ResolvedShape,
    transform: Option<SceneAffine>,
) -> Option<(ScenePoint, NumericProvenance, NumericProvenance)> {
    let (x, x_provenance) = number_with_provenance(shape, "LocPinX")?;
    let (y, y_provenance) = number_with_provenance(shape, "LocPinY")?;
    Some((transform?.apply_point(x, y), x_provenance, y_provenance))
}
