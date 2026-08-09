//! Resolved glue records and connection-point positions.

use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};
use vsdx_parse::Connect;

use crate::{Lookup, ResolvedShape, Resolver};

/// A scene-space point in Visio inches (with Visio's Y-up convention).
#[derive(Clone, Copy, Debug, PartialEq, Serialize, Deserialize)]
pub struct ScenePoint {
    pub x: f64,
    pub y: f64,
}

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
pub struct ConnectionPoint {
    pub row: u32,
    pub position: ScenePoint,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
pub enum ConnectivityDiagnostic {
    MissingFromShape { shape_id: u32 },
    MissingToShape { shape_id: u32 },
    MissingConnectionPoint { shape_id: u32, row: u32 },
    UnsupportedFromCell { shape_id: u32, cell: String },
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
        for connect in page.connects() {
            self.add_connectivity_record(connect, &shapes, &mut out);
        }
        Ok(out)
    }

    fn add_connectivity_record(
        &self,
        connect: &Connect,
        shapes: &BTreeMap<u32, ResolvedShape>,
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
        let Some(endpoint) = connect.from_cell.as_deref().and_then(endpoint_name) else {
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
        let to = match shapes.get(&connect.to_sheet) {
            None => {
                diagnostics.push(ConnectivityDiagnostic::MissingToShape {
                    shape_id: connect.to_sheet,
                });
                None
            }
            Some(target) => {
                let point = connect
                    .to_cell
                    .as_deref()
                    .and_then(connection_row)
                    .and_then(|row| {
                        connection_point(target, row).or_else(|| {
                            diagnostics.push(ConnectivityDiagnostic::MissingConnectionPoint {
                                shape_id: connect.to_sheet,
                                row,
                            });
                            None
                        })
                    });
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
            endpoint,
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
    let Lookup::Found(value) = shape.cell(name)? else {
        return None;
    };
    value
        .cell
        .value
        .as_deref()?
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())
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
fn connection_point(shape: &ResolvedShape, row: u32) -> Option<ConnectionPoint> {
    let section = shape.sections.get("Connection")?;
    let resolved_row = section.rows.get(&format!("IX:{row}"))?;
    let value = |name: &str| match resolved_row.cells.get(name)? {
        Lookup::Found(cell) => cell.cell.value.as_deref()?.parse::<f64>().ok(),
        _ => None,
    };
    let x = value("X")?;
    let y = value("Y")?;
    let width = number(shape, "Width")?;
    let height = number(shape, "Height")?;
    let pin_x = number(shape, "PinX")?;
    let pin_y = number(shape, "PinY")?;
    let loc_x = number(shape, "LocPinX").unwrap_or(width / 2.0);
    let loc_y = number(shape, "LocPinY").unwrap_or(height / 2.0);
    let angle = number(shape, "Angle").unwrap_or(0.0);
    let flip_x = number(shape, "FlipX").unwrap_or(0.0) != 0.0;
    let flip_y = number(shape, "FlipY").unwrap_or(0.0) != 0.0;
    let (sin, cos) = angle.sin_cos();
    let x = (x - loc_x) * if flip_x { -1.0 } else { 1.0 };
    let y = (y - loc_y) * if flip_y { -1.0 } else { 1.0 };
    Some(ConnectionPoint {
        row,
        position: ScenePoint {
            x: pin_x + cos * x - sin * y,
            y: pin_y + sin * x + cos * y,
        },
    })
}
