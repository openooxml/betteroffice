use std::fs;

use vsdx_parse::{
    Cell, Connect, ConnectsChild, Row, RowChild, Section, SectionChild, Shape, ShapeChild, Sheet,
    SheetChild, TextToken, VsdxPackage, parse_vsdx,
};

use crate::{
    ConnectivityDiagnostic, ConnectorEndpoint, Lookup, Provenance, ResolveError, ResolvedTextToken,
    Resolver,
};

fn package() -> VsdxPackage {
    serde_json::from_value(serde_json::json!({
        "documentPartPath": "", "pagesPartPath": null, "mastersPartPath": null,
        "pagePartPaths": [], "masterPartPaths": [], "themePartPaths": [], "windowsPartPath": null,
        "relationships": {}, "documentSheet": null, "styleSheets": [], "colors": [], "faceNames": [],
        "pageSheets": {}, "masterSheets": {}, "pagePartIds": {}, "masterPartIds": {},
        "pageContents": {}, "masterContents": {}
    }))
    .unwrap()
}

fn cell(name: &str, value: &str) -> Cell {
    Cell {
        name: name.into(),
        formula: None,
        value: Some(value.into()),
        unit: None,
        del: false,
        other_attrs: vec![],
    }
}

fn formula_cell(name: &str, formula: &str) -> Cell {
    Cell {
        name: name.into(),
        formula: Some(formula.into()),
        value: None,
        unit: None,
        del: false,
        other_attrs: vec![],
    }
}

fn deleted_cell(name: &str) -> Cell {
    Cell {
        name: name.into(),
        formula: None,
        value: None,
        unit: None,
        del: true,
        other_attrs: vec![],
    }
}

fn shape(id: u32, children: Vec<ShapeChild>) -> Shape {
    Shape {
        id,
        name: None,
        name_u: None,
        shape_type: None,
        master: None,
        master_shape: None,
        line_style: None,
        fill_style: None,
        text_style: None,
        children,
        del: false,
        other_attrs: vec![],
    }
}

fn sheet(id: Option<u32>, children: Vec<SheetChild>) -> Sheet {
    Sheet {
        id,
        children,
        other_attrs: vec![],
    }
}

fn section(name: &str, rows: Vec<Row>) -> Section {
    Section {
        name: name.into(),
        index: None,
        del: false,
        children: rows.into_iter().map(SectionChild::Row).collect(),
        other_attrs: vec![],
    }
}

fn row(index: u32, cells: Vec<Cell>) -> Row {
    Row {
        index: Some(index),
        name: None,
        local_name: None,
        row_type: None,
        del: false,
        children: cells.into_iter().map(RowChild::Cell).collect(),
        other_attrs: vec![],
    }
}

fn deleted_row(index: u32) -> Row {
    deleted_row_with_cells(index, vec![])
}

fn deleted_row_with_cells(index: u32, cells: Vec<Cell>) -> Row {
    Row {
        index: Some(index),
        name: None,
        local_name: None,
        row_type: None,
        del: true,
        children: cells.into_iter().map(RowChild::Cell).collect(),
        other_attrs: vec![],
    }
}

fn endpoint_cells() -> Vec<ShapeChild> {
    vec![
        ShapeChild::Cell(cell("BeginX", "1")),
        ShapeChild::Cell(cell("BeginY", "2")),
        ShapeChild::Cell(cell("EndX", "4")),
        ShapeChild::Cell(cell("EndY", "2")),
    ]
}

fn connectivity_with_glued_source(source: Vec<ShapeChild>) -> crate::PageConnectivity {
    let mut package = package();
    package.page_contents.insert(
        "page".into(),
        sheet(
            None,
            vec![
                SheetChild::Shapes(vec![
                    vsdx_parse::ShapesChild::Shape(shape(1, source)),
                    vsdx_parse::ShapesChild::Shape(shape(
                        2,
                        vec![
                            ShapeChild::Cell(cell("Width", "1")),
                            ShapeChild::Cell(cell("Height", "1")),
                            ShapeChild::Cell(cell("PinX", "0")),
                            ShapeChild::Cell(cell("PinY", "0")),
                            ShapeChild::Cell(cell("LocPinX", "0")),
                            ShapeChild::Cell(cell("LocPinY", "0")),
                            ShapeChild::Section(section(
                                "Connection",
                                vec![row(0, vec![cell("X", "0.5"), cell("Y", "0.5")])],
                            )),
                        ],
                    )),
                ]),
                SheetChild::Connects(vec![ConnectsChild::Connect(Connect {
                    from_sheet: 1,
                    from_cell: Some("BeginX".into()),
                    from_part: Some(9),
                    to_sheet: 2,
                    to_cell: Some("Connections.X1".into()),
                    to_part: Some(100),
                    other_attrs: vec![],
                })]),
            ],
        ),
    );
    Resolver::new(&package)
        .resolve_page_connectivity("page")
        .unwrap()
}

fn connectivity_to_connection(rows: Vec<Row>, to_cell: &str) -> crate::PageConnectivity {
    let mut package = package();
    package.page_contents.insert(
        "page".into(),
        sheet(
            None,
            vec![
                SheetChild::Shapes(vec![
                    vsdx_parse::ShapesChild::Shape(shape(
                        1,
                        vec![ShapeChild::Cell(cell("OneD", "1"))],
                    )),
                    vsdx_parse::ShapesChild::Shape(shape(
                        2,
                        vec![
                            ShapeChild::Cell(cell("Width", "1")),
                            ShapeChild::Cell(cell("Height", "1")),
                            ShapeChild::Cell(cell("PinX", "0")),
                            ShapeChild::Cell(cell("PinY", "0")),
                            ShapeChild::Cell(cell("LocPinX", "0")),
                            ShapeChild::Cell(cell("LocPinY", "0")),
                            ShapeChild::Section(section("Connection", rows)),
                        ],
                    )),
                ]),
                SheetChild::Connects(vec![ConnectsChild::Connect(Connect {
                    from_sheet: 1,
                    from_cell: Some("BeginX".into()),
                    from_part: None,
                    to_sheet: 2,
                    to_cell: Some(to_cell.into()),
                    to_part: None,
                    other_attrs: vec![],
                })]),
            ],
        ),
    );
    Resolver::new(&package)
        .resolve_page_connectivity("page")
        .unwrap()
}

#[test]
fn resolves_first_connection_row_from_one_based_ordinal() {
    let connectivity = connectivity_to_connection(
        vec![row(0, vec![cell("X", "3"), cell("Y", "4")])],
        "Connections.X1",
    );

    let point = connectivity.connectors[&1].glue[0]
        .to
        .as_ref()
        .unwrap()
        .connection_point
        .as_ref()
        .unwrap();
    assert_eq!(point.row, 0);
    assert_eq!(point.position, crate::ScenePoint { x: 3.0, y: 4.0 });
    assert!(connectivity.diagnostics.is_empty());
}

#[test]
fn resolves_last_connection_row_from_one_based_ordinal() {
    let connectivity = connectivity_to_connection(
        vec![
            row(0, vec![cell("X", "10"), cell("Y", "10")]),
            row(1, vec![cell("X", "20"), cell("Y", "20")]),
            row(2, vec![cell("X", "30"), cell("Y", "30")]),
            row(3, vec![cell("X", "40"), cell("Y", "40")]),
            row(4, vec![cell("X", "50"), cell("Y", "50")]),
            row(5, vec![cell("X", "60"), cell("Y", "60")]),
        ],
        "Connections.X6",
    );

    let point = connectivity.connectors[&1].glue[0]
        .to
        .as_ref()
        .unwrap()
        .connection_point
        .as_ref()
        .unwrap();
    assert_eq!(point.row, 5);
    assert_eq!(point.position, crate::ScenePoint { x: 60.0, y: 60.0 });
    assert!(connectivity.diagnostics.is_empty());
}

#[test]
fn rejects_zero_connection_ordinal() {
    let connectivity = connectivity_to_connection(
        vec![row(0, vec![cell("X", "3"), cell("Y", "4")])],
        "Connections.X0",
    );

    assert!(
        connectivity.connectors[&1].glue[0]
            .to
            .as_ref()
            .unwrap()
            .connection_point
            .is_none()
    );
    assert_eq!(
        connectivity.diagnostics,
        vec![ConnectivityDiagnostic::UnsupportedToCell {
            shape_id: 2,
            cell: "Connections.X0".into(),
        }]
    );
}

#[test]
fn reports_out_of_range_connection_ordinal() {
    let connectivity = connectivity_to_connection(
        vec![
            row(0, vec![cell("X", "3"), cell("Y", "4")]),
            row(1, vec![cell("X", "5"), cell("Y", "6")]),
        ],
        "Connections.X9",
    );

    assert!(
        connectivity.connectors[&1].glue[0]
            .to
            .as_ref()
            .unwrap()
            .connection_point
            .is_none()
    );
    assert_eq!(
        connectivity.diagnostics,
        vec![ConnectivityDiagnostic::MissingConnectionPoint {
            shape_id: 2,
            row: 8,
        }]
    );
}

#[test]
fn endpoint_cells_without_one_d_resolve_as_a_glued_connector() {
    let connectivity = connectivity_with_glued_source(endpoint_cells());

    let connector = connectivity.connectors.get(&1).unwrap();
    assert!(connector.is_1d);
    assert_eq!(connector.begin, Some(crate::ScenePoint { x: 1.0, y: 2.0 }));
    assert_eq!(connector.end, Some(crate::ScenePoint { x: 4.0, y: 2.0 }));
    assert_eq!(connector.glue.len(), 1);
    assert!(connectivity.diagnostics.is_empty());
}

#[test]
fn one_d_zero_overrides_endpoint_cells() {
    let mut source = endpoint_cells();
    source.insert(0, ShapeChild::Cell(cell("OneD", "0")));

    let connectivity = connectivity_with_glued_source(source);

    assert!(!connectivity.connectors[&1].is_1d);
    assert_eq!(
        connectivity.diagnostics,
        vec![ConnectivityDiagnostic::UnsupportedFromCell {
            shape_id: 1,
            cell: "BeginX".into(),
        }]
    );
}

#[test]
fn one_d_one_remains_a_connector() {
    let mut source = endpoint_cells();
    source.insert(0, ShapeChild::Cell(cell("OneD", "1")));

    let connectivity = connectivity_with_glued_source(source);

    assert!(connectivity.connectors[&1].is_1d);
    assert!(connectivity.diagnostics.is_empty());
}

#[test]
fn incomplete_endpoint_cells_do_not_make_a_connector() {
    let mut source = endpoint_cells();
    source.pop();

    let connectivity = connectivity_with_glued_source(source);

    assert!(!connectivity.connectors[&1].is_1d);
    assert_eq!(
        connectivity.diagnostics,
        vec![ConnectivityDiagnostic::UnsupportedFromCell {
            shape_id: 1,
            cell: "BeginX".into(),
        }]
    );
}

#[test]
fn resolves_glue_connection_points_and_part_fields() {
    let mut package = package();
    package.page_contents.insert(
        "page".into(),
        sheet(
            None,
            vec![
                SheetChild::Shapes(vec![
                    vsdx_parse::ShapesChild::Shape(shape(
                        1,
                        vec![
                            ShapeChild::Cell(cell("OneD", "1")),
                            ShapeChild::Cell(cell("BeginX", "1")),
                            ShapeChild::Cell(cell("BeginY", "2")),
                            ShapeChild::Cell(cell("EndX", "4")),
                            ShapeChild::Cell(cell("EndY", "2")),
                        ],
                    )),
                    vsdx_parse::ShapesChild::Shape(shape(
                        2,
                        vec![
                            ShapeChild::Cell(cell("Width", "4")),
                            ShapeChild::Cell(cell("Height", "2")),
                            ShapeChild::Cell(cell("PinX", "10")),
                            ShapeChild::Cell(cell("PinY", "5")),
                            ShapeChild::Section(section(
                                "Connection",
                                vec![row(
                                    0,
                                    vec![
                                        formula_cell("X", "Width*0.5"),
                                        formula_cell("Y", "Height/2"),
                                    ],
                                )],
                            )),
                        ],
                    )),
                ]),
                SheetChild::Connects(vec![ConnectsChild::Connect(Connect {
                    from_sheet: 1,
                    from_cell: Some("BeginX".into()),
                    from_part: Some(9),
                    to_sheet: 2,
                    to_cell: Some("Connections.X1".into()),
                    to_part: Some(100),
                    other_attrs: vec![],
                })]),
            ],
        ),
    );
    let connectivity = Resolver::new(&package)
        .resolve_page_connectivity("page")
        .unwrap();
    let connector = connectivity.connectors.get(&1).unwrap();
    assert_eq!(connector.begin.unwrap().x, 1.0);
    assert_eq!(connector.glue[0].endpoint, ConnectorEndpoint::Begin);
    assert_eq!(connector.glue[0].from_part, Some(9));
    let target = connector.glue[0].to.as_ref().unwrap();
    assert_eq!(target.part, Some(100));
    assert_eq!(target.connection_point.as_ref().unwrap().position.x, 10.0);
    assert_eq!(target.connection_point.as_ref().unwrap().position.y, 5.0);
    assert_eq!(
        target.connection_point.as_ref().unwrap().x_provenance,
        crate::NumericProvenance::Formula
    );
    assert_eq!(
        target.connection_point.as_ref().unwrap().y_provenance,
        crate::NumericProvenance::Formula
    );
}

#[test]
fn connectivity_with_resolved_shapes_matches_page_resolution() {
    let package = parse_vsdx(include_bytes!(
        "../../vsdx-parse/tests/fixtures/grouped-glue.vsdx"
    ))
    .unwrap();
    let page = &package.page_part_paths[0];
    let resolver = Resolver::new(&package);
    let shapes = resolver.resolve_page_shapes(page).unwrap();

    assert_eq!(
        resolver.resolve_page_connectivity(page).unwrap(),
        resolver
            .resolve_page_connectivity_with(page, &shapes)
            .unwrap()
    );
}

#[test]
fn glue_numeric_provenance_prefers_supported_formulas_and_falls_back_to_cached_values() {
    let mut package = package();
    package.page_contents.insert(
        "page".into(),
        sheet(
            None,
            vec![
                SheetChild::Shapes(vec![
                    vsdx_parse::ShapesChild::Shape(shape(
                        1,
                        vec![ShapeChild::Cell(cell("OneD", "1"))],
                    )),
                    vsdx_parse::ShapesChild::Shape(shape(
                        2,
                        vec![
                            ShapeChild::Cell(cell("Width", "1")),
                            ShapeChild::Cell(cell("Height", "1")),
                            ShapeChild::Cell(cell("PinX", "0")),
                            ShapeChild::Cell(cell("PinY", "0")),
                            ShapeChild::Cell(cell("LocPinX", "0")),
                            ShapeChild::Cell(cell("LocPinY", "0")),
                            ShapeChild::Section(section(
                                "Connection",
                                vec![row(
                                    0,
                                    vec![
                                        Cell {
                                            name: "X".into(),
                                            formula: Some("2+3".into()),
                                            value: Some("99".into()),
                                            unit: None,
                                            del: false,
                                            other_attrs: vec![],
                                        },
                                        Cell {
                                            name: "Y".into(),
                                            formula: Some("GUARD(4)".into()),
                                            value: Some("7".into()),
                                            unit: None,
                                            del: false,
                                            other_attrs: vec![],
                                        },
                                    ],
                                )],
                            )),
                        ],
                    )),
                ]),
                SheetChild::Connects(vec![ConnectsChild::Connect(Connect {
                    from_sheet: 1,
                    from_cell: Some("BeginX".into()),
                    from_part: None,
                    to_sheet: 2,
                    to_cell: Some("Connections.X1".into()),
                    to_part: None,
                    other_attrs: vec![],
                })]),
            ],
        ),
    );
    let connectivity = Resolver::new(&package)
        .resolve_page_connectivity("page")
        .unwrap();
    let point = connectivity.connectors[&1].glue[0]
        .to
        .as_ref()
        .unwrap()
        .connection_point
        .as_ref()
        .unwrap();
    assert_eq!(point.position, crate::ScenePoint { x: 5.0, y: 7.0 });
    assert_eq!(point.x_provenance, crate::NumericProvenance::Formula);
    assert_eq!(point.y_provenance, crate::NumericProvenance::CachedValue);
}

#[test]
fn reports_dangling_connects_without_dropping_them() {
    let mut package = package();
    package.page_contents.insert(
        "page".into(),
        sheet(
            None,
            vec![
                SheetChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(shape(
                    1,
                    vec![ShapeChild::Cell(cell("OneD", "1"))],
                ))]),
                SheetChild::Connects(vec![ConnectsChild::Connect(Connect {
                    from_sheet: 1,
                    from_cell: Some("EndX".into()),
                    from_part: None,
                    to_sheet: 99,
                    to_cell: Some("Connections.X7".into()),
                    to_part: None,
                    other_attrs: vec![],
                })]),
            ],
        ),
    );
    let connectivity = Resolver::new(&package)
        .resolve_page_connectivity("page")
        .unwrap();
    assert_eq!(connectivity.connectors.get(&1).unwrap().glue.len(), 1);
    assert_eq!(
        connectivity.diagnostics,
        vec![ConnectivityDiagnostic::MissingToShape { shape_id: 99 }]
    );
}

#[test]
fn grouped_glue_connection_points_use_scene_transforms() {
    let package = parse_vsdx(include_bytes!(
        "../../vsdx-parse/tests/fixtures/grouped-glue.vsdx"
    ))
    .unwrap();
    let page = &package.page_part_paths[0];
    let connectivity = Resolver::new(&package)
        .resolve_page_connectivity(page)
        .unwrap();
    let points = connectivity.connectors[&1]
        .glue
        .iter()
        .filter(|glue| glue.to.as_ref().unwrap().cell.as_deref() != Some("PinX"))
        .map(|glue| {
            let target = glue.to.as_ref().unwrap();
            (
                target.shape_id,
                target.connection_point.as_ref().unwrap().position,
            )
        })
        .collect::<std::collections::BTreeMap<_, _>>();

    // Target 11: rotate the local (0.5, 0.5) by +90° and scale by 2 around (10, 10):
    // (-1, 1) + (10, 10) = (9, 11). Target 21 scales (0.5, 0.5) by (2, 2) at
    // (20, 10) = (21, 11). Target 32 is scaled by 2 in each nested group: (0.5, 0.5)
    // becomes (2, 2), then the outer group's origin maps it to (32, 2).
    assert_eq!(points[&11].x, 9.0);
    assert_eq!(points[&11].y, 11.0);
    assert_eq!(points[&21].x, 21.0);
    assert_eq!(points[&21].y, 11.0);
    assert_eq!(points[&32].x, 32.0);
    assert_eq!(points[&32].y, 2.0);

    let direct_pins = connectivity.connectors[&1]
        .glue
        .iter()
        .filter(|glue| glue.to.as_ref().unwrap().cell.as_deref() == Some("PinX"))
        .map(|glue| {
            let target = glue.to.as_ref().unwrap();
            (
                target.shape_id,
                target.connection_point.as_ref().unwrap().position,
            )
        })
        .collect::<std::collections::BTreeMap<_, _>>();

    // A target pin is its local LocPin transformed through every containing group.
    assert_eq!(direct_pins[&11], crate::ScenePoint { x: 10.0, y: 10.0 });
    assert_eq!(direct_pins[&21], crate::ScenePoint { x: 20.0, y: 10.0 });
    assert_eq!(direct_pins[&32], crate::ScenePoint { x: 30.0, y: 0.0 });
}

#[test]
fn deleted_style_rows_do_not_contribute_cells() {
    let mut package = package();
    package.style_sheets = vec![sheet(
        Some(1),
        vec![SheetChild::Section(section(
            "Character",
            vec![deleted_row_with_cells(
                0,
                vec![cell("Leaked", "style"), deleted_cell("Shared")],
            )],
        ))],
    )];
    package.page_sheets.insert(
        1,
        sheet(
            None,
            vec![SheetChild::Section(section(
                "Character",
                vec![row(0, vec![cell("Shared", "page")])],
            ))],
        ),
    );
    let mut local = shape(
        1,
        vec![ShapeChild::Section(section(
            "Character",
            vec![row(0, vec![])],
        ))],
    );
    local.text_style = Some(1);
    add_page(&mut package, local);

    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    let cells = &resolved.sections["Character"].rows["IX:0"].cells;
    assert!(!cells.contains_key("Leaked"));
    match &cells["Shared"] {
        Lookup::Found(value) => {
            assert_eq!(value.cell.value.as_deref(), Some("page"));
            assert_eq!(value.provenance, Provenance::Page);
        }
        value => panic!("expected page cell, got {value:?}"),
    }
}

#[test]
fn deleted_master_rows_do_not_suppress_lower_live_cells() {
    let mut package = package();
    package.page_sheets.insert(
        1,
        sheet(
            None,
            vec![SheetChild::Section(section(
                "Character",
                vec![row(0, vec![cell("Char", "page")])],
            ))],
        ),
    );
    let mut local = shape(
        1,
        vec![ShapeChild::Section(section(
            "Character",
            vec![row(0, vec![])],
        ))],
    );
    local.master = Some(1);
    add_page(&mut package, local);
    add_master(
        &mut package,
        1,
        shape(
            1,
            vec![ShapeChild::Section(section(
                "Character",
                vec![deleted_row_with_cells(0, vec![cell("Char", "master")])],
            ))],
        ),
    );

    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    match &resolved.sections["Character"].rows["IX:0"].cells["Char"] {
        Lookup::Found(value) => {
            assert_eq!(value.cell.value.as_deref(), Some("page"));
            assert_eq!(value.provenance, Provenance::Page);
        }
        value => panic!("expected page cell, got {value:?}"),
    }
}

#[test]
fn row_deletion_uses_the_highest_priority_row_and_preserves_its_state() {
    let mut package = package();
    package.style_sheets = vec![sheet(
        Some(1),
        vec![SheetChild::Section(section(
            "Character",
            vec![deleted_row(0)],
        ))],
    )];

    let mut local = shape(
        1,
        vec![ShapeChild::Section(section(
            "Character",
            vec![row(0, vec![cell("Char", "local")])],
        ))],
    );
    local.text_style = Some(1);
    let mut deleted = shape(
        2,
        vec![ShapeChild::Section(section(
            "Character",
            vec![deleted_row(0)],
        ))],
    );
    deleted.text_style = Some(1);
    let empty = shape(
        3,
        vec![ShapeChild::Section(section(
            "Character",
            vec![row(0, vec![])],
        ))],
    );
    package.page_part_ids.insert("page".into(), 1);
    package.page_contents.insert(
        "page".into(),
        sheet(
            None,
            vec![SheetChild::Shapes(vec![
                vsdx_parse::ShapesChild::Shape(local),
                vsdx_parse::ShapesChild::Shape(deleted),
                vsdx_parse::ShapesChild::Shape(empty),
            ])],
        ),
    );

    let resolver = Resolver::new(&package);
    let local = resolver.resolve_shape("page", 1).unwrap();
    let local_row = &local.sections["Character"].rows["IX:0"];
    assert!(!local_row.deleted);
    match &local_row.cells["Char"] {
        Lookup::Found(value) => assert_eq!(value.provenance, Provenance::Local),
        value => panic!("expected local row cell, got {value:?}"),
    }

    let deleted = resolver.resolve_shape("page", 2).unwrap();
    let deleted_row = &deleted.sections["Character"].rows["IX:0"];
    assert!(deleted_row.deleted);
    assert!(deleted_row.cells.is_empty());

    let empty = resolver.resolve_shape("page", 3).unwrap();
    let empty_row = &empty.sections["Character"].rows["IX:0"];
    assert!(!empty_row.deleted);
    assert!(empty_row.cells.is_empty());
}

fn found<'a>(shape: &'a crate::ResolvedShape, name: &str) -> (&'a str, Provenance) {
    match shape.cells.get(name) {
        Some(Lookup::Found(value)) => (value.cell.value.as_deref().unwrap(), value.provenance),
        value => panic!("expected {name} to be found, got {value:?}"),
    }
}

fn found_row<'a>(
    section: &'a crate::ResolvedSection,
    row_key: &str,
    name: &str,
) -> (&'a str, Provenance) {
    match section.rows[row_key].cells.get(name) {
        Some(Lookup::Found(value)) => (value.cell.value.as_deref().unwrap(), value.provenance),
        value => panic!("expected {row_key}.{name} to be found, got {value:?}"),
    }
}

fn add_page(package: &mut VsdxPackage, value: Shape) {
    package.page_part_ids.insert("page".into(), 1);
    package.page_contents.insert(
        "page".into(),
        sheet(
            None,
            vec![SheetChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
                value,
            )])],
        ),
    );
}

fn add_master(package: &mut VsdxPackage, id: u32, value: Shape) {
    let path = format!("master{id}");
    package.master_part_ids.insert(path.clone(), id);
    package.master_contents.insert(
        path,
        sheet(
            None,
            vec![SheetChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
                value,
            )])],
        ),
    );
}

fn add_master_shapes(package: &mut VsdxPackage, id: u32, values: Vec<Shape>) {
    let path = format!("master{id}");
    package.master_part_ids.insert(path.clone(), id);
    package.master_contents.insert(
        path,
        sheet(
            None,
            vec![SheetChild::Shapes(
                values
                    .into_iter()
                    .map(vsdx_parse::ShapesChild::Shape)
                    .collect(),
            )],
        ),
    );
}

#[test]
fn page_shape_tree_matches_by_id_resolution_for_nested_group_leaf() {
    let mut package = package();
    package.page_part_ids.insert("page".into(), 1);
    package.page_sheets.insert(
        1,
        sheet(None, vec![SheetChild::Cell(cell("PageValue", "page"))]),
    );
    let leaf = shape(4, vec![ShapeChild::Cell(cell("LeafValue", "leaf"))]);
    let inner = shape(
        3,
        vec![ShapeChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
            leaf,
        )])],
    );
    let middle = shape(
        2,
        vec![ShapeChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
            inner,
        )])],
    );
    let outer = shape(
        1,
        vec![ShapeChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
            middle,
        )])],
    );
    package.page_contents.insert(
        "page".into(),
        sheet(
            None,
            vec![SheetChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
                outer,
            )])],
        ),
    );

    let resolver = Resolver::new(&package);
    let resolved_by_tree = resolver.resolve_page_shapes("page").unwrap();
    let resolved_by_id = resolver.resolve_shape("page", 4).unwrap();

    assert_eq!(resolved_by_tree[&4], resolved_by_id);
    assert_eq!(
        found(&resolved_by_tree[&4], "PageValue"),
        ("page", Provenance::Page)
    );
}

#[test]
fn master_without_master_shape_inherits_from_master_root() {
    let mut package = package();
    let mut local = shape(10, vec![]);
    local.master = Some(5);
    add_page(&mut package, local);
    add_master_shapes(
        &mut package,
        5,
        vec![
            shape(50, vec![ShapeChild::Cell(cell("PinX", "root"))]),
            shape(51, vec![ShapeChild::Cell(cell("PinX", "other"))]),
        ],
    );

    let resolved = Resolver::new(&package).resolve_shape("page", 10).unwrap();
    assert_eq!(found(&resolved, "PinX"), ("root", Provenance::Master));
}

#[test]
fn master_ignores_master_shape_and_inherits_from_master_root() {
    let mut package = package();
    let mut local = shape(10, vec![]);
    local.master = Some(5);
    local.master_shape = Some(51);
    add_page(&mut package, local);
    add_master_shapes(
        &mut package,
        5,
        vec![
            shape(50, vec![ShapeChild::Cell(cell("PinX", "root"))]),
            shape(51, vec![ShapeChild::Cell(cell("PinX", "specified"))]),
        ],
    );

    let resolved = Resolver::new(&package).resolve_shape("page", 10).unwrap();
    assert_eq!(found(&resolved, "PinX"), ("root", Provenance::Master));
}

#[test]
fn master_shape_inherits_from_the_enclosing_masters_subshape() {
    let mut package = package();
    let mut group = shape(
        10,
        vec![ShapeChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
            shape(11, vec![]),
        )])],
    );
    group.master = Some(5);
    let ShapeChild::Shapes(children) = &mut group.children[0] else {
        panic!("expected group children");
    };
    let vsdx_parse::ShapesChild::Shape(local) = &mut children[0] else {
        panic!("expected local subshape");
    };
    local.master_shape = Some(51);
    add_page(&mut package, group);
    add_master_shapes(
        &mut package,
        5,
        vec![shape(
            50,
            vec![ShapeChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
                shape(51, vec![ShapeChild::Cell(cell("PinX", "subshape"))]),
            )])],
        )],
    );

    let resolved = Resolver::new(&package).resolve_shape("page", 11).unwrap();
    assert_eq!(
        found(&resolved, "PinX"),
        ("subshape", Provenance::MasterShape)
    );
}

#[test]
fn missing_master_reports_a_diagnostic() {
    let mut package = package();
    let mut local = shape(10, vec![]);
    local.master = Some(5);
    add_page(&mut package, local);

    assert_eq!(
        Resolver::new(&package).resolve_shape("page", 10),
        Err(ResolveError::MissingMaster(5))
    );
}

#[test]
fn master_inheritance_walks_deeply_and_local_overrides() {
    let mut package = package();
    let mut local = shape(10, vec![]);
    local.master = Some(1);
    add_page(&mut package, local);
    let mut first = shape(1, vec![]);
    first.master = Some(2);
    add_master(&mut package, 1, first);
    let mut second = shape(2, vec![]);
    second.master = Some(3);
    add_master(&mut package, 2, second);
    add_master(
        &mut package,
        3,
        shape(
            3,
            vec![
                ShapeChild::Cell(cell("PinX", "furthest")),
                ShapeChild::Cell(cell("PinY", "master-shape")),
            ],
        ),
    );
    let resolved = Resolver::new(&package).resolve_shape("page", 10).unwrap();
    assert_eq!(found(&resolved, "PinX"), ("furthest", Provenance::Master));
    let mut direct = shape(12, vec![]);
    direct.master = Some(3);
    add_page(&mut package, direct);
    assert_eq!(
        found(
            &Resolver::new(&package).resolve_shape("page", 12).unwrap(),
            "PinY"
        ),
        ("master-shape", Provenance::Master)
    );

    let mut local = shape(11, vec![ShapeChild::Cell(cell("PinX", "local"))]);
    local.master = Some(4);
    add_page(&mut package, local);
    add_master(
        &mut package,
        4,
        shape(4, vec![ShapeChild::Cell(cell("PinX", "master"))]),
    );
    let resolved = Resolver::new(&package).resolve_shape("page", 11).unwrap();
    assert_eq!(found(&resolved, "PinX"), ("local", Provenance::Local));
}

#[test]
fn style_slices_and_based_on_chains_resolve_independently() {
    let mut package = package();
    package.style_sheets = vec![
        sheet(Some(1), vec![SheetChild::Cell(cell("LineColor", "line"))]),
        sheet(Some(2), vec![SheetChild::Cell(cell("FillForegnd", "fill"))]),
        sheet(Some(3), vec![SheetChild::Cell(cell("Text", "text"))]),
        sheet(
            Some(4),
            vec![SheetChild::Cell(cell("LineWeight", "ancestor"))],
        ),
        Sheet {
            id: Some(5),
            children: vec![SheetChild::Cell(cell("LineWeight", "child"))],
            other_attrs: vec![("BasedOn".into(), "4".into())],
        },
        sheet(
            Some(6),
            vec![SheetChild::Cell(cell("LineWeight", "only-ancestor"))],
        ),
        Sheet {
            id: Some(7),
            children: vec![],
            other_attrs: vec![("BasedOn".into(), "6".into())],
        },
    ];
    let mut value = shape(1, vec![]);
    value.line_style = Some(1);
    value.fill_style = Some(2);
    value.text_style = Some(3);
    add_page(&mut package, value);
    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    assert_eq!(
        found(&resolved, "LineColor"),
        ("line", Provenance::StyleLine)
    );
    assert_eq!(
        found(&resolved, "FillForegnd"),
        ("fill", Provenance::StyleFill)
    );
    assert_eq!(found(&resolved, "Text"), ("text", Provenance::StyleText));

    let mut inherited = shape(2, vec![]);
    inherited.line_style = Some(7);
    add_page(&mut package, inherited);
    assert_eq!(
        found(
            &Resolver::new(&package).resolve_shape("page", 2).unwrap(),
            "LineWeight"
        ),
        ("only-ancestor", Provenance::StyleLine)
    );
    let mut overridden = shape(3, vec![]);
    overridden.line_style = Some(5);
    add_page(&mut package, overridden);
    assert_eq!(
        found(
            &Resolver::new(&package).resolve_shape("page", 3).unwrap(),
            "LineWeight"
        ),
        ("child", Provenance::StyleLine)
    );
}

#[test]
fn style_refs_fall_back_to_the_master_when_the_local_shape_has_none() {
    let mut package = package();
    package.style_sheets = vec![
        sheet(Some(1), vec![SheetChild::Cell(cell("LineColor", "line"))]),
        sheet(Some(2), vec![SheetChild::Cell(cell("FillForegnd", "fill"))]),
        sheet(Some(3), vec![SheetChild::Cell(cell("Text", "text"))]),
    ];
    let mut master_value = shape(4, vec![]);
    master_value.line_style = Some(1);
    master_value.fill_style = Some(2);
    master_value.text_style = Some(3);
    add_master(&mut package, 4, master_value);

    let mut local = shape(10, vec![]);
    local.master = Some(4);
    add_page(&mut package, local);

    let resolved = Resolver::new(&package).resolve_shape("page", 10).unwrap();
    assert_eq!(
        found(&resolved, "LineColor"),
        ("line", Provenance::StyleLine)
    );
    assert_eq!(
        found(&resolved, "FillForegnd"),
        ("fill", Provenance::StyleFill)
    );
    assert_eq!(found(&resolved, "Text"), ("text", Provenance::StyleText));
}

#[test]
fn sheet_style_attributes_resolve_style_to_style_inheritance() {
    let mut package = package();
    package.style_sheets = vec![
        Sheet {
            id: Some(1),
            children: vec![SheetChild::Cell(formula_cell("LineWeight", "Inh"))],
            other_attrs: vec![("lInEsTyLe".into(), "2".into())],
        },
        Sheet {
            id: Some(2),
            children: vec![SheetChild::Cell(formula_cell("LineWeight", "Inh"))],
            other_attrs: vec![("BasedOn".into(), "3".into())],
        },
        sheet(Some(3), vec![SheetChild::Cell(cell("LineWeight", "3"))]),
    ];

    let resolved = Resolver::new(&package)
        .resolve_sheet(&package.style_sheets[0])
        .unwrap();
    assert_eq!(found(&resolved, "LineWeight"), ("3", Provenance::StyleLine));
}

#[test]
fn sheets_without_style_attributes_preserve_unresolved_inh() {
    let resolved = Resolver::new(&package())
        .resolve_sheet(&sheet(
            None,
            vec![SheetChild::Cell(formula_cell("LineWeight", "Inh"))],
        ))
        .unwrap();

    match &resolved.cells["LineWeight"] {
        Lookup::Found(value) => {
            assert_eq!(value.cell.formula.as_deref(), Some("Inh"));
            assert_eq!(value.provenance, Provenance::Local);
        }
        value => panic!("expected unresolved Inh, got {value:?}"),
    }
}

#[test]
fn unparseable_sheet_style_attributes_are_ignored() {
    let unresolved = sheet(
        None,
        vec![SheetChild::Cell(formula_cell("LineWeight", "Inh"))],
    );
    let sheet = Sheet {
        other_attrs: vec![("LineStyle".into(), "not-a-style-id".into())],
        ..unresolved
    };

    let resolved = Resolver::new(&package()).resolve_sheet(&sheet).unwrap();
    match &resolved.cells["LineWeight"] {
        Lookup::Found(value) => assert_eq!(value.cell.formula.as_deref(), Some("Inh")),
        value => panic!("expected unresolved Inh, got {value:?}"),
    }
}

#[test]
fn inh_skips_each_inherited_layer_until_a_concrete_cell() {
    let mut package = package();
    let mut local = shape(1, vec![ShapeChild::Cell(formula_cell("PinX", "Inh"))]);
    local.master = Some(1);
    add_page(&mut package, local);
    let mut master = shape(1, vec![ShapeChild::Cell(formula_cell("PinX", "Inh"))]);
    master.master = Some(2);
    add_master(&mut package, 1, master);
    add_master(
        &mut package,
        2,
        shape(2, vec![ShapeChild::Cell(cell("PinX", "4"))]),
    );
    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    assert_eq!(found(&resolved, "PinX"), ("4", Provenance::Master));

    package.style_sheets = vec![
        Sheet {
            id: Some(1),
            children: vec![SheetChild::Cell(formula_cell("LineWeight", "Inh"))],
            other_attrs: vec![("BasedOn".into(), "2".into())],
        },
        sheet(Some(2), vec![SheetChild::Cell(cell("LineWeight", "3"))]),
    ];
    let mut styled = shape(3, vec![ShapeChild::Cell(formula_cell("LineWeight", "Inh"))]);
    styled.line_style = Some(1);
    add_page(&mut package, styled);
    let resolved = Resolver::new(&package).resolve_shape("page", 3).unwrap();
    assert_eq!(found(&resolved, "LineWeight"), ("3", Provenance::StyleLine));

    let unresolved = shape(4, vec![ShapeChild::Cell(formula_cell("PinY", "Inh"))]);
    add_page(&mut package, unresolved);
    let resolved = Resolver::new(&package).resolve_shape("page", 4).unwrap();
    match &resolved.cells["PinY"] {
        Lookup::Found(cell) => assert_eq!(cell.cell.formula.as_deref(), Some("Inh")),
        value => panic!("expected unresolved Inh, got {value:?}"),
    }

    let defaulted = shape(5, vec![ShapeChild::Cell(formula_cell("LocPinX", "Inh"))]);
    add_page(&mut package, defaulted);
    let resolved = Resolver::new(&package).resolve_shape("page", 5).unwrap();
    match &resolved.cells["LocPinX"] {
        Lookup::Found(cell) => {
            assert_eq!(cell.cell.formula.as_deref(), Some("Width * 0.5"));
            assert_eq!(cell.provenance, Provenance::Default);
        }
        value => panic!("expected documented default, got {value:?}"),
    }

    let concrete = shape(6, vec![ShapeChild::Cell(formula_cell("LocPinX", "Inh"))]);
    add_page(&mut package, concrete);
    package
        .document_sheet
        .get_or_insert_with(|| sheet(None, vec![]))
        .children
        .push(SheetChild::Cell(cell("LocPinX", "7")));
    let resolved = Resolver::new(&package).resolve_shape("page", 6).unwrap();
    assert_eq!(found(&resolved, "LocPinX"), ("7", Provenance::Document));
}

#[test]
fn inherited_master_cell_beats_documented_default() {
    let mut package = package();
    let mut local = shape(1, vec![ShapeChild::Cell(formula_cell("LocPinX", "Inh"))]);
    local.master = Some(1);
    add_page(&mut package, local);
    add_master(
        &mut package,
        1,
        shape(1, vec![ShapeChild::Cell(cell("LocPinX", "master"))]),
    );

    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    assert_eq!(found(&resolved, "LocPinX"), ("master", Provenance::Master));
}

#[test]
fn geometry_rows_without_ix_all_realize_in_source_order() {
    let package = parse_vsdx(include_bytes!(
        "../../vsdx-parse/tests/fixtures/geometry-anonymous-rows.vsdx"
    ))
    .unwrap();
    let page = &package.page_part_paths[0];
    let resolved = Resolver::new(&package).resolve_shape(page, 1).unwrap();
    let section = &resolved.sections["Geometry"];
    assert_eq!(section.row_order.len(), 2);
    let geometry = crate::realize_geometry(section, 1.0, 1.0);
    assert_eq!(
        geometry.commands,
        vec![
            ooxml_drawingml::GeometryPathCommand::Move { x: 1.0, y: 2.0 },
            ooxml_drawingml::GeometryPathCommand::Line { x: 3.0, y: 4.0 },
        ]
    );
}

#[test]
fn geometry_rows_with_duplicate_ix_all_realize_in_source_order() {
    let package = parse_vsdx(include_bytes!(
        "../../vsdx-parse/tests/fixtures/geometry-duplicate-ix-rows.vsdx"
    ))
    .unwrap();
    let page = &package.page_part_paths[0];
    let resolved = Resolver::new(&package).resolve_shape(page, 1).unwrap();
    let section = &resolved.sections["Geometry"];
    assert_eq!(section.row_order.len(), 2);
    let geometry = crate::realize_geometry(section, 1.0, 1.0);
    assert_eq!(
        geometry.commands,
        vec![
            ooxml_drawingml::GeometryPathCommand::Move { x: 1.0, y: 2.0 },
            ooxml_drawingml::GeometryPathCommand::Line { x: 3.0, y: 4.0 },
        ]
    );
}

#[test]
fn section_rows_inherit_by_name_or_ix_and_preserve_duplicate_occurrences() {
    let mut package = package();
    let mut local = shape(
        1,
        vec![ShapeChild::Section(section(
            "Geometry",
            vec![row(1, vec![cell("X", "local-one")])],
        ))],
    );
    local.master = Some(1);
    add_page(&mut package, local);
    add_master(
        &mut package,
        1,
        shape(
            1,
            vec![ShapeChild::Section(section(
                "Geometry",
                vec![
                    row(0, vec![cell("X", "master-zero")]),
                    row(1, vec![cell("X", "master-one")]),
                ],
            ))],
        ),
    );
    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    let geometry = &resolved.sections["Geometry"];
    assert_eq!(geometry.row_order, vec!["IX:1", "IX:0"]);
    assert_eq!(
        found_row(geometry, "IX:1", "X"),
        ("local-one", Provenance::Local)
    );
    assert_eq!(
        found_row(geometry, "IX:0", "X"),
        ("master-zero", Provenance::Master)
    );

    let mut reordered = shape(
        2,
        vec![ShapeChild::Section(section(
            "Geometry",
            vec![
                row(0, vec![cell("X", "local-zero")]),
                row(1, vec![cell("X", "local-one")]),
            ],
        ))],
    );
    reordered.master = Some(2);
    add_page(&mut package, reordered);
    add_master(
        &mut package,
        2,
        shape(
            2,
            vec![ShapeChild::Section(section(
                "Geometry",
                vec![
                    row(1, vec![cell("X", "master-one")]),
                    row(0, vec![cell("X", "master-zero")]),
                ],
            ))],
        ),
    );
    let resolved = Resolver::new(&package).resolve_shape("page", 2).unwrap();
    let geometry = &resolved.sections["Geometry"];
    assert_eq!(
        found_row(geometry, "IX:0", "X"),
        ("local-zero", Provenance::Local)
    );
    assert_eq!(
        found_row(geometry, "IX:1", "X"),
        ("local-one", Provenance::Local)
    );

    let mut named = row(0, vec![cell("X", "local")]);
    named.name = Some("TextPosition".into());
    let mut local = shape(
        3,
        vec![ShapeChild::Section(section("Character", vec![named]))],
    );
    local.master = Some(3);
    add_page(&mut package, local);
    let mut named = row(7, vec![cell("X", "master")]);
    named.name = Some("TextPosition".into());
    add_master(
        &mut package,
        3,
        shape(
            3,
            vec![ShapeChild::Section(section("Character", vec![named]))],
        ),
    );
    let resolved = Resolver::new(&package).resolve_shape("page", 3).unwrap();
    assert_eq!(
        found_row(&resolved.sections["Character"], "N:TextPosition", "X"),
        ("local", Provenance::Local)
    );
}

#[test]
fn unequal_duplicate_and_anonymous_rows_preserve_each_occurrence() {
    for anonymous in [false, true] {
        let mut first_package = package();
        let mut local_row = row(0, vec![cell("X", "local-first")]);
        let mut master_second = row(0, vec![cell("X", "master-second")]);
        if anonymous {
            local_row.index = None;
            master_second.index = None;
        }
        let mut local = shape(
            1,
            vec![ShapeChild::Section(section("Geometry", vec![local_row]))],
        );
        local.master = Some(1);
        add_page(&mut first_package, local);
        let mut master_first = row(0, vec![cell("X", "master-first")]);
        if anonymous {
            master_first.index = None;
        }
        add_master(
            &mut first_package,
            1,
            shape(
                1,
                vec![ShapeChild::Section(section(
                    "Geometry",
                    vec![master_first, master_second],
                ))],
            ),
        );
        let first_section = &Resolver::new(&first_package)
            .resolve_shape("page", 1)
            .unwrap()
            .sections["Geometry"];
        assert_eq!(first_section.row_order.len(), 2);
        assert_eq!(
            found_row(first_section, &first_section.row_order[0], "X").0,
            "local-first"
        );
        assert_eq!(
            found_row(first_section, &first_section.row_order[1], "X").0,
            "master-second"
        );

        let mut package = package();
        let mut first = row(0, vec![cell("X", "local-first")]);
        let mut second = row(0, vec![cell("X", "local-second")]);
        if anonymous {
            first.index = None;
            second.index = None;
        }
        let mut local = shape(
            1,
            vec![ShapeChild::Section(section(
                "Geometry",
                vec![first, second],
            ))],
        );
        local.master = Some(1);
        add_page(&mut package, local);
        let mut master = row(0, vec![cell("X", "master-first")]);
        if anonymous {
            master.index = None;
        }
        add_master(
            &mut package,
            1,
            shape(
                1,
                vec![ShapeChild::Section(section("Geometry", vec![master]))],
            ),
        );
        let section = &Resolver::new(&package)
            .resolve_shape("page", 1)
            .unwrap()
            .sections["Geometry"];
        assert_eq!(section.row_order.len(), 2);
        assert_eq!(
            found_row(section, &section.row_order[0], "X").0,
            "local-first"
        );
        assert_eq!(
            found_row(section, &section.row_order[1], "X").0,
            "local-second"
        );
    }
}

#[test]
fn inh_section_cells_skip_to_master_and_text_style() {
    let mut package = package();
    let mut local = shape(
        1,
        vec![ShapeChild::Section(section(
            "Geometry",
            vec![row(0, vec![formula_cell("X", "Inh")])],
        ))],
    );
    local.master = Some(1);
    add_master(
        &mut package,
        1,
        shape(
            1,
            vec![ShapeChild::Section(section(
                "Geometry",
                vec![row(0, vec![cell("X", "4")])],
            ))],
        ),
    );
    package.style_sheets = vec![sheet(
        Some(2),
        vec![SheetChild::Section(section(
            "Character",
            vec![row(0, vec![cell("Font", "3")])],
        ))],
    )];
    let mut styled = shape(
        2,
        vec![ShapeChild::Section(section(
            "Character",
            vec![row(0, vec![formula_cell("Font", "Inh")])],
        ))],
    );
    styled.text_style = Some(2);
    package.page_part_ids.insert("page".into(), 1);
    package.page_contents.insert(
        "page".into(),
        sheet(
            None,
            vec![SheetChild::Shapes(vec![
                vsdx_parse::ShapesChild::Shape(local),
                vsdx_parse::ShapesChild::Shape(styled),
            ])],
        ),
    );

    let resolver = Resolver::new(&package);
    let geometry = resolver.resolve_shape("page", 1).unwrap();
    match &geometry.sections["Geometry"].rows["IX:0"].cells["X"] {
        Lookup::Found(value) => {
            assert_eq!(value.cell.value.as_deref(), Some("4"));
            assert_eq!(value.provenance, Provenance::Master);
        }
        value => panic!("expected inherited geometry cell, got {value:?}"),
    }
    let character = resolver.resolve_shape("page", 2).unwrap();
    match &character.sections["Character"].rows["IX:0"].cells["Font"] {
        Lookup::Found(value) => {
            assert_eq!(value.cell.value.as_deref(), Some("3"));
            assert_eq!(value.provenance, Provenance::StyleText);
        }
        value => panic!("expected inherited character cell, got {value:?}"),
    }
}

#[test]
fn deletions_block_inheritance_while_absence_inherits() {
    let mut package = package();
    let mut master = shape(
        100,
        vec![
            ShapeChild::Cell(cell("PinX", "master")),
            ShapeChild::Section(section("Geometry", vec![row(0, vec![cell("X", "1")])])),
        ],
    );
    master.master = None;
    add_master_shapes(&mut package, 1, vec![master]);
    for (id, local) in [
        (1, shape(1, vec![ShapeChild::Cell(deleted_cell("PinX"))])),
        (
            2,
            shape(
                2,
                vec![ShapeChild::Section(Section {
                    name: "Geometry".into(),
                    index: None,
                    del: true,
                    children: vec![],
                    other_attrs: vec![],
                })],
            ),
        ),
        (
            3,
            shape(
                3,
                vec![ShapeChild::Section(section(
                    "Geometry",
                    vec![Row {
                        index: Some(0),
                        name: None,
                        local_name: None,
                        row_type: None,
                        del: true,
                        children: vec![],
                        other_attrs: vec![],
                    }],
                ))],
            ),
        ),
    ] {
        let mut local = local;
        local.master = Some(1);
        add_page(&mut package, local);
        let resolved = Resolver::new(&package).resolve_shape("page", id).unwrap();
        if id == 1 {
            assert_eq!(resolved.cells["PinX"], Lookup::Deleted);
        }
        if id == 2 {
            assert!(resolved.sections["Geometry"].deleted);
        }
        if id == 3 {
            assert!(resolved.sections["Geometry"].rows["IX:0"].cells.is_empty());
        }
    }
    let mut absent = shape(4, vec![]);
    absent.master = Some(1);
    add_page(&mut package, absent);
    let resolved = Resolver::new(&package).resolve_shape("page", 4).unwrap();
    assert_eq!(found(&resolved, "PinX"), ("master", Provenance::Master));
    assert_ne!(Lookup::Deleted, Lookup::Absent);
}

#[test]
fn shape_deletion_and_all_provenance_layers_are_exposed() {
    let mut package = package();
    package.document_sheet = Some(sheet(None, vec![SheetChild::Cell(cell("Doc", "document"))]));
    package
        .page_sheets
        .insert(1, sheet(None, vec![SheetChild::Cell(cell("Page", "page"))]));
    package.style_sheets = vec![
        sheet(Some(1), vec![SheetChild::Cell(cell("LineColor", "line"))]),
        sheet(Some(2), vec![SheetChild::Cell(cell("FillForegnd", "fill"))]),
        sheet(Some(3), vec![SheetChild::Cell(cell("Text", "text"))]),
    ];
    let mut local = shape(1, vec![ShapeChild::Cell(cell("Local", "local"))]);
    local.master = Some(1);
    local.line_style = Some(1);
    local.fill_style = Some(2);
    local.text_style = Some(3);
    add_page(&mut package, local);
    add_master(
        &mut package,
        1,
        shape(1, vec![ShapeChild::Cell(cell("Master", "master"))]),
    );
    let resolved = Resolver::new(&package)
        .with_defaults([cell("Default", "default")])
        .resolve_shape("page", 1)
        .unwrap();
    for (name, expected) in [
        ("Local", Provenance::Local),
        ("Master", Provenance::Master),
        ("LineColor", Provenance::StyleLine),
        ("FillForegnd", Provenance::StyleFill),
        ("Text", Provenance::StyleText),
        ("Page", Provenance::Page),
        ("Doc", Provenance::Document),
        ("Default", Provenance::Default),
    ] {
        assert_eq!(found(&resolved, name).1, expected);
    }
    let mut deleted = shape(2, vec![]);
    deleted.del = true;
    add_page(&mut package, deleted);
    assert!(
        Resolver::new(&package)
            .resolve_shape("page", 2)
            .unwrap()
            .deleted
    );
}

#[test]
fn cycles_return_diagnostics() {
    let mut cycle_package = package();
    cycle_package.style_sheets = vec![Sheet {
        id: Some(1),
        children: vec![],
        other_attrs: vec![("BasedOn".into(), "1".into())],
    }];
    let mut local = shape(1, vec![]);
    local.line_style = Some(1);
    add_page(&mut cycle_package, local);
    assert!(matches!(
        Resolver::new(&cycle_package).resolve_shape("page", 1),
        Err(ResolveError::Cycle(_))
    ));

    let mut master_loop_package = package();
    let mut local = shape(1, vec![]);
    local.master = Some(1);
    add_page(&mut master_loop_package, local);
    let mut first = shape(1, vec![]);
    first.master = Some(2);
    add_master(&mut master_loop_package, 1, first);
    let mut second = shape(2, vec![]);
    second.master = Some(1);
    add_master(&mut master_loop_package, 2, second);
    assert!(matches!(
        Resolver::new(&master_loop_package).resolve_shape("page", 1),
        Err(ResolveError::Cycle(_))
    ));

    let mut deep_package = package();
    let mut local = shape(9, vec![]);
    local.master = Some(1);
    add_page(&mut deep_package, local);
    for id in 1..=65 {
        let mut ancestor = shape(id, vec![]);
        ancestor.master = (id < 65).then_some(id + 1);
        add_master(&mut deep_package, id, ancestor);
    }
    assert!(matches!(
        Resolver::new(&deep_package).resolve_shape("page", 9),
        Err(ResolveError::Cycle(message)) if message == "maximum inheritance depth"
    ));
}

#[test]
fn text_markers_fields_and_style_rows_are_merged() {
    let mut package = package();
    package.style_sheets = vec![sheet(
        Some(1),
        vec![
            SheetChild::Section(section(
                "Character",
                vec![row(0, vec![cell("Font", "style")])],
            )),
            SheetChild::Section(section(
                "Paragraph",
                vec![row(0, vec![cell("IndFirst", "style")])],
            )),
        ],
    )];
    let mut value = shape(
        1,
        vec![ShapeChild::Text(vec![
            TextToken::CharacterRun(0),
            TextToken::Literal("hello".into()),
            TextToken::ParagraphRun(0),
            TextToken::Field(0),
        ])],
    );
    value.text_style = Some(1);
    add_page(&mut package, value.clone());
    let resolver = Resolver::new(&package);
    let empty = sheet(None, vec![]);
    let resolved = resolver.resolve_shape_in_sheet(&value, &empty).unwrap();
    let tokens = resolver
        .resolve_text_in_context(&value, &empty, &resolved)
        .unwrap();
    assert!(
        matches!(tokens[0], ResolvedTextToken::CharacterRun { ref properties, .. } if matches!(properties["Font"], Lookup::Found(_)))
    );
    assert_eq!(tokens[1], ResolvedTextToken::Literal("hello".into()));
    assert!(
        matches!(tokens[2], ResolvedTextToken::ParagraphRun { ref properties, .. } if matches!(properties["IndFirst"], Lookup::Found(_)))
    );
    assert!(matches!(
        tokens[3],
        ResolvedTextToken::Field { index: 0, .. }
    ));
}

#[test]
fn text_uses_effective_page_or_document_rows_and_master_stream() {
    let mut package = package();
    package.page_sheets.insert(
        1,
        sheet(
            Some(1),
            vec![SheetChild::Section(section(
                "Character",
                vec![row(0, vec![cell("Font", "page")])],
            ))],
        ),
    );
    package.document_sheet = Some(sheet(
        None,
        vec![SheetChild::Section(section(
            "Paragraph",
            vec![row(0, vec![cell("IndLeft", "document")])],
        ))],
    ));
    let mut local = shape(1, vec![]);
    local.master = Some(7);
    add_page(&mut package, local.clone());
    add_master(
        &mut package,
        7,
        shape(
            7,
            vec![ShapeChild::Text(vec![
                TextToken::CharacterRun(0),
                TextToken::Literal("master text".into()),
                TextToken::ParagraphRun(0),
            ])],
        ),
    );
    let contents = package.page_contents.get("page").unwrap();
    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    let tokens = Resolver::new(&package)
        .resolve_text_in_context(&local, contents, &resolved)
        .unwrap();
    assert!(
        matches!(tokens[0], ResolvedTextToken::CharacterRun { ref properties, .. } if matches!(&properties["Font"], Lookup::Found(cell) if cell.cell.value.as_deref() == Some("page")))
    );
    assert_eq!(tokens[1], ResolvedTextToken::Literal("master text".into()));
    assert!(
        matches!(tokens[2], ResolvedTextToken::ParagraphRun { ref properties, .. } if matches!(&properties["IndLeft"], Lookup::Found(cell) if cell.cell.value.as_deref() == Some("document")))
    );
}

#[test]
fn section_references_use_one_based_indices_and_user_values() {
    let mut package = package();
    let scratch = row(0, vec![cell("X", "3")]);
    let mut user = row(0, vec![cell("Value", "4")]);
    user.name = Some("ScaleFactor".into());
    let character = row(0, vec![cell("Case", "5")]);
    let shape = shape(
        1,
        vec![
            ShapeChild::Section(section("Scratch", vec![scratch])),
            ShapeChild::Section(section("User", vec![user])),
            ShapeChild::Section(section("Character", vec![character])),
        ],
    );
    add_page(&mut package, shape);
    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    assert!(matches!(
        resolved.cell("Scratch.X1"),
        Some(Lookup::Found(_))
    ));
    assert!(matches!(
        resolved.cell("User.ScaleFactor"),
        Some(Lookup::Found(_))
    ));
    assert!(matches!(
        resolved.cell("Character.Case"),
        Some(Lookup::Found(_))
    ));
}

#[test]
fn corpus_shapes_resolve_without_silent_empty_results() {
    let Some(dir) = std::env::var_os("VSDX_CORPUS_DIR") else {
        eprintln!("warning: VSDX_CORPUS_DIR is unset; skipping corpus resolver test");
        return;
    };
    let files: Vec<_> = fs::read_dir(dir)
        .unwrap()
        .filter_map(Result::ok)
        .map(|entry| entry.path())
        .filter(|path| {
            path.extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case("vsdx"))
        })
        .collect();
    assert!(files.len() >= 2, "expected both corpus files");
    for path in files {
        let package = parse_vsdx(&fs::read(&path).unwrap()).unwrap();
        let resolver = Resolver::new(&package);
        for page in &package.page_part_paths {
            let page_sheet = &package.page_contents[page];
            for value in page_sheet.shapes() {
                let resolved = resolver.resolve_shape(page, value.id).unwrap();
                for name in ["PinX", "PinY", "Width", "Height"] {
                    if let Some(lookup) = resolved.cells.get(name) {
                        assert!(
                            !matches!(lookup, Lookup::Absent),
                            "{} shape {} has silent {name}",
                            path.display(),
                            value.id
                        );
                    }
                }
            }
        }
    }
}

#[test]
fn corpus_connectivity_accounts_for_every_glue_record() {
    let Some(dir) = std::env::var_os("VSDX_CORPUS_DIR") else {
        eprintln!("warning: VSDX_CORPUS_DIR is unset; skipping corpus connectivity test");
        return;
    };
    let mut total = 0;
    let mut resolved = 0;
    let mut missing = 0;
    for file in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
        let package =
            parse_vsdx(&fs::read(std::path::Path::new(&dir).join(file)).unwrap()).unwrap();
        let resolver = Resolver::new(&package);
        for page in &package.page_part_paths {
            let connectivity = resolver.resolve_page_connectivity(page).unwrap();
            if file == "soundplan.vsdx" && page.ends_with("page1.xml") {
                let glue = connectivity.connectors[&1306]
                    .glue
                    .iter()
                    .find(|glue| {
                        glue.endpoint == ConnectorEndpoint::Begin
                            && glue.to.as_ref().is_some_and(|target| {
                                target.shape_id == 1159 && target.cell.as_deref() == Some("PinX")
                            })
                    })
                    .unwrap();
                assert_eq!(
                    glue.to
                        .as_ref()
                        .unwrap()
                        .connection_point
                        .as_ref()
                        .unwrap()
                        .position,
                    crate::ScenePoint {
                        x: 16.872_047_239_024_92,
                        y: 16.281_496_360_318_24,
                    }
                );
            }
            for connector in connectivity.connectors.values() {
                total += connector.glue.len();
                resolved += connector
                    .glue
                    .iter()
                    .filter(|glue| {
                        glue.to
                            .as_ref()
                            .and_then(|target| target.connection_point.as_ref())
                            .is_some()
                    })
                    .count();
            }
            missing += connectivity
                .diagnostics
                .iter()
                .filter(|diagnostic| {
                    matches!(
                        diagnostic,
                        ConnectivityDiagnostic::MissingConnectionPoint { .. }
                    )
                })
                .count();
        }
    }
    // Record 121 is soundplan.vsdx visio/pages/page1.xml's Connect FromSheet=1306,
    // FromCell=BeginX, ToSheet=1159, ToCell=PinX. Its target resolves via PinX/PinY, so it
    // is a valid direct-pin glue record. All 151 records now resolve, including the 30 that
    // previously did not because Connections.XN was used as the row index instead of N-1 (3954514).
    assert_eq!((total, resolved, missing), (151, 151, 0));
}

#[test]
fn nested_group_members_and_nested_master_shapes_resolve() {
    let mut package = package();
    let member = shape(2, vec![ShapeChild::Cell(cell("PinX", "member"))]);
    add_page(
        &mut package,
        shape(
            1,
            vec![ShapeChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
                member,
            )])],
        ),
    );
    let resolved = Resolver::new(&package).resolve_shape("page", 2).unwrap();
    assert_eq!(found(&resolved, "PinX"), ("member", Provenance::Local));

    let nested_master = shape(5, vec![ShapeChild::Cell(cell("PinY", "nested-master"))]);
    add_master(
        &mut package,
        1,
        shape(
            1,
            vec![ShapeChild::Shapes(vec![vsdx_parse::ShapesChild::Shape(
                nested_master,
            )])],
        ),
    );
    let mut instance = shape(3, vec![]);
    instance.master = Some(1);
    instance.master_shape = Some(5);
    add_page(&mut package, instance);
    let resolved = Resolver::new(&package).resolve_shape("page", 3).unwrap();
    assert_eq!(
        found(&resolved, "PinY"),
        ("nested-master", Provenance::MasterShape)
    );
}

#[test]
fn style_references_supplied_by_a_master_are_consulted() {
    let mut package = package();
    package.style_sheets = vec![
        sheet(Some(1), vec![SheetChild::Cell(cell("LineColor", "master"))]),
        sheet(Some(2), vec![SheetChild::Cell(cell("FillForegnd", "fill"))]),
        sheet(Some(3), vec![SheetChild::Cell(cell("Text", "text"))]),
        sheet(Some(4), vec![SheetChild::Cell(cell("LineColor", "local"))]),
    ];
    let mut master = shape(1, vec![]);
    master.line_style = Some(1);
    master.fill_style = Some(2);
    master.text_style = Some(3);
    add_master(&mut package, 1, master);

    let mut instance = shape(1, vec![]);
    instance.master = Some(1);
    add_page(&mut package, instance);
    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    assert_eq!(
        found(&resolved, "LineColor"),
        ("master", Provenance::StyleLine)
    );
    assert_eq!(
        found(&resolved, "FillForegnd"),
        ("fill", Provenance::StyleFill)
    );
    assert_eq!(found(&resolved, "Text"), ("text", Provenance::StyleText));

    let mut overriding = shape(2, vec![]);
    overriding.master = Some(1);
    overriding.line_style = Some(4);
    add_page(&mut package, overriding);
    assert_eq!(
        found(
            &Resolver::new(&package).resolve_shape("page", 2).unwrap(),
            "LineColor"
        ),
        ("local", Provenance::StyleLine)
    );
}

const GROUP_PAGE: &str = "visio/pages/page1.xml";

fn group_master_package() -> &'static VsdxPackage {
    static PACKAGE: std::sync::LazyLock<VsdxPackage> = std::sync::LazyLock::new(|| {
        parse_vsdx(include_bytes!(
            "../../vsdx-parse/tests/fixtures/group-master-shape.vsdx"
        ))
        .unwrap()
    });
    &PACKAGE
}

#[test]
fn group_subshape_master_shape_resolves_one_level_with_page_sheet() {
    let resolver = Resolver::new(group_master_package());
    let shapes = resolver.resolve_page_shapes(GROUP_PAGE).unwrap();
    assert_eq!(found(&shapes[&2], "PinX"), ("1", Provenance::MasterShape));
    assert_eq!(found(&shapes[&2], "Width"), ("2", Provenance::MasterShape));
    assert_eq!(found(&shapes[&2], "PageValue"), ("23", Provenance::Page));
    assert_eq!(found(&shapes[&2], "LocalValue"), ("11", Provenance::Local));
    assert_eq!(shapes[&2], resolver.resolve_shape(GROUP_PAGE, 2).unwrap());
}

#[test]
fn group_subshape_master_shape_resolves_two_levels_with_page_sheet() {
    let resolver = Resolver::new(group_master_package());
    let shapes = resolver.resolve_page_shapes(GROUP_PAGE).unwrap();
    assert_eq!(found(&shapes[&4], "PinY"), ("2", Provenance::MasterShape));
    assert_eq!(found(&shapes[&4], "PageValue"), ("23", Provenance::Page));
    assert_eq!(shapes[&4], resolver.resolve_shape(GROUP_PAGE, 4).unwrap());
}

#[test]
fn top_level_master_shape_is_unchanged_with_page_sheet() {
    let resolver = Resolver::new(group_master_package());
    let shapes = resolver.resolve_page_shapes(GROUP_PAGE).unwrap();
    assert_eq!(found(&shapes[&1], "PinX"), ("2", Provenance::Master));
    assert_eq!(found(&shapes[&1], "PageValue"), ("23", Provenance::Page));
}

#[test]
fn master_internal_group_lookup_is_unchanged() {
    let package = group_master_package();
    let resolver = Resolver::new(package);
    let sheet = &package.master_contents["visio/masters/master7.xml"];
    let child = sheet.shapes().next().unwrap().shapes().next().unwrap();
    let resolved = resolver.resolve_shape_in_sheet(child, sheet).unwrap();
    assert_eq!(found(&resolved, "PinX"), ("1", Provenance::MasterShape));
    let sheet = &package.master_contents["visio/masters/master8.xml"];
    let resolved = resolver.resolve_sheet(sheet).unwrap();
    assert_eq!(found(&resolved, "MasterValue"), ("41", Provenance::Local));
    assert_eq!(
        found(&resolved, "DocumentValue"),
        ("37", Provenance::Document)
    );
}

fn collect_shapes<'a>(shape: &'a Shape, out: &mut Vec<&'a Shape>) {
    out.push(shape);
    for child in shape.shapes() {
        collect_shapes(child, out);
    }
}

#[derive(Default, Debug)]
struct LookupTally {
    lost: usize,
    gained: usize,
    changed: usize,
}

fn resolved_cells(
    shape: &crate::ResolvedShape,
) -> std::collections::BTreeMap<(&str, &str, &str), &crate::ResolvedCell> {
    shape
        .cells
        .iter()
        .map(|(name, value)| (("", "", name.as_str()), value))
        .chain(shape.sections.iter().flat_map(|(section, value)| {
            value.rows.iter().flat_map(move |(row, value)| {
                value.cells.iter().map(move |(name, value)| {
                    ((section.as_str(), row.as_str(), name.as_str()), value)
                })
            })
        }))
        .filter_map(|(key, value)| match value {
            Lookup::Found(cell) => Some((key, cell)),
            _ => None,
        })
        .collect()
}

fn tally_shape(old: &crate::ResolvedShape, new: &crate::ResolvedShape, tally: &mut LookupTally) {
    let old = resolved_cells(old);
    let new = resolved_cells(new);
    tally.lost += old.keys().filter(|key| !new.contains_key(*key)).count();
    tally.gained += new.keys().filter(|key| !old.contains_key(*key)).count();
    tally.changed += old
        .iter()
        .filter(|(key, value)| new.get(*key).is_some_and(|new| new != *value))
        .count();
}

fn tally_group_lookups(package: &VsdxPackage, tally: &mut LookupTally) -> usize {
    let resolver = Resolver::new(package);
    let mut subshapes = 0;
    for (page_part, contents) in &package.page_contents {
        let inherit = package
            .page_part_ids
            .get(page_part)
            .and_then(|id| package.page_sheets.get(id))
            .unwrap_or(contents);
        let fixed = resolver.resolve_page_shapes(page_part).unwrap();
        let mut shapes = Vec::new();
        for shape in contents.shapes() {
            collect_shapes(shape, &mut shapes);
        }
        for shape in shapes {
            subshapes += usize::from(shape.master_shape.is_some() && shape.master.is_none());
            let legacy = resolver.resolve_shape_in_sheet(shape, inherit).unwrap();
            tally_shape(&legacy, &fixed[&shape.id], tally);
        }
    }
    subshapes
}

#[test]
fn group_lookup_adds_and_changes_cells_without_losing_any() {
    let mut tally = LookupTally::default();
    assert_eq!(tally_group_lookups(group_master_package(), &mut tally), 2);
    assert_eq!(tally.lost, 0);
    assert!(tally.gained > 0);
    assert!(tally.changed > 0);
}

#[test]
fn corpus_group_lookup_adds_cells_without_losing_any() {
    let Some(dir) = std::env::var_os("VSDX_CORPUS_DIR") else {
        eprintln!("skipping group lookup corpus test: VSDX_CORPUS_DIR is unset");
        return;
    };
    let files: Vec<_> = fs::read_dir(dir)
        .unwrap()
        .map(|entry| entry.unwrap().path())
        .filter(|path| {
            path.extension()
                .is_some_and(|ext| ext.eq_ignore_ascii_case("vsdx"))
        })
        .collect();
    assert!(!files.is_empty(), "expected VSDX corpus files");
    let mut tally = LookupTally::default();
    let mut subshapes = 0;
    for file in files {
        let package = parse_vsdx(&fs::read(file).unwrap()).unwrap();
        subshapes += tally_group_lookups(&package, &mut tally);
    }
    assert_eq!(tally.lost, 0);
    if subshapes == 0 {
        eprintln!(
            "skipping group lookup corpus test: no sub-shapes with MasterShape and no Master"
        );
        return;
    }
    eprintln!("VSDX corpus group lookup: subshapes={subshapes} {tally:?}");
    assert!(tally.gained > 0);
}

#[test]
fn geometry_controls_follow_section_inheritance() {
    let geometry = |formula: &str| {
        let mut geometry = section("Geometry", Vec::new());
        geometry.index = Some(1);
        geometry
            .children
            .push(SectionChild::Unknown(vsdx_parse::OpaqueXml {
                name: "Cell".into(),
                attributes: vec![("N".into(), "NoShow".into()), ("F".into(), formula.into())],
                children: Vec::new(),
            }));
        ShapeChild::Section(geometry)
    };
    for (formula, active) in [("Inh", true), ("0", false)] {
        let mut package = package();
        let mut local = shape(10, vec![geometry(formula)]);
        local.master = Some(5);
        add_page(&mut package, local);
        add_master(&mut package, 5, shape(50, vec![geometry("1")]));
        let resolved = Resolver::new(&package).resolve_shape("page", 10).unwrap();
        let section = resolved
            .sections
            .values()
            .find(|section| section.index == Some(1))
            .unwrap();
        assert!(section.unsupported_controls.is_empty());
        assert_eq!(section.controls.no_show, active);
        let realized = crate::realize_geometry(section, 1.0, 1.0);
        assert_eq!(realized.controls.no_show, active);
        assert!(realized.issues.is_empty());
    }
}

#[test]
fn geometry_unevaluable_control_reports_uncertainty_without_hiding() {
    let mut geometry = section("Geometry", Vec::new());
    geometry.index = Some(1);
    geometry
        .children
        .push(SectionChild::Unknown(vsdx_parse::OpaqueXml {
            name: "Cell".into(),
            attributes: vec![
                ("N".into(), "NoFill".into()),
                ("F".into(), "Unknown(1)".into()),
                ("V".into(), "0".into()),
            ],
            children: Vec::new(),
        }));
    let mut package = package();
    add_page(&mut package, shape(10, vec![ShapeChild::Section(geometry)]));
    let resolved = Resolver::new(&package).resolve_shape("page", 10).unwrap();
    let section = resolved
        .sections
        .values()
        .find(|section| section.index == Some(1))
        .unwrap();
    assert!(!section.controls.no_fill);
    assert_eq!(section.unsupported_controls, vec!["NoFill".to_owned()]);
    let realized = crate::realize_geometry(section, 1.0, 1.0);
    assert!(!realized.controls.no_fill);
    assert_eq!(
        realized.issues,
        vec![crate::GeometryIssue::UnsupportedSectionControl(
            "NoFill".into()
        )]
    );
}

#[test]
fn geometry_controls_evaluate_formulas_before_cached_values() {
    for control in ["NoFill", "NoLine", "NoShow"] {
        for (formula, cached, active) in [
            (Some("0"), "1", false),
            (Some("FALSE"), "1", false),
            (Some("TRUE"), "0", true),
            (Some("-2"), "0", true),
            (Some("1-1"), "1", false),
            (None, "1", true),
        ] {
            let mut attributes = vec![("N".into(), control.into()), ("V".into(), cached.into())];
            if let Some(formula) = formula {
                attributes.push(("F".into(), formula.into()));
            }
            let mut geometry = section("Geometry", Vec::new());
            geometry
                .children
                .push(SectionChild::Unknown(vsdx_parse::OpaqueXml {
                    name: "Cell".into(),
                    attributes,
                    children: vec![],
                }));
            let mut package = package();
            add_page(&mut package, shape(10, vec![ShapeChild::Section(geometry)]));
            let resolved = Resolver::new(&package).resolve_shape("page", 10).unwrap();
            let section = resolved
                .sections
                .values()
                .find(|section| section.name == "Geometry")
                .unwrap();
            assert!(section.unsupported_controls.is_empty());
            assert_eq!(
                [
                    section.controls.no_fill,
                    section.controls.no_line,
                    section.controls.no_show
                ],
                [
                    active && control == "NoFill",
                    active && control == "NoLine",
                    active && control == "NoShow"
                ]
            );
        }
    }
}

#[test]
fn geometry_editing_cells_are_not_unsupported_controls() {
    for control in ["NoSnap", "NoQuickDrag"] {
        let mut geometry = section("Geometry", Vec::new());
        geometry
            .children
            .push(SectionChild::Unknown(vsdx_parse::OpaqueXml {
                name: "Cell".into(),
                attributes: vec![("N".into(), control.into()), ("V".into(), "1".into())],
                children: vec![],
            }));
        let mut package = package();
        add_page(&mut package, shape(10, vec![ShapeChild::Section(geometry)]));
        let resolved = Resolver::new(&package).resolve_shape("page", 10).unwrap();
        let section = resolved
            .sections
            .values()
            .find(|section| section.name == "Geometry")
            .unwrap();
        assert!(section.unsupported_controls.is_empty(), "{control}");
        assert_eq!(
            section.controls,
            crate::GeometrySectionControls::default(),
            "{control}"
        );
    }
}

#[test]
fn geometry_unknown_controls_report_issues_after_inheritance() {
    let mut geometry = section("Geometry", Vec::new());
    geometry
        .children
        .push(SectionChild::Unknown(vsdx_parse::OpaqueXml {
            name: "Cell".into(),
            attributes: vec![
                ("N".into(), "NoSuchControl".into()),
                ("V".into(), "1".into()),
            ],
            children: vec![],
        }));
    let mut package = package();
    let mut local = shape(10, vec![]);
    local.master = Some(5);
    add_page(&mut package, local);
    add_master(
        &mut package,
        5,
        shape(50, vec![ShapeChild::Section(geometry)]),
    );
    let resolved = Resolver::new(&package).resolve_shape("page", 10).unwrap();
    let section = resolved
        .sections
        .values()
        .find(|section| section.name == "Geometry")
        .unwrap();
    assert_eq!(
        crate::realize_geometry(section, 1.0, 1.0).issues,
        vec![crate::GeometryIssue::UnsupportedSectionControl(
            "NoSuchControl".into()
        )]
    );
}
