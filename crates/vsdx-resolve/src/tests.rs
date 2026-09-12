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
                                    1,
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
                                    1,
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
    let geometry = crate::realize_geometry(section);
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
    let geometry = crate::realize_geometry(section);
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
    let tokens = Resolver::new(&package)
        .resolve_text(&value, &sheet(None, vec![]))
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
    let page = package.page_sheets.get(&1).unwrap();
    let resolved = Resolver::new(&package).resolve_shape("page", 1).unwrap();
    let tokens = Resolver::new(&package)
        .resolve_text_in_context(&local, page, &resolved)
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
    // FromCell=BeginX, ToSheet=1159, ToCell=PinX. Its target's PinX/PinY resolve to
    // (16.87204723902492, 16.28149636031824), so it is a valid direct-pin glue record;
    // the remaining 30 records lack Connection rows.
    assert_eq!((total, resolved, missing), (151, 121, 30));
}
