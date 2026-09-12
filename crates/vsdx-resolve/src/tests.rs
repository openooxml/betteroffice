use std::fs;

use vsdx_parse::{
    Cell, Row, RowChild, Section, SectionChild, Shape, ShapeChild, Sheet, SheetChild, TextToken,
    VsdxPackage, parse_vsdx,
};

use crate::{Lookup, Provenance, ResolveError, ResolvedTextToken, Resolver};

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

#[test]
fn master_inheritance_walks_deeply_and_local_overrides() {
    let mut package = package();
    let mut local = shape(10, vec![]);
    local.master = Some(1);
    local.master_shape = Some(1);
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
    direct.master_shape = Some(3);
    add_page(&mut package, direct);
    assert_eq!(
        found(
            &Resolver::new(&package).resolve_shape("page", 12).unwrap(),
            "PinY"
        ),
        ("master-shape", Provenance::MasterShape)
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
fn deletions_block_inheritance_while_absence_inherits() {
    let mut package = package();
    let mut master = shape(
        1,
        vec![
            ShapeChild::Cell(cell("PinX", "master")),
            ShapeChild::Section(section("Geometry", vec![row(0, vec![cell("X", "1")])])),
        ],
    );
    master.master = None;
    add_master(&mut package, 1, master);
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
