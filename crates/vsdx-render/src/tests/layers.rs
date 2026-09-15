use super::*;

fn layer_row(index: u32, name: &str, visible: &str) -> Row {
    row(
        index,
        "Layer",
        vec![
            cell("Name", name),
            cell("Color", "255"),
            cell("Status", "0"),
            cell("Visible", visible),
            cell("Print", "1"),
            cell("Active", "0"),
            cell("Lock", "0"),
        ],
    )
}

fn layer_package(shapes: Vec<Shape>, rows: Vec<Row>) -> VsdxPackage {
    let mut package = package(shapes);
    package.page_sheets.insert(
        1,
        Sheet {
            id: None,
            children: vec![
                SheetChild::Cell(cell("PageWidth", "10")),
                SheetChild::Cell(cell("PageHeight", "8")),
                SheetChild::Section(Section {
                    name: "Layer".into(),
                    index: None,
                    del: false,
                    children: rows.into_iter().map(SectionChild::Row).collect(),
                    other_attrs: vec![],
                }),
            ],
            other_attrs: vec![],
        },
    );
    package
}

fn membered(id: u32, member: &str) -> Shape {
    let mut shape = shape(id, f64::from(id), 1.0);
    shape
        .children
        .push(ShapeChild::Cell(cell("LayerMember", member)));
    shape
}

fn layer_package_two_layers(shapes: Vec<Shape>) -> VsdxPackage {
    layer_package(
        shapes,
        vec![layer_row(0, "Trussing", "1"), layer_row(1, "Lighting", "0")],
    )
}

fn painted_ids(list: &VsdxDisplayList) -> Vec<String> {
    let mut ids = Vec::new();
    collect_shape_ids(&list.primitives, &mut ids);
    ids.sort();
    ids
}

fn collect_shape_ids(primitives: &[Primitive], ids: &mut Vec<String>) {
    for primitive in primitives {
        match primitive {
            Primitive::Shape { id, .. } => ids.push(id.clone()),
            Primitive::Group { primitives, .. } => collect_shape_ids(primitives, ids),
            Primitive::Image { .. } | Primitive::TextBox { .. } | Primitive::Placeholder { .. } => {
            }
        }
    }
}

#[test]
fn hidden_layer_shapes_do_not_paint() {
    let package = layer_package_two_layers(vec![
        membered(1, "1"),
        membered(2, "0"),
        membered(3, ""),
        membered(4, "0;1"),
        shape(5, 5.0, 1.0),
    ]);
    let list = Renderer::default().layout_page(&package, "page").unwrap();
    assert_eq!(painted_ids(&list), ["page:2", "page:3", "page:4", "page:5"]);
}

#[test]
fn layer_visibility_multiplies_with_no_show() {
    let mut hidden_by_cell = membered(1, "0");
    with_cell(&mut hidden_by_cell, "NoShow", "1");
    let mut hidden_by_both = membered(2, "1");
    with_cell(&mut hidden_by_both, "NoShow", "1");
    let package = layer_package_two_layers(vec![
        membered(3, "0"),
        hidden_by_cell,
        membered(4, "1"),
        hidden_by_both,
    ]);
    let list = Renderer::default().layout_page(&package, "page").unwrap();
    assert_eq!(painted_ids(&list), ["page:3"]);
}

#[test]
fn hidden_group_hides_its_children() {
    let mut hidden = group(1, 1.0, 1.0, vec![membered(2, "0")]);
    hidden
        .children
        .push(ShapeChild::Cell(cell("LayerMember", "1")));
    let package = layer_package_two_layers(vec![hidden, membered(3, "0")]);
    let list = Renderer::default().layout_page(&package, "page").unwrap();
    assert_eq!(painted_ids(&list), ["page:3"]);
}

#[test]
fn layer_override_shadows_the_page_sheet() {
    let package = layer_package_two_layers(vec![membered(1, "1"), membered(2, "0")]);
    let mut renderer = Renderer::default();
    assert_eq!(
        painted_ids(&renderer.layout_page(&package, "page").unwrap()),
        ["page:2"]
    );
    renderer.set_layer_override("page", 1, true);
    renderer.set_layer_override("page", 0, false);
    assert_eq!(
        painted_ids(&renderer.layout_page(&package, "page").unwrap()),
        ["page:1"]
    );
    assert_eq!(
        renderer.effective_page_layers(&package, "page"),
        vec![
            vsdx_resolve::PageLayer {
                index: 0,
                name: "Trussing".into(),
                visible: false,
                print: true,
                lock: false,
                active: false,
                color: "255".into(),
                status: "0".into(),
            },
            vsdx_resolve::PageLayer {
                index: 1,
                name: "Lighting".into(),
                visible: true,
                print: true,
                lock: false,
                active: false,
                color: "255".into(),
                status: "0".into(),
            },
        ]
    );
    renderer.clear_layer_overrides();
    assert_eq!(
        painted_ids(&renderer.layout_page(&package, "page").unwrap()),
        ["page:2"]
    );
}
