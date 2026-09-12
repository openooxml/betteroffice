use betteroffice_vsdx::Diagram;

#[test]
fn opens_pages_and_inspects_shapes() {
    let diagram = Diagram::open(include_bytes!(
        "../../vsdx-parse/tests/fixtures/foundation.vsdx"
    ))
    .unwrap();
    let page = diagram.pages().next().unwrap();
    let shape = page.shapes().next().unwrap();
    assert!(!shape.resolved().unwrap().cells.is_empty());
}
