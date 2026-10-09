use betteroffice_vsdx::{CellLocator, CellRow, Diagram};
use ooxml_opc::{rezip_parts, unzip_parts};
use vsdx_edit::{DiagramSession, EditCtx};

const PAGE: &str = "\
<PageContents><Shapes><Shape ID='1'>\
<Section N='Property'>\
<Row N='Device'><Cell N='Label' V='Device name'/><Cell N='Value' V='Old' F='\"Old\"'/><Cell N='Type' V='0'/><Cell N='Invisible' V='0'/><Cell N='SortKey' V=''/></Row>\
<Row N='Hidden'><Cell N='Label' V='Hidden'/><Cell N='Value' V='x'/><Cell N='Type' V='0'/><Cell N='Invisible' V='1'/></Row>\
</Section></Shape></Shapes></PageContents>";

#[test]
fn shape_data_lists_visible_properties() {
    let diagram = Diagram::open(&package_with_page(PAGE)).unwrap();
    let page = diagram.pages().next().unwrap();
    let shape = page.shapes().next().unwrap();
    let properties = shape.shape_data().unwrap();
    assert_eq!(properties.len(), 2);
    assert_eq!(properties[0].row, "Device");
    assert_eq!(properties[0].label, "Device name");
    assert!(!properties[0].invisible);
    assert_eq!(
        properties[0].value.value.as_deref(),
        Some("Old"),
        "{properties:?}"
    );
    assert_eq!(properties[1].row, "Hidden");
    assert!(properties[1].invisible);
}

#[test]
fn property_value_edits_round_trip_through_save() {
    let bytes = package_with_page(PAGE);
    let session = DiagramSession::open(&bytes, 7).unwrap();
    let snapshot = session.snapshot().unwrap();
    let page_id = snapshot.pages[0].id.clone();
    let shape_id = snapshot.pages[0].shapes[0].id.clone();
    session
        .set_cell_formula_at(
            &EditCtx::local("test"),
            &page_id,
            &shape_id,
            CellLocator {
                sheet: vsdx_cell_sheet(&page_id),
                shape_id: None,
                section: Some("Property".to_owned()),
                section_index: None,
                row: Some(CellRow::Name("Device".to_owned())),
                cell_name: "Value".to_owned(),
            },
            "\"New\"",
        )
        .unwrap();
    let diagram = Diagram::open(&bytes).unwrap();
    let saved = diagram
        .save_cell_edits(&session.semantic_cell_edits().unwrap())
        .unwrap();
    let reopened = Diagram::open(&saved).unwrap();
    let page = reopened.pages().next().unwrap();
    let shape = page.shapes().next().unwrap();
    let properties = shape.shape_data().unwrap();
    let device = properties
        .iter()
        .find(|property| property.row == "Device")
        .unwrap();
    assert_eq!(
        device.value.formula.as_deref(),
        Some("\"New\""),
        "{properties:?}"
    );
}

fn vsdx_cell_sheet(page_id: &str) -> betteroffice_vsdx::CellSheet {
    let id = page_id
        .trim_start_matches("page:")
        .parse()
        .unwrap_or_default();
    betteroffice_vsdx::CellSheet::Page(id)
}

fn package_with_page(page_xml: &str) -> Vec<u8> {
    let fixture = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
    let package = vsdx_parse::parse_vsdx(fixture).unwrap();
    let page_path = package.page_part_paths[0].clone();
    let mut archive = unzip_parts(fixture).unwrap();
    archive
        .iter_mut()
        .find(|(path, _)| path == &page_path)
        .unwrap()
        .1 = page_xml.as_bytes().to_vec();
    rezip_parts(&archive).unwrap()
}
