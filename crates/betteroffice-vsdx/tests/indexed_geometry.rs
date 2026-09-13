use betteroffice_vsdx::{
    CellLocator, CellRow, CellSheet, Diagram, MutationGesture, SemanticCellEdit,
};
use vsdx_edit::{DiagramSession, EditCtx, ShapeDraft};
use vsdx_resolve::{Lookup, Resolver};

const SOURCE: &[u8] = include_bytes!("../../vsdx-parse/tests/fixtures/indexed-geometry.vsdx");

fn locator(index: u32) -> CellLocator {
    CellLocator {
        sheet: CellSheet::Page(1),
        shape_id: Some(42),
        section: Some("Geometry".into()),
        section_index: Some(index),
        row: Some(CellRow::Index(1)),
        cell_name: "X".into(),
    }
}

fn assert_geometry(bytes: &[u8], first: &str, second: &str) {
    let package = vsdx_parse::parse_vsdx(bytes).unwrap();
    let resolved = Resolver::new(&package)
        .resolve_shape("visio/pages/page1.xml", 42)
        .unwrap();
    assert_eq!(resolved.sections.len(), 2);
    for (reference, expected) in [("Geometry1.X1", first), ("Geometry2.X1", second)] {
        let Some(Lookup::Found(cell)) = resolved.cell(reference) else {
            panic!("{reference}")
        };
        assert_eq!(cell.cell.formula.as_deref(), Some(expected), "{reference}");
    }
}

#[test]
fn indexed_geometry_sections_resolve_and_edit_independently() {
    assert_geometry(SOURCE, "1", "2");
    let diagram = Diagram::open(SOURCE).unwrap();
    let saved = diagram
        .save_cell_edits(&[SemanticCellEdit {
            locator: locator(1),
            gesture: MutationGesture::CellEdit,
            formula: Some("7".into()),
            value: None,
        }])
        .unwrap();
    assert_geometry(&saved, "1", "7");
}

#[test]
fn indexed_geometry_survives_collaboration_and_save() {
    let session = DiagramSession::open(SOURCE, 7).unwrap();
    let peer = DiagramSession::open(SOURCE, 8).unwrap();
    let before = session.snapshot().unwrap();
    let shape = &before.pages[0].shapes[0];
    assert_eq!(
        shape
            .cells
            .iter()
            .filter(|cell| cell.locator.section.is_some())
            .count(),
        8
    );
    session
        .set_cell_formula_at(
            &EditCtx::local("test"),
            "page:1",
            &shape.id,
            locator(1),
            "7",
        )
        .unwrap();
    peer.apply_update_v1(
        &session
            .encode_diff_v1(&peer.encode_state_vector_v1())
            .unwrap(),
    )
    .unwrap();
    assert_eq!(session.snapshot().unwrap(), peer.snapshot().unwrap());
    assert_geometry(&session.save().unwrap(), "1", "7");
    assert_geometry(&peer.save().unwrap(), "1", "7");
}

#[test]
fn indexed_geometry_drafts_keep_distinct_sections_when_added() {
    let session = DiagramSession::open(SOURCE, 7).unwrap();
    let snapshot = session.snapshot().unwrap();
    let draft = ShapeDraft {
        name: Some("copy".into()),
        cells: snapshot.pages[0].shapes[0]
            .cells
            .iter()
            .cloned()
            .map(|mut cell| {
                cell.value = None;
                cell
            })
            .collect(),
    };
    session
        .add_shape(&EditCtx::local("test"), "page:1", &draft)
        .unwrap();
    let saved = session.save().unwrap();
    let package = vsdx_parse::parse_vsdx(&saved).unwrap();
    let added = package.page_contents["visio/pages/page1.xml"]
        .shapes()
        .last()
        .unwrap();
    assert_eq!(
        added
            .sections()
            .map(|section| section.index)
            .collect::<Vec<_>>(),
        vec![Some(0), Some(1)]
    );
    let reopened = DiagramSession::open(&saved, 9).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    let added = snapshot.pages[0].shapes.last().unwrap();
    assert_eq!(
        added
            .cells
            .iter()
            .filter(|cell| cell.locator.section.is_some())
            .count(),
        8
    );
}

#[test]
fn concurrent_layer_reorders_and_deletion_remain_serializable() {
    let seed = DiagramSession::open(SOURCE, 7).unwrap();
    let mut cells = seed.snapshot().unwrap().pages[0].shapes[0].cells.clone();
    for cell in &mut cells {
        cell.value = None;
    }
    seed.add_shape(
        &EditCtx::local("seed"),
        "page:1",
        &ShapeDraft {
            name: Some("copy".into()),
            cells,
        },
    )
    .unwrap();
    let source = seed.save().unwrap();
    for delete in [false, true] {
        let left = DiagramSession::open(&source, 8).unwrap();
        let right = DiagramSession::open(&source, 9).unwrap();
        let id = left.snapshot().unwrap().pages[0].shapes[0].id.clone();
        left.reorder_shape(&EditCtx::local("left"), "page:1", &id, 1)
            .unwrap();
        if delete {
            right
                .delete_shape(&EditCtx::local("right"), "page:1", &id)
                .unwrap();
        } else {
            right
                .reorder_shape(&EditCtx::local("right"), "page:1", &id, 1)
                .unwrap();
        }
        let left_update = left
            .encode_diff_v1(&right.encode_state_vector_v1())
            .unwrap();
        let right_update = right
            .encode_diff_v1(&left.encode_state_vector_v1())
            .unwrap();
        left.apply_update_v1(&right_update).unwrap();
        right.apply_update_v1(&left_update).unwrap();
        assert_eq!(left.snapshot().unwrap(), right.snapshot().unwrap());
        let left_saved = vsdx_parse::parse_vsdx(&left.save().unwrap()).unwrap();
        let right_saved = vsdx_parse::parse_vsdx(&right.save().unwrap()).unwrap();
        assert_eq!(left_saved.page_contents, right_saved.page_contents);
        assert_eq!(
            left_saved.page_contents["visio/pages/page1.xml"]
                .shapes()
                .count(),
            if delete { 1 } else { 2 }
        );
    }
}
