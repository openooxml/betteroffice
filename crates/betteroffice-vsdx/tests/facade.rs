use betteroffice_vsdx::{
    CellLocator, CellRow, CellSheet, Diagram, MutationGesture, SemanticCellEdit, StructuralEdit,
};
use ooxml_opc::{rezip_parts, unzip_parts};
use std::collections::BTreeMap;
use vsdx_edit::{CellSnapshot, DiagramSession, EditCtx, ShapeDraft};

fn diagram_with_page(xml: &str) -> (Vec<u8>, Diagram, u32) {
    let source = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
    let package = vsdx_parse::parse_vsdx(source).unwrap();
    let path = package.page_part_paths[0].clone();
    let page_id = package.page_part_ids[&path];
    let mut parts = unzip_parts(source).unwrap();
    parts
        .iter_mut()
        .find(|(candidate, _)| candidate == &path)
        .unwrap()
        .1 = xml.as_bytes().to_vec();
    let source = rezip_parts(&parts).unwrap();
    let diagram = Diagram::open(&source).unwrap();
    (source, diagram, page_id)
}

fn diagram_with_parts(replacements: &[(&str, &str)]) -> (Vec<u8>, Diagram, u32) {
    let source = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
    let package = vsdx_parse::parse_vsdx(source).unwrap();
    let page_id = *package.page_part_ids.values().next().unwrap();
    let mut parts = unzip_parts(source).unwrap();
    for (path, xml) in replacements {
        parts
            .iter_mut()
            .find(|(candidate, _)| candidate == path)
            .unwrap()
            .1 = xml.as_bytes().to_vec();
    }
    let source = rezip_parts(&parts).unwrap();
    let diagram = Diagram::open(&source).unwrap();
    (source, diagram, page_id)
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    unzip_parts(bytes).unwrap().into_iter().collect()
}

fn assert_only_part_changed(before: &[u8], after: &[u8], changed: &str) {
    let before = parts(before);
    let after = parts(after);
    assert_eq!(before.len(), after.len());
    for (path, bytes) in before {
        if path == changed {
            assert_ne!(after[&path], bytes, "{path} must change");
        } else {
            assert_eq!(after[&path], bytes, "{path}");
        }
    }
}

#[test]
fn saves_crdt_session_edits_for_every_corpus_file() {
    let Ok(corpus) = std::env::var("VSDX_CORPUS_DIR") else {
        eprintln!("warning: VSDX_CORPUS_DIR is unset; skipping corpus save bridge test");
        return;
    };
    for file in ["lichtsysteme.vsdx", "soundplan.vsdx"] {
        let source = std::fs::read(std::path::Path::new(&corpus).join(file)).unwrap();
        let session = DiagramSession::open(&source, 7).unwrap();
        let snapshot = session.snapshot().unwrap();
        let page = snapshot.pages.first().unwrap();
        let shape = page
            .shapes
            .iter()
            .find(|shape| shape.cells.iter().any(|cell| cell.formula.is_some()))
            .unwrap();
        let cell = shape
            .cells
            .iter()
            .find(|cell| cell.formula.is_some())
            .unwrap();
        session
            .set_cell_formula_at(
                &EditCtx::local("test"),
                &page.id,
                &shape.id,
                cell.locator.clone(),
                "42",
            )
            .unwrap();
        let saved = Diagram::open(&source)
            .unwrap()
            .save_session(&session)
            .unwrap();
        let reopened = DiagramSession::open(&saved, 8).unwrap();
        let reopened_snapshot = reopened.snapshot().unwrap();
        let changed = reopened_snapshot
            .pages
            .iter()
            .find(|candidate| candidate.id == page.id)
            .unwrap()
            .shapes
            .iter()
            .find(|candidate| candidate.id == shape.id)
            .unwrap()
            .cells
            .iter()
            .find(|candidate| candidate.locator == cell.locator)
            .unwrap();
        assert_eq!(changed.formula.as_deref(), Some("42"), "{file}");
        assert_only_part_changed(&source, &saved, &page.source_part_path);
    }
}

fn edit(
    page_id: u32,
    shape_id: u32,
    name: &str,
    formula: &str,
    gesture: MutationGesture,
) -> SemanticCellEdit {
    SemanticCellEdit {
        locator: CellLocator {
            sheet: CellSheet::Page(page_id),
            shape_id: Some(shape_id),
            section: None,
            row: None,
            cell_name: name.to_owned(),
        },
        gesture,
        formula: Some(formula.to_owned()),
        value: None,
    }
}

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

#[test]
fn guarded_edit_refuses_without_writing() {
    let (_, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' F='GUARD(1)' V='1'/></Shape></Shapes></PageContents>",
    );
    let page_path = diagram.package().page_part_paths[0].clone();
    let before = diagram.package().part_bytes(&page_path).unwrap().to_vec();
    assert!(
        diagram
            .save_cell_edits(&[edit(page_id, 1, "Width", "2", MutationGesture::ResizeWidth)])
            .is_err()
    );
    assert_eq!(diagram.package().part_bytes(&page_path).unwrap(), before);
}

#[test]
fn setatref_redirects_to_the_referenced_cell() {
    let (_, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' F='SETATREF(Target)' V='1'/><Cell N='Target' V='1'/></Shape></Shapes></PageContents>",
    );
    let saved = diagram
        .save_cell_edits(&[edit(page_id, 1, "Width", "7", MutationGesture::ResizeWidth)])
        .unwrap();
    let reopened = Diagram::open(&saved).unwrap();
    let page = reopened.pages().next().unwrap();
    let shape = page.shapes().next().unwrap();
    assert_eq!(
        shape
            .model()
            .cells()
            .find(|cell| cell.name == "Width")
            .unwrap()
            .formula
            .as_deref(),
        Some("SETATREF(Target)")
    );
    let target = shape
        .model()
        .cells()
        .find(|cell| cell.name == "Target")
        .unwrap();
    assert_eq!(target.formula.as_deref(), Some("7"));
    assert_eq!(target.value.as_deref(), Some("7"));
}

#[test]
fn setatref_resolves_a_sheet_reference() {
    let (_, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' F='SETATREF(Sheet.2!Target)' V='1'/></Shape><Shape ID='2'><Cell N='Target' V='3'/></Shape></Shapes></PageContents>",
    );
    let saved = diagram
        .save_cell_edits(&[edit(page_id, 1, "Width", "9", MutationGesture::ResizeWidth)])
        .unwrap();
    let reopened = Diagram::open(&saved).unwrap();
    let page = reopened.pages().next().unwrap();
    let target = page.shapes().find(|shape| shape.model().id == 2).unwrap();
    assert_eq!(
        target.model().cells().next().unwrap().formula.as_deref(),
        Some("9")
    );
}

#[test]
fn setatref_bypasses_are_rejected_without_writing_the_source() {
    for formula in [
        "SETATREF(Target)+1",
        "SETATREF(Target)+SETATREF(Other)",
        "IF(1,SETATREF(Target),0)",
    ] {
        let (_, diagram, page_id) = diagram_with_page(&format!(
            "<PageContents><Shapes><Shape ID='1'><Cell N='Width' F='{formula}' V='1'/><Cell N='Target' V='1'/><Cell N='Other' V='1'/></Shape></Shapes></PageContents>"
        ));
        let page_path = diagram.package().page_part_paths[0].clone();
        let before = diagram.package().part_bytes(&page_path).unwrap().to_vec();
        assert!(
            diagram
                .save_cell_edits(&[edit(page_id, 1, "Width", "7", MutationGesture::ResizeWidth)])
                .is_err()
        );
        assert_eq!(diagram.package().part_bytes(&page_path).unwrap(), before);
    }
}

#[test]
fn redirects_reapply_guards_and_reject_missing_targets() {
    for formula in ["SETATREF(Target)", "SETATREF(Missing)"] {
        let target = if formula.contains("Missing") {
            ""
        } else {
            "<Cell N='Target' F='GUARD(1)' V='1'/>"
        };
        let (_, diagram, page_id) = diagram_with_page(&format!(
            "<PageContents><Shapes><Shape ID='1'><Cell N='Width' F='{formula}' V='1'/>{target}</Shape></Shapes></PageContents>"
        ));
        assert!(
            diagram
                .save_cell_edits(&[edit(page_id, 1, "Width", "7", MutationGesture::ResizeWidth)])
                .is_err()
        );
    }
}

#[test]
fn locks_refuse_only_their_matching_gesture() {
    let (_, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='LockWidth' V='1'/><Cell N='LockMoveX' V='1'/><Cell N='Width' V='1'/><Cell N='PinX' V='1'/><Cell N='PinY' V='1'/></Shape></Shapes></PageContents>",
    );
    assert!(
        diagram
            .save_cell_edits(&[edit(page_id, 1, "Width", "2", MutationGesture::ResizeWidth)])
            .is_err()
    );
    assert!(
        diagram
            .save_cell_edits(&[edit(page_id, 1, "PinX", "2", MutationGesture::MoveX)])
            .is_err()
    );
    assert!(
        diagram
            .save_cell_edits(&[edit(page_id, 1, "PinY", "2", MutationGesture::MoveY)])
            .is_ok()
    );
}

#[test]
fn lock_literals_are_evaluated_instead_of_treated_as_nonzero_text() {
    for (formula, expected) in [("FALSE", true), ("0", true), ("TRUE", false), ("1", false)] {
        let (_, diagram, page_id) = diagram_with_page(&format!(
            "<PageContents><Shapes><Shape ID='1'><Cell N='LockWidth' F='{formula}' V='stale'/><Cell N='Width' V='1'/></Shape></Shapes></PageContents>"
        ));
        assert_eq!(
            diagram
                .save_cell_edits(&[edit(page_id, 1, "Width", "2", MutationGesture::ResizeWidth)])
                .is_ok(),
            expected
        );
    }
}

#[test]
fn unevaluable_lock_formula_is_unsupported() {
    let (_, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='LockWidth' F='Unknown(1)' V='0'/><Cell N='Width' V='1'/></Shape></Shapes></PageContents>",
    );
    assert!(
        diagram
            .save_cell_edits(&[edit(page_id, 1, "Width", "2", MutationGesture::ResizeWidth)])
            .is_err()
    );
}

#[test]
fn a_refused_edit_aborts_the_entire_set() {
    let (_, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/><Cell N='Height' F='GUARD(1)' V='1'/></Shape></Shapes></PageContents>",
    );
    let page_path = diagram.package().page_part_paths[0].clone();
    let before = diagram.package().part_bytes(&page_path).unwrap().to_vec();
    assert!(
        diagram
            .save_cell_edits(&[
                edit(page_id, 1, "Width", "2", MutationGesture::ResizeWidth),
                edit(page_id, 1, "Height", "2", MutationGesture::ResizeHeight),
            ])
            .is_err()
    );
    assert_eq!(diagram.package().part_bytes(&page_path).unwrap(), before);
}

#[test]
fn writing_a_master_only_cell_creates_a_local_formula_override() {
    let (source, diagram, page_id) = diagram_with_parts(&[
        (
            "visio/pages/page1.xml",
            "<PageContents><Shapes><Shape ID='1' Master='1' MasterShape='10'/></Shapes></PageContents>",
        ),
        (
            "visio/masters/master1.xml",
            "<MasterContents><Shapes><Shape ID='10'><Cell N='LineWeight' V='1'/></Shape></Shapes></MasterContents>",
        ),
    ]);
    let saved = diagram
        .save_cell_edits(&[edit(
            page_id,
            1,
            "LineWeight",
            "4",
            MutationGesture::CellEdit,
        )])
        .unwrap();
    assert_only_part_changed(&source, &saved, "visio/pages/page1.xml");
    let before = parts(&source);
    let after = parts(&saved);
    let before_page = std::str::from_utf8(&before["visio/pages/page1.xml"]).unwrap();
    assert_eq!(
        after["visio/pages/page1.xml"],
        before_page
            .replacen("/>", "><Cell N='LineWeight' F='4' V='4'/></Shape>", 1,)
            .as_bytes()
    );
    let reopened = Diagram::open(&saved).unwrap();
    let page = reopened.pages().next().unwrap();
    let shape = page.shapes().next().unwrap();
    let cell = shape.resolved().unwrap().cells["LineWeight"].clone();
    assert!(
        matches!(cell, vsdx_resolve::Lookup::Found(ref cell) if cell.provenance == vsdx_resolve::Provenance::Local && cell.cell.formula.as_deref() == Some("4"))
    );
}

#[test]
fn writing_a_style_only_cell_creates_a_local_formula_override() {
    let (source, diagram, page_id) = diagram_with_parts(&[
        (
            "visio/document.xml",
            "<VisioDocument><StyleSheets><StyleSheet ID='2'><Cell N='LineWeight' V='1'/></StyleSheet></StyleSheets><DocumentSheet><Cell N='PageWidth' V='8.5'/></DocumentSheet></VisioDocument>",
        ),
        (
            "visio/pages/page1.xml",
            "<PageContents><Shapes><Shape ID='1' LineStyle='2'/></Shapes></PageContents>",
        ),
    ]);
    let saved = diagram
        .save_cell_edits(&[edit(
            page_id,
            1,
            "LineWeight",
            "4",
            MutationGesture::CellEdit,
        )])
        .unwrap();
    assert_only_part_changed(&source, &saved, "visio/pages/page1.xml");
    let reopened = Diagram::open(&saved).unwrap();
    let page = reopened.pages().next().unwrap();
    let shape = page.shapes().next().unwrap();
    assert!(matches!(
        shape.resolved().unwrap().cells["LineWeight"],
        vsdx_resolve::Lookup::Found(ref cell)
            if cell.provenance == vsdx_resolve::Provenance::Local
                && cell.cell.formula.as_deref() == Some("4")
    ));
}

#[test]
fn setatref_redirects_to_page_and_document_sheets_without_touching_the_source() {
    for (reference, target_part) in [
        ("ThePage!PageWidth", "visio/pages/pages.xml"),
        ("TheDoc!PageWidth", "visio/document.xml"),
    ] {
        let (source, diagram, page_id) = diagram_with_page(&format!(
            "<PageContents><Shapes><Shape ID='1'><Cell N='Width' F='SETATREF({reference})' V='1'/></Shape></Shapes></PageContents>"
        ));
        let saved = diagram
            .save_cell_edits(&[edit(page_id, 1, "Width", "7", MutationGesture::ResizeWidth)])
            .unwrap();
        assert_only_part_changed(&source, &saved, target_part);
        let after = parts(&saved);
        let target = std::str::from_utf8(&after[target_part]).unwrap();
        assert!(target.contains("F='7'") && target.contains("V='7'"));
    }
}

#[test]
fn mixed_mutation_failures_never_produce_a_partial_package() {
    for (failing, formula) in [
        ("<Cell N='PinX' F='GUARD(1)' V='1'/>", "4"),
        ("<Cell N='PinX' V='1'/>", "Unknown(1)"),
        ("<Cell N='PinX' F='SETATREF(Target,1)' V='1'/>", "4"),
    ] {
        let (source, diagram, page_id) = diagram_with_page(&format!(
            "<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/><Cell N='Height' F='SETATREF(Target)' V='1'/><Cell N='Target' V='1'/>{failing}</Shape></Shapes></PageContents>"
        ));
        assert!(
            diagram
                .save_cell_edits(&[
                    edit(page_id, 1, "Width", "2", MutationGesture::ResizeWidth),
                    edit(page_id, 1, "Height", "3", MutationGesture::ResizeHeight),
                    edit(page_id, 1, "PinX", formula, MutationGesture::MoveX),
                ])
                .is_err()
        );
        let original = parts(&source);
        for (path, bytes) in original {
            assert_eq!(
                diagram.package().part_bytes(&path).unwrap(),
                bytes,
                "{path}"
            );
        }
    }
}

#[test]
fn saves_a_semantic_cell_edit_without_source_spans() {
    let source = include_bytes!("../../vsdx-parse/tests/fixtures/foundation.vsdx");
    let diagram = Diagram::open(source).unwrap();
    let page_id = *diagram.package().page_part_ids.values().next().unwrap();
    let saved = diagram
        .save_cell_edits(&[SemanticCellEdit {
            locator: CellLocator {
                sheet: CellSheet::Page(page_id),
                shape_id: Some(1),
                section: None,
                row: None,
                cell_name: "Both".to_owned(),
            },
            gesture: MutationGesture::CellEdit,
            formula: Some("42".to_owned()),
            value: None,
        }])
        .unwrap();
    let saved = Diagram::open(&saved).unwrap();
    let page = saved.pages().next().unwrap();
    assert!(page.shapes().any(|shape| {
        shape.model().cells().any(|cell| {
            cell.name == "Both"
                && cell.formula.as_deref() == Some("42")
                && cell.value.as_deref() == Some("42")
        })
    }));
}

#[test]
fn mixed_cell_and_structural_edits_are_atomic() {
    let (_, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/></Shape></Shapes></PageContents>",
    );
    let saved = diagram
        .save_edits(
            &[edit(page_id, 1, "Width", "2", MutationGesture::ResizeWidth)],
            &[StructuralEdit::AddShape {
                page_id,
                shape_xml: b"<Shape><Cell N='Width' V='3'/></Shape>".to_vec(),
            }],
        )
        .unwrap();
    let reopened = Diagram::open(&saved).unwrap();
    assert_eq!(reopened.pages().next().unwrap().shapes().count(), 2);
    let page = reopened.pages().next().unwrap();
    let first = page.shapes().next().unwrap();
    assert!(first.model().cells().any(|cell| {
        cell.name == "Width"
            && cell.formula.as_deref() == Some("2")
            && cell.value.as_deref() == Some("2")
    }));

    let (locked_source, locked, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/><Cell N='LockDelete' F='1' V='1'/></Shape></Shapes></PageContents>",
    );
    assert!(
        locked
            .save_edits(
                &[edit(page_id, 1, "Width", "2", MutationGesture::ResizeWidth)],
                &[StructuralEdit::DeleteShape {
                    page_id,
                    shape_id: 1,
                }],
            )
            .is_err()
    );
    assert_eq!(
        locked
            .package()
            .part_bytes("visio/pages/page1.xml")
            .unwrap(),
        parts(&locked_source)["visio/pages/page1.xml"]
    );
}

#[test]
fn mixed_edits_reject_invalid_structure_without_changing_the_source() {
    let (source, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/></Shape><Shape ID='1'/></Shapes></PageContents>",
    );
    assert!(
        diagram
            .save_edits(
                &[edit(page_id, 1, "Width", "2", MutationGesture::ResizeWidth)],
                &[StructuralEdit::ReorderShape {
                    page_id,
                    shape_id: 1,
                    before_shape_id: None,
                }],
            )
            .is_err()
    );
    for (path, bytes) in parts(&source) {
        assert_eq!(
            diagram.package().part_bytes(&path).unwrap(),
            bytes,
            "{path}"
        );
    }
}

#[test]
fn mixed_edits_rollback_when_cell_patching_rejects_invalid_xml() {
    let (source, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/></Shape></Shapes></PageContents>",
    );
    assert!(
        diagram
            .save_edits(
                &[edit(
                    page_id,
                    1,
                    "Width",
                    "\"bad\u{1}\"",
                    MutationGesture::CellEdit,
                )],
                &[StructuralEdit::AddShape {
                    page_id,
                    shape_xml: b"<Shape><Cell N='Width' V='3'/></Shape>".to_vec(),
                }],
            )
            .is_err()
    );
    for (path, bytes) in parts(&source) {
        assert_eq!(
            diagram.package().part_bytes(&path).unwrap(),
            bytes,
            "{path}"
        );
    }
}

#[test]
fn lock_delete_refusals_leave_the_facade_package_unchanged() {
    let local = "<PageContents><Shapes><Shape ID='1'><Cell N='LockDelete' F='1' V='1'/></Shape></Shapes></PageContents>";
    let (_, diagram, page_id) = diagram_with_page(local);
    let before = diagram
        .package()
        .part_bytes("visio/pages/page1.xml")
        .unwrap()
        .to_vec();
    assert!(
        diagram
            .save_structural_edits(&[StructuralEdit::DeleteShape {
                page_id,
                shape_id: 1
            }])
            .is_err()
    );
    assert_eq!(
        diagram
            .package()
            .part_bytes("visio/pages/page1.xml")
            .unwrap(),
        before
    );

    let (_, diagram, page_id) = diagram_with_parts(&[
        (
            "visio/pages/page1.xml",
            "<PageContents><Shapes><Shape ID='1' Master='1' MasterShape='10'/></Shapes></PageContents>",
        ),
        (
            "visio/masters/master1.xml",
            "<MasterContents><Shapes><Shape ID='10'><Cell N='LockDelete' F='1' V='1'/></Shape></Shapes></MasterContents>",
        ),
    ]);
    let before = diagram
        .package()
        .part_bytes("visio/pages/page1.xml")
        .unwrap()
        .to_vec();
    assert!(
        diagram
            .save_structural_edits(&[StructuralEdit::DeleteShape {
                page_id,
                shape_id: 1
            }])
            .is_err()
    );
    assert_eq!(
        diagram
            .package()
            .part_bytes("visio/pages/page1.xml")
            .unwrap(),
        before
    );

    let (_, diagram, page_id) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='LockDelete' F='Unknown(1)' V='0'/></Shape></Shapes></PageContents>",
    );
    let before = diagram
        .package()
        .part_bytes("visio/pages/page1.xml")
        .unwrap()
        .to_vec();
    assert!(
        diagram
            .save_structural_edits(&[StructuralEdit::DeleteShape {
                page_id,
                shape_id: 1
            }])
            .is_err()
    );
    assert_eq!(
        diagram
            .package()
            .part_bytes("visio/pages/page1.xml")
            .unwrap(),
        before
    );
}

fn draft_cell(
    name: &str,
    formula: &str,
    section: Option<&str>,
    row: Option<CellRow>,
) -> CellSnapshot {
    CellSnapshot {
        locator: CellLocator {
            sheet: CellSheet::Page(1),
            shape_id: Some(99),
            section: section.map(str::to_owned),
            row,
            cell_name: name.to_owned(),
        },
        name: name.to_owned(),
        formula: Some(formula.to_owned()),
        value: None,
    }
}

#[test]
fn session_added_shapes_reach_the_saved_package() {
    let (source, diagram, _) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/></Shape></Shapes></PageContents>",
    );
    let session = DiagramSession::open(&source, 7).unwrap();
    let page = session.snapshot().unwrap().pages[0].id.clone();
    let receipt = session
        .add_shape(
            &EditCtx::local("a"),
            &page,
            &ShapeDraft {
                source_id: 99,
                name: Some("Added".to_owned()),
                cells: vec![
                    draft_cell("Width", "5", None, None),
                    draft_cell("X", "2", Some("Geometry"), Some(CellRow::Index(0))),
                ],
            },
        )
        .unwrap();
    session
        .set_cell_formula(&EditCtx::local("a"), &page, &receipt.shape_id, "Width", "7")
        .unwrap();

    let saved = diagram.save_session(&session).unwrap();
    let reopened = Diagram::open(&saved).unwrap();
    let page = reopened.pages().next().unwrap();
    assert_eq!(page.shapes().count(), 2);
    let added = page
        .shapes()
        .find(|shape| shape.model().name.as_deref() == Some("Added"))
        .unwrap();
    let width = added
        .model()
        .cells()
        .find(|cell| cell.name == "Width")
        .unwrap();
    assert_eq!(width.formula.as_deref(), Some("7"));
    assert_eq!(width.value.as_deref(), Some("7"));
    let resolved = added.resolved().unwrap();
    assert!(matches!(
        resolved.sections["Geometry"].rows["IX:0"].cells["X"],
        vsdx_resolve::Lookup::Found(ref cell) if cell.cell.formula.as_deref() == Some("2")
    ));
}

#[test]
fn session_added_shapes_do_not_leak_into_cell_edits() {
    let (source, _, _) = diagram_with_page(
        "<PageContents><Shapes><Shape ID='1'><Cell N='Width' V='1'/></Shape></Shapes></PageContents>",
    );
    let session = DiagramSession::open(&source, 7).unwrap();
    let page = session.snapshot().unwrap().pages[0].id.clone();
    session
        .add_shape(
            &EditCtx::local("a"),
            &page,
            &ShapeDraft {
                source_id: 99,
                name: None,
                cells: vec![draft_cell("Width", "5", None, None)],
            },
        )
        .unwrap();
    let export = session.export().unwrap();
    assert!(export.cell_edits.is_empty());
    assert_eq!(export.added_shapes.len(), 1);
    assert_eq!(export.added_shapes[0].cells.len(), 1);
}
