//! Round-trip gates through the package oracles, which parse bytes with
//! their own reader so a model gap cannot hide a save loss.

mod common;

use betteroffice_docx::Document;
use common::{
    Parts, parts_of, roundtrip_report, sample_docx, save_unedited, with_document_xml, with_part,
};
use ooxml_fidelity::wml::{Difference, diff_digests, semantic_digest};
use ooxml_fidelity::{element_census, losses};

fn digest_diff(before: &Parts, after: &Parts) -> Vec<Difference> {
    diff_digests(
        &semantic_digest(before).unwrap(),
        &semantic_digest(after).unwrap(),
    )
}

#[test]
fn an_unedited_round_trip_passes_both_oracles_and_the_byte_rules() {
    let original = sample_docx();
    let saved = save_unedited(&original);
    assert_eq!(
        roundtrip_report(&parts_of(&original), &parts_of(&saved)),
        Vec::<String>::new()
    );
}

#[test]
fn save_reopen_save_is_byte_identical() {
    let first = save_unedited(&sample_docx());
    let second = save_unedited(&first);
    assert_eq!(first, second);
}

#[test]
fn an_edited_round_trip_loses_nothing_beyond_its_footprint() {
    let original = sample_docx();
    let mut document = Document::open(&original).unwrap();
    document
        .replace_paragraph_text("22222222", "Edited natively")
        .unwrap();
    let edited = document.save().unwrap();
    let before = parts_of(&original);
    let after = parts_of(&edited);
    assert_eq!(
        losses(
            &element_census(&before).unwrap(),
            &element_census(&after).unwrap()
        ),
        vec![]
    );
    assert_eq!(
        digest_diff(&before, &after),
        vec![Difference {
            path: "word/document.xml block[1].text".to_owned(),
            before: "Cell text".to_owned(),
            after: "Edited natively".to_owned(),
        }]
    );
}

#[test]
fn the_guard_has_teeth() {
    let saved = save_unedited(&sample_docx());
    let broken = with_document_xml(&saved, |xml| {
        xml.replace(r#"<w:bookmarkStart w:id="1" w:name="mark"/>"#, "")
    });
    let before = parts_of(&saved);
    let after = parts_of(&broken);
    let census_losses = losses(
        &element_census(&before).unwrap(),
        &element_census(&after).unwrap(),
    );
    assert_eq!(census_losses.len(), 1);
    assert_eq!(census_losses[0].local, "bookmarkStart");
    assert_ne!(digest_diff(&before, &after), vec![]);
}

#[test]
fn an_unknown_element_in_a_modelled_part_survives_the_round_trip() {
    let original = with_document_xml(&sample_docx(), |xml| {
        xml.replace(
            "<w:body>",
            r#"<w:body><cx:block xmlns:cx="urn:custom-x"><cx:y cx:z="1"/></cx:block>"#,
        )
        .replace(
            "<w:r><w:t>DOCX</w:t></w:r>",
            r#"<cx:tag xmlns:cx="urn:custom-x" cx:k="v"/><w:r><w:t>DOCX</w:t></w:r>"#,
        )
    });
    let saved = save_unedited(&original);
    assert_eq!(
        losses(
            &element_census(&parts_of(&original)).unwrap(),
            &element_census(&parts_of(&saved)).unwrap(),
        ),
        vec![]
    );
    assert_eq!(digest_diff(&parts_of(&original), &parts_of(&saved)), vec![]);
}

/// Foreign markup in a header, declared on the story root rather than
/// inline, must keep resolving after a save re-emits the root.
#[test]
fn root_declared_foreign_markup_in_a_header_survives_the_round_trip() {
    let original = with_part(&sample_docx(), "word/header1.xml", |xml| {
        xml.replace(
            r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#,
            r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:bofx="urn:betteroffice-fixture-x""#,
        )
        .replace("<w:p ", r#"<bofx:hmark/><w:p bofx:flag="1" "#)
    });
    let saved = save_unedited(&original);
    assert_eq!(
        roundtrip_report(&parts_of(&original), &parts_of(&saved)),
        Vec::<String>::new()
    );
}

#[test]
fn an_unknown_element_in_a_table_cell_survives_the_round_trip() {
    let original = with_document_xml(&sample_docx(), |xml| {
        xml.replace(
            r#"<w:p w14:paraId="22222222">"#,
            r#"<cx:cellmark xmlns:cx="urn:custom-x"/><w:p w14:paraId="22222222">"#,
        )
    });
    let saved = save_unedited(&original);
    assert_eq!(
        roundtrip_report(&parts_of(&original), &parts_of(&saved)),
        Vec::<String>::new()
    );
}

/// The story roots do not declare the chartex family, so an authored `cx*`
/// declaration is a custom binding, not boilerplate to filter.
#[test]
fn chartex_declarations_on_story_roots_survive_the_round_trip() {
    let with_cx = |xml: String| {
        xml.replace(
            r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main""#,
            r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:cx1="http://schemas.microsoft.com/office/drawing/2015/9/8/chartex""#,
        )
        .replace("<w:p ", r#"<cx1:mark/><w:p "#)
    };
    for part in ["word/document.xml", "word/header1.xml"] {
        let original = with_part(&sample_docx(), part, with_cx);
        let saved = save_unedited(&original);
        assert_eq!(
            roundtrip_report(&parts_of(&original), &parts_of(&saved)),
            Vec::<String>::new(),
            "{part}"
        );
    }
}

/// The census is blind to attributes; the digest is the net here.
#[test]
fn an_unknown_attribute_on_a_known_element_survives_the_round_trip() {
    let original = with_document_xml(&sample_docx(), |xml| {
        xml.replace(
            r#"<w:p w14:paraId="33333333">"#,
            r#"<w:p w14:paraId="33333333" cx:flag="1" xmlns:cx="urn:custom-x">"#,
        )
    });
    let saved = save_unedited(&original);
    assert_eq!(digest_diff(&parts_of(&original), &parts_of(&saved)), vec![]);
}
