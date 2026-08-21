//! Ceilings: every defect the ledger enumerates is probed here and pinned with
//! exact equality — never `<=`. Fixing one without lowering its ceiling fails,
//! and so does a probe that stops reproducing while the ledger still lists it,
//! so a fix must claim its fix and a regression cannot hide in headroom.
//! Governed by `openspec/changes/docx-word-fidelity/specs/fidelity-meta`.

mod common;

use std::collections::BTreeSet;
use std::path::{Path, PathBuf};

use betteroffice_docx::{Document, Error};
use common::{
    parts_of, roundtrip_report, sample_docx, save_unedited, with_document_xml, with_part,
};
use docx_edit::{EditingDoc, project_story, seed_from_docx};

/// One enumerated defect and the observation that still reproduces it.
struct Probe {
    criterion: &'static str,
    /// Verbatim the ledger's defect text.
    defect: &'static str,
    reproduces: fn() -> bool,
}

const PROBES: &[Probe] = &[
    Probe {
        criterion: "pkg.unknown-xml",
        defect: "raw foreign nodes do not travel the collaborative edit lane; an editor session \
                 drops them",
        reproduces: the_edit_lane_cannot_see_foreign_markup,
    },
    Probe {
        criterion: "pkg.unknown-xml",
        defect: "foreign root declarations on note and comment parts are not yet carried through \
                 save",
        reproduces: a_note_part_loses_its_foreign_root_declaration,
    },
    Probe {
        criterion: "edit.typing",
        defect: "replace_paragraph_text refuses a paragraph carrying bookmark markers \
                 (UnsupportedParagraphEdit)",
        reproduces: typing_refuses_a_bookmarked_paragraph,
    },
];

/// The save path preserves foreign markup (`fidelity.rs` gates that), but the
/// collaborative model is seeded without it: nothing the editor session can
/// project mentions the foreign namespace, so an edit round trip cannot put
/// back what it never held.
fn the_edit_lane_cannot_see_foreign_markup() -> bool {
    let package = with_document_xml(&sample_docx(), |xml| {
        xml.replace(
            "<w:r><w:t>DOCX</w:t></w:r>",
            r#"<cx:tag xmlns:cx="urn:custom-x" cx:k="v"/><w:r><w:t>DOCX</w:t></w:r>"#,
        )
    });
    let document = EditingDoc::new(1);
    seed_from_docx(&document, &package).unwrap();
    !project_story(&document, "body")
        .unwrap()
        .iter()
        .any(|item| format!("{item:?}").contains("custom-x"))
}

fn a_note_part_loses_its_foreign_root_declaration() -> bool {
    let original = with_part(&comprehensive_fixture(), "word/footnotes.xml", |xml| {
        xml.replacen(
            r#"<w:footnotes "#,
            r#"<w:footnotes xmlns:bofx="urn:betteroffice-fixture-x" bofx:flag="1" "#,
            1,
        )
    });
    let saved = save_unedited(&original);
    roundtrip_report(&parts_of(&original), &parts_of(&saved))
        .iter()
        .any(|finding| finding.contains("word/footnotes.xml"))
}

fn typing_refuses_a_bookmarked_paragraph() -> bool {
    let mut document = Document::open(&sample_docx()).unwrap();
    matches!(
        document.replace_paragraph_text("11111111", "Retyped"),
        Err(Error::UnsupportedParagraphEdit(_))
    )
}

fn comprehensive_fixture() -> Vec<u8> {
    let path: PathBuf = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/corpus/fixtures/wordprocessingml-comprehensive.docx");
    std::fs::read(path).unwrap()
}

fn reproducing(criterion: &str) -> Vec<&'static str> {
    PROBES
        .iter()
        .filter(|probe| probe.criterion == criterion && (probe.reproduces)())
        .map(|probe| probe.defect)
        .collect()
}

fn ledger_defects() -> BTreeSet<(String, String)> {
    let ledger: serde_json::Value = serde_json::from_str(
        &std::fs::read_to_string(
            Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/scorecard/ledger.json"),
        )
        .unwrap(),
    )
    .unwrap();
    ledger["criteria"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|criterion| {
            let id = criterion["id"].as_str().unwrap().to_owned();
            criterion["defects"]
                .as_array()
                .unwrap()
                .iter()
                .map(move |defect| (id.clone(), defect.as_str().unwrap().to_owned()))
        })
        .collect()
}

#[test]
fn unknown_xml_holds_at_two_defects() {
    assert_eq!(reproducing("pkg.unknown-xml").len(), 2);
}

#[test]
fn typing_holds_at_one_defect() {
    assert_eq!(reproducing("edit.typing").len(), 1);
}

/// The count alone would let a fixed defect pay for a new one, so the names
/// bind too: the ledger's enumeration is exactly what still reproduces.
#[test]
fn the_ledger_enumerates_exactly_the_defects_that_still_reproduce() {
    let observed: BTreeSet<(String, String)> = PROBES
        .iter()
        .filter(|probe| (probe.reproduces)())
        .map(|probe| (probe.criterion.to_owned(), probe.defect.to_owned()))
        .collect();
    assert_eq!(observed, ledger_defects());
}
