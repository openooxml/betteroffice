use std::sync::Arc;

use docx_edit::{EditingDoc, seed_from_docx};
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{Any, Out, ReadTxn, Text, Transact};

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);

const FORMATTED: &str = r#"<w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:b/><w:color w:val="C00000"/><w:highlight w:val="yellow"/></w:rPr>"#;

const NORMAL_STYLE: &str = r#"<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>"#;

const SHARED_PARAGRAPH_STYLE: &str = concat!(
    r#"<w:style w:type="paragraph" w:styleId="Shared"><w:basedOn w:val="Normal"/>"#,
    r#"<w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:b/></w:rPr></w:style>"#
);

fn package(body: &str, styles: &str) -> Vec<u8> {
    let content_types = r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#;
    let root_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    let document_rels = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdStyles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let document = format!(r#"<w:document {NS}><w:body>{body}<w:sectPr/></w:body></w:document>"#);
    let styles = format!(r#"<w:styles {NS}>{NORMAL_STYLE}{styles}</w:styles>"#);
    let parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            content_types.as_bytes().to_vec(),
        ),
        ("_rels/.rels".to_owned(), root_rels.as_bytes().to_vec()),
        (
            "word/_rels/document.xml.rels".to_owned(),
            document_rels.as_bytes().to_vec(),
        ),
        ("word/document.xml".to_owned(), document.into_bytes()),
        ("word/styles.xml".to_owned(), styles.into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// `First`, `Second` and `Third` carry the same direct formatting; the plain run between the
/// first two makes the third start its formatting afresh rather than continue the first's.
fn repeated_formatting() -> Vec<u8> {
    package(
        &format!(
            concat!(
                r#"<w:p w14:paraId="00000001"><w:r>{rpr}<w:t>First</w:t></w:r>"#,
                r#"<w:r><w:t xml:space="preserve"> plain </w:t></w:r>"#,
                r#"<w:r>{rpr}<w:t>Second</w:t></w:r></w:p>"#,
                r#"<w:p w14:paraId="00000002"><w:r>{rpr}<w:t>Third</w:t></w:r></w:p>"#,
            ),
            rpr = FORMATTED
        ),
        "",
    )
}

/// Three paragraphs with the same style and spacing, so their marks carry equal payload values.
fn repeated_paragraphs() -> Vec<u8> {
    let body: String = (1..=3)
        .map(|index| {
            format!(
                r#"<w:p w14:paraId="0000000{index}"><w:pPr><w:pStyle w:val="Shared"/><w:spacing w:before="120" w:after="240"/></w:pPr><w:r><w:t>Paragraph {index}</w:t></w:r></w:p>"#
            )
        })
        .collect();
    package(&body, SHARED_PARAGRAPH_STYLE)
}

fn seeded(bytes: &[u8]) -> EditingDoc {
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, bytes).unwrap();
    doc
}

fn run_attributes(doc: &EditingDoc, texts: &[&str]) -> Vec<Attrs> {
    let txn = doc.yrs_doc().transact();
    let stories = txn.get_map("stories").unwrap();
    let Some(Out::YText(body)) = yrs::Map::get(&stories, &txn, "body") else {
        panic!("the body story is text");
    };
    let diff = body.diff(&txn, YChange::identity);
    texts
        .iter()
        .map(|text| {
            diff.iter()
                .find_map(|op| match &op.insert {
                    Out::Any(Any::String(chunk)) if chunk.contains(*text) => {
                        Some(op.attributes.as_deref().cloned().unwrap_or_default())
                    }
                    _ => None,
                })
                .unwrap_or_else(|| panic!("no body text holds {text:?}"))
        })
        .collect()
}

/// A text diff rebuilds each attribute key from `&str` but clones the stored values, so value
/// allocations are observable through it and key allocations are not.
#[test]
fn equally_formatted_runs_share_their_attribute_values() {
    let doc = seeded(&repeated_formatting());
    let runs = run_attributes(&doc, &["First", "Second", "Third"]);
    let (first, rest) = runs.split_first().unwrap();
    let mut strings = 0;
    let mut maps = 0;
    for other in rest {
        for (key, value) in first {
            if other.get(key) != Some(value) {
                continue;
            }
            match (value, &other[key]) {
                (Any::String(left), Any::String(right)) => {
                    strings += 1;
                    assert!(
                        Arc::ptr_eq(left, right),
                        "{key}: equal string values are separate allocations"
                    );
                }
                (Any::Map(left), Any::Map(right)) => {
                    maps += 1;
                    assert!(
                        Arc::ptr_eq(left, right),
                        "{key}: equal map values are separate allocations"
                    );
                }
                _ => {}
            }
        }
    }
    assert!(
        strings > 0 && maps > 0,
        "the runs share no string or no map attribute: {first:?}"
    );
}

/// Guard: a replica decodes every attribute value into its own allocation, so its encoding is
/// the unshared layout's; sharing must not change a single byte of the seed.
#[test]
fn shared_attribute_values_encode_like_unshared_ones() {
    let doc = seeded(&repeated_formatting());
    let shared = doc.encode_state_as_update_v1();
    let replica = EditingDoc::new(2);
    replica.apply_update_v1(&shared).unwrap();
    let separate = run_attributes(&replica, &["First", "Third"]);
    assert!(
        !matches!(
            (separate[0].get("highlight"), separate[1].get("highlight")),
            (Some(Any::String(left)), Some(Any::String(right))) if Arc::ptr_eq(left, right)
        ),
        "the replica shares values, so it does not stand for the unshared layout"
    );
    assert!(replica.encode_state_as_update_v1() == shared);
}

type MarkEntry = (usize, String, Any);

/// Each body paragraph mark's entries, with the address of the stored key.
fn paragraph_marks(doc: &EditingDoc) -> Vec<Vec<MarkEntry>> {
    let txn = doc.yrs_doc().transact();
    let stories = txn.get_map("stories").unwrap();
    let Some(Out::YText(body)) = yrs::Map::get(&stories, &txn, "body") else {
        panic!("the body story is text");
    };
    body.diff(&txn, YChange::identity)
        .into_iter()
        .filter_map(|op| match op.insert {
            Out::YMap(mark) => Some(mark),
            _ => None,
        })
        .filter(|mark| {
            matches!(
                yrs::Map::get(mark, &txn, "_kind"),
                Some(Out::Any(Any::String(kind))) if &*kind == "pilcrow"
            )
        })
        .map(|mark| {
            yrs::Map::iter(&mark, &txn)
                .filter_map(|(key, value)| match value {
                    Out::Any(value) => Some((key.as_ptr() as usize, key.to_owned(), value)),
                    _ => None,
                })
                .collect()
        })
        .collect()
}

fn mark_entry<'a>(mark: &'a [MarkEntry], key: &str) -> Option<&'a MarkEntry> {
    mark.iter().find(|(_, name, _)| name == key)
}

/// Map iteration borrows each key from the stored item, so key allocations are observable too.
#[test]
fn equally_styled_paragraph_marks_share_their_payload_keys_and_values() {
    let doc = seeded(&repeated_paragraphs());
    let marks = paragraph_marks(&doc);
    let styled: Vec<&Vec<MarkEntry>> = marks
        .iter()
        .filter(|mark| {
            mark_entry(mark, "pStyle").is_some_and(|(_, _, value)| *value == Any::from("Shared"))
        })
        .collect();
    assert_eq!(styled.len(), 3, "the paragraph marks: {marks:?}");
    let (first, rest) = styled.split_first().unwrap();
    let mut strings = 0;
    let mut maps = 0;
    for other in rest {
        for key in ["pStyle", "defaultTextFormatting", "_originalFormatting"] {
            let (Some((left_key, _, left)), Some((right_key, _, right))) =
                (mark_entry(first, key), mark_entry(other, key))
            else {
                continue;
            };
            assert_eq!(left, right, "{key}: the marks' values differ");
            assert_eq!(
                left_key, right_key,
                "{key}: equal payload keys are separate allocations"
            );
            match (left, right) {
                (Any::String(left), Any::String(right)) => {
                    strings += 1;
                    assert!(
                        Arc::ptr_eq(left, right),
                        "{key}: equal string values are separate allocations"
                    );
                }
                (Any::Map(left), Any::Map(right)) => {
                    maps += 1;
                    assert!(
                        Arc::ptr_eq(left, right),
                        "{key}: equal map values are separate allocations"
                    );
                }
                _ => {}
            }
        }
    }
    assert!(
        strings > 0 && maps > 0,
        "the marks share no string or no map payload value: {first:?}"
    );
}
