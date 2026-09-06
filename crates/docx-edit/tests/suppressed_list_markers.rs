use docx_edit::bridge::{RenderEnv, yrs_doc_to_layout_blocks};
use docx_edit::{EditingDoc, seed_from_docx};
use docx_layout::types::{LayoutBlock, ParagraphBlock};

const DOCX: &[u8] = include_bytes!("fixtures/suppressed-list-markers.docx");

fn paragraphs() -> Vec<ParagraphBlock> {
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, DOCX).unwrap();
    yrs_doc_to_layout_blocks(&doc, "body", &RenderEnv::default())
        .unwrap()
        .into_iter()
        .filter_map(|block| match block {
            LayoutBlock::Paragraph(paragraph) => Some(paragraph),
            _ => None,
        })
        .collect()
}

#[test]
fn none_suppresses_direct_and_inherited_placeholders_at_each_level_and_suffix() {
    let paragraphs = paragraphs();
    assert_eq!(paragraphs.len(), 25);
    for (index, paragraph) in paragraphs[1..13].iter().enumerate() {
        let attrs = paragraph.attrs.as_ref().unwrap();
        let num_pr = attrs.num_pr.as_ref().unwrap();
        assert_eq!(num_pr.num_id, Some((41 + index / 4) as f64));
        assert_eq!(num_pr.ilvl, Some(((index % 4) / 2) as f64));
        assert_eq!(attrs.list_is_bullet, Some(false));
        assert_eq!(
            attrs.list_marker_suffix.as_deref(),
            Some(["nothing", "space", "tab"][index / 4])
        );
        assert_eq!(attrs.list_marker, None, "paragraph {}", index + 1);
    }
}

#[test]
fn none_with_missing_or_empty_text_does_not_synthesize_a_marker() {
    for paragraph in &paragraphs()[13..17] {
        assert_eq!(paragraph.attrs.as_ref().unwrap().list_marker, None);
    }
}

#[test]
fn decimal_counters_and_bullet_markers_are_preserved() {
    let paragraphs = paragraphs();
    let markers: Vec<_> = paragraphs[17..21]
        .iter()
        .map(|paragraph| paragraph.attrs.as_ref().unwrap().list_marker.as_deref())
        .collect();
    assert_eq!(markers, [Some("1."), Some("2."), Some("•"), Some("•")]);
}

#[test]
fn none_preserves_literal_text_but_resolves_embedded_placeholders() {
    let paragraphs = paragraphs();
    let markers: Vec<_> = paragraphs[21..]
        .iter()
        .map(|paragraph| paragraph.attrs.as_ref().unwrap().list_marker.as_deref())
        .collect();
    assert_eq!(
        markers,
        [
            Some("Chapter"),
            Some("Chapter"),
            Some("Literal  label"),
            Some("Literal  label")
        ]
    );
}
