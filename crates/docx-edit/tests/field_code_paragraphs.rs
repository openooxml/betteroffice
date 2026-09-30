//! A paragraph mark inside a complex field's code is part of the code, so
//! Word hides it and joins the paragraph with the next one. Each fixture
//! holds IF fields whose code spans paragraph marks, followed by their
//! results, on A4 at an exact 12pt pitch. The page breaks and lines were
//! exported from Word 16.113.

use docx_edit::bridge::{RenderEnv, RevisionPreview};
use docx_edit::structured::{
    Anchor, FragmentSlice, PageDiagnosticCode, PageExportOptions, RevisionView, StorySelection,
};
use docx_edit::{
    EditCtx, EditTextView, EngineSession, FormatPolicy, MergeDirection, RawOp, ReadParagraphsRequest,
    StoryRange, seed_from_docx,
};
use serde_json::{Value, json};
use yrs::Any;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn fixture(name: &str) -> Vec<u8> {
    std::fs::read(
        std::path::Path::new(env!("CARGO_MANIFEST_DIR"))
            .join("tests/fixtures/field-code-paragraphs")
            .join(format!("{name}.docx")),
    )
    .unwrap()
}

fn request(bytes: &[u8]) -> String {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    json!({
        "bodyStory": "body", "renderEnv": {},
        "regions": {"sections": [{"properties": package.document.final_section_properties}], "settings": package.settings},
        "measurement": {"fontChains": {"arial|0|0": [font]}, "defaults": {"fontFamily": "Arial", "fontSize": 12}}
    })
    .to_string()
}

fn seeded(bytes: &[u8]) -> EngineSession {
    let engine = EngineSession::new(76500);
    seed_from_docx(engine.doc(), bytes).unwrap();
    engine
}

fn field_paragraph_ids(engine: &EngineSession) -> [String; 4] {
    let paragraphs = engine.doc().paragraphs("body").unwrap();
    let owner = paragraphs
        .iter()
        .position(|paragraph| paragraph.text.starts_with("Line 041"))
        .unwrap();
    paragraphs[owner..owner + 4]
        .iter()
        .map(|paragraph| paragraph.para_id.clone())
        .collect::<Vec<_>>()
        .try_into()
        .unwrap()
}

fn lower(engine: &EngineSession, env: &RenderEnv) -> Value {
    serde_json::from_str(&engine.lower_story_json("body", env).unwrap()).unwrap()
}

fn block_text(block: &Value) -> String {
    block["runs"]
        .as_array()
        .unwrap()
        .iter()
        .filter_map(|run| run["text"].as_str())
        .collect()
}

/// Each page's body lines, as text.
fn pages(bytes: &[u8]) -> Vec<Vec<String>> {
    let output: Value = serde_json::from_str(
        &seeded(bytes)
            .layout_document_with_regions_json(&request(bytes))
            .unwrap(),
    )
    .unwrap();
    output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|page| {
            page["fragments"]
                .as_array()
                .unwrap()
                .iter()
                .flat_map(|fragment| {
                    fragment["resolvedLines"]
                        .as_array()
                        .cloned()
                        .unwrap_or_default()
                })
                .map(|line| {
                    line["segments"]
                        .as_array()
                        .cloned()
                        .unwrap_or_default()
                        .iter()
                        .filter_map(|segment| segment["text"].as_str().map(str::to_owned))
                        .collect::<String>()
                })
                .collect()
        })
        .collect()
}

fn page_ranges(pages: &[Vec<String>]) -> Vec<(String, String)> {
    pages
        .iter()
        .map(|lines| (lines[0].clone(), lines.last().unwrap().clone()))
        .collect()
}

#[test]
fn a_header_field_code_over_paragraph_marks_takes_one_line() {
    let pages = pages(&fixture("header-field-code-paragraphs"));
    let starts: Vec<_> = page_ranges(&pages)
        .into_iter()
        .map(|(first, _)| first)
        .collect();
    assert_eq!(
        starts,
        [
            "Line 001", "Line 055", "Line 109", "Line 163", "Line 217", "Line 271"
        ]
    );
}

#[test]
fn a_body_field_code_over_paragraph_marks_joins_its_paragraphs() {
    let pages = pages(&fixture("body-field-code-paragraphs"));
    assert!(
        pages[0]
            .iter()
            .any(|line| line.trim_end() == "Line 041 yes")
    );
    let starts: Vec<_> = page_ranges(&pages)
        .into_iter()
        .map(|(first, _)| first)
        .collect();
    assert_eq!(
        starts,
        ["Line 001", "Line 059", "Line 117", "Line 175", "Line 233"]
    );
}

#[test]
fn adjacent_chained_and_nested_field_codes_join_their_paragraphs() {
    for (name, joined) in [
        ("body-field-code-adjacent", "Line 041 yes"),
        ("body-field-code-chained", "Line 041 yes and again"),
        ("body-field-code-nested", "Line 041 yes"),
    ] {
        let pages = pages(&fixture(name));
        assert!(
            pages[0].iter().any(|line| line.trim_end() == joined),
            "{name}: {:?}",
            &pages[0][38..44]
        );
        let starts: Vec<_> = page_ranges(&pages)
            .into_iter()
            .map(|(first, _)| first)
            .collect();
        assert_eq!(
            starts,
            ["Line 001", "Line 059", "Line 117", "Line 175", "Line 233"],
            "{name}"
        );
    }
}

#[test]
fn a_field_code_over_paragraph_marks_in_a_table_cell_joins_its_paragraphs() {
    let pages = pages(&fixture("cell-field-code-paragraphs"));
    let starts: Vec<_> = page_ranges(&pages)
        .into_iter()
        .map(|(first, _)| first)
        .collect();
    assert_eq!(
        starts,
        ["Line 001", "Line 059", "Line 117", "Line 175", "Line 233"]
    );
}

#[test]
fn the_page_map_places_both_joined_paragraphs() {
    let bytes = fixture("body-field-code-paragraphs");
    let engine = seeded(&bytes);
    engine
        .layout_document_with_regions_retained_json(&request(&bytes))
        .unwrap();
    let content = engine
        .export_structured_with_pages(&PageExportOptions {
            stories: Some(vec![StorySelection::Body]),
            ..PageExportOptions::new(RevisionView::Accepted)
        })
        .unwrap()
        .content;
    let paragraphs = engine
        .doc()
        .read_paragraphs(&ReadParagraphsRequest {
            story: None,
            para_ids: None,
            view: EditTextView::Accepted,
        })
        .unwrap()
        .paragraphs;
    let owner = paragraphs
        .iter()
        .position(|paragraph| paragraph.text.starts_with("Line 041"))
        .unwrap();
    for paragraph in [&paragraphs[owner], &paragraphs[owner + 3]] {
        let block = content.structured.stories[0]
            .blocks
            .iter()
            .find(|block| {
                matches!(&block.anchor, Anchor::Paragraph { para_id, .. } if *para_id == paragraph.para_id)
            })
            .unwrap();
        assert!(
            content
                .layout
                .fragments
                .iter()
                .any(|fragment| fragment.node_id == block.id),
            "{:?} has a page fragment",
            paragraph.text
        );
        assert!(
            !content.layout.diagnostics.iter().any(|diagnostic| {
                diagnostic.code == PageDiagnosticCode::NotLaidOut
                    && diagnostic.node_id.as_deref() == Some(block.id.as_str())
            }),
            "{:?} is laid out",
            paragraph.text
        );
    }
}

#[test]
fn merging_the_result_backward_keeps_its_visible_content() {
    let engine = seeded(&fixture("body-field-code-paragraphs"));
    let [owner, _, hidden, target] = field_paragraph_ids(&engine);
    engine
        .doc()
        .merge_paragraphs(&EditCtx::local("", ""), &target, MergeDirection::Backward)
        .unwrap();
    let blocks = lower(&engine, &RenderEnv::default());
    let blocks = blocks.as_array().unwrap();
    assert!(blocks.iter().any(|block| block["id"] == owner));
    let survivor = blocks.iter().find(|block| block["id"] == hidden).unwrap();
    assert!(block_text(survivor).contains("yes"));
}

#[test]
fn a_missing_bound_code_paragraph_cancels_the_join() {
    let engine = seeded(&fixture("body-field-code-paragraphs"));
    let [owner, hidden, _, target] = field_paragraph_ids(&engine);
    engine
        .doc()
        .merge_paragraphs(&EditCtx::local("", ""), &hidden, MergeDirection::Forward)
        .unwrap();
    let blocks = lower(&engine, &RenderEnv::default());
    let blocks = blocks.as_array().unwrap();
    assert!(blocks.iter().any(|block| block["id"] == owner));
    let result = blocks.iter().find(|block| block["id"] == target).unwrap();
    assert_eq!(block_text(result), "yes");
}

#[test]
fn a_rejected_spanning_field_does_not_join_its_paragraphs() {
    let engine = seeded(&fixture("body-field-code-paragraphs"));
    let [owner, first_hidden, last_hidden, target] = field_paragraph_ids(&engine);
    let field = engine.doc().paragraph_mark_position(&owner).unwrap().index - 1;
    engine
        .doc()
        .apply_raw_ops(
            "body",
            vec![RawOp::Format {
                index: field,
                len: 1,
                attrs: [(
                    "ins".into(),
                    Any::Map(std::sync::Arc::new(std::collections::HashMap::from([
                        ("id".to_owned(), Any::from("field-insertion")),
                    ]))),
                )]
                .into(),
            }],
            &EditCtx::local("", ""),
        )
        .unwrap();
    let joined = lower(&engine, &RenderEnv::default());
    assert!(
        !joined
            .as_array()
            .unwrap()
            .iter()
            .any(|block| block["id"] == owner)
    );
    let rejected = lower(
        &engine,
        &RenderEnv::default().with_revision_preview("field-insertion", RevisionPreview::Rejected),
    );
    for id in [owner, first_hidden, last_hidden, target.clone()] {
        assert!(
            rejected
                .as_array()
                .unwrap()
                .iter()
                .any(|block| block["id"] == id)
        );
    }
    let result = rejected
        .as_array()
        .unwrap()
        .iter()
        .find(|block| block["id"] == target)
        .unwrap();
    assert_eq!(block_text(result), "yes");
}

#[test]
fn only_visible_inline_drawings_block_a_join() {
    for (kind, source, anchored, hidden, rejected) in [
        ("shape", 0, true, false, false),
        ("shape", 1, true, false, false),
        ("shape", 3, true, false, false),
        ("shape", 0, false, true, false),
        ("shape", 0, false, false, true),
        ("shape", 0, false, false, false),
        ("image", 0, true, false, false),
        ("image", 1, true, false, false),
        ("image", 3, true, false, false),
        ("image", 0, false, true, false),
        ("image", 0, false, false, true),
        ("image", 1, false, false, false),
    ] {
        let engine = seeded(&fixture("body-field-code-paragraphs"));
        let ids = field_paragraph_ids(&engine);
        let at = engine
            .doc()
            .paragraph_mark_position(&ids[source])
            .unwrap()
            .index;
        let mut attrs = yrs::types::Attrs::from([("hidden".into(), Any::Bool(hidden))]);
        if rejected {
            attrs.insert(
                "ins".into(),
                Any::Map(std::sync::Arc::new(std::collections::HashMap::from([
                    ("id".to_owned(), Any::from("drawing-insertion")),
                ]))),
            );
        }
        let payload = if kind == "shape" {
            vec![(
                "shapeJson".to_owned(),
                Any::from(
                    json!({
                        "shapeType": "textBox",
                        "wrap": anchored.then(|| json!({"type": "square"})),
                    })
                    .to_string(),
                ),
            )]
        } else {
            vec![
                ("src".to_owned(), Any::from("image.png")),
                (
                    "wrapType".to_owned(),
                    Any::from(if anchored { "square" } else { "inline" }),
                ),
                (
                    "displayMode".to_owned(),
                    Any::from(if anchored { "float" } else { "inline" }),
                ),
            ]
        };
        engine
            .doc()
            .apply_raw_ops(
                "body",
                vec![RawOp::InsertEmbed {
                    index: at,
                    kind: kind.to_owned(),
                    payload,
                    attrs,
                }],
                &EditCtx::local("", ""),
            )
            .unwrap();
        let blocks = lower(
            &engine,
            &RenderEnv::default()
                .with_revision_preview("drawing-insertion", RevisionPreview::Rejected),
        );
        let blocks = blocks.as_array().unwrap();
        let target = blocks.iter().find(|block| block["id"] == ids[3]).unwrap();
        assert_eq!(
            block_text(target).contains("Line 041"),
            anchored || hidden || rejected
        );
        assert_eq!(
            blocks.iter().any(|block| {
                block["kind"] == "shape"
                    || block["runs"]
                        .as_array()
                        .is_some_and(|runs| runs.iter().any(|run| run["kind"] == "image"))
            }),
            !hidden && !rejected
        );
    }
}

#[test]
fn inserting_text_beside_a_nested_code_field_cancels_the_join() {
    let engine = seeded(&fixture("body-field-code-nested"));
    let paragraphs = engine.doc().paragraphs("body").unwrap();
    let owner = paragraphs
        .iter()
        .position(|paragraph| paragraph.text.starts_with("Line 041"))
        .unwrap();
    let hidden = &paragraphs[owner + 1].para_id;
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            engine.doc().paragraph_mark_position(hidden).unwrap(),
            "inserted",
            FormatPolicy::Plain,
        )
        .unwrap();
    let blocks = lower(&engine, &RenderEnv::default());
    let blocks = blocks.as_array().unwrap();
    assert!(
        blocks
            .iter()
            .any(|block| block["id"] == paragraphs[owner].para_id)
    );
    let survivor = blocks.iter().find(|block| block["id"] == *hidden).unwrap();
    assert!(block_text(survivor).contains("inserted"));
}

#[test]
fn a_trailing_soft_break_keeps_its_blank_line_page_fragment() {
    let mut parts = ooxml_opc::unzip_parts(&fixture("body-field-code-paragraphs")).unwrap();
    let (_, document) = parts
        .iter_mut()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap();
    let lines = r#"<w:r><w:t>line</w:t><w:br/></w:r>"#.repeat(6);
    *document = format!(
        r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/><w:widowControl w:val="0"/></w:pPr>{lines}</w:p><w:sectPr><w:pgSz w:w="5760" w:h="2880"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="0" w:footer="0" w:gutter="0"/></w:sectPr></w:body></w:document>"#
    )
    .into_bytes();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let engine = seeded(&bytes);
    let paragraph = engine.doc().paragraphs("body").unwrap()[0].para_id.clone();
    engine
        .layout_document_with_regions_retained_json(&request(&bytes))
        .unwrap();
    let content = engine
        .export_structured_with_pages(&PageExportOptions {
            stories: Some(vec![StorySelection::Body]),
            ..PageExportOptions::new(RevisionView::Accepted)
        })
        .unwrap()
        .content;
    let fragments: Vec<_> = content
        .layout
        .fragments
        .iter()
        .filter(|fragment| {
            matches!(&fragment.anchor, Anchor::Paragraph { para_id, .. } if *para_id == paragraph)
                && matches!(fragment.slice, FragmentSlice::Block)
        })
        .collect();
    assert_eq!(fragments.len(), 2);
    assert_eq!(fragments[0].page_index, 0);
    assert!(!fragments[0].continued_from_previous);
    assert!(fragments[0].continued_on_next);
    assert_eq!(fragments[1].page_index, 1);
    assert!(fragments[1].continued_from_previous);
    assert!(!fragments[1].continued_on_next);
    assert!(!content.layout.fragments.iter().any(|fragment| {
        fragment.block_id == fragments[1].block_id
            && fragment.page_index == 1
            && !matches!(fragment.slice, FragmentSlice::Block)
    }));
}

#[test]
fn retained_display_refreshes_joined_run_positions() {
    let bytes = fixture("body-field-code-paragraphs");
    let engine = seeded(&bytes);
    let [_, hidden, _, _] = field_paragraph_ids(&engine);
    let code = engine.doc().paragraph_mark_position(&hidden).unwrap().index;
    engine
        .doc()
        .apply_raw_ops(
            "body",
            vec![RawOp::Insert {
                index: code,
                text: "code".to_owned(),
                attrs: [("hidden".into(), Any::Bool(true))].into(),
            }],
            &EditCtx::local("", ""),
        )
        .unwrap();
    let request = request(&bytes);
    let config: Value = serde_json::from_str(&request).unwrap();
    let extras = json!({"fontChains": config["measurement"]["fontChains"]}).to_string();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    engine.build_display_list_frame(&extras, 0).unwrap();
    engine
        .doc()
        .delete_range(
            &EditCtx::local("", ""),
            StoryRange::new("body", code, code + 4),
        )
        .unwrap();
    let last = engine
        .doc()
        .paragraphs("body")
        .unwrap()
        .last()
        .unwrap()
        .para_id
        .clone();
    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            engine.doc().paragraph_mark_position(&last).unwrap(),
            "x",
            FormatPolicy::Plain,
        )
        .unwrap();
    engine
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    engine.build_display_list_frame(&extras, 1).unwrap();
    let fresh = EngineSession::new(76501);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    fresh
        .layout_document_with_regions_retained_json(&request)
        .unwrap();
    fresh.build_display_list_frame(&extras, 0).unwrap();
    assert!(engine.stats().incremental_display_builds > 0);
    let display = |engine: &EngineSession| {
        engine
            .with_display_list(|list| serde_json::to_value(list).unwrap())
            .unwrap()
    };
    assert_eq!(display(&engine), display(&fresh));
}

#[test]
fn joined_sources_have_their_own_page_continuations() {
    let bytes = fixture("body-field-code-paragraphs");
    let engine = seeded(&bytes);
    let [owner, _, _, target] = field_paragraph_ids(&engine);
    let owner_end = engine.doc().paragraph_mark_position(&owner).unwrap().index - 1;
    engine
        .doc()
        .apply_raw_ops(
            "body",
            vec![RawOp::Insert {
                index: owner_end,
                text: "carried\u{000b}".repeat(90),
                attrs: Default::default(),
            }],
            &EditCtx::local("", ""),
        )
        .unwrap();
    let target_end = engine.doc().paragraph_mark_position(&target).unwrap();
    engine
        .doc()
        .apply_raw_ops(
            "body",
            vec![RawOp::Insert {
                index: target_end.index,
                text: "result\u{000b}".repeat(90),
                attrs: Default::default(),
            }],
            &EditCtx::local("", ""),
        )
        .unwrap();
    engine
        .layout_document_with_regions_retained_json(&request(&bytes))
        .unwrap();
    let content = engine
        .export_structured_with_pages(&PageExportOptions {
            stories: Some(vec![StorySelection::Body]),
            ..PageExportOptions::new(RevisionView::Accepted)
        })
        .unwrap()
        .content;
    let fragments = |id: &str| {
        content
            .layout
            .fragments
            .iter()
            .filter(|fragment| {
                matches!(&fragment.anchor, Anchor::Paragraph { para_id, .. } if para_id == id)
                    && matches!(fragment.slice, FragmentSlice::Block)
            })
            .collect::<Vec<_>>()
    };
    let owner_fragments = fragments(&owner);
    let target_fragments = fragments(&target);
    for fragments in [&owner_fragments, &target_fragments] {
        assert!(fragments.len() >= 2);
        for (index, fragment) in fragments.iter().enumerate() {
            assert_eq!(fragment.continued_from_previous, index > 0);
            assert_eq!(fragment.continued_on_next, index + 1 < fragments.len());
        }
    }
    assert!(owner_fragments[0].page_index < target_fragments[0].page_index);
    assert!(
        owner_fragments.last().unwrap().page_index < target_fragments.last().unwrap().page_index
    );
}
