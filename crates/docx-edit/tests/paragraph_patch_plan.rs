use std::collections::BTreeMap;
use std::sync::Arc;

use docx_edit::{
    EditCtx, EditHistory, EditOperation, EditRequest, EditSource, EditStep, EditingDoc,
    FormatPolicy, InlineFormatDelta, ParaAttrDelta, ParaSelector, ParagraphRef, ParagraphTarget,
    Patch, Position, StoryRange, StrikePatch, UnderlinePatch, UndoSession, seed_from_docx,
};
use docx_parse::paragraph_patch::{
    Attrs, DonorV1, ParagraphPatchV1, PatchErrorCode, SlotV1, UnitV1,
};
use serde_json::json;
use yrs::Any;

#[path = "../src/paragraph_patch_hash.rs"]
mod paragraph_patch_hash;

use paragraph_patch_hash::{result_sha256, unit_sha256};

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);

fn open(content: &str) -> EditingDoc {
    let document = format!(
        r#"<w:document {NS}><w:body><w:p w14:paraId="00000001">{content}</w:p><w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr></w:body></w:document>"#
    );
    let styles = format!(
        r#"<w:styles {NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style><w:style w:type="paragraph" w:styleId="Slant"><w:name w:val="Slant"/><w:rPr><w:i/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Strong"><w:name w:val="Strong"/><w:pPr><w:keepNext/></w:pPr><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style><w:style w:type="paragraph" w:styleId="Right"><w:name w:val="Right"/><w:pPr><w:jc w:val="right"/></w:pPr></w:style></w:styles>"#
    );
    let header = format!(
        r#"<w:hdr {NS}><w:p w14:paraId="00000002"><w:r><w:t>abc</w:t></w:r></w:p></w:hdr>"#
    );
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_owned(),
        ),
        ("word/document.xml", document),
        ("word/styles.xml", styles),
        ("word/header1.xml", header),
    ];
    let bytes = ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap();
    let doc = EditingDoc::new(71);
    seed_from_docx(&doc, &bytes).unwrap();
    doc
}

fn text(value: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{value}</w:t></w:r>"#)
}

fn ctx() -> EditCtx {
    EditCtx::local("", "2026-10-21T00:00:00Z")
}

fn patch(doc: &EditingDoc) -> ParagraphPatchV1 {
    let plan = doc.paragraph_save_plan();
    let part = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/document.xml")
        .unwrap();
    assert_eq!(part.changed, vec![0]);
    assert!(part.legacy_paragraphs.is_empty());
    assert_eq!(part.paragraph_patches.len(), 1);
    part.paragraph_patches[0].clone()
}

fn replace(doc: &EditingDoc, from: u32, to: u32, value: &str) {
    doc.replace_range(&ctx(), StoryRange::new("body", from, to), value)
        .unwrap();
}

fn format(doc: &EditingDoc, from: u32, to: u32, delta: InlineFormatDelta) {
    doc.format_range(&ctx(), StoryRange::new("body", from, to), &delta)
        .unwrap();
}

fn target(value: &str) -> Vec<UnitV1> {
    vec![UnitV1::Text {
        text: value.to_owned(),
        attrs: Attrs::new(),
    }]
}

fn assert_text_edit(
    doc: &EditingDoc,
    from: u32,
    to: u32,
    value: &str,
    donor: u32,
    target: Vec<UnitV1>,
) {
    let patch = patch(doc);
    assert_eq!(patch.text_edits.len(), 1);
    let edit = &patch.text_edits[0];
    assert_eq!((edit.from, edit.to), (from, to));
    assert_eq!(edit.donor, DonorV1::At { at: donor });
    assert_eq!(
        edit.insert
            .iter()
            .map(|run| run.text.as_str())
            .collect::<String>(),
        value
    );
    assert!(patch.run_edits.is_empty());
    assert!(patch.para_deltas.is_empty());
    let expected_seed_hash = unit_sha256(&patch.seed.units);
    let expected_hash = result_sha256(&target, &patch.seed.properties);
    assert_eq!(patch.seed.unit_sha256, expected_seed_hash);
    assert_eq!(patch.result_sha256, expected_hash);
}

#[test]
fn replace_a_word() {
    let doc = open(&text("one cat"));
    replace(&doc, 4, 7, "dog");
    assert_text_edit(&doc, 4, 7, "dog", 4, target("one dog"));
    assert_eq!(
        serde_json::to_value(patch(&doc).text_edits).unwrap(),
        json!([{
            "from": 4,
            "to": 7,
            "insert": [{ "text": "dog", "attrs": {}, "inherit": {} }],
            "donor": { "kind": "at", "at": 4 }
        }])
    );
}

#[test]
fn insert_at_start_middle_and_end() {
    for (at, donor, result) in [(0, 0, "!abc"), (1, 0, "a!bc"), (3, 2, "abc!")] {
        let doc = open(&text("abc"));
        doc.insert_text(
            &ctx(),
            Position::new("body", at),
            "!",
            FormatPolicy::Inherit,
        )
        .unwrap();
        assert_text_edit(&doc, at, at, "!", donor, target(result));
    }
}

#[test]
fn insert_into_empty_paragraph_uses_pilcrow() {
    let doc = open("");
    doc.insert_text(&ctx(), Position::new("body", 0), "x", FormatPolicy::Inherit)
        .unwrap();
    assert_eq!(patch(&doc).text_edits[0].donor, DonorV1::Pilcrow);
}

#[test]
fn delete_a_range() {
    let doc = open(&text("abcXYZdef"));
    doc.delete_range(&ctx(), StoryRange::new("body", 3, 6))
        .unwrap();
    assert_text_edit(&doc, 3, 6, "", 3, target("abcdef"));
}

#[test]
fn replacement_spanning_runs_uses_first_text_donor() {
    let doc = open(
        r#"<w:r><w:rPr><w:b/></w:rPr><w:t>abc</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>def</w:t></w:r>"#,
    );
    replace(&doc, 1, 5, "XYZ");
    assert_text_edit(
        &doc,
        1,
        5,
        "XYZ",
        1,
        vec![
            UnitV1::Text {
                text: "aXYZ".to_owned(),
                attrs: Attrs::from([("bold".to_owned(), json!(true))]),
            },
            UnitV1::Text {
                text: "f".to_owned(),
                attrs: Attrs::from([("italic".to_owned(), json!(true))]),
            },
        ],
    );
    let patch = patch(&doc);
    assert_eq!(
        patch.text_edits[0].insert[0].attrs.get("bold"),
        Some(&json!(true))
    );
    assert!(!patch.text_edits[0].insert[0].attrs.contains_key("italic"));
}

#[test]
fn bold_subrange_has_target_and_inheritance() {
    let doc = open(&text("abcd"));
    format(
        &doc,
        1,
        3,
        InlineFormatDelta {
            bold: Patch::Set(true),
            ..Default::default()
        },
    );
    let patch = patch(&doc);
    assert!(patch.text_edits.is_empty());
    assert_eq!(patch.run_edits.len(), 1);
    let edit = &patch.run_edits[0];
    assert_eq!((edit.from, edit.to), (1, 3));
    assert_eq!(edit.target.get("bold"), Some(&json!(true)));
    assert!(!edit.inherit.contains_key("bold"));
    assert_eq!(edit.deltas.len(), 1);
    assert_eq!(edit.deltas[0].path, vec!["bold"]);
    assert_eq!(edit.deltas[0].before, SlotV1::Missing);
    assert_eq!(edit.deltas[0].after, SlotV1::Present { value: json!(true) });
}

#[test]
fn clearing_inherited_italic_carries_explicit_off_context() {
    let doc = open(&format!(
        r#"<w:pPr><w:pStyle w:val="Slant"/></w:pPr>{}"#,
        text("abc")
    ));
    format(
        &doc,
        0,
        3,
        InlineFormatDelta {
            italic: Patch::Clear,
            ..Default::default()
        },
    );
    let patch = patch(&doc);
    let edit = &patch.run_edits[0];
    assert_eq!(edit.inherit.get("italic"), Some(&json!(true)));
    assert!(!edit.target.contains_key("italic"));
    assert_eq!(edit.deltas[0].path, vec!["italic"]);
    assert_eq!(edit.deltas[0].after, SlotV1::Missing);
}

#[test]
fn font_change_diffs_only_one_slot() {
    let doc =
        open(r#"<w:r><w:rPr><w:rFonts w:ascii="Old" w:hAnsi="Keep"/></w:rPr><w:t>abc</w:t></w:r>"#);
    let slots = Any::Map(Arc::new(std::collections::HashMap::from([
        ("ascii".to_owned(), Any::from("New")),
        ("hAnsi".to_owned(), Any::from("Keep")),
    ])));
    format(
        &doc,
        0,
        3,
        InlineFormatDelta {
            other: BTreeMap::from([("fontFamily".to_owned(), Some(slots))]),
            ..Default::default()
        },
    );
    let patch = patch(&doc);
    assert_eq!(patch.run_edits[0].deltas.len(), 1);
    assert_eq!(
        patch.run_edits[0].deltas[0].path,
        vec!["fontFamily", "ascii"]
    );
}

#[test]
fn supplementary_plane_offsets_are_utf16() {
    let doc = open(&text("a😀bc"));
    replace(&doc, 3, 4, "X");
    assert_text_edit(&doc, 3, 4, "X", 3, target("a😀Xc"));
}

#[test]
fn chunk_split_without_semantic_change_has_no_edits() {
    let doc = open(&text("abcd"));
    format(
        &doc,
        1,
        3,
        InlineFormatDelta {
            bold: Patch::Set(true),
            ..Default::default()
        },
    );
    format(
        &doc,
        1,
        3,
        InlineFormatDelta {
            bold: Patch::Clear,
            ..Default::default()
        },
    );
    let plan = doc.paragraph_save_plan();
    for part in plan.spliced_parts {
        for patch in part.paragraph_patches {
            assert!(patch.text_edits.is_empty());
            assert!(patch.run_edits.is_empty());
            assert!(patch.para_deltas.is_empty());
        }
    }
}

#[test]
fn paragraph_style_change_includes_property_resets_and_run_sweep() {
    let doc = open(
        r#"<w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:rPr><w:i/></w:rPr><w:t>abc</w:t></w:r>"#,
    );
    let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
    doc.apply_edits(
        &EditRequest {
            expect_version: doc.version(),
            source: EditSource::Host,
            history: EditHistory::Separate,
            steps: vec![EditStep::new(EditOperation::SetParagraphStyle {
                target: ParagraphTarget {
                    story: "body".to_owned(),
                    para_id,
                },
                style_id: "Strong".to_owned(),
            })],
        },
        &UndoSession::new(),
    )
    .unwrap()
    .unwrap();
    let patch = patch(&doc);
    assert!(
        patch
            .para_deltas
            .iter()
            .any(|delta| delta.path == ["pStyle"])
    );
    assert!(
        patch
            .para_deltas
            .iter()
            .any(|delta| delta.path == ["alignment"])
    );
    assert!(patch.run_edits.iter().any(|edit| {
        edit.deltas.iter().any(|delta| delta.path == ["italic"])
            && edit.target.get("bold") == Some(&json!(true))
    }));
    assert_eq!(patch.para_inherit.get("keepNext"), Some(&json!(true)));
}

fn assert_legacy(doc: &EditingDoc, code: PatchErrorCode, prefix: &str) {
    let plan = doc.paragraph_save_plan();
    let part = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/document.xml")
        .unwrap();
    assert_eq!(part.changed, vec![0]);
    assert!(part.paragraph_patches.is_empty());
    assert_eq!(part.legacy_paragraphs.len(), 1);
    assert_eq!(part.legacy_paragraphs[0].ordinal, 0);
    assert_eq!(part.legacy_paragraphs[0].code, code);
    assert!(part.legacy_paragraphs[0].detail.starts_with(prefix));
}

#[test]
fn empty_seeded_font_family_checks_changed_leaves() {
    for (key, value) in [("hint", "cs"), ("ascii", "Serif")] {
        let doc = open(r#"<w:r><w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr><w:t>abc</w:t></w:r>"#);
        let slots = Any::Map(Arc::new(std::collections::HashMap::from([(
            key.to_owned(),
            Any::from(value),
        )])));
        format(
            &doc,
            0,
            3,
            InlineFormatDelta {
                other: BTreeMap::from([("fontFamily".to_owned(), Some(slots))]),
                ..Default::default()
            },
        );
        if key == "hint" {
            assert_legacy(&doc, PatchErrorCode::UnsupportedKey, "key:fontFamily.hint");
        } else {
            let patch = patch(&doc);
            assert_eq!(
                serde_json::to_value(&patch.run_edits).unwrap(),
                json!([{
                    "from": 0, "to": 3,
                    "deltas": [{
                        "path": ["fontFamily"],
                        "before": { "state": "present", "value": {} },
                        "after": { "state": "present", "value": { "ascii": "Serif" } }
                    }],
                    "target": { "fontFamily": { "ascii": "Serif" } }, "inherit": {}
                }])
            );
        }
    }
}

#[test]
fn unknown_other_key_is_legacy() {
    let doc = open(&text("abc"));
    format(
        &doc,
        0,
        3,
        InlineFormatDelta {
            other: BTreeMap::from([("unknown".to_owned(), Some(Any::Bool(true)))]),
            ..Default::default()
        },
    );
    assert_legacy(&doc, PatchErrorCode::UnsupportedKey, "key:unknown");
}

#[test]
fn field_paragraph_is_legacy() {
    let doc = open(
        r#"<w:r><w:t>abc</w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>"#,
    );
    replace(&doc, 0, 1, "X");
    assert_legacy(&doc, PatchErrorCode::UnsupportedStructure, "field:");
}

#[test]
fn repeated_text_alignment_is_legacy() {
    let doc = open(&text("aaaa"));
    doc.delete_range(&ctx(), StoryRange::new("body", 1, 2))
        .unwrap();
    assert_legacy(&doc, PatchErrorCode::AmbiguousAlignment, "ambiguous:");
}

#[test]
fn successive_save_plans_keep_the_opening_seed() {
    let doc = open(&text("abc"));
    replace(&doc, 1, 2, "X");
    let first = patch(&doc);
    assert!(
        doc.record_saved_paragraph_ids(&[(first.session_key.clone(), "00000001".to_owned())])
            .is_empty()
    );
    replace(&doc, 1, 2, "Y");
    let second = patch(&doc);
    assert_eq!(first.seed, second.seed);
    assert!(matches!(&second.seed.units[0], UnitV1::Text { text, .. } if text == "abc"));
    assert_text_edit(&doc, 1, 2, "Y", 1, target("aYc"));
}

#[test]
fn changed_header_paragraph_gets_a_patch() {
    let doc = open(&text("body"));
    let identity = doc.paragraph_identities();
    let story = identity
        .paragraphs
        .iter()
        .find_map(|identity| match &identity.paragraph {
            ParagraphRef::Session { story, .. } if story != "body" => Some(story.clone()),
            _ => None,
        })
        .unwrap();
    doc.replace_range(&ctx(), StoryRange::new(story, 1, 2), "X")
        .unwrap();
    let plan = doc.paragraph_save_plan();
    let header = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/header1.xml")
        .unwrap();
    assert_eq!(header.changed, vec![0]);
    assert_eq!(header.paragraph_patches.len(), 1);
    assert!(header.legacy_paragraphs.is_empty());
    assert_eq!(header.paragraph_patches[0].text_edits[0].from, 1);
}

#[test]
fn deleting_all_text_keeps_an_empty_target() {
    let doc = open(&text("abc"));
    doc.delete_range(&ctx(), StoryRange::new("body", 0, 3))
        .unwrap();
    assert_text_edit(&doc, 0, 3, "", 0, Vec::new());
    assert!(patch(&doc).text_edits[0].insert.is_empty());
}

#[test]
fn unexplained_carrier_change_is_legacy() {
    let doc = open(&text("abc"));
    let key = doc.paragraphs("body").unwrap()[0].para_id.clone();
    doc.set_paragraph_attr(&key, "_unknownCarrier", Any::Bool(true))
        .unwrap();
    assert_legacy(&doc, PatchErrorCode::UnsupportedKey, "key:_unknownCarrier");
}

#[test]
fn note_atom_is_retained_between_text_edits() {
    let doc = open(
        r#"<w:r><w:t>abc</w:t></w:r><w:r><w:footnoteReference w:id="1"/></w:r><w:r><w:t>def</w:t></w:r>"#,
    );
    replace(&doc, 1, 2, "X");
    let patch = patch(&doc);
    assert_eq!(patch.text_edits[0].from, 1);
    assert!(
        patch.seed.units.iter().any(|unit| {
            matches!(unit, UnitV1::Atom { atom_kind, .. } if atom_kind == "noteRef")
        })
    );
}

#[test]
fn hyperlink_mutation_is_legacy() {
    let doc = open(&text("abc"));
    let link = Any::Map(Arc::new(std::collections::HashMap::from([(
        "href".to_owned(),
        Any::from("#target"),
    )])));
    use yrs::{Map, Out, ReadTxn, Text, Transact};

    let mut txn = doc.yrs_doc().transact_mut();
    let stories = txn.get_map("stories").unwrap();
    let Out::YText(story) = stories.get(&txn, "body").unwrap() else {
        panic!("missing body");
    };
    story.format(
        &mut txn,
        0,
        3,
        [(Arc::from("hyperlink"), link)].into_iter().collect(),
    );
    drop(txn);
    assert_legacy(&doc, PatchErrorCode::UnsupportedStructure, "hyperlink:");
}

#[test]
fn numeric_marks_survive_the_json_transport() {
    let doc = open(&text("abc"));
    format(
        &doc,
        0,
        3,
        InlineFormatDelta {
            font_size: Patch::Set(15.0),
            ..Default::default()
        },
    );
    let patch = patch(&doc);
    let edit = &patch.run_edits[0];
    assert_eq!(
        edit.target.get("fontSize"),
        Some(&json!({ "size": 30, "sizeCs": 30 }))
    );
    assert_eq!(edit.deltas[0].after, SlotV1::Present { value: json!(30) });
}

#[test]
fn edits_strictly_between_bookmarks_keep_the_markers() {
    let source = concat!(
        r#"<w:bookmarkStart w:id="1" w:name="mark"/>"#,
        "<w:r><w:t>abc</w:t></w:r>",
        r#"<w:bookmarkEnd w:id="1"/>"#,
    );
    let doc = open(source);
    replace(&doc, 1, 2, "X");
    assert_text_edit(&doc, 1, 2, "X", 1, target("aXc"));
    let inserted = open(source);
    inserted
        .insert_text(&ctx(), Position::new("body", 1), "!", FormatPolicy::Inherit)
        .unwrap();
    assert_legacy(&inserted, PatchErrorCode::UnsupportedStructure, "marker:");
    let touching = open(source);
    replace(&touching, 0, 1, "X");
    assert_legacy(&touching, PatchErrorCode::UnsupportedStructure, "marker:");
}

#[test]
fn style_change_cannot_hide_a_direct_alignment_removal() {
    let doc = open(&format!(
        r#"<w:pPr><w:jc w:val="right"/></w:pPr>{}"#,
        text("abc")
    ));
    let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
    doc.apply_edits(
        &EditRequest {
            expect_version: doc.version(),
            source: EditSource::Host,
            history: EditHistory::Separate,
            steps: vec![EditStep::new(EditOperation::SetParagraphStyle {
                target: ParagraphTarget {
                    story: "body".to_owned(),
                    para_id,
                },
                style_id: "Right".to_owned(),
            })],
        },
        &UndoSession::new(),
    )
    .unwrap()
    .unwrap();
    assert_legacy(&doc, PatchErrorCode::UnsupportedKey, "key:alignment");
}

#[test]
fn insertion_before_bookmarks_rejects_stale_live_offsets() {
    let doc = open(concat!(
        "<w:r><w:t>a</w:t></w:r>",
        r#"<w:bookmarkStart w:id="1" w:name="mark"/>"#,
        "<w:r><w:t>bc</w:t></w:r>",
        r#"<w:bookmarkEnd w:id="1"/>"#,
        "<w:r><w:t>d</w:t></w:r>",
    ));
    doc.insert_text(&ctx(), Position::new("body", 0), "X", FormatPolicy::Inherit)
        .unwrap();
    assert_legacy(&doc, PatchErrorCode::UnsupportedStructure, "marker:");
}

#[test]
fn numbering_changes_have_exact_leaf_deltas_and_hash() {
    let doc = open(&format!(
        r#"<w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr></w:pPr>{}"#,
        text("abc")
    ));
    let key = doc.paragraphs("body").unwrap()[0].para_id.clone();
    doc.set_paragraph_attr(
        &key,
        "numPr",
        Any::Map(Arc::new(std::collections::HashMap::from([
            ("numId".to_owned(), Any::Number(4.0)),
            ("ilvl".to_owned(), Any::Number(2.0)),
        ]))),
    )
    .unwrap();
    let patch = patch(&doc);
    assert_eq!(
        serde_json::to_value(&patch.para_deltas).unwrap(),
        json!([
            { "path": ["numPr", "ilvl"],
              "before": { "state": "present", "value": 1 },
              "after": { "state": "present", "value": 2 } },
            { "path": ["numPr", "numId"],
              "before": { "state": "present", "value": 3 },
              "after": { "state": "present", "value": 4 } }
        ])
    );
    let mut properties = patch.seed.properties.clone();
    properties.insert("numPr".to_owned(), json!({ "numId": 4, "ilvl": 2 }));
    let expected_hash = result_sha256(&target("abc"), &properties);
    assert_eq!(patch.result_sha256, expected_hash);
}

#[test]
fn underline_and_strike_have_exact_scripts_and_hashes() {
    for (key, delta, expected) in [
        (
            "underline",
            InlineFormatDelta {
                underline: Patch::Set(UnderlinePatch {
                    style: Some("double".to_owned()),
                    color: None,
                }),
                ..Default::default()
            },
            json!({ "path": ["underline", "style"],
                "before": { "state": "missing" },
                "after": { "state": "present", "value": "double" } }),
        ),
        (
            "strike",
            InlineFormatDelta {
                strike: Patch::Set(StrikePatch { double: true }),
                ..Default::default()
            },
            json!({ "path": ["strike", "double"],
                "before": { "state": "missing" },
                "after": { "state": "present", "value": true } }),
        ),
    ] {
        let doc = open(&text("abc"));
        format(&doc, 0, 3, delta);
        let patch = patch(&doc);
        let attrs = if key == "underline" {
            json!({ "underline": { "style": "double" } })
        } else {
            json!({ "strike": { "double": true } })
        };
        assert_eq!(
            serde_json::to_value(&patch.run_edits).unwrap(),
            json!([{
                "from": 0, "to": 3, "deltas": [expected], "target": attrs, "inherit": {}
            }])
        );
        let units = vec![UnitV1::Text {
            text: "abc".to_owned(),
            attrs: serde_json::from_value(attrs).unwrap(),
        }];
        let expected_hash = result_sha256(&units, &patch.seed.properties);
        assert_eq!(patch.result_sha256, expected_hash);
    }
}

#[test]
fn spacing_and_indent_clears_have_exact_scripts_and_hash() {
    let doc = open(&format!(
        r#"<w:pPr><w:spacing w:before="120"/><w:ind w:left="720"/></w:pPr>{}"#,
        text("abc")
    ));
    let key = doc.paragraphs("body").unwrap()[0].para_id.clone();
    doc.set_paragraph_attrs(
        &ctx(),
        &ParaSelector::One(key),
        &ParaAttrDelta {
            space_before: Patch::Clear,
            indent_left: Patch::Clear,
            ..Default::default()
        },
    )
    .unwrap();
    let patch = patch(&doc);
    assert_eq!(
        serde_json::to_value(&patch.para_deltas).unwrap(),
        json!([
            { "path": ["indentLeft"],
              "before": { "state": "present", "value": 720 },
              "after": { "state": "missing" } },
            { "path": ["spaceBefore"],
              "before": { "state": "present", "value": 120 },
              "after": { "state": "missing" } }
        ])
    );
    let mut properties = patch.seed.properties.clone();
    properties.remove("indentLeft");
    properties.remove("spaceBefore");
    let expected_hash = result_sha256(&target("abc"), &properties);
    assert_eq!(patch.result_sha256, expected_hash);
}

#[test]
fn unsupported_source_structures_have_typed_reason_prefixes() {
    for (content, prefix) in [
        (
            r#"<w:r><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:rect style="width:20pt;height:10pt"/></w:pict></w:r>"#,
            "drawing:",
        ),
        (
            r#"<w:sdt><w:sdtPr><w:id w:val="1"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>x</w:t></w:r></w:sdtContent></w:sdt>"#,
            "sdt:",
        ),
        (
            r#"<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>x</m:t></m:r></m:oMath>"#,
            "math:",
        ),
        (
            r#"<w:ins w:id="1" w:author="editor" w:date="2026-10-21T00:00:00Z"><w:r><w:t>x</w:t></w:r></w:ins>"#,
            "tc:",
        ),
        (
            r#"<w:r><w:rPr><w:rPrChange w:id="1" w:author="editor"><w:rPr><w:b/></w:rPr></w:rPrChange></w:rPr><w:t>x</w:t></w:r>"#,
            "tc:",
        ),
    ] {
        let doc = open(&format!("{}{content}", text("abc")));
        replace(&doc, 1, 2, "X");
        assert_legacy(&doc, PatchErrorCode::UnsupportedStructure, prefix);
    }
}
