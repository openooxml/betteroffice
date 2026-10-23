use std::collections::BTreeMap;
use std::sync::Arc;

use docx_edit::{
    EditCtx, EditHistory, EditOperation, EditRequest, EditSource, EditStep, EditingDoc,
    FormatPolicy, InlineFormatDelta, ParaAttrDelta, ParaSelector, ParagraphRef, ParagraphTarget,
    Patch, Position, StoryRange, StrikePatch, StyleResolver, UnderlinePatch, UndoSession,
    seed_from_docx,
};
use docx_parse::paragraph_patch::{
    Attrs, DonorV1, ParagraphPatchV1, PatchErrorCode, SlotV1, UnitV1, result_sha256, unit_sha256,
};
use docx_parse::serializer::{
    S13SaveRequest, SerializerContext, SerializerDeterminism, element_span, serialize_paragraph,
    write_docx_s13_with_warnings,
};
use serde_json::json;
use yrs::Any;

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);

fn open(content: &str) -> EditingDoc {
    let doc = EditingDoc::new(71);
    seed_from_docx(&doc, &package(content)).unwrap();
    doc
}

fn package(content: &str) -> Vec<u8> {
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
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
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

fn save_document(
    source: &docx_parse::S9DocumentBodyWire,
    edit_model: impl FnOnce(&mut serde_json::Value),
) -> serde_json::Value {
    let mut document = serde_json::to_value(source).unwrap();
    edit_model(&mut document);
    if let Some(sections) = &source.sections {
        let sections: Vec<_> = sections
            .iter()
            .map(|section| {
                let content = document["content"].as_array().unwrap();
                json!({
                    "id": section.id,
                    "properties": section.properties,
                    "content": &content[section.content_start..section.content_end]
                })
            })
            .collect();
        document["sections"] = json!(sections);
    }
    document
}

fn patched_parts(plan: &docx_edit::ParagraphSavePlan) -> Vec<serde_json::Value> {
    plan.patched_parts
        .iter()
        .map(|(part, para_ids)| json!({"part": part, "paraIds": para_ids}))
        .collect()
}

fn save_paragraph(
    doc: &EditingDoc,
    content: &str,
    edit_model: impl FnOnce(&mut serde_json::Value),
) -> (String, Vec<String>) {
    let bytes = package(content);
    let plan = doc.paragraph_save_plan();
    let part = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/document.xml")
        .unwrap();
    let package = docx_parse::parse_docx_s9_wire(
        &bytes,
        docx_parse::S9ParseOptions {
            source_ordinals: true,
            ..Default::default()
        },
    )
    .unwrap()
    .document
    .package;
    let document = save_document(&package.document, |document| {
        edit_model(&mut document["content"][0]);
    });
    let request = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": document,
        "headerEntries": package.header_entries.unwrap_or_default(),
        "relationshipEntries": package.relationship_entries,
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"patchedParts": patched_parts(&plan), "splicedParts": [{
            "part": part.part, "sha256": part.sha256,
            "paragraphs": [0], "changed": part.changed,
            "paragraphPatches": part.paragraph_patches,
            "legacyParagraphs": part.legacy_paragraphs
        }]}
    }))
    .unwrap();
    let (saved, warnings) = write_docx_s13_with_warnings(request, &bytes).unwrap();
    let (_, xml) = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    (String::from_utf8(xml).unwrap(), warnings)
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
fn text_edit_keeps_paragraph_mark_formatting_in_verification_seed() {
    for (properties, expected) in [
        (r#"<w:sz w:val="32"/>"#, json!({ "fontSize": 32 })),
        (
            concat!(
                r#"<w:b/><w:sz w:val="32"/>"#,
                r#"<w:rFonts w:ascii="Serif" w:hAnsi="Keep"/>"#,
            ),
            json!({
                "bold": true, "fontSize": 32,
                "fontFamily": { "ascii": "Serif", "hAnsi": "Keep" }
            }),
        ),
    ] {
        let doc = open(&format!(
            r#"<w:pPr><w:pStyle w:val="Normal"/><w:rPr>{properties}</w:rPr></w:pPr>{}"#,
            text("abc")
        ));
        replace(&doc, 1, 2, "X");
        let patch = patch(&doc);
        assert_eq!(patch.seed.properties.get("pStyle"), Some(&json!("Normal")));
        assert_eq!(
            patch.seed.properties.get("defaultTextFormatting"),
            Some(&expected)
        );
        assert_text_edit(&doc, 1, 2, "X", 1, target("aXc"));
    }
}

#[test]
fn paragraph_mark_formatting_change_is_legacy() {
    for content in ["", "<w:r><w:t>abc</w:t></w:r>"] {
        let doc = open(&format!(
            r#"<w:pPr><w:rPr><w:sz w:val="32"/><w:b/></w:rPr></w:pPr>{content}"#
        ));
        let key = doc.paragraphs("body").unwrap()[0].para_id.clone();
        doc.set_paragraph_default_format(
            &ctx(),
            &ParaSelector::One(key),
            Some(&BTreeMap::from([
                ("fontSize".to_owned(), Any::Number(40.0)),
                ("bold".to_owned(), Any::Bool(true)),
            ])),
        )
        .unwrap();
        assert_legacy(
            &doc,
            PatchErrorCode::UnsupportedKey,
            "key:defaultTextFormatting.fontSize",
        );
    }
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
fn font_slot_change_on_theme_bound_run_is_legacy() {
    for theme in ["asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"] {
        for slot in ["ascii", "hAnsi", "eastAsia", "cs"] {
            let doc = open(&format!(
                r#"<w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:eastAsia="Arial" w:cs="Arial" w:{theme}="minorHAnsi"/></w:rPr><w:t>abc</w:t></w:r>"#
            ));
            let theme_key = if theme == "cstheme" { "csTheme" } else { theme };
            let mut slots = std::collections::HashMap::from([
                ("ascii".to_owned(), Any::from("Arial")),
                ("hAnsi".to_owned(), Any::from("Arial")),
                ("eastAsia".to_owned(), Any::from("Arial")),
                ("cs".to_owned(), Any::from("Arial")),
                (theme_key.to_owned(), Any::from("minorHAnsi")),
            ]);
            slots.insert(slot.to_owned(), Any::from("Courier New"));
            format(
                &doc,
                0,
                3,
                InlineFormatDelta {
                    other: BTreeMap::from([(
                        "fontFamily".to_owned(),
                        Some(Any::Map(Arc::new(slots))),
                    )]),
                    ..Default::default()
                },
            );
            assert_ambiguous(&doc, &format!("ambiguous: theme font {slot}"));
        }
    }
}

const THEME_FONTS: &str = concat!(
    "<w:rFonts  w:hAnsiTheme='minorHAnsi' w:asciiTheme = 'minorHAnsi' ",
    "w:eastAsiaTheme=\"minorEastAsia\" w:cstheme='minorBidi'/>"
);

fn theme_run() -> String {
    format!(
        "<w:r><w:rPr>{THEME_FONTS}</w:rPr><w:t>one cat today</w:t></w:r>{}",
        text("!")
    )
}

fn assert_theme_fonts_unchanged(output: &str) {
    assert!(output.contains(THEME_FONTS), "{output}");
    assert_eq!(
        output.matches("<w:rFonts").count(),
        output.matches(THEME_FONTS).count(),
        "{output}"
    );
}

#[test]
fn theme_run_text_correction_still_patches() {
    let content = theme_run();
    let doc = open(&content);
    replace(&doc, 4, 7, "dog");
    let patch = patch(&doc);
    assert_eq!(patch.text_edits.len(), 1);
    assert!(patch.run_edits.is_empty());
    assert_eq!(patch.text_edits[0].donor, DonorV1::At { at: 4 });
    let (output, warnings) = save_paragraph(&doc, &content, |paragraph| {
        paragraph["content"][0]["content"][0]["text"] = json!("one dog today");
    });
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains(">one dog today</w:t>"), "{output}");
    assert_theme_fonts_unchanged(&output);
}

#[test]
fn theme_run_boundary_insertion_still_patches() {
    let content = theme_run();
    let doc = open(&content);
    doc.insert_text(
        &ctx(),
        Position::new("body", 13),
        "#",
        FormatPolicy::Inherit,
    )
    .unwrap();
    let patch = patch(&doc);
    assert_eq!(patch.text_edits.len(), 1);
    assert!(patch.run_edits.is_empty());
    assert_eq!(patch.text_edits[0].donor, DonorV1::At { at: 12 });
    let UnitV1::Text { attrs, .. } = &patch.seed.units[0] else {
        panic!("missing theme run");
    };
    assert_eq!(patch.text_edits[0].insert[0].attrs, *attrs);
    let (output, warnings) = save_paragraph(&doc, &content, |paragraph| {
        paragraph["content"][0]["content"][0]["text"] = json!("one cat today#");
    });
    assert!(warnings.is_empty(), "{warnings:?}");
    assert!(output.contains(">one cat today#</w:t>"), "{output}");
    assert_theme_fonts_unchanged(&output);
}

fn assert_theme_run_format_still_patches(key: &str, delta: InlineFormatDelta) {
    let content = theme_run();
    let doc = open(&content);
    format(&doc, 4, 7, delta);
    let patch = patch(&doc);
    assert!(patch.text_edits.is_empty());
    assert_eq!(patch.run_edits.len(), 1);
    let edit = &patch.run_edits[0];
    assert_eq!((edit.from, edit.to), (4, 7));
    assert_eq!(edit.deltas.len(), 1);
    assert_eq!(edit.deltas[0].path, vec![key]);
    let (output, warnings) = save_paragraph(&doc, &content, |paragraph| {
        let run = paragraph["content"][0].clone();
        let mut before = run.clone();
        let mut phrase = run.clone();
        let mut after = run;
        before["content"][0]["text"] = json!("one ");
        phrase["content"][0]["text"] = json!("cat");
        phrase["formatting"][key] = json!(true);
        after["content"][0]["text"] = json!(" today");
        paragraph["content"] = json!([before, phrase, after, paragraph["content"][1]]);
    });
    assert!(warnings.is_empty(), "{warnings:?}");
    let tag = if key == "bold" { "b" } else { "i" };
    assert!(output.contains(&format!("<w:{tag}/>")), "{output}");
    assert!(output.contains(">cat</w:t>"), "{output}");
    assert_theme_fonts_unchanged(&output);
}

#[test]
fn theme_run_bold_phrase_still_patches() {
    assert_theme_run_format_still_patches(
        "bold",
        InlineFormatDelta {
            bold: Patch::Set(true),
            ..Default::default()
        },
    );
}

#[test]
fn theme_run_italic_phrase_still_patches() {
    assert_theme_run_format_still_patches(
        "italic",
        InlineFormatDelta {
            italic: Patch::Set(true),
            ..Default::default()
        },
    );
}

#[test]
fn bold_toggle_off_on_non_theme_run_still_patches() {
    let doc = open(r#"<w:r><w:rPr><w:b/></w:rPr><w:t>abc</w:t></w:r>"#);
    format(
        &doc,
        0,
        3,
        InlineFormatDelta {
            bold: Patch::Set(false),
            ..Default::default()
        },
    );
    let patch = patch(&doc);
    assert_eq!(patch.run_edits.len(), 1);
    assert_eq!(patch.run_edits[0].deltas[0].path, vec!["bold"]);
    assert_eq!(patch.run_edits[0].deltas[0].after, SlotV1::Missing);
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

const RESTYLE_NEIGHBOURS: [&str; 3] = [
    "<w:p w:rsidR='11111111'><w:r><w:t>before</w:t></w:r></w:p>",
    "<w:p w:rsidR='22222222'><w:r><w:t>after</w:t></w:r></w:p>",
    "<w:p w:rsidR='33333333'><w:r><w:t>untouched</w:t></w:r></w:p>",
];

fn restyle_source(content: &str, a: &str, b: &str) -> (EditingDoc, Vec<u8>, StyleResolver) {
    let mut parts = ooxml_opc::unzip_parts(&package(content)).unwrap();
    for (name, bytes) in &mut parts {
        if name == "word/document.xml" {
            *bytes = format!(
                "<w:document {NS}><w:body>{}<w:p w14:paraId=\"00000001\">\
                 {content}</w:p>{}{}</w:body></w:document>",
                RESTYLE_NEIGHBOURS[0], RESTYLE_NEIGHBOURS[1], RESTYLE_NEIGHBOURS[2]
            )
            .into_bytes();
        } else if name == "word/styles.xml" {
            *bytes = format!(
                "<w:styles {NS}><w:style w:type=\"paragraph\" w:default=\"1\" \
                 w:styleId=\"Normal\"><w:name w:val=\"Normal\"/></w:style>\
                 <w:style w:type=\"paragraph\" w:styleId=\"A\">\
                 <w:name w:val=\"A\"/>{a}</w:style>\
                 <w:style w:type=\"paragraph\" w:styleId=\"B\">\
                 <w:name w:val=\"B\"/>{b}</w:style>\
                 <w:style w:type=\"character\" w:styleId=\"C\">\
                 <w:name w:val=\"C\"/></w:style></w:styles>"
            )
            .into_bytes();
        }
    }
    let styles = StyleResolver::from_parts(&parts).unwrap();
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    let doc = EditingDoc::new(71);
    seed_from_docx(&doc, &bytes).unwrap();
    (doc, bytes, styles)
}

fn seed_format_request(bytes: &[u8]) -> S13SaveRequest {
    let doc = EditingDoc::new(71);
    seed_from_docx(&doc, bytes).unwrap();
    replace(&doc, 7, 8, "y");
    let plan = doc.paragraph_save_plan();
    let part = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/document.xml")
        .unwrap();
    assert_eq!(part.changed, vec![1]);
    assert!(part.legacy_paragraphs.is_empty());
    assert_eq!(part.paragraph_patches.len(), 1);
    let package = docx_parse::parse_docx_s9_wire(
        bytes,
        docx_parse::S9ParseOptions {
            source_ordinals: true,
            ..Default::default()
        },
    )
    .unwrap()
    .document
    .package;
    let document = save_document(&package.document, |document| {
        document["content"][1]["content"][0]["content"][0]["text"] = json!("y");
    });
    serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": document,
        "headerEntries": package.header_entries.unwrap_or_default(),
        "relationshipEntries": package.relationship_entries,
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"patchedParts": patched_parts(&plan), "splicedParts": [{
            "part": part.part, "sha256": part.sha256,
            "paragraphs": [0, 1, 2, 3], "changed": [1],
            "paragraphPatches": part.paragraph_patches
        }]}
    }))
    .unwrap()
}

fn assert_seed_format_replace(bytes: &[u8], inspect: impl FnOnce(&ParagraphPatchV1)) {
    let request = seed_format_request(bytes);
    let part = &request.paragraph_ids.as_ref().unwrap().spliced_parts[0];
    inspect(&part.paragraph_patches[0]);
    let (saved, warnings) = write_docx_s13_with_warnings(request, bytes).unwrap();
    assert_eq!(
        warnings,
        ["DOCX_MINIMAL_PATCH replace counts {\"ambiguous\":1}"]
    );
    let saved_parts = ooxml_opc::unzip_parts(&saved).unwrap();
    for (name, original) in ooxml_opc::unzip_parts(bytes).unwrap() {
        let output = &saved_parts
            .iter()
            .find(|(path, _)| path == &name)
            .unwrap()
            .1;
        if name == "word/document.xml" {
            let original = std::str::from_utf8(&original).unwrap();
            let output = std::str::from_utf8(output).unwrap();
            let start = original.find("<w:p w14:paraId=\"00000001\">").unwrap();
            let end = start + original[start..].find("</w:p>").unwrap() + "</w:p>".len();
            assert!(output.starts_with(&original[..start]), "{output}");
            assert!(output.ends_with(&original[end..]), "{output}");
            assert!(output.contains(">y</w:t>"), "{output}");
        } else {
            assert_eq!(output, &original, "{name}");
        }
    }
}

#[test]
fn seed_format_inherited_color_with_direct_auto_uses_counted_replace() {
    let (_, bytes, _) = restyle_source(
        concat!(
            "<w:pPr><w:pStyle w:val=\"A\"/></w:pPr>",
            "<w:r><w:rPr><w:color w:val=\"auto\"/></w:rPr><w:t>x</w:t></w:r>",
        ),
        "<w:rPr><w:color w:val=\"FF0000\"/></w:rPr>",
        "",
    );
    assert_seed_format_replace(&bytes, |patch| {
        let UnitV1::Text { attrs, .. } = &patch.seed.units[0] else {
            panic!("expected text seed");
        };
        assert_eq!(attrs["textColor"]["rgb"], json!("FF0000"));
    });
}

#[test]
fn seed_format_inherited_highlight_with_direct_shading_uses_counted_replace() {
    let (_, bytes, _) = restyle_source(
        concat!(
            "<w:pPr><w:pStyle w:val=\"A\"/></w:pPr>",
            "<w:r><w:rPr><w:shd w:val=\"clear\" w:fill=\"FF0000\"/></w:rPr>",
            "<w:t>x</w:t></w:r>",
        ),
        "<w:rPr><w:highlight w:val=\"yellow\"/></w:rPr>",
        "",
    );
    assert_seed_format_replace(&bytes, |patch| {
        let UnitV1::Text { attrs, .. } = &patch.seed.units[0] else {
            panic!("expected text seed");
        };
        assert_eq!(attrs["highlight"], json!("yellow"));
    });
}

#[test]
fn seed_format_numbering_indent_with_direct_zero_uses_counted_replace() {
    let (_, bytes, _) = restyle_source(
        concat!(
            "<w:pPr><w:numPr><w:ilvl w:val=\"0\"/><w:numId w:val=\"1\"/></w:numPr>",
            "<w:ind w:firstLine=\"0\"/></w:pPr><w:r><w:t>x</w:t></w:r>",
        ),
        "",
        "",
    );
    let mut parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    for (path, bytes) in &mut parts {
        if path == "word/_rels/document.xml.rels" {
            *bytes = std::str::from_utf8(bytes)
                .unwrap()
                .replace(
                    "</Relationships>",
                    concat!(
                        "<Relationship Id=\"numbering\" Type=\"http://schemas.openxmlformats.org/",
                        "officeDocument/2006/relationships/numbering\" Target=\"numbering.xml\"/>",
                        "</Relationships>",
                    ),
                )
                .into_bytes();
        } else if path == "[Content_Types].xml" {
            *bytes = std::str::from_utf8(bytes)
                .unwrap()
                .replace(
                    "</Types>",
                    concat!(
                        "<Override PartName=\"/word/numbering.xml\" ContentType=\"application/",
                        "vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml\"/>",
                        "</Types>",
                    ),
                )
                .into_bytes();
        }
    }
    let numbering = format!(
        "<w:numbering {NS}><w:abstractNum w:abstractNumId=\"0\"><w:lvl w:ilvl=\"0\">\
         <w:start w:val=\"1\"/><w:numFmt w:val=\"decimal\"/><w:lvlText w:val=\"%1.\"/>\
         <w:pPr><w:ind w:left=\"720\" w:hanging=\"360\"/></w:pPr></w:lvl></w:abstractNum>\
         <w:num w:numId=\"1\"><w:abstractNumId w:val=\"0\"/></w:num></w:numbering>"
    );
    parts.push(("word/numbering.xml".to_owned(), numbering.into_bytes()));
    let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
    assert_seed_format_replace(&bytes, |patch| {
        assert_eq!(patch.seed.properties["indentFirstLine"], json!(-360));
        assert_eq!(patch.seed.properties["hangingIndent"], json!(true));
    });
}

fn assert_restyle_replace(
    doc: &EditingDoc,
    bytes: &[u8],
    edit_model: impl FnOnce(&mut serde_json::Value),
) {
    let plan = doc.paragraph_save_plan();
    let part = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/document.xml")
        .unwrap();
    assert_eq!(part.changed, vec![1]);
    assert!(part.paragraph_patches.is_empty());
    assert_eq!(part.legacy_paragraphs.len(), 1);
    assert_eq!(part.legacy_paragraphs[0].ordinal, 1);
    assert_eq!(
        part.legacy_paragraphs[0].code,
        PatchErrorCode::UnsupportedStructure
    );
    assert!(part.legacy_paragraphs[0].detail.starts_with("restyle:"));
    let package = docx_parse::parse_docx_s9_wire(
        bytes,
        docx_parse::S9ParseOptions {
            source_ordinals: true,
            ..Default::default()
        },
    )
    .unwrap()
    .document
    .package;
    let document = save_document(&package.document, |document| {
        edit_model(&mut document["content"][1]);
    });
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": document,
        "headerEntries": package.header_entries.unwrap_or_default(),
        "relationshipEntries": package.relationship_entries,
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"patchedParts": patched_parts(&plan), "splicedParts": [{
            "part": part.part, "sha256": part.sha256,
            "paragraphs": [0, 1, 2, 3], "changed": [1],
            "paragraphPatches": part.paragraph_patches,
            "legacyParagraphs": part.legacy_paragraphs
        }]}
    }))
    .unwrap();
    let (saved, warnings) = write_docx_s13_with_warnings(request, bytes).unwrap();
    assert_eq!(
        warnings,
        ["DOCX_MINIMAL_PATCH replace counts {\"restyle\":1}"]
    );
    let saved_parts = ooxml_opc::unzip_parts(&saved).unwrap();
    for (name, original) in ooxml_opc::unzip_parts(bytes).unwrap() {
        let output = &saved_parts
            .iter()
            .find(|(path, _)| path == &name)
            .unwrap()
            .1;
        if name == "word/document.xml" {
            let output = std::str::from_utf8(output).unwrap();
            for neighbour in RESTYLE_NEIGHBOURS {
                assert!(output.contains(neighbour), "{output}");
            }
        } else {
            assert_eq!(output, &original, "{name}");
        }
    }
}

fn assert_paragraph_restyle(content: &str, a: &str, b: &str) {
    let (doc, bytes, styles) = restyle_source(content, a, b);
    let para_id = doc.paragraphs("body").unwrap()[1].para_id.clone();
    doc.restyle_paragraphs(&ctx(), &ParaSelector::One(para_id), "B", &styles)
        .unwrap();
    assert_restyle_replace(&doc, &bytes, |paragraph| {
        paragraph["formatting"]["styleId"] = json!("B");
    });
}

#[test]
fn restyle_complex_script_toggles_replace() {
    for toggle in ["bCs", "iCs"] {
        assert_paragraph_restyle(
            &format!(
                "<w:pPr><w:pStyle w:val=\"A\"/></w:pPr>\
                 <w:r><w:rPr><w:{toggle}/></w:rPr><w:t>x</w:t></w:r>"
            ),
            &format!("<w:rPr><w:{toggle}/></w:rPr>"),
            &format!("<w:rPr><w:{toggle} w:val=\"0\"/></w:rPr>"),
        );
    }
}

#[test]
fn restyle_paragraph_mark_font_slots_replace() {
    assert_paragraph_restyle(
        concat!(
            "<w:pPr><w:pStyle w:val=\"A\"/><w:rPr>",
            "<w:rFonts w:ascii=\"Courier New\"/></w:rPr></w:pPr>",
        ),
        "<w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\"/></w:rPr>",
        concat!(
            "<w:rPr><w:rFonts w:ascii=\"Times New Roman\" ",
            "w:hAnsi=\"Times New Roman\"/></w:rPr>",
        ),
    );
}

#[test]
fn restyle_shading_under_new_highlight_replaces() {
    assert_paragraph_restyle(
        concat!(
            "<w:pPr><w:pStyle w:val=\"A\"/></w:pPr>",
            "<w:r><w:rPr><w:shd w:val=\"clear\" w:fill=\"FF0000\"/>",
            "</w:rPr><w:t>x</w:t></w:r>",
        ),
        "",
        "<w:rPr><w:highlight w:val=\"yellow\"/></w:rPr>",
    );
}

#[test]
fn restyle_unsupported_paragraph_inheritance_replaces() {
    assert_paragraph_restyle(
        "<w:pPr><w:pStyle w:val=\"A\"/></w:pPr><w:r><w:t>x</w:t></w:r>",
        "",
        "<w:pPr><w:bidi/></w:pPr>",
    );
}

#[test]
fn restyle_retained_single_strike_replaces() {
    assert_paragraph_restyle(
        concat!(
            "<w:pPr><w:pStyle w:val=\"A\"/></w:pPr>",
            "<w:r><w:rPr><w:dstrike w:val=\"0\"/><w:noProof/></w:rPr>",
            "<w:t>x</w:t></w:r>",
        ),
        "<w:rPr><w:dstrike/></w:rPr>",
        "<w:rPr><w:strike/></w:rPr>",
    );
}

#[test]
fn numbering_change_replaces() {
    let (doc, bytes, _) = restyle_source("<w:r><w:t>x</w:t></w:r>", "", "");
    let para_id = doc.paragraphs("body").unwrap()[1].para_id.clone();
    doc.set_paragraph_attr(
        &para_id,
        "numPr",
        Any::Map(Arc::new(std::collections::HashMap::from([
            ("numId".to_owned(), Any::Number(4.0)),
            ("ilvl".to_owned(), Any::Number(2.0)),
        ]))),
    )
    .unwrap();
    assert_restyle_replace(&doc, &bytes, |paragraph| {
        paragraph["formatting"] = json!({"numPr": {"numId": 4, "ilvl": 2}});
    });
}

#[test]
fn run_style_only_change_replaces() {
    let (doc, bytes, _) = restyle_source("<w:r><w:t>x</w:t></w:r>", "", "");
    let style = Any::Map(Arc::new(std::collections::HashMap::from([(
        "styleId".to_owned(),
        Any::from("C"),
    )])));
    format(
        &doc,
        7,
        8,
        InlineFormatDelta {
            other: BTreeMap::from([("runStyle".to_owned(), Some(style))]),
            ..Default::default()
        },
    );
    assert_restyle_replace(&doc, &bytes, |paragraph| {
        paragraph["content"][0]["formatting"] = json!({"styleId": "C"});
    });
}

#[test]
fn paragraph_style_change_replaces() {
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
    assert_legacy(&doc, PatchErrorCode::UnsupportedStructure, "restyle:");
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

fn assert_ambiguous(doc: &EditingDoc, detail: &str) {
    assert_ambiguous_at(doc, 0, detail);
}

fn assert_ambiguous_at(doc: &EditingDoc, ordinal: u32, detail: &str) {
    let plan = doc.paragraph_save_plan();
    let part = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/document.xml")
        .unwrap();
    assert_eq!(part.changed, vec![ordinal]);
    assert!(part.paragraph_patches.is_empty());
    assert_eq!(part.legacy_paragraphs.len(), 1);
    assert_eq!(part.legacy_paragraphs[0].ordinal, ordinal);
    assert_eq!(
        part.legacy_paragraphs[0].code,
        PatchErrorCode::AmbiguousAlignment
    );
    assert_eq!(part.legacy_paragraphs[0].detail, detail);
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
    assert_insertion_patch(&inserted, 1);
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
    assert_legacy(&doc, PatchErrorCode::UnsupportedStructure, "restyle:");
}

fn assert_insertion_patch(doc: &EditingDoc, at: u32) {
    let patch = patch(doc);
    assert_eq!(patch.text_edits.len(), 1);
    assert_eq!((patch.text_edits[0].from, patch.text_edits[0].to), (at, at));
}

#[test]
fn insertion_before_bookmarks_patches_with_shifted_offsets() {
    let doc = open(concat!(
        "<w:r><w:t>a</w:t></w:r>",
        r#"<w:bookmarkStart w:id="1" w:name="mark"/>"#,
        "<w:r><w:t>bc</w:t></w:r>",
        r#"<w:bookmarkEnd w:id="1"/>"#,
        "<w:r><w:t>d</w:t></w:r>",
    ));
    doc.insert_text(&ctx(), Position::new("body", 0), "X", FormatPolicy::Inherit)
        .unwrap();
    assert_insertion_patch(&doc, 0);
}

#[test]
fn numbering_leaf_changes_replace() {
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
    assert_legacy(&doc, PatchErrorCode::UnsupportedStructure, "restyle:");
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
fn hanging_indent_kind_conversion_uses_counted_replace() {
    for (attribute, amount, flag) in [
        ("hanging", "indentFirstLine", "hangingIndent"),
        ("hangingChars", "indentFirstLineChars", "hangingIndentChars"),
    ] {
        let content = format!(
            r#"<w:pPr><w:ind w:{attribute}="100"/></w:pPr>{}"#,
            text("abc")
        );
        let doc = open(&content);
        let paragraph = &doc.paragraphs("body").unwrap()[0];
        let Any::Map(original) = &paragraph.properties["_originalFormatting"] else {
            panic!("missing original formatting");
        };
        assert_eq!(original.get(amount), Some(&Any::Number(-100.0)));
        let mut original = (**original).clone();
        original.insert(flag.to_owned(), Any::Bool(false));
        doc.set_paragraph_attrs(
            &ctx(),
            &ParaSelector::One(paragraph.para_id.clone()),
            &ParaAttrDelta {
                other: BTreeMap::from([(
                    "_originalFormatting".to_owned(),
                    Some(Any::Map(Arc::new(original))),
                )]),
                ..Default::default()
            },
        )
        .unwrap();
        assert_ambiguous(&doc, &format!("ambiguous: indent kind {flag}"));
        let mut expected = String::new();
        let (output, warnings) = save_paragraph(&doc, &content, |paragraph| {
            paragraph["formatting"][flag] = json!(false);
            let paragraph = serde_json::from_value(paragraph.clone()).unwrap();
            let mut context = SerializerContext::new(&SerializerDeterminism {
                seed: "0".repeat(64),
                now: "2000-01-01T00:00:00.000Z".to_owned(),
            })
            .unwrap();
            expected = serialize_paragraph(&paragraph, &mut context).unwrap();
        });
        assert_eq!(
            warnings,
            ["DOCX_MINIMAL_PATCH replace counts {\"ambiguous\":1}"]
        );
        let span = element_span(&output, &[0, 0]).unwrap();
        assert_eq!(&output[span], expected);
        assert!(!output.contains(&format!("w:{attribute}=")), "{output}");
    }
}

#[test]
fn first_line_indent_sign_changes_are_legacy() {
    for (attribute, key) in [
        ("firstLine", "indentFirstLine"),
        ("firstLineChars", "indentFirstLineChars"),
    ] {
        for (before, after) in [(100, -200.0), (-100, -200.0), (-100, 200.0)] {
            for carrier in [false, true] {
                let doc = open(&format!(
                    r#"<w:pPr><w:ind w:{attribute}="{before}"/></w:pPr>{}"#,
                    text("abc")
                ));
                let paragraph = &doc.paragraphs("body").unwrap()[0];
                let (delta_key, value) = if carrier {
                    let Any::Map(original) = &paragraph.properties["_originalFormatting"] else {
                        panic!("missing original formatting");
                    };
                    let mut original = (**original).clone();
                    original.insert(key.to_owned(), Any::Number(after));
                    ("_originalFormatting", Any::Map(Arc::new(original)))
                } else {
                    (key, Any::Number(after))
                };
                doc.set_paragraph_attrs(
                    &ctx(),
                    &ParaSelector::One(paragraph.para_id.clone()),
                    &ParaAttrDelta {
                        other: BTreeMap::from([(delta_key.to_owned(), Some(value))]),
                        ..Default::default()
                    },
                )
                .unwrap();
                assert_ambiguous(&doc, &format!("ambiguous: indent kind {key}"));
            }
        }
    }
}

#[test]
fn hanging_indent_flag_changes_are_legacy() {
    for (attribute, flag) in [
        ("hanging", "hangingIndent"),
        ("hangingChars", "hangingIndentChars"),
    ] {
        let doc = open(&format!(
            r#"<w:pPr><w:ind w:{attribute}="100"/></w:pPr>{}"#,
            text("abc")
        ));
        let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
        doc.set_paragraph_attrs(
            &ctx(),
            &ParaSelector::One(para_id),
            &ParaAttrDelta {
                other: BTreeMap::from([(flag.to_owned(), Some(Any::Bool(false)))]),
                ..Default::default()
            },
        )
        .unwrap();
        assert_ambiguous(&doc, &format!("ambiguous: indent kind {flag}"));
    }
}

#[test]
fn concrete_first_line_and_hanging_value_changes_still_patch() {
    for (attribute, key, value) in [
        ("firstLine", "indentFirstLine", 200.0),
        ("firstLineChars", "indentFirstLineChars", 200.0),
        ("hanging", "indentFirstLine", -200.0),
        ("hangingChars", "indentFirstLineChars", -200.0),
    ] {
        let content = format!(
            r#"<w:pPr><w:ind w:{attribute}="100"/></w:pPr>{}"#,
            text("abc")
        );
        let doc = open(&content);
        let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
        doc.set_paragraph_attrs(
            &ctx(),
            &ParaSelector::One(para_id),
            &ParaAttrDelta {
                other: BTreeMap::from([(key.to_owned(), Some(Any::Number(value)))]),
                ..Default::default()
            },
        )
        .unwrap();
        let patch = patch(&doc);
        assert_eq!(patch.para_deltas.len(), 1);
        assert_eq!(patch.para_deltas[0].path, vec![key]);
        let (output, warnings) = save_paragraph(&doc, &content, |paragraph| {
            paragraph["formatting"][key] = json!(value);
        });
        assert!(warnings.is_empty(), "{warnings:?}");
        assert!(
            output.contains(&format!(r#"w:{attribute}="200""#)),
            "{output}"
        );
    }
}

#[test]
fn spacing_and_indent_clears_are_legacy() {
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
    assert_ambiguous(&doc, "ambiguous: paragraph property clear indentLeft");
}

#[test]
fn clearing_direct_widow_control_is_legacy() {
    let (doc, _, _) = restyle_source(
        r#"<w:pPr><w:pStyle w:val="A"/><w:widowControl w:val="0"/></w:pPr><w:r><w:t>abc</w:t></w:r>"#,
        "<w:pPr><w:widowControl/></w:pPr>",
        "",
    );
    let key = doc.paragraphs("body").unwrap()[1].para_id.clone();
    doc.set_paragraph_attrs(
        &ctx(),
        &ParaSelector::One(key),
        &ParaAttrDelta {
            other: BTreeMap::from([("widowControl".to_owned(), None)]),
            ..Default::default()
        },
    )
    .unwrap();
    assert_ambiguous_at(&doc, 1, "ambiguous: paragraph property clear widowControl");
}

#[test]
fn clearing_inherited_paragraph_properties_is_legacy() {
    for key in ["outlineLevel", "beforeAutospacing", "afterAutospacing"] {
        let (doc, _, _) = restyle_source(
            r#"<w:pPr><w:pStyle w:val="A"/></w:pPr><w:r><w:t>abc</w:t></w:r>"#,
            r#"<w:pPr><w:outlineLvl w:val="0"/><w:spacing w:beforeAutospacing="1" w:afterAutospacing="1"/></w:pPr>"#,
            "",
        );
        let para_id = doc.paragraphs("body").unwrap()[1].para_id.clone();
        doc.set_paragraph_attrs(
            &ctx(),
            &ParaSelector::One(para_id),
            &ParaAttrDelta {
                other: BTreeMap::from([(key.to_owned(), None)]),
                ..Default::default()
            },
        )
        .unwrap();
        assert_ambiguous_at(
            &doc,
            1,
            &format!("ambiguous: paragraph property clear {key}"),
        );
    }
}

#[test]
fn concrete_paragraph_property_change_still_patches() {
    let doc = open(&format!(
        r#"<w:pPr><w:jc w:val="left"/></w:pPr>{}"#,
        text("abc")
    ));
    let key = doc.paragraphs("body").unwrap()[0].para_id.clone();
    doc.set_paragraph_attrs(
        &ctx(),
        &ParaSelector::One(key),
        &ParaAttrDelta {
            alignment: Patch::Set("right".to_owned()),
            ..Default::default()
        },
    )
    .unwrap();
    let patch = patch(&doc);
    assert_eq!(patch.para_deltas.len(), 1);
    assert_eq!(patch.para_deltas[0].path, vec!["alignment"]);
    assert_eq!(
        patch.para_deltas[0].after,
        SlotV1::Present {
            value: json!("right")
        }
    );
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
