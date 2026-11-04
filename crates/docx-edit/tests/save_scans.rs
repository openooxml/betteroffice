use docx_edit::{EditCtx, EditingDoc, StoryRange, seed_from_docx};
use docx_parse::paragraph_identity::take_tagged_bytes;
use docx_parse::serializer::write_docx_s13_with_warnings;
use serde_json::json;

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml""#
);
const PARAGRAPHS: usize = 200;

fn document_xml() -> String {
    let body: String = (0..PARAGRAPHS)
        .map(|i| {
            format!(
                r#"<w:p w14:paraId="{:08X}"><w:r><w:t xml:space="preserve">Line {i:04}</w:t></w:r></w:p>"#,
                i + 1
            )
        })
        .collect();
    format!(
        r#"<w:document {NS}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>"#
    )
}

fn package() -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned(),
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="styles" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_owned(),
        ),
        ("word/document.xml", document_xml()),
        (
            "word/styles.xml",
            format!(
                r#"<w:styles {NS}><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>"#
            ),
        ),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

/// Edits the first `edited` paragraphs, saves, and returns the bytes the save tokenized.
fn tokenized_saving(edited: usize) -> usize {
    let bytes = package();
    let doc = EditingDoc::new(71);
    seed_from_docx(&doc, &bytes).unwrap();
    let ctx = EditCtx::local("", "2026-10-21T00:00:00Z");
    for i in 0..edited {
        let at = (i * 10) as u32;
        doc.replace_range(&ctx, StoryRange::new("body", at, at + 1), "M")
            .unwrap();
    }
    let plan = doc.paragraph_save_plan();
    let part = plan
        .spliced_parts
        .iter()
        .find(|part| part.part == "word/document.xml")
        .unwrap();
    assert_eq!(part.paragraph_patches.len(), edited);
    assert!(part.legacy_paragraphs.is_empty());
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
    let mut document = serde_json::to_value(&package.document).unwrap();
    for i in 0..edited {
        document["content"][i]["content"][0]["content"][0]["text"] = json!(format!("Mine {i:04}"));
    }
    if let Some(sections) = &package.document.sections {
        let content = document["content"].as_array().unwrap().clone();
        document["sections"] = json!(
            sections
                .iter()
                .map(|section| json!({
                    "id": section.id,
                    "properties": section.properties,
                    "content": &content[section.content_start..section.content_end]
                }))
                .collect::<Vec<_>>()
        );
    }
    let ordinals: Vec<u32> = part
        .paragraphs
        .iter()
        .map(|(ordinal, _)| *ordinal)
        .collect();
    let patched: Vec<_> = plan
        .patched_parts
        .iter()
        .map(|(part, para_ids)| json!({"part": part, "paraIds": para_ids}))
        .collect();
    let request = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": document,
        "headerEntries": package.header_entries.unwrap_or_default(),
        "relationshipEntries": package.relationship_entries,
        "options": {"updateModifiedDate": false},
        "paragraphIds": {"patchedParts": patched, "splicedParts": [{
            "part": part.part, "sha256": part.sha256,
            "paragraphs": ordinals, "changed": part.changed,
            "paragraphPatches": part.paragraph_patches,
            "legacyParagraphs": part.legacy_paragraphs
        }]}
    }))
    .unwrap();
    take_tagged_bytes();
    let (saved, warnings) = write_docx_s13_with_warnings(request, &bytes).unwrap();
    let tokenized = take_tagged_bytes();
    assert!(warnings.is_empty(), "{warnings:?}");
    let (_, xml) = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    let xml = String::from_utf8(xml).unwrap();
    for i in 0..edited {
        assert!(
            xml.contains(&format!(">Mine {i:04}</w:t>")),
            "paragraph {i} not saved"
        );
    }
    assert!(xml.contains(&format!(">Line {:04}</w:t>", PARAGRAPHS - 1)));
    tokenized
}

#[test]
fn saving_more_patched_paragraphs_does_not_rescan_the_part_for_each() {
    let part = document_xml().len();
    let one = tokenized_saving(1);
    let many = tokenized_saving(100);
    let extra = many.saturating_sub(one);
    let patched = 99;
    assert!(
        extra * 10 < part * patched,
        "patching {patched} more paragraphs tokenized {extra} more bytes, {:.2} times the \
         {part}-byte part for each, where a tenth is the limit",
        extra as f64 / (part * patched) as f64
    );
}
