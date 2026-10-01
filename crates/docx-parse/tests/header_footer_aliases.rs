use std::sync::Arc;

use docx_parse::block::BlockContent;
use docx_parse::inline::{InlineNode, RunContent};
use docx_parse::paragraph::ParagraphContent;
use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{HeaderFooter, S9ParseOptions, parse_docx_s9_wire};
use serde_json::json;
use sha2::{Digest, Sha256};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";

fn fixture(kind: &str, sections: bool) -> Vec<u8> {
    let root = if kind == "header" { "hdr" } else { "ftr" };
    let reference =
        |id: &str, page: &str| format!(r#"<w:{kind}Reference w:type="{page}" r:id="{id}"/>"#);
    let first = reference("rId8", "default");
    let second = reference("rId9", if sections { "default" } else { "first" });
    let body = if sections {
        format!(
            "<w:p><w:pPr><w:sectPr>{first}</w:sectPr></w:pPr><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr>{second}</w:sectPr>"
        )
    } else {
        format!(
            "<w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr>{first}{second}<w:titlePg/></w:sectPr>"
        )
    };
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/{kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/></Types>"#).into_bytes()),
        ("_rels/.rels".to_owned(), format!(r#"<Relationships xmlns="{RELS}"><Relationship Id="office" Type="{R}/officeDocument" Target="word/document.xml"/></Relationships>"#).into_bytes()),
        ("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="{RELS}"><Relationship Id="rId8" Type="{R}/{kind}" Target="{kind}1.xml"/><Relationship Id="rId9" Type="{R}/{kind}" Target="./{kind}1.xml"/></Relationships>"#).into_bytes()),
        ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="{W}" xmlns:r="{R}"><w:body>{body}</w:body></w:document>"#).into_bytes()),
        (format!("word/{kind}1.xml"), format!(r#"<w:{root} xmlns:w="{W}">
  <w:p><w:r><w:t>Synthetic story</w:t></w:r></w:p>
</w:{root}>"#).into_bytes()),
    ]).unwrap()
}

fn request(bytes: &[u8]) -> S13SaveRequest {
    let package = parse_docx_s9_wire(bytes, S9ParseOptions::default())
        .unwrap()
        .document
        .package;
    serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {
            "content": package.document.content,
            "finalSectionProperties": package.document.final_section_properties,
            "customRootBindings": package.document.custom_root_bindings,
        },
        "headerEntries": package.header_entries.unwrap_or_default(),
        "footerEntries": package.footer_entries.unwrap_or_default(),
        "relationshipEntries": package.relationship_entries,
        "options": {"updateModifiedDate": false},
    }))
    .unwrap()
}

fn edit(story: &mut HeaderFooter, replacement: &str) {
    let BlockContent::Paragraph(paragraph) = &mut story.content[0] else {
        panic!("paragraph")
    };
    let ParagraphContent::Inline(InlineNode::Run(run)) = &mut Arc::make_mut(paragraph).content[0]
    else {
        panic!("run")
    };
    let RunContent::Text { text, .. } = &mut run.content[0] else {
        panic!("text")
    };
    *text = replacement.to_owned();
}

fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == name)
        .unwrap()
        .1
}

#[test]
fn saves_edits_through_either_header_or_footer_alias() {
    for kind in ["header", "footer"] {
        for sections in [false, true] {
            for alias in [0, 1] {
                let source = fixture(kind, sections);
                let mut save = request(&source);
                let entries = if kind == "header" {
                    &mut save.header_entries
                } else {
                    &mut save.footer_entries
                };
                edit(&mut entries[alias].1, "Edited synthetic story");
                let saved = write_docx_s13(save, &source).unwrap();
                assert!(
                    String::from_utf8(part(&saved, &format!("word/{kind}1.xml")))
                        .unwrap()
                        .contains("Edited synthetic story")
                );
                let reopened = request(&saved);
                let entries = if kind == "header" {
                    reopened.header_entries
                } else {
                    reopened.footer_entries
                };
                assert!(entries.iter().all(|(_, story)| {
                    serde_json::to_string(story)
                        .unwrap()
                        .contains("Edited synthetic story")
                }));
            }
        }
    }
}

#[test]
fn preserves_untouched_alias_parts_and_relationships() {
    for kind in ["header", "footer"] {
        let source = fixture(kind, false);
        let saved = write_docx_s13(request(&source), &source).unwrap();
        for path in [
            format!("word/{kind}1.xml"),
            "word/_rels/document.xml.rels".to_owned(),
        ] {
            assert_eq!(part(&saved, &path), part(&source, &path));
        }
    }
}

#[test]
fn resolves_conflicting_alias_edits_by_last_changed_entry() {
    for kind in ["header", "footer"] {
        let source = fixture(kind, false);
        for reverse_entries in [false, true] {
            for edit_order in [[0, 1], [1, 0]] {
                let mut save = request(&source);
                let entries = if kind == "header" {
                    &mut save.header_entries
                } else {
                    &mut save.footer_entries
                };
                if reverse_entries {
                    entries.reverse();
                }
                for index in edit_order {
                    edit(
                        &mut entries[index].1,
                        if index == 0 {
                            "First entry"
                        } else {
                            "Last entry"
                        },
                    );
                }
                let saved = write_docx_s13(save.clone(), &source).unwrap();
                assert_eq!(saved, write_docx_s13(save, &source).unwrap());
                let xml = String::from_utf8(part(&saved, &format!("word/{kind}1.xml"))).unwrap();
                assert!(xml.contains("Last entry"));
                assert!(!xml.contains("First entry"));
            }
        }
    }
}

fn replace_part(bytes: &[u8], path: &str, replacement: Vec<u8>) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(bytes).unwrap();
    parts.iter_mut().find(|(name, _)| name == path).unwrap().1 = replacement;
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn header_and_footer_aliases_do_not_overwrite_a_header_edit() {
    let source = fixture("header", false);
    let relationships = String::from_utf8(part(&source, "word/_rels/document.xml.rels"))
        .unwrap()
        .replace(
            &format!(r#"Id="rId9" Type="{R}/header""#),
            &format!(r#"Id="rId9" Type="{R}/footer""#),
        );
    let source = replace_part(
        &source,
        "word/_rels/document.xml.rels",
        relationships.into_bytes(),
    );
    let mut save = request(&source);
    assert_eq!(save.header_entries.len(), 1);
    assert_eq!(save.footer_entries.len(), 1);
    assert!(save.header_entries[0].1.source_alias.is_some());
    assert!(save.footer_entries[0].1.source_alias.is_some());
    edit(
        &mut save.header_entries[0].1,
        "Edited header with footer alias",
    );
    let saved = write_docx_s13(save, &source).unwrap();
    assert!(
        String::from_utf8(part(&saved, "word/header1.xml"))
            .unwrap()
            .contains("Edited header with footer alias")
    );
}

#[test]
fn writes_untouched_aliases_from_the_model_when_source_bytes_change() {
    for kind in ["header", "footer"] {
        let source = fixture(kind, false);
        let path = format!("word/{kind}1.xml");
        let swapped = replace_part(
            &source,
            &path,
            String::from_utf8(part(&source, &path))
                .unwrap()
                .replace("Synthetic story", "Swapped source story")
                .into_bytes(),
        );
        let saved = write_docx_s13(request(&source), &swapped).unwrap();
        let xml = String::from_utf8(part(&saved, &path)).unwrap();
        assert!(xml.contains("Synthetic story"));
        assert!(!xml.contains("Swapped source story"));
    }
}

#[test]
fn case_variant_targets_are_not_aliases() {
    let source = fixture("header", false);
    let relationships = String::from_utf8(part(&source, "word/_rels/document.xml.rels"))
        .unwrap()
        .replace("./header1.xml", "Header1.xml");
    let source = replace_part(
        &source,
        "word/_rels/document.xml.rels",
        relationships.into_bytes(),
    );
    let mut save = request(&source);
    assert!(
        save.header_entries
            .iter()
            .all(|(_, story)| story.source_alias.is_none())
    );
    edit(&mut save.header_entries[0].1, "Lowercase target edit");
    let saved = write_docx_s13(save, &source).unwrap();
    assert!(
        String::from_utf8(part(&saved, "word/header1.xml"))
            .unwrap()
            .contains("Lowercase target edit")
    );
    assert!(
        String::from_utf8(part(&saved, "word/Header1.xml"))
            .unwrap()
            .contains("Synthetic story")
    );
}

#[test]
fn aliased_math_with_a_custom_prefix_opens_and_saves_source_bytes() {
    let source = fixture("header", false);
    let source = replace_part(&source, "word/header1.xml", format!(r#"<w:hdr xmlns:w="{W}" xmlns:q="http://schemas.openxmlformats.org/officeDocument/2006/math"><w:p><q:oMath><q:r><q:t>x</q:t></q:r></q:oMath></w:p></w:hdr>"#).into_bytes());
    let mut save = request(&source);
    assert_eq!(save.header_entries.len(), 2);
    for (_, story) in &save.header_entries {
        assert!(story.source_alias.is_none());
        let BlockContent::Paragraph(paragraph) = &story.content[0] else {
            panic!("paragraph")
        };
        assert!(matches!(
            &paragraph.content[0],
            ParagraphContent::Inline(InlineNode::Math(_))
        ));
    }
    save.selective = Some(
        serde_json::from_value(json!({
            "sourceParagraphs": {
                "partSha256": format!("{:x}", Sha256::digest(part(&source, "word/document.xml"))),
                "paragraphs": [],
            },
        }))
        .unwrap(),
    );
    let saved = write_docx_s13(save, &source).unwrap();
    assert_eq!(
        part(&saved, "word/header1.xml"),
        part(&source, "word/header1.xml")
    );
    assert_eq!(request(&saved).header_entries.len(), 2);
}
