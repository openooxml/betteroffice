use std::sync::Arc;

use docx_parse::block::BlockContent;
use docx_parse::inline::{InlineNode, RunContent};
use docx_parse::paragraph::ParagraphContent;
use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{HeaderFooter, S9ParseOptions, parse_docx_s9_wire};
use serde_json::json;

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

fn revision(story: &mut HeaderFooter, revision: u64) {
    let mut value = serde_json::to_value(&*story).unwrap();
    let alias = value
        .as_object_mut()
        .unwrap()
        .entry("sourceAlias")
        .or_insert_with(|| json!({"part": "word/header1.xml", "fingerprint": ""}));
    alias["revision"] = json!(revision);
    *story = serde_json::from_value(value).unwrap();
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
fn resolves_conflicting_alias_edits_by_revision_then_entry_order() {
    let source = fixture("header", false);
    for latest in [0, 1] {
        let mut save = request(&source);
        edit(&mut save.header_entries[1 - latest].1, "Earlier edit");
        revision(&mut save.header_entries[1 - latest].1, 1);
        edit(&mut save.header_entries[latest].1, "Later edit");
        revision(&mut save.header_entries[latest].1, 2);
        let saved = write_docx_s13(save, &source).unwrap();
        let xml = String::from_utf8(part(&saved, "word/header1.xml")).unwrap();
        assert!(xml.contains("Later edit"));
        assert!(!xml.contains("Earlier edit"));
    }
    let mut save = request(&source);
    edit(&mut save.header_entries[0].1, "First entry");
    edit(&mut save.header_entries[1].1, "Second entry");
    let saved = write_docx_s13(save, &source).unwrap();
    assert!(
        String::from_utf8(part(&saved, "word/header1.xml"))
            .unwrap()
            .contains("Second entry")
    );
}
