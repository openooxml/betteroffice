use docx_parse::block::BlockContent;
use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use quick_xml::NsReader;
use quick_xml::events::Event;
use quick_xml::name::ResolveResult;
use serde_json::json;

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

fn source(kind: &str, note_type: &str, attributes: &str) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        ("[Content_Types].xml".to_owned(), br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("word/document.xml".to_owned(), format!(r#"<w:document xmlns:w="{W}"><w:body><w:p/></w:body></w:document>"#).into_bytes()),
        (format!("word/{kind}s.xml"), format!(r#"<w:{kind}s xmlns:w="{W}" xmlns:ext="urn:custom"><w:{kind} w:id="-1" w:type="{note_type}" {attributes}><w:p><w:r><w:t>Original separator</w:t></w:r></w:p></w:{kind}></w:{kind}s>"#).into_bytes()),
    ])
    .unwrap()
}

fn save(bytes: &[u8], edit: impl FnOnce(&mut S13SaveRequest)) -> Vec<u8> {
    let package = parse_docx_s9_wire(bytes, S9ParseOptions::default())
        .unwrap()
        .document
        .package;
    let mut request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": {"content": package.document.content},
        "footnoteSeparators": package.footnote_separators.unwrap_or_default(),
        "endnoteSeparators": package.endnote_separators.unwrap_or_default(),
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    edit(&mut request);
    write_docx_s13(request, bytes).unwrap()
}

fn part(bytes: &[u8], kind: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| *path == format!("word/{kind}s.xml"))
        .unwrap()
        .1
}

#[test]
fn full_save_keeps_separator_attribute_prefixes_bound() {
    for kind in ["footnote", "endnote"] {
        for note_type in ["separator", "continuationSeparator", "continuationNotice"] {
            let bytes = source(kind, note_type, r#"ext:flag="1""#);
            let xml = part(&save(&bytes, |_| {}), kind);
            assert!(String::from_utf8_lossy(&xml).contains("Original separator"));
            let mut reader = NsReader::from_reader(xml.as_slice());
            loop {
                match reader.read_resolved_event().unwrap() {
                    (namespace, Event::Start(element) | Event::Empty(element)) => {
                        assert!(!matches!(namespace, ResolveResult::Unknown(_)));
                        for attribute in element.attributes() {
                            let attribute = attribute.unwrap();
                            let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
                            assert!(!matches!(namespace, ResolveResult::Unknown(_)));
                        }
                    }
                    (_, Event::Eof) => break,
                    _ => {}
                }
            }
        }
    }
}

#[test]
fn full_save_writes_separator_model_edits() {
    for kind in ["footnote", "endnote"] {
        for note_type in ["separator", "continuationSeparator", "continuationNotice"] {
            let bytes = source(kind, note_type, "");
            let saved = save(&bytes, |request| {
                for note in request
                    .footnote_separators
                    .iter_mut()
                    .chain(&mut request.endnote_separators)
                {
                    let BlockContent::Paragraph(paragraph) = &mut note.content[0] else {
                        panic!("expected a separator paragraph");
                    };
                    let paragraph = std::sync::Arc::make_mut(paragraph);
                    paragraph.content = serde_json::from_value(json!([{
                        "type": "run",
                        "formatting": {},
                        "content": [{"type": "text", "text": "Edited separator"}]
                    }]))
                    .unwrap();
                    paragraph
                        .formatting
                        .get_or_insert_with(Default::default)
                        .space_after = Some(240.0);
                }
            });
            let xml = part(&saved, kind);
            assert!(String::from_utf8_lossy(&xml).contains(">Edited separator</w:t>"));
            assert!(!String::from_utf8_lossy(&xml).contains("Original separator"));
            let package = parse_docx_s9_wire(&saved, S9ParseOptions::default())
                .unwrap()
                .document
                .package;
            let notes = if kind == "footnote" {
                package.footnote_separators
            } else {
                package.endnote_separators
            }
            .unwrap();
            let BlockContent::Paragraph(paragraph) = &notes[0].content[0] else {
                panic!("expected a saved separator paragraph");
            };
            assert_eq!(
                paragraph.formatting.as_ref().unwrap().space_after,
                Some(240.0)
            );
        }
    }
}
