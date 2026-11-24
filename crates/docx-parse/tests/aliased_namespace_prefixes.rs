use docx_parse::document::{DocumentBody, Section};
use docx_parse::s9::{S9PackageWire, S9ParseOptions, parse_docx_s9_wire};
use docx_parse::section::StoryReference;
use docx_parse::serializer::{
    S13SaveOptions, S13SaveRequest, SerializerDeterminism, write_docx_s13,
};
use serde_json::Value;
use sha2::{Digest, Sha256};

const W: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const W14: &str = "http://schemas.microsoft.com/office/word/2010/wordml";

fn aliased_document() -> String {
    format!(
        r#"<?xml version='1.0' encoding='utf-8'?>
<ns0:document xmlns:ns0="{W}" xmlns:ns1="{W14}" xmlns:ns4="{R}"><ns0:body><ns0:p ns1:paraId="00000001"><ns0:pPr><ns0:jc ns0:val="center"/></ns0:pPr><ns0:r><ns0:rPr><ns0:b/></ns0:rPr><ns0:t>Hello</ns0:t></ns0:r></ns0:p><ns0:sectPr><ns0:headerReference ns0:type="default" ns4:id="rId6"/><ns0:footerReference ns0:type="first" ns4:id="rId8"/><ns0:titlePg/></ns0:sectPr></ns0:body></ns0:document>"#
    )
}

fn package(document_xml: &str) -> Vec<u8> {
    let part = |name: &str, xml: String| (name.to_owned(), xml.into_bytes());
    ooxml_opc::rezip_parts(&[
        part(
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>"#.to_owned(),
        ),
        part(
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned(),
        ),
        part(
            "word/_rels/document.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/><Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/></Relationships>"#.to_owned(),
        ),
        part("word/document.xml", document_xml.to_owned()),
        part(
            "word/header1.xml",
            format!(r#"<ns0:hdr xmlns:ns0="{W}"><ns0:p><ns0:r><ns0:t>Head</ns0:t></ns0:r></ns0:p></ns0:hdr>"#),
        ),
        part(
            "word/footer1.xml",
            format!(r#"<x:ftr xmlns:x="{W}"><x:p><x:r><x:t>Foot</x:t></x:r></x:p></x:ftr>"#),
        ),
    ])
    .unwrap()
}

fn parsed(bytes: &[u8]) -> S9PackageWire {
    parse_docx_s9_wire(bytes, S9ParseOptions::default())
        .unwrap()
        .document
        .package
}

fn reference(reference_type: &str, relationship_id: &str) -> StoryReference {
    StoryReference {
        reference_type: reference_type.to_owned(),
        relationship_id: relationship_id.to_owned(),
    }
}

fn texts(value: &Value, out: &mut Vec<String>) {
    match value {
        Value::Object(map) => {
            if map.get("type").and_then(Value::as_str) == Some("text")
                && let Some(text) = map.get("text").and_then(Value::as_str)
            {
                out.push(text.to_owned());
            }
            map.values().for_each(|child| texts(child, out));
        }
        Value::Array(items) => items.iter().for_each(|child| texts(child, out)),
        _ => {}
    }
}

fn text_of(value: impl serde::Serialize) -> Vec<String> {
    let mut out = Vec::new();
    texts(&serde_json::to_value(value).unwrap(), &mut out);
    out
}

fn assert_reads_as_standard(package: &S9PackageWire) {
    let properties = package
        .document
        .final_section_properties
        .as_ref()
        .expect("final section properties");
    assert_eq!(
        properties.header_references.as_deref(),
        Some(&[reference("default", "rId6")][..])
    );
    assert_eq!(
        properties.footer_references.as_deref(),
        Some(&[reference("first", "rId8")][..])
    );
    assert_eq!(properties.title_pg, Some(true));
    assert_eq!(text_of(&package.document.content), ["Hello"]);
    let headers = package.header_entries.as_deref().unwrap_or_default();
    assert_eq!(headers.len(), 1);
    assert_eq!(headers[0].0, "rId6");
    assert_eq!(text_of(&headers[0].1.content), ["Head"]);
    let footers = package.footer_entries.as_deref().unwrap_or_default();
    assert_eq!(footers.len(), 1);
    assert_eq!(footers[0].0, "rId8");
    assert_eq!(text_of(&footers[0].1.content), ["Foot"]);
}

#[test]
fn attributes_under_aliased_prefixes_read_as_their_standard_namespace() {
    let aliased = parsed(&package(&aliased_document()));
    assert_reads_as_standard(&aliased);
    let standard = parsed(&package(
        &aliased_document()
            .replace("ns0:", "w:")
            .replace("ns1:", "w14:")
            .replace("ns4:", "r:")
            .replace("xmlns:ns0=", "xmlns:w=")
            .replace("xmlns:ns1=", "xmlns:w14=")
            .replace("xmlns:ns4=", "xmlns:r="),
    ));
    assert_eq!(
        aliased.document.final_section_properties,
        standard.document.final_section_properties
    );
}

#[test]
fn a_package_with_aliased_prefixes_saves_and_reopens() {
    let original = package(&aliased_document());
    let package = parsed(&original);
    let body = package.document;
    let request = S13SaveRequest {
        determinism: SerializerDeterminism {
            seed: format!("{:x}", Sha256::digest(&original)),
            now: "1970-01-01T00:00:00.000Z".to_owned(),
        },
        document: DocumentBody {
            sections: body.sections.map(|sections| {
                sections
                    .into_iter()
                    .map(|section| Section {
                        id: section.id,
                        properties: section.properties,
                        content: body.content[section.content_start..section.content_end].to_vec(),
                    })
                    .collect()
            }),
            content: body.content,
            final_section_properties: body.final_section_properties,
            custom_root_bindings: body.custom_root_bindings,
            comments: body.comments,
        },
        header_entries: package.header_entries.unwrap_or_default(),
        footer_entries: package.footer_entries.unwrap_or_default(),
        footnotes: package.footnotes.unwrap_or_default(),
        endnotes: package.endnotes.unwrap_or_default(),
        footnote_separators: package.footnote_separators.unwrap_or_default(),
        endnote_separators: package.endnote_separators.unwrap_or_default(),
        relationship_entries: package.relationship_entries,
        numbering: Some(package.numbering),
        options: S13SaveOptions {
            update_modified_date: false,
            modified_by: None,
        },
        selective: None,
        paragraph_ids: None,
    };
    let saved = write_docx_s13(request, &original).unwrap();
    assert_reads_as_standard(&parsed(&saved));
}
