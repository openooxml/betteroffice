//! TIFF media must reach the renderer as PNG while the saved package keeps the
//! original TIFF bytes.
#![cfg(feature = "tiff")]

use docx_parse::document::{DocumentBody, Section};
use docx_parse::s9::{S9ParseOptions, parse_docx_s9_wire};
use docx_parse::serializer::{
    S13SaveOptions, S13SaveRequest, SerializerDeterminism, write_docx_s13,
};
use sha2::{Digest, Sha256};

const TIFF_PATH: &str = "word/media/image1.tif";

/// 2x1 little-endian uncompressed RGB TIFF.
fn tiff_bytes() -> Vec<u8> {
    tiff_with_photometric(2)
}

/// Same TIFF with a CIELab photometric the decoder does not support.
fn unsupported_tiff_bytes() -> Vec<u8> {
    tiff_with_photometric(8)
}

fn tiff_with_photometric(photometric: u32) -> Vec<u8> {
    let entries: [(u16, u16, u32, u32); 9] = [
        (256, 3, 1, 2),
        (257, 3, 1, 1),
        (258, 3, 3, 122),
        (259, 3, 1, 1),
        (262, 3, 1, photometric),
        (273, 4, 1, 128),
        (277, 3, 1, 3),
        (278, 4, 1, 1),
        (279, 4, 1, 6),
    ];
    let mut data = vec![0u8; 134];
    data[..4].copy_from_slice(b"II\x2a\x00");
    data[4..8].copy_from_slice(&8u32.to_le_bytes());
    data[8..10].copy_from_slice(&(entries.len() as u16).to_le_bytes());
    for (index, (tag, kind, count, value)) in entries.iter().enumerate() {
        let at = 10 + index * 12;
        data[at..at + 2].copy_from_slice(&tag.to_le_bytes());
        data[at + 2..at + 4].copy_from_slice(&kind.to_le_bytes());
        data[at + 4..at + 8].copy_from_slice(&count.to_le_bytes());
        data[at + 8..at + 12].copy_from_slice(&value.to_le_bytes());
    }
    for index in 0..3 {
        let at = 122 + index * 2;
        data[at..at + 2].copy_from_slice(&8u16.to_le_bytes());
    }
    data[128..134].copy_from_slice(&[0xff, 0x00, 0x00, 0x00, 0x80, 0xff]);
    data
}

fn package() -> Vec<u8> {
    package_with_media(tiff_bytes())
}

fn package_with_media(media: Vec<u8>) -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="tif" ContentType="image/tiff"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.tif"/></Relationships>"#.to_vec(),
        ),
        (TIFF_PATH.to_owned(), media),
        (
            "word/document.xml".to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body><w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p></w:body></w:document>"#.to_vec(),
        ),
    ])
    .unwrap()
}

fn find_images(value: &serde_json::Value) -> Vec<&serde_json::Value> {
    let mut found = Vec::new();
    let mut stack = vec![value];
    while let Some(node) = stack.pop() {
        match node {
            serde_json::Value::Object(map) => {
                if map.get("type").and_then(|kind| kind.as_str()) == Some("image") {
                    found.push(node);
                }
                stack.extend(map.values());
            }
            serde_json::Value::Array(items) => stack.extend(items),
            _ => {}
        }
    }
    found
}

fn part(data: &[u8], path: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(data)
        .unwrap()
        .into_iter()
        .find(|(candidate, _)| candidate == path)
        .unwrap()
        .1
}

#[test]
fn inline_tiff_images_resolve_to_a_png_data_url() {
    let original = package();
    let wire = parse_docx_s9_wire(&original, S9ParseOptions::default()).unwrap();
    let json = serde_json::to_value(&wire).unwrap();
    let images = find_images(&json);
    assert_eq!(images.len(), 1, "one inline picture: {images:?}");
    let image = images[0];
    assert_eq!(image["rId"], "rId5");
    assert_eq!(image["mimeType"], "image/png");
    assert!(
        image["src"]
            .as_str()
            .unwrap()
            .starts_with("data:image/png;base64,iVBORw0KGgo"),
        "image src must be a PNG data url: {}",
        image["src"].as_str().unwrap()
    );
    let media = &json["document"]["package"]["mediaEntries"];
    let entry = media
        .as_array()
        .unwrap()
        .iter()
        .find(|entry| entry[0] == TIFF_PATH)
        .unwrap();
    assert_eq!(entry[1]["mimeType"], "image/png");
    assert_eq!(entry[1]["path"], TIFF_PATH);
    assert_eq!(entry[1]["filename"], "image1.tif");
}

fn save_request(original: &[u8]) -> S13SaveRequest {
    let wire = parse_docx_s9_wire(original, S9ParseOptions::default()).unwrap();
    let package = wire.document.package;
    let body = package.document;
    let sections = body.sections.map(|sections| {
        sections
            .into_iter()
            .map(|section| Section {
                id: section.id,
                properties: section.properties,
                content: body.content[section.content_start..section.content_end].to_vec(),
            })
            .collect()
    });
    S13SaveRequest {
        determinism: SerializerDeterminism {
            seed: format!("{:x}", Sha256::digest(original)),
            now: "1970-01-01T00:00:00.000Z".to_owned(),
        },
        document: DocumentBody {
            content: body.content,
            sections,
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
    }
}

#[test]
fn saving_keeps_the_original_tiff_part_byte_identical() {
    let original = package();
    let saved = write_docx_s13(save_request(&original), &original).unwrap();
    assert_eq!(part(&saved, TIFF_PATH), tiff_bytes());
    assert!(
        String::from_utf8(part(&saved, "word/document.xml"))
            .unwrap()
            .contains(r#"r:embed="rId5""#),
        "the saved drawing must still point at the TIFF relationship"
    );
    assert!(
        !ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .iter()
            .any(|(path, _)| path.ends_with(".png")),
        "the display transcode must not add a media part"
    );
}

#[test]
fn an_unsupported_tiff_encoding_warns_and_keeps_the_original_source() {
    let original = package_with_media(unsupported_tiff_bytes());
    let wire = parse_docx_s9_wire(&original, S9ParseOptions::default()).unwrap();
    let warnings = wire.document.warnings.clone().unwrap_or_default();
    assert_eq!(
        warnings.len(),
        1,
        "one warning per media part: {warnings:?}"
    );
    assert!(
        warnings[0].starts_with(&format!("TIFF image {TIFF_PATH} could not be decoded")),
        "the warning must name the part: {}",
        warnings[0]
    );
    assert!(
        warnings[0].contains("unsupported"),
        "the warning must carry the decoder reason: {}",
        warnings[0]
    );

    let json = serde_json::to_value(&wire).unwrap();
    let images = find_images(&json);
    assert_eq!(images.len(), 1);
    assert_eq!(images[0]["mimeType"], "image/tiff");
    assert!(
        images[0]["src"]
            .as_str()
            .unwrap()
            .starts_with("data:image/tiff;base64,"),
        "an undecodable TIFF keeps its own source for decoders that handle it"
    );
    let entry = json["document"]["package"]["mediaEntries"]
        .as_array()
        .unwrap()
        .iter()
        .find(|entry| entry[0] == TIFF_PATH)
        .unwrap()
        .clone();
    assert_eq!(entry[1]["mimeType"], "image/tiff");
}

#[test]
fn a_decodable_tiff_warns_about_nothing() {
    let wire = parse_docx_s9_wire(&package(), S9ParseOptions::default()).unwrap();
    assert_eq!(wire.document.warnings, None);
}

#[test]
fn an_unsupported_tiff_still_round_trips_its_bytes() {
    let original = package_with_media(unsupported_tiff_bytes());
    let saved = write_docx_s13(save_request(&original), &original).unwrap();
    assert_eq!(part(&saved, TIFF_PATH), unsupported_tiff_bytes());
}
