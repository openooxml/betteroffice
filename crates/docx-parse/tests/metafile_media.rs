//! EMF and WMF media must reach the renderer as SVG while the saved package
//! keeps the original metafile bytes.
#![cfg(feature = "metafile")]

use base64::Engine as _;
use docx_parse::document::{DocumentBody, Section};
use docx_parse::s9::{S9ParseOptions, parse_docx_s9_wire};
use docx_parse::serializer::{
    S13SaveOptions, S13SaveRequest, SerializerDeterminism, write_docx_s13,
};
use sha2::{Digest, Sha256};

#[path = "../../ooxml-metafile/tests/common/mod.rs"]
mod metafile;

const EMF: &[u8] = include_bytes!("../../ooxml-metafile/tests/fixtures/shapes.emf");
const WMF: &[u8] = include_bytes!("../../ooxml-metafile/tests/fixtures/shapes.wmf");

const DRAWING: &str = r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="609600"/><wp:docPr id="1" name="Picture 1"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#;
const VML: &str = r#"<w:r><w:pict><v:shape id="Object 1" style="width:72pt;height:48pt"><v:imagedata r:id="rId5" o:title=""/></v:shape></w:pict></w:r>"#;

fn package(path: &str, media: &[u8], run: &str) -> Vec<u8> {
    let extension = path.rsplit('.').next().unwrap();
    let content_type = if extension == "wmf" {
        "image/x-wmf"
    } else {
        "image/x-emf"
    };
    let target = path.strip_prefix("word/").unwrap();
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="{extension}" ContentType="{content_type}"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#).into_bytes(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="{target}"/></Relationships>"#).into_bytes(),
        ),
        (path.to_owned(), media.to_vec()),
        (
            "word/document.xml".to_owned(),
            format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"><w:body><w:p>{run}</w:p></w:body></w:document>"#).into_bytes(),
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

fn svg_of(src: &str) -> String {
    let encoded = src
        .strip_prefix("data:image/svg+xml;base64,")
        .unwrap_or_else(|| {
            panic!(
                "the picture must be an SVG data URL: {}",
                &src[..40.min(src.len())]
            )
        });
    String::from_utf8(
        base64::engine::general_purpose::STANDARD
            .decode(encoded)
            .unwrap(),
    )
    .unwrap()
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
fn emf_and_wmf_pictures_resolve_to_svg_data_urls() {
    for (path, media, run) in [
        ("word/media/image1.emf", EMF, DRAWING),
        ("word/media/image1.wmf", WMF, DRAWING),
        ("word/media/image1.emf", EMF, VML),
    ] {
        let wire =
            parse_docx_s9_wire(&package(path, media, run), S9ParseOptions::default()).unwrap();
        assert_eq!(wire.document.warnings, None, "{path} converts cleanly");
        let json = serde_json::to_value(&wire).unwrap();
        let images = find_images(&json);
        assert_eq!(images.len(), 1, "{path}: {images:?}");
        assert_eq!(images[0]["mimeType"], "image/svg+xml");
        let svg = svg_of(images[0]["src"].as_str().unwrap());
        assert!(svg.starts_with("<svg xmlns=\"http://www.w3.org/2000/svg\""));
        assert!(svg.contains("<path"), "{path} draws its shapes");
        let entry = json["document"]["package"]["mediaEntries"]
            .as_array()
            .unwrap()
            .iter()
            .find(|entry| entry[0] == path)
            .unwrap()
            .clone();
        assert_eq!(entry[1]["mimeType"], "image/svg+xml");
        assert_eq!(entry[1]["filename"], path.rsplit('/').next().unwrap());
    }
}

#[test]
fn a_metafile_that_cannot_be_replayed_shows_a_placeholder_and_warns() {
    let mut broken = EMF.to_vec();
    broken[88..92].copy_from_slice(&250u32.to_le_bytes());
    let path = "word/media/image1.emf";
    let wire =
        parse_docx_s9_wire(&package(path, &broken, DRAWING), S9ParseOptions::default()).unwrap();
    let warnings = wire.document.warnings.clone().unwrap_or_default();
    assert_eq!(warnings.len(), 1, "{warnings:?}");
    assert!(
        warnings[0].starts_with(&format!(
            "EMF image {path} could not be converted for display: "
        )),
        "{}",
        warnings[0]
    );
    let json = serde_json::to_value(&wire).unwrap();
    let svg = svg_of(find_images(&json)[0]["src"].as_str().unwrap());
    assert!(
        svg.contains(r##"fill="#f1f3f4""##),
        "a neutral placeholder stands in: {svg}"
    );
}

#[test]
fn compressed_metafile_parts_exhaust_the_document_pixel_budget() {
    let metafile = metafile::cropped_rle_emf(1024);
    assert!(metafile.len() < 256);
    let mut parts =
        ooxml_opc::unzip_parts(&package("word/media/image0.emf", &metafile, DRAWING)).unwrap();
    for index in 1..66 {
        parts.push((format!("word/media/image{index}.emf"), metafile.clone()));
    }
    let wire = parse_docx_s9_wire(
        &ooxml_opc::rezip_parts(&parts).unwrap(),
        S9ParseOptions::default(),
    )
    .unwrap();
    let warnings = wire.document.warnings.clone().unwrap_or_default();
    assert_eq!(warnings.len(), 2, "{warnings:?}");
    for index in 64..66 {
        assert!(warnings.iter().any(|warning| {
            warning.contains(&format!("image{index}.emf")) && warning.contains("replay limits")
        }));
    }
    let json = serde_json::to_value(&wire).unwrap();
    let entries = json["document"]["package"]["mediaEntries"]
        .as_array()
        .unwrap();
    for index in 0..66 {
        let path = format!("word/media/image{index}.emf");
        let entry = entries.iter().find(|entry| entry[0] == path).unwrap();
        let svg = svg_of(entry[1]["dataUrl"].as_str().unwrap());
        if index < 64 {
            assert!(svg.contains("<image"), "{path}: {svg}");
        } else {
            assert!(svg.contains(r##"fill="#f1f3f4""##), "{path}: {svg}");
        }
    }
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
        paragraph_ids: None,
    }
}

#[test]
fn saving_keeps_the_original_metafile_part_byte_identical() {
    for (path, media) in [
        ("word/media/image1.emf", EMF),
        ("word/media/image1.wmf", WMF),
    ] {
        let original = package(path, media, DRAWING);
        let saved = write_docx_s13(save_request(&original), &original).unwrap();
        assert_eq!(part(&saved, path), media);
        assert!(
            !ooxml_opc::unzip_parts(&saved)
                .unwrap()
                .iter()
                .any(|(path, _)| path.ends_with(".svg")),
            "the display transcode must not add a media part"
        );
    }
}

#[test]
fn a_metafile_part_with_a_damaged_signature_shows_a_placeholder_and_warns() {
    let mut damaged = WMF.to_vec();
    damaged[..4].copy_from_slice(b"\0\0\0\0");
    let path = "word/media/image1.wmf";
    let wire =
        parse_docx_s9_wire(&package(path, &damaged, DRAWING), S9ParseOptions::default()).unwrap();
    let warnings = wire.document.warnings.clone().unwrap_or_default();
    assert_eq!(warnings.len(), 1, "{warnings:?}");
    assert!(
        warnings[0].starts_with(&format!(
            "WMF image {path} could not be converted for display: "
        )),
        "{}",
        warnings[0]
    );
    let json = serde_json::to_value(&wire).unwrap();
    let svg = svg_of(find_images(&json)[0]["src"].as_str().unwrap());
    assert!(svg.contains(r##"fill="#f1f3f4""##), "{svg}");
}

#[test]
fn a_raster_or_svg_saved_under_a_metafile_name_passes_through() {
    let png: &[u8] = b"\x89PNG\r\n\x1a\n\0\0\0\rIHDR";
    let svg: &[u8] =
        b"\xEF\xBB\xBF\n<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"4\" height=\"4\"/>";
    for (data, extension, mime) in [
        (png, "emf", "image/x-emf"),
        (svg, "emf", "image/x-emf"),
        (svg, "wmf", "image/x-wmf"),
    ] {
        let path = format!("word/media/image1.{extension}");
        let bytes = package(&path, data, DRAWING);
        let wire = parse_docx_s9_wire(&bytes, S9ParseOptions::default()).unwrap();
        assert_eq!(wire.document.warnings, None);
        let json = serde_json::to_value(&wire).unwrap();
        let src = find_images(&json)[0]["src"].as_str().unwrap().to_owned();
        assert!(src.starts_with(&format!("data:{mime};base64,")), "{src}");
        let package = ooxml_opc::RetainedPackage::new(bytes.into()).unwrap();
        let table = docx_parse::media::MediaTable::new(package).unwrap();
        assert!(table.warnings().is_empty());
        assert_eq!(table.bytes(0).unwrap().as_ref(), data);
    }
}
