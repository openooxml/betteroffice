#![cfg(feature = "manager")]

use betteroffice_fonts::embedded_font_definitions;

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const RELS: &[u8] = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="face" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" Target="../fonts/face.odttf"/></Relationships>"#;
const XML: &[u8] = br#"<p:embeddedFont xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:font typeface="Embedded"/><p:regular r:id="face"/></p:embeddedFont>"#;

#[test]
fn shared_embedded_parts_are_decoded_once_and_definitions_are_bounded() {
    let mut xml = String::from("<w:fonts xmlns:w='w' xmlns:r='r'>");
    for index in 0..5000 {
        xml.push_str(&format!(
            "<w:font w:name='Face {index}'><w:embedRegular r:id='face'/></w:font>"
        ));
    }
    xml.push_str("</w:fonts>");
    let reads = std::cell::Cell::new(0);
    let font = include_bytes!("fixtures/Synthetic-Regular.ttf");
    let faces = embedded_font_definitions("word/fontTable.xml", xml.as_bytes(), RELS, |_| {
        reads.set(reads.get() + 1);
        Some(font)
    });
    assert_eq!(faces.len(), 256);
    assert_eq!(reads.get(), 1);
}

#[test]
fn duplicate_embedded_definitions_are_skipped() {
    let xml = format!(
        "<w:fonts xmlns:w='w' xmlns:r='r'>{}</w:fonts>",
        "<w:font w:name='Same'><w:embedRegular r:id='face'/></w:font>".repeat(5000)
    );
    let font = include_bytes!("fixtures/Synthetic-Regular.ttf");
    let faces =
        embedded_font_definitions("word/fontTable.xml", xml.as_bytes(), RELS, |_| Some(font));
    assert_eq!(faces.len(), 1);
}

#[test]
fn embedded_definition_bytes_include_each_retained_copy() {
    let mut font = include_bytes!("fixtures/Synthetic-Regular.ttf").to_vec();
    font.resize(32 * 1024 * 1024, 0);
    let mut xml = String::from("<w:fonts xmlns:w='w' xmlns:r='r'>");
    for index in 0..5 {
        xml.push_str(&format!(
            "<w:font w:name='Face {index}'><w:embedRegular r:id='face'/></w:font>"
        ));
    }
    xml.push_str("</w:fonts>");
    let faces =
        embedded_font_definitions("word/fontTable.xml", xml.as_bytes(), RELS, |_| Some(&font));
    assert_eq!(faces.len(), 4);
    assert_eq!(
        faces.iter().map(|face| face.data.len()).sum::<usize>(),
        128 * 1024 * 1024
    );
}

#[test]
fn nested_eot_headers_are_skipped() {
    let depth = 4096;
    let mut encoded = vec![0; depth * 82];
    for index in 0..depth {
        let offset = index * 82;
        let length = (depth - index - 1) * 82 + FONT.len();
        encoded[offset + 4..offset + 8].copy_from_slice(&(length as u32).to_le_bytes());
        encoded[offset + 34..offset + 36].copy_from_slice(&[0x4c, 0x50]);
    }
    encoded.extend_from_slice(FONT);
    let faces = embedded_font_definitions("ppt/presentation.xml", XML, RELS, |_| Some(&encoded));
    assert!(faces.is_empty());
}

#[test]
fn oversized_embedded_fonts_are_skipped() {
    let mut bytes = vec![0; 32 * 1024 * 1024 + 1];
    bytes[..4].copy_from_slice(&[0, 1, 0, 0]);
    let faces = embedded_font_definitions("ppt/presentation.xml", XML, RELS, |_| Some(&bytes));
    assert!(faces.is_empty());
}

#[test]
fn eot_payloads_require_a_font_signature() {
    let mut encoded = vec![0; 86];
    encoded[4..8].copy_from_slice(&4_u32.to_le_bytes());
    encoded[34..36].copy_from_slice(&[0x4c, 0x50]);
    encoded[82..].copy_from_slice(b"font");
    let faces = embedded_font_definitions("ppt/presentation.xml", XML, RELS, |_| Some(&encoded));
    assert!(faces.is_empty());
}

#[test]
fn decodes_obfuscated_document_faces_and_relative_relationships() {
    let key = "00112233-4455-6677-8899-AABBCCDDEEFF";
    let mask: [u8; 16] = [
        0, 0x11, 0x22, 0x33, 0x44, 0x55, 0x66, 0x77, 0x88, 0x99, 0xaa, 0xbb, 0xcc, 0xdd, 0xee, 0xff,
    ];
    let mut encoded = FONT.to_vec();
    for (index, byte) in encoded[..32].iter_mut().enumerate() {
        *byte ^= mask[15 - index % 16];
    }
    let xml = format!(
        r#"<w:fonts xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:font w:name="Embedded"><w:embedBold r:id="face" w:fontKey="{key}"/></w:font></w:fonts>"#
    );
    let faces = embedded_font_definitions("word/tables/fonts.xml", xml.as_bytes(), RELS, |path| {
        (path == "word/fonts/face.odttf").then_some(encoded.as_slice())
    });
    assert_eq!(faces.len(), 1);
    assert_eq!(faces[0].family, "Embedded");
    assert!(faces[0].bold);
    assert_eq!(faces[0].data, FONT);
}

#[test]
fn extracts_uncompressed_embedded_font_payloads() {
    let mut encoded = vec![0; 82];
    encoded[4..8].copy_from_slice(&(FONT.len() as u32).to_le_bytes());
    encoded[34..36].copy_from_slice(&[0x4c, 0x50]);
    encoded.extend_from_slice(FONT);
    let xml = br#"<p:embeddedFontLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:embeddedFont><p:font typeface="Embedded"/><p:regular r:id="face"/></p:embeddedFont></p:embeddedFontLst>"#;
    let faces = embedded_font_definitions("ppt/presentation.xml", xml, RELS, |_| {
        Some(encoded.as_slice())
    });
    assert_eq!(faces.len(), 1);
    assert_eq!(faces[0].data, FONT);
}

#[test]
fn external_font_relationships_are_inert() {
    let xml = br#"<p:embeddedFont xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:font typeface="Embedded"/><p:regular r:id="face"/></p:embeddedFont>"#;
    let relationships = br#"<Relationships><Relationship Id="face" Target="https://invalid/font.ttf" TargetMode="External"/></Relationships>"#;
    let faces = embedded_font_definitions("ppt/presentation.xml", xml, relationships, |_| {
        panic!("external font was read")
    });
    assert!(faces.is_empty());
}

#[test]
fn unresolvable_references_are_bounded() {
    let mut xml = String::from("<w:fonts xmlns:w='w' xmlns:r='r'>");
    for index in 0..5000 {
        xml.push_str(&format!(
            "<w:font w:name='Missing {index}'><w:embedRegular r:id='missing'/></w:font>"
        ));
    }
    xml.push_str("<w:font w:name='Present'><w:embedRegular r:id='face'/></w:font></w:fonts>");
    let rels = br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="missing" Target="../fonts/missing.odttf"/><Relationship Id="face" Target="../fonts/face.odttf"/></Relationships>"#;
    let font = include_bytes!("fixtures/Synthetic-Regular.ttf");
    let reads = std::cell::Cell::new(0);
    let faces = embedded_font_definitions("word/fontTable.xml", xml.as_bytes(), rels, |path| {
        reads.set(reads.get() + 1);
        path.ends_with("face.odttf").then_some(&font[..])
    });
    assert!(faces.is_empty());
    assert_eq!(reads.get(), 1);
}

#[test]
fn oversized_family_names_are_skipped() {
    let xml = format!(
        "<w:fonts xmlns:w='w' xmlns:r='r'><w:font w:name='{}'><w:embedRegular r:id='face'/></w:font></w:fonts>",
        "F".repeat(4096)
    );
    let font = include_bytes!("fixtures/Synthetic-Regular.ttf");
    let faces =
        embedded_font_definitions("word/fontTable.xml", xml.as_bytes(), RELS, |_| Some(font));
    assert!(faces.is_empty());
}
