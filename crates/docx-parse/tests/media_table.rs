//! A parse against a media table must equal the default parse once its
//! `media:{n}` tokens are read back as the `data:` URLs they stand for.

use std::sync::Arc;

use docx_parse::ParseLimits;
use docx_parse::media::{build_media_map_with_warnings, media_token, media_token_index};
use docx_parse::s9::{
    S9ParseOptions, media_table_parts, media_table_parts_within, parse_docx_s9_preview_from_parts,
    parse_docx_s9_preview_with_media_table, parse_docx_s9_wire_parts_with_limits,
    parse_docx_s9_wire_with_media_table,
};
use serde_json::Value;

const W: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office""#;

fn png(width: u32, height: u32, seed: u8) -> Vec<u8> {
    let mut bytes = b"\x89PNG\r\n\x1a\n\0\0\0\x0dIHDR".to_vec();
    bytes.extend(width.to_be_bytes());
    bytes.extend(height.to_be_bytes());
    bytes.extend((0..200).map(|index: u32| (index as u8).wrapping_mul(seed)));
    bytes
}

/// 2x1 little-endian uncompressed RGB TIFF.
fn tiff() -> Vec<u8> {
    let entries: [(u16, u16, u32, u32); 9] = [
        (256, 3, 1, 2),
        (257, 3, 1, 1),
        (258, 3, 3, 122),
        (259, 3, 1, 1),
        (262, 3, 1, 2),
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

fn inline_picture(relationship: &str, id: u32) -> String {
    format!(
        r#"<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="{id}" name="Picture {id}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="{relationship}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"#
    )
}

fn rels(entries: &[(&str, &str, &str)]) -> Vec<u8> {
    let mut xml =
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#
            .to_owned();
    for (id, kind, target) in entries {
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#
        ));
    }
    xml.push_str("</Relationships>");
    xml.into_bytes()
}

fn package() -> Vec<u8> {
    let body = format!(
        r#"<w:document {W}><w:body><w:p>{}</w:p><w:p><w:commentRangeStart w:id="0"/><w:r><w:t>text</w:t></w:r><w:commentRangeEnd w:id="0"/><w:r><w:commentReference w:id="0"/></w:r></w:p><w:p><w:r><w:pict><v:shape id="_x0000_i1025" style="position:absolute"><v:imagedata r:id="rId5"/></v:shape></w:pict></w:r></w:p><w:p>{}</w:p><w:p><w:r><w:drawing><wp:anchor behindDoc="0" relativeHeight="3" simplePos="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="914400" cy="914400"/><wp:wrapNone/><wp:docPr id="9" name="Shape 9"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="rId5"/><a:stretch><a:fillRect/></a:stretch></a:blipFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rId7"/></w:sectPr></w:body></w:document>"#,
        inline_picture("rId5", 1),
        inline_picture("rId6", 2),
    );
    let header = format!(
        r##"<w:hdr {W}><w:p><w:r><w:pict><v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t75" style="position:absolute;margin-left:0;margin-top:0;width:100pt;height:50pt;z-index:-251657216;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"><v:imagedata r:id="rId1" o:title="mark" gain="19661f" blacklevel="22938f"/></v:shape></w:pict></w:r></w:p></w:hdr>"##
    );
    let comments = format!(
        r#"<w:comments {W}><w:comment w:id="0" w:author="A" w:date="2024-01-01T00:00:00Z"><w:p>{}</w:p></w:comment></w:comments>"#,
        inline_picture("rId1", 3)
    );
    let parts: Vec<(String, Vec<u8>)> = vec![
        (
            "[Content_Types].xml".into(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Default Extension="tif" ContentType="image/tiff"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".into(),
            rels(&[("rId1", "officeDocument", "word/document.xml")]),
        ),
        (
            "word/_rels/document.xml.rels".into(),
            rels(&[
                ("rId5", "image", "media/image1.png"),
                ("rId6", "image", "media/image3.tif"),
                ("rId7", "header", "header1.xml"),
                ("rId8", "comments", "comments.xml"),
            ]),
        ),
        ("word/media/image1.png".into(), png(40, 20, 3)),
        ("word/media/image2.png".into(), png(8, 8, 5)),
        ("word/document.xml".into(), body.into_bytes()),
        (
            "word/_rels/header1.xml.rels".into(),
            rels(&[("rId1", "image", "media/image1.png")]),
        ),
        ("word/header1.xml".into(), header.into_bytes()),
        (
            "word/_rels/comments.xml.rels".into(),
            rels(&[("rId1", "image", "media/image1.png")]),
        ),
        ("word/comments.xml".into(), comments.into_bytes()),
        ("word/media/image3.tif".into(), tiff()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn options() -> S9ParseOptions {
    S9ParseOptions {
        source_ordinals: true,
        determinism_seed: Some("7".repeat(64)),
        ..S9ParseOptions::default()
    }
}

/// Replaces each token in `value` with what `resolve` makes of it; how many.
fn resolve_tokens(value: &mut Value, resolve: &dyn Fn(usize) -> String) -> usize {
    match value {
        Value::String(text) => match media_token_index(text) {
            Some(index) => {
                *text = resolve(index);
                1
            }
            None => 0,
        },
        Value::Array(items) => items
            .iter_mut()
            .map(|item| resolve_tokens(item, resolve))
            .sum(),
        Value::Object(fields) => fields
            .values_mut()
            .map(|field| resolve_tokens(field, resolve))
            .sum(),
        _ => 0,
    }
}

fn count_data_urls(value: &Value) -> usize {
    match value {
        Value::String(text) => usize::from(text.starts_with("data:")),
        Value::Array(items) => items.iter().map(count_data_urls).sum(),
        Value::Object(fields) => fields.values().map(count_data_urls).sum(),
        _ => 0,
    }
}

#[test]
fn a_media_table_parse_equals_the_default_one_read_through_its_tokens() {
    let bytes: Arc<[u8]> = package().into();
    let limits = ParseLimits::default();
    let (eager, eager_parts) =
        parse_docx_s9_wire_parts_with_limits(&bytes, options(), &limits).unwrap();
    let (tokens, parts, table) =
        parse_docx_s9_wire_with_media_table(Arc::clone(&bytes), options(), &limits).unwrap();

    assert_eq!(table.len(), 3);
    assert_eq!(
        (0..3)
            .map(|index| table.path(index).unwrap())
            .collect::<Vec<_>>(),
        [
            "word/media/image1.png",
            "word/media/image2.png",
            "word/media/image3.tif"
        ]
    );
    assert_eq!(
        parts,
        eager_parts
            .iter()
            .filter(|(path, _)| !path.starts_with("word/media/"))
            .cloned()
            .collect::<Vec<_>>()
    );
    assert!(tokens.document.package.media_entries.is_empty());

    let (media, warnings) = build_media_map_with_warnings(&eager_parts);
    assert_eq!(table.warnings(), warnings.as_slice());
    for index in 0..table.len() {
        let file = &media[table.path(index).unwrap()];
        assert_eq!(table.data_url(index).unwrap(), file.data_url);
        assert_eq!(table.mime_type(index).unwrap(), file.mime_type);
        assert_eq!(table.resolve(&media_token(index)).unwrap(), file.data_url);
    }
    assert_eq!(table.resolve(&media_token(3)), None);
    assert_eq!(table.resolve("media:01"), None);

    let mut expected = serde_json::to_value(&eager).unwrap();
    expected["document"]["package"]["mediaEntries"] = Value::Array(Vec::new());
    let mut actual = serde_json::to_value(&tokens).unwrap();
    let replaced = resolve_tokens(&mut actual, &|index| table.data_url(index).unwrap());
    // The body's inline and VML pictures, the TIFF, the shape's picture fill
    // and the comment's picture, twice; the watermark arrives resolved.
    assert_eq!(replaced, 6);
    assert_eq!(count_data_urls(&serde_json::to_value(&tokens).unwrap()), 1);
    assert_eq!(actual, expected);
}

#[test]
fn a_preview_parse_against_a_media_table_resolves_the_same_way() {
    let bytes: Arc<[u8]> = package().into();
    let limits = ParseLimits::default();
    let eager_parts = ooxml_opc::unzip_parts(&bytes).unwrap();
    let eager = parse_docx_s9_preview_from_parts(&eager_parts, 2, options(), &limits)
        .unwrap()
        .unwrap();
    let (parts, table) = media_table_parts(&bytes).unwrap();
    let tokens = parse_docx_s9_preview_with_media_table(&parts, &table, 2, options(), &limits)
        .unwrap()
        .unwrap();
    let mut expected = serde_json::to_value(&eager).unwrap();
    expected["document"]["package"]["mediaEntries"] = Value::Array(Vec::new());
    let mut actual = serde_json::to_value(&tokens).unwrap();
    assert!(resolve_tokens(&mut actual, &|index| table.data_url(index).unwrap()) > 0);
    assert_eq!(actual, expected);
}

#[test]
fn images_declared_past_the_container_budget_refuse_the_package() {
    let bytes: Arc<[u8]> = package().into();
    let package = ooxml_opc::RetainedPackage::new(Arc::clone(&bytes)).unwrap();
    let table = docx_parse::media::MediaTable::new(package).unwrap();
    assert!(
        table
            .check_budget(ooxml_opc::MAX_TOTAL_UNCOMPRESSED_BYTES - 1000)
            .is_ok()
    );
    let error = table
        .check_budget(ooxml_opc::MAX_TOTAL_UNCOMPRESSED_BYTES - 100)
        .unwrap_err();
    assert!(error.contains("inflated size exceeds"), "{error}");
}

#[test]
fn the_other_parts_inflate_within_what_compressed_images_leave_of_the_budget() {
    let image = png(4, 4, 1);
    let mut document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec();
    document.resize(document.len() + 64, b' ');
    let budget = (image.len() + document.len()) as u64 - 1;
    let bytes: Arc<[u8]> = ooxml_opc::rezip_parts(&[
        ("word/media/image1.png".to_owned(), image.clone()),
        ("word/document.xml".to_owned(), document),
    ])
    .unwrap()
    .into();
    let error = media_table_parts_within(&bytes, budget).unwrap_err();
    let remaining = budget - image.len() as u64;
    assert!(matches!(
        error,
        docx_parse::ParseError::Container(message)
            if message == format!("inflated size exceeds {remaining} bytes")
    ));
    assert!(media_table_parts_within(&bytes, budget + 1).is_ok());
}

#[cfg(feature = "tiff")]
#[test]
fn tiff_and_non_media_parts_share_a_budget_before_transcoding() {
    let mut image = tiff();
    image[4..8].copy_from_slice(&u32::MAX.to_le_bytes());
    let mut document = br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#.to_vec();
    document.resize(image.len(), b' ');
    let budget = image.len() as u64;
    let eager_parts = vec![
        ("word/media/fallback.tif".to_owned(), image),
        ("word/document.xml".to_owned(), document),
    ];
    for part in &eager_parts {
        let bytes: Arc<[u8]> = ooxml_opc::rezip_parts(std::slice::from_ref(part))
            .unwrap()
            .into();
        assert!(media_table_parts_within(&bytes, budget).is_ok());
    }
    let bytes: Arc<[u8]> = ooxml_opc::rezip_parts(&eager_parts).unwrap().into();
    let error = media_table_parts_within(&bytes, budget).unwrap_err();
    assert!(matches!(
        error,
        docx_parse::ParseError::Container(message)
            if message == format!("inflated size exceeds {budget} bytes")
    ));

    let (parts, table) = media_table_parts_within(&bytes, budget * 2).unwrap();
    assert_eq!(parts, eager_parts[1..]);
    assert_eq!(table.bytes(0).unwrap().as_ref(), eager_parts[0].1);
    let (media, warnings) = build_media_map_with_warnings(&eager_parts);
    assert_eq!(warnings.len(), 1);
    assert_eq!(table.warnings(), warnings.as_slice());
    assert_eq!(
        table.data_url(0).unwrap(),
        media[table.path(0).unwrap()].data_url
    );
    let package = ooxml_opc::RetainedPackage::new(Arc::clone(&bytes)).unwrap();
    let standalone = docx_parse::media::MediaTable::new(package).unwrap();
    assert_eq!(standalone.warnings(), table.warnings());
    assert_eq!(standalone.data_url(0).unwrap(), table.data_url(0).unwrap());

    let limits = ParseLimits::default();
    let eager = parse_docx_s9_preview_from_parts(&eager_parts, 1, options(), &limits)
        .unwrap()
        .unwrap();
    let tokens = parse_docx_s9_preview_with_media_table(&parts, &table, 1, options(), &limits)
        .unwrap()
        .unwrap();
    assert_eq!(tokens.document.warnings, eager.document.warnings);
    let mut expected = serde_json::to_value(&eager).unwrap();
    expected["document"]["package"]["mediaEntries"] = Value::Array(Vec::new());
    assert_eq!(serde_json::to_value(&tokens).unwrap(), expected);
}

#[test]
fn markup_under_word_media_inflates_with_the_package() {
    let parts = vec![
        ("word/media/image1.png".to_owned(), png(4, 4, 1)),
        ("word/media/notes.bin".to_owned(), b"<w:document/>".to_vec()),
        (
            "word/media/header.bin".to_owned(),
            format!("<!--{} EMF--><w:hdr/>", "x".repeat(36)).into_bytes(),
        ),
    ];
    let bytes: Arc<[u8]> = ooxml_opc::rezip_parts(&parts).unwrap().into();
    let (inflated, table) = media_table_parts(&bytes).unwrap();
    assert_eq!(table.len(), 3);
    assert_eq!(inflated, parts[1..]);
    assert!(table.keeps_compressed("word/media/image1.png"));
    assert!(!table.keeps_compressed("word/media/notes.bin"));
    assert!(!table.keeps_compressed("word/media/header.bin"));
}
