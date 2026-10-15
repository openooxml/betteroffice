#[path = "common/ole_vml.rs"]
mod fixture;

use pptx_parse::{
    GraphicFrameData, ParseLimits, Picture, PptxPackage, ShapeNode, parse_pptx,
    parse_pptx_with_limits, write_pptx,
};

fn picture(package: &PptxPackage) -> Option<&Picture> {
    let ShapeNode::GraphicFrame(frame) = &package.slides[0].shapes[0] else {
        panic!("expected a graphic frame");
    };
    let GraphicFrameData::Unknown { picture, .. } = &frame.data else {
        panic!("expected an OLE frame");
    };
    picture.as_deref()
}

#[test]
fn ole_preview_resolves_media_through_the_vml_relationships() {
    let bytes = fixture::deck(fixture::VML, fixture::OBJECT);
    let package = parse_pptx(&bytes).unwrap();
    let preview = picture(&package).expect("expected the legacy preview");
    assert_eq!(preview.media_part_path.as_deref(), Some(fixture::MEDIA));
    assert_eq!(preview.relationship_id, None);
    assert!(preview.crop.is_whole());
    assert_eq!(preview.geometry, "rect");
    let ShapeNode::GraphicFrame(frame) = &package.slides[0].shapes[0] else {
        unreachable!();
    };
    assert_eq!(preview.base, frame.base);
    assert!(
        package
            .media
            .iter()
            .any(|media| media.part_path == fixture::MEDIA)
    );
    assert_eq!(
        ooxml_opc::unzip_parts(&write_pptx(&package).unwrap()).unwrap(),
        ooxml_opc::unzip_parts(&bytes).unwrap()
    );
}

#[test]
fn ole_fallback_resolves_office_spid_and_relationship_id() {
    let vml = fixture::VML
        .replace(
            "id=\"_x0000_s3075\"",
            "id=\"local\" o:spid=\"_x0000_s3075\"",
        )
        .replace("o:relid=", "r:id=");
    let object = format!(
        "<mc:AlternateContent><mc:Choice Requires=\"v\"><p:oleObj spid=\"wrong\"/></mc:Choice><mc:Fallback>{}</mc:Fallback></mc:AlternateContent>",
        fixture::OBJECT
    );
    let package = parse_pptx(&fixture::deck(&vml, &object)).unwrap();
    assert_eq!(
        picture(&package).unwrap().media_part_path.as_deref(),
        Some(fixture::MEDIA)
    );
}

#[test]
fn an_unmatched_vml_shape_keeps_the_placeholder() {
    let vml = fixture::VML.replace("_x0000_s3075", "_x0000_s9999");
    let package = parse_pptx(&fixture::deck(&vml, fixture::OBJECT)).unwrap();
    assert!(picture(&package).is_none());
}

#[test]
fn malformed_vml_does_not_fail_the_open() {
    let vml = fixture::VML.replace("</xml>", "<v:shape></xml>");
    let package = parse_pptx(&fixture::deck(&vml, fixture::OBJECT)).unwrap();
    assert!(picture(&package).is_none());
}

#[test]
fn incomplete_vml_previews_keep_the_placeholder() {
    for missing in [
        "ppt/drawings/vmlDrawing1.vml",
        "ppt/drawings/_rels/vmlDrawing1.vml.rels",
        fixture::MEDIA,
    ] {
        let mut parts = fixture::parts(fixture::VML, fixture::OBJECT);
        parts.retain(|(path, _)| path != missing);
        let package = parse_pptx(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
        assert!(picture(&package).is_none(), "{missing}");
    }
    for vml in [
        fixture::VML.replace("<v:imagedata o:relid=\"rId1\" o:title=\"\"/>", ""),
        fixture::VML.replace("o:relid=\"rId1\"", "o:relid=\"missing\""),
    ] {
        let package = parse_pptx(&fixture::deck(&vml, fixture::OBJECT)).unwrap();
        assert!(picture(&package).is_none());
    }
    let mut parts = fixture::parts(fixture::VML, fixture::OBJECT);
    let (_, rels) = parts
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/_rels/slide1.xml.rels")
        .unwrap();
    *rels = String::from_utf8(rels.clone())
        .unwrap()
        .replace("/vmlDrawing", "/unrelated")
        .into_bytes();
    let package = parse_pptx(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    assert!(picture(&package).is_none());
}

#[test]
fn vml_over_the_parse_limits_keeps_the_placeholder_and_opens() {
    for (limits, vml) in [
        (
            ParseLimits {
                max_xml_bytes: 8192,
                ..ParseLimits::default()
            },
            fixture::VML.replace("</xml>", &format!("{}</xml>", " ".repeat(16384))),
        ),
        (
            ParseLimits {
                max_xml_events: 128,
                ..ParseLimits::default()
            },
            fixture::VML.replace("</xml>", &format!("{}</xml>", "<v:shape/>".repeat(256))),
        ),
        (
            ParseLimits {
                max_xml_depth: 16,
                ..ParseLimits::default()
            },
            fixture::VML.replace(
                "</xml>",
                &format!(
                    "{}{}</xml>",
                    "<v:group>".repeat(32),
                    "</v:group>".repeat(32)
                ),
            ),
        ),
    ] {
        let package = parse_pptx_with_limits(&fixture::deck(&vml, fixture::OBJECT), &limits)
            .expect("optional VML must not reject the slide");
        assert!(picture(&package).is_none());
    }
}

#[test]
fn an_inline_ole_picture_takes_precedence_over_vml() {
    let object = fixture::OBJECT.replace(
        "<p:embed/>",
        r#"<p:embed/><p:pic><p:blipFill><a:blip r:embed="rId1"/></p:blipFill></p:pic>"#,
    );
    let package = parse_pptx(&fixture::deck(fixture::VML, &object)).unwrap();
    assert_eq!(
        picture(&package).unwrap().media_part_path.as_deref(),
        Some("ppt/media/wrong.png")
    );
}
