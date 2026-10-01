#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EngineSession, seed_from_docx};
use fixture::{FONT, region_request};
use serde_json::{Value, json};

const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture""#
);
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PNG: &[u8] = &[
    137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 2, 0, 0, 0, 2, 8, 2, 0,
    0, 0, 253, 212, 154, 115, 0, 0, 0, 20, 73, 68, 65, 84, 120, 156, 99, 248, 207, 192, 192, 192,
    240, 159, 225, 63, 3, 195, 127, 0, 21, 246, 3, 253, 154, 140, 125, 85, 0, 0, 0, 0, 73, 69, 78,
    68, 174, 66, 96, 130,
];
const TX1_OUTLINE: &str =
    r#"<a:ln w="3175"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill></a:ln>"#;

fn picture(outline: &str, anchored: bool) -> String {
    let graphic = format!(
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="one.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdImage"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom>{outline}</pic:spPr></pic:pic></a:graphicData></a:graphic>"#
    );
    let extent = r#"<wp:extent cx="914400" cy="457200"/>"#;
    let doc_pr = r#"<wp:docPr id="1" name="Picture"/>"#;
    let drawing = if anchored {
        format!(
            r#"<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>{extent}<wp:wrapSquare wrapText="bothSides"/>{doc_pr}{graphic}</wp:anchor>"#
        )
    } else {
        format!(r#"<wp:inline>{extent}{doc_pr}{graphic}</wp:inline>"#)
    };
    format!(
        r#"<w:p><w:r><w:drawing>{drawing}</w:drawing></w:r><w:r><w:t>Picture</w:t></w:r></w:p>"#
    )
}

fn docx(outline: &str, anchored: bool, theme: Option<&str>) -> Vec<u8> {
    let theme_type = if theme.is_some() {
        r#"<Override PartName="/word/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>"#
    } else {
        ""
    };
    let theme_rel = if theme.is_some() {
        format!(r#"<Relationship Id="rIdTheme" Type="{REL}/theme" Target="theme/theme1.xml"/>"#)
    } else {
        String::new()
    };
    let mut parts: Vec<_> = [
        (
            "[Content_Types].xml",
            format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{theme_type}</Types>"#),
        ),
        (
            "_rels/.rels",
            format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdDoc" Type="{REL}/officeDocument" Target="word/document.xml"/></Relationships>"#),
        ),
        (
            "word/_rels/document.xml.rels",
            format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="{REL}/image" Target="media/one.png"/>{theme_rel}</Relationships>"#),
        ),
        (
            "word/document.xml",
            format!(
                r#"<w:document {NS}><w:body>{}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#,
                picture(outline, anchored)
            ),
        ),
    ]
    .into_iter()
    .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
    .collect();
    parts.push(("word/media/one.png".to_owned(), PNG.to_vec()));
    if let Some(theme) = theme {
        parts.push((
            "word/theme/theme1.xml".to_owned(),
            theme.as_bytes().to_vec(),
        ));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn painted_image(outline: &str, anchored: bool, theme: Option<&str>) -> Value {
    let bytes = docx(outline, anchored, theme);
    let engine = EngineSession::new(76700);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let mut request = region_request(&engine, &bytes, font);
    if theme.is_some() {
        let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
            .unwrap()
            .document
            .package;
        request["renderEnv"]["themeColors"] = json!(package.theme.color_scheme);
    }
    engine
        .layout_document_with_regions_retained_json(&request.to_string())
        .unwrap();
    engine.build_display_list_frame("{}", 0).unwrap();
    let display = engine
        .with_display_list(|display| serde_json::to_value(display).unwrap())
        .unwrap();
    let images: Vec<_> = display["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["primitives"].as_array().unwrap())
        .filter(|primitive| primitive["kind"] == "image")
        .collect();
    assert_eq!(images.len(), 1);
    images[0].clone()
}

fn assert_border(image: &Value, width: f64, color: &str, style: &str) {
    let border = &image["border"];
    let actual = border["width"].as_f64().unwrap();
    assert!((actual - width).abs() < 0.01, "{actual} vs {width}");
    assert_eq!(border["color"], color);
    assert_eq!(border["style"], style);
}

#[test]
fn inline_picture_outlines_paint_their_width_color_and_style() {
    for (outline, width, color, style) in [
        (TX1_OUTLINE, 3175.0 / 9525.0, "#000000", "solid"),
        (
            r#"<a:ln><a:solidFill><a:schemeClr val="bg1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill></a:ln>"#,
            1.0,
            "#BFBFBF",
            "solid",
        ),
        (
            r#"<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln>"#,
            1.33,
            "#FF0000",
            "dashed",
        ),
    ] {
        assert_border(&painted_image(outline, false, None), width, color, style);
    }
}

#[test]
fn picture_outlines_without_a_visible_color_do_not_paint() {
    for outline in [
        r#"<a:ln w="12700"><a:noFill/></a:ln>"#,
        r#"<a:ln w="12700"/>"#,
        "",
    ] {
        assert!(painted_image(outline, false, None).get("border").is_none());
    }
}

#[test]
fn anchored_picture_outlines_paint_the_same_border_as_inline_pictures() {
    assert_border(
        &painted_image(TX1_OUTLINE, true, None),
        3175.0 / 9525.0,
        "#000000",
        "solid",
    );
}

#[test]
fn picture_outlines_resolve_colors_from_the_document_theme() {
    let theme = r#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom"><a:themeElements><a:clrScheme name="Custom"><a:dk1><a:srgbClr val="112233"/></a:dk1></a:clrScheme></a:themeElements></a:theme>"#;
    assert_border(
        &painted_image(TX1_OUTLINE, false, Some(theme)),
        3175.0 / 9525.0,
        "#112233",
        "solid",
    );
}
