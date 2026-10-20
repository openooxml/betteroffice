use serde_json::{Value, json};

use crate::EngineSession;

const FONT: &[u8] = include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#;
const PICTURE: &str = r#"<w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="7" name="Picture 7" descr="Figure description"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId9"/><a:srcRect t="10000" r="20000" b="30000" l="40000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#;

const ANCHOR: &str = r#"<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="1371600"/><wp:wrapNone/><wp:docPr id="1" name="Box"/>"#;

fn document(content: &str, boxed: bool) -> Vec<u8> {
    let paragraph = format!("<w:p><w:r>{content}</w:r></w:p>");
    document_paragraphs(&paragraph, boxed)
}

fn document_paragraphs(paragraphs: &str, boxed: bool) -> Vec<u8> {
    let body = if boxed {
        format!(
            r#"<w:p><w:r><w:drawing>{ANCHOR}<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>{paragraphs}</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>"#
        )
    } else {
        paragraphs.to_owned()
    };
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture.png"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document {NAMESPACES}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#)),
    ];
    let mut parts: Vec<_> = parts
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value.into_bytes()))
        .collect();
    parts.push((
        "word/media/picture.png".to_owned(),
        vec![
            137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1,
            8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 11, 73, 68, 65, 84, 120, 156, 99, 96, 0, 2,
            0, 0, 5, 0, 1, 169, 245, 232, 39, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
        ],
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn primitives(content: &str, boxed: bool, media_tokens: bool) -> Vec<Value> {
    let display = display(&document(content, boxed), media_tokens);
    assert_eq!(display["pages"].as_array().unwrap().len(), 1);
    display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .clone()
}

fn primitives_for_paragraphs(paragraphs: &str, boxed: bool, media_tokens: bool) -> Vec<Value> {
    let display = display(&document_paragraphs(paragraphs, boxed), media_tokens);
    assert_eq!(display["pages"].as_array().unwrap().len(), 1);
    display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .clone()
}

pub(super) fn display(bytes: &[u8], media_tokens: bool) -> Value {
    display_with_layout(bytes, media_tokens).0
}

fn display_with_layout(bytes: &[u8], media_tokens: bool) -> (Value, Value) {
    display_with_layout_edit(bytes, media_tokens, |_| {})
}

fn display_with_layout_edit(
    bytes: &[u8],
    media_tokens: bool,
    edit: impl FnOnce(&EngineSession),
) -> (Value, Value) {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75301);
    crate::seed::seed_with_layout_tokens(engine.doc(), bytes).unwrap();
    edit(&engine);
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body", "renderEnv": {"mediaTokens": media_tokens},
                "regions": {"sections": [{"properties": package.document.final_section_properties}]},
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    (
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap(),
        serde_json::from_str(&output).unwrap(),
    )
}

fn images(primitives: &[Value]) -> Vec<&Value> {
    primitives
        .iter()
        .filter(|primitive| primitive["kind"] == "image")
        .collect()
}

fn text(primitives: &[Value]) -> String {
    primitives
        .iter()
        .filter(|primitive| matches!(primitive["kind"].as_str(), Some("text" | "glyphRun")))
        .filter_map(|primitive| primitive["text"].as_str())
        .collect()
}

fn text_primitive<'a>(primitives: &'a [Value], word: &str) -> &'a Value {
    let first = &word[..1];
    primitives
        .iter()
        .find(|p| {
            matches!(p["kind"].as_str(), Some("text" | "glyphRun"))
                && p["text"]
                    .as_str()
                    .is_some_and(|text| text.starts_with(first))
        })
        .unwrap_or_else(|| panic!("no text starting with {first}"))
}

fn baseline_y(primitive: &Value) -> f64 {
    primitive["baselineY"]
        .as_f64()
        .or_else(|| primitive["glyphs"][0]["y"].as_f64())
        .unwrap()
}

fn start_x(primitive: &Value) -> f64 {
    primitive["x"]
        .as_f64()
        .or_else(|| primitive["glyphs"][0]["x"].as_f64())
        .unwrap()
}

fn shown_anchor_pages(display: &Value, layout: &Value, anchor_text: &str) -> Vec<usize> {
    let table = layout["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|item| item["block"]["kind"] == "table")
        .unwrap();
    let blocks = table["block"]["rows"][0]["cells"][0]["blocks"]
        .as_array()
        .unwrap();
    let anchor = blocks
        .iter()
        .position(|block| {
            block["kind"] == "paragraph"
                && block["runs"].as_array().unwrap().iter().any(|run| {
                    run["text"]
                        .as_str()
                        .is_some_and(|text| text.contains(anchor_text))
                })
        })
        .unwrap();
    let anchor_id = &blocks[anchor]["id"];
    let identity_field = match anchor_id {
        Value::String(_) => "blockKey",
        Value::Number(_) => "blockId",
        _ => panic!("anchor paragraph has no block identity"),
    };
    let line = &table["measure"]["rows"][0]["cells"][0]["blocks"][anchor]["lines"][0];
    let ascent = line["ascent"].as_f64().unwrap();
    let height = line["lineHeight"].as_f64().unwrap();
    display["pages"]
        .as_array()
        .unwrap()
        .iter()
        .enumerate()
        .filter_map(|(page, value)| {
            value["primitives"]
                .as_array()
                .unwrap()
                .iter()
                .any(|primitive| {
                    if !matches!(primitive["kind"].as_str(), Some("text" | "glyphRun"))
                        || &primitive[identity_field] != anchor_id
                        || primitive["lineIndex"].as_u64() != Some(0)
                    {
                        return false;
                    }
                    let baseline = primitive["baselineY"].as_f64().unwrap_or_else(|| {
                        primitive["glyphs"]
                            .as_array()
                            .unwrap()
                            .iter()
                            .map(|glyph| glyph["y"].as_f64().unwrap())
                            .fold(f64::NEG_INFINITY, f64::max)
                    });
                    let top = baseline - ascent;
                    let clip = &primitive["clipGroup"]["clip"];
                    let clip_top = clip["y"].as_f64().unwrap();
                    let clip_bottom = clip_top + clip["h"].as_f64().unwrap();
                    let inside = (top + height).min(clip_bottom) - top.max(clip_top);
                    inside >= 1.0
                })
                .then_some(page)
        })
        .collect()
}

#[test]
fn vml_group_pictures_paint_at_mapped_positions() {
    let content = r#"<w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:group coordorigin="100,200" coordsize="1200,600" style="position:absolute;margin-left:12pt;margin-top:24pt;width:300pt;height:90pt;mso-position-horizontal-relative:page;mso-position-vertical-relative:page"><v:shape alt="first" style="position:absolute;left:100;top:200;width:300;height:150"><v:imagedata r:id="rId9"/></v:shape><v:shape alt="second" style="position:absolute;left:500;top:250;width:360;height:200"><v:imagedata r:id="rId9"/></v:shape><v:shape alt="third" style="position:absolute;left:900;top:400;width:240;height:300"><v:imagedata r:id="rId9"/></v:shape></v:group></w:pict><w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape alt="separate" style="width:12pt;height:6pt"><v:imagedata r:id="rId9"/></v:shape></w:pict>"#;
    for media_tokens in [true, false] {
        let painted = primitives(content, false, media_tokens);
        let pictures = images(&painted);
        assert_eq!(pictures.len(), 4);
        for (alt, x, y, width, height) in [
            ("first", 16.0, 32.0, 100.0, 30.0),
            ("second", 16.0 + 400.0 / 3.0, 42.0, 120.0, 40.0),
            ("third", 16.0 + 800.0 / 3.0, 72.0, 80.0, 60.0),
        ] {
            let image = pictures
                .iter()
                .find(|image| image["altText"] == alt)
                .unwrap();
            for (key, expected) in [("x", x), ("y", y), ("w", width), ("h", height)] {
                assert!((image[key].as_f64().unwrap() - expected).abs() <= 0.5);
            }
            if media_tokens {
                assert_eq!(image["relId"], "media:0");
            } else {
                assert!(
                    image["relId"]
                        .as_str()
                        .unwrap()
                        .starts_with("data:image/png;base64,")
                );
            }
        }
        let separate = pictures
            .iter()
            .find(|image| image["altText"] == "separate")
            .unwrap();
        assert_eq!(separate["w"], 16);
        assert_eq!(separate["h"], 8);
    }
}

#[test]
fn inline_picture_in_text_box_paints_with_body_image_properties() {
    for media_tokens in [true, false] {
        let boxed = primitives(PICTURE, true, media_tokens);
        let body = primitives(PICTURE, false, media_tokens);
        let boxed_images = images(&boxed);
        let body_images = images(&body);
        assert_eq!(boxed_images.len(), 1);
        assert_eq!(body_images.len(), 1);
        let image = boxed_images[0];
        for key in ["relId", "w", "h", "crop", "altText"] {
            assert_eq!(image[key], body_images[0][key], "{key}");
        }
        if media_tokens {
            assert_eq!(image["relId"], "media:0");
        } else {
            assert!(
                image["relId"]
                    .as_str()
                    .unwrap()
                    .starts_with("data:image/png;base64,")
            );
        }
        assert_eq!(image["w"], 192);
        assert_eq!(image["h"], 96);
        assert_eq!(image["altText"], "Figure description");
        assert_eq!(
            image["crop"],
            json!({"top": 0.1, "right": 0.2, "bottom": 0.3, "left": 0.4})
        );
        let shape = boxed
            .iter()
            .find(|primitive| primitive["kind"] == "shape")
            .unwrap();
        for (position, size) in [("x", "w"), ("y", "h")] {
            let start = image[position].as_f64().unwrap();
            let end = start + image[size].as_f64().unwrap();
            let box_start = shape[position].as_f64().unwrap();
            let box_end = box_start + shape[size].as_f64().unwrap();
            assert!(start >= box_start && end <= box_end);
        }
    }
}

#[test]
fn chartex_choice_paints_its_picture_fallback_once() {
    let alternate = format!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><mc:Choice Requires="cx"><w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="8" name="Chart"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart r:id="rIdChartEx"/></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback>{PICTURE}</mc:Fallback></mc:AlternateContent>"#
    );
    for boxed in [false, true] {
        let mut parts = ooxml_opc::unzip_parts(&document(&alternate, boxed)).unwrap();
        let relationships = &mut parts
            .iter_mut()
            .find(|(path, _)| path == "word/_rels/document.xml.rels")
            .unwrap()
            .1;
        *relationships = String::from_utf8(relationships.clone())
            .unwrap()
            .replace(
                "</Relationships>",
                r#"<Relationship Id="rIdChartEx" Type="http://schemas.microsoft.com/office/2014/relationships/chartEx" Target="charts/chartEx1.xml"/></Relationships>"#,
            )
            .into_bytes();
        parts.push((
            "word/charts/chartEx1.xml".to_owned(),
            br#"<cx:chartSpace xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex"><cx:chart/></cx:chartSpace>"#.to_vec(),
        ));
        let bytes = ooxml_opc::rezip_parts(&parts).unwrap();
        for media_tokens in [true, false] {
            let display = display(&bytes, media_tokens);
            let painted: Vec<_> = display["pages"]
                .as_array()
                .unwrap()
                .iter()
                .flat_map(|page| page["primitives"].as_array().unwrap())
                .cloned()
                .collect();
            let pictures = images(&painted);
            assert_eq!(
                pictures.len(),
                1,
                "boxed={boxed}, media_tokens={media_tokens}"
            );
            if media_tokens {
                assert_eq!(pictures[0]["relId"], "media:0");
            } else {
                assert!(
                    pictures[0]["relId"]
                        .as_str()
                        .unwrap()
                        .starts_with("data:image/png;base64,")
                );
            }
            let fallback = primitives(PICTURE, boxed, media_tokens);
            let expected = images(&fallback);
            assert_eq!(expected.len(), 1);
            for key in ["relId", "w", "h", "crop", "altText"] {
                assert_eq!(pictures[0][key], expected[0][key], "{key}");
            }
            assert!(painted.iter().all(|primitive| primitive["chart"].is_null()));
        }
    }
}

#[test]
fn text_around_a_text_box_picture_still_paints() {
    let display = primitives(
        &format!("<w:t>Before</w:t>{PICTURE}<w:t>After</w:t>"),
        true,
        true,
    );
    assert_eq!(images(&display).len(), 1);
    assert_eq!(text(&display), "BeforeAfter");
}

#[test]
fn text_only_text_box_still_paints() {
    let display = primitives("<w:t>Plain text</w:t>", true, true);
    assert!(images(&display).is_empty());
    assert_eq!(text(&display), "Plain text");
}

#[test]
fn text_box_paragraph_spacing_collapses_between_paragraphs() {
    let unspaced = primitives_for_paragraphs(
        "<w:p><w:r><w:t>First</w:t></w:r></w:p><w:p><w:r><w:t>Second</w:t></w:r></w:p>",
        true,
        true,
    );
    let spaced = primitives_for_paragraphs(
        r#"<w:p><w:pPr><w:spacing w:before="240" w:after="240"/></w:pPr><w:r><w:t>First</w:t></w:r></w:p><w:p><w:pPr><w:spacing w:before="240"/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>"#,
        true,
        true,
    );
    let unspaced_first = baseline_y(text_primitive(&unspaced, "First"));
    let spaced_first = baseline_y(text_primitive(&spaced, "First"));
    let unspaced_second = baseline_y(text_primitive(&unspaced, "Second"));
    let spaced_second = baseline_y(text_primitive(&spaced, "Second"));
    assert!(spaced_second > unspaced_second);
    assert!((spaced_first - unspaced_first - 16.0).abs() < 1e-3);
    assert!((spaced_second - unspaced_second - 32.0).abs() < 1e-3);
}

#[test]
fn text_box_paragraph_exact_line_spacing_gives_32px_line_pitch() {
    let display = primitives_for_paragraphs(
        r#"<w:p><w:pPr><w:spacing w:line="480" w:lineRule="exact"/></w:pPr><w:r><w:t>First</w:t><w:br/><w:t>Second</w:t></w:r></w:p>"#,
        true,
        true,
    );
    let first = baseline_y(text_primitive(&display, "First"));
    let second = baseline_y(text_primitive(&display, "Second"));
    assert!((second - first - 32.0).abs() < 1e-3);
}

#[test]
fn text_box_paragraph_before_lines_spacing_lowers_first_baseline() {
    let unspaced = primitives_for_paragraphs("<w:p><w:r><w:t>First</w:t></w:r></w:p>", true, true);
    let twips = primitives_for_paragraphs(
        r#"<w:p><w:pPr><w:spacing w:before="240"/></w:pPr><w:r><w:t>First</w:t></w:r></w:p>"#,
        true,
        true,
    );
    let lines = primitives_for_paragraphs(
        r#"<w:p><w:pPr><w:spacing w:beforeLines="100"/></w:pPr><w:r><w:t>First</w:t></w:r></w:p>"#,
        true,
        true,
    );
    let both = primitives_for_paragraphs(
        r#"<w:p><w:pPr><w:spacing w:before="240" w:beforeLines="100"/></w:pPr><w:r><w:t>First</w:t></w:r></w:p>"#,
        true,
        true,
    );
    let unspaced_y = baseline_y(text_primitive(&unspaced, "First"));
    let twips_y = baseline_y(text_primitive(&twips, "First"));
    let lines_y = baseline_y(text_primitive(&lines, "First"));
    let both_y = baseline_y(text_primitive(&both, "First"));
    assert!((lines_y - unspaced_y - 16.0).abs() < 1e-3);
    assert!((lines_y - twips_y).abs() < 1e-3);
    assert!((both_y - lines_y).abs() < 1e-3);
}

#[test]
fn text_box_paragraph_hanging_indent_offsets_first_and_following_lines() {
    let lines = "<w:r><w:t>First</w:t><w:br/><w:t>Second</w:t></w:r>";
    let plain = primitives_for_paragraphs(&format!("<w:p>{lines}</w:p>"), true, true);
    let indented = primitives_for_paragraphs(
        &format!(r#"<w:p><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr>{lines}</w:p>"#),
        true,
        true,
    );
    let shift = |needle: &str| {
        start_x(text_primitive(&indented, needle)) - start_x(text_primitive(&plain, needle))
    };
    assert!((shift("First") - 24.0).abs() < 1e-3, "{}", shift("First"));
    assert!((shift("Second") - 48.0).abs() < 1e-3, "{}", shift("Second"));
}

pub(super) fn table_boundary_document(
    bytes: &[u8],
    before_twips: u32,
    anchor_lines: usize,
) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    let source = String::from_utf8(xml.clone()).unwrap();
    let (prefix, body) = source.split_once("<w:body>").unwrap();
    let prefix = prefix.replace(
        "<w:document ",
        r#"<w:document xmlns:v="urn:schemas-microsoft-com:vml" "#,
    );
    let prefix = if prefix.contains("xmlns:mc=") {
        prefix
    } else {
        prefix.replace("<w:document ", r#"<w:document xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#)
    };
    let (boxed, section) = body.split_once("<w:sectPr>").unwrap();
    let boxed = boxed
        .replace("<w:r><w:t>Before</w:t></w:r>", "")
        .replace("<w:r><w:t>After</w:t></w:r>", "");
    let boxed = if let Some((_, inner)) = boxed.split_once("<w:txbxContent>") {
        let inner = inner.split_once("</w:txbxContent>").unwrap().0;
        let drawing = boxed.split_once("<w:drawing>").unwrap().1;
        let drawing = drawing.split_once("</wp:anchor></w:drawing>").unwrap().0;
        format!(
            r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>{drawing}</wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape id="fallback-box" style="width:216pt;height:108pt"><v:textbox><w:txbxContent>{inner}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r></w:p>"#
        )
    } else {
        boxed
    };
    let continuation: String = (1..anchor_lines)
        .map(|line| format!("<w:br/><w:t>Anchor continuation {line:02};</w:t>"))
        .collect();
    let boxed = format!(
        "{}<w:r><w:t>Anchor paragraph</w:t>{continuation}</w:r></w:p>",
        boxed.strip_suffix("</w:p>").unwrap()
    );
    let tail: String = (0..20)
        .map(|line| {
            format!(
                r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>Cell tail {line:02};</w:t></w:r></w:p>"#
            )
        })
        .collect();
    let document = format!(
        r#"{prefix}<w:body><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{before_twips}" w:lineRule="exact"/></w:pPr><w:r><w:t>Before table</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>{boxed}{tail}</w:tc></w:tr></w:tbl><w:sectPr>{section}"#
    );
    *xml = document.into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn assert_table_boundary_drawing(bytes: &[u8], kinds: &[&str]) {
    for anchor_lines in [1, 30] {
        for before in [10080, 10800, 11520, 12000, 12240, 12480, 12600] {
            let bytes = table_boundary_document(bytes, before, anchor_lines);
            let (display, layout) = display_with_layout(&bytes, true);
            let pages = display["pages"].as_array().unwrap();
            let anchor_pages = shown_anchor_pages(&display, &layout, "Anchor paragraph");
            assert_eq!(
                anchor_pages.len(),
                1,
                "before={before}, anchor_lines={anchor_lines}: anchor paragraph"
            );
            for kind in kinds {
                let painted: Vec<_> = pages
                    .iter()
                    .enumerate()
                    .filter_map(|(page, value)| {
                        let primitives = value["primitives"].as_array().unwrap();
                        let paints = if *kind == "text" {
                            text(primitives).contains("Plain text")
                        } else {
                            primitives.iter().any(|p| p["kind"] == *kind)
                        };
                        paints.then_some(page)
                    })
                    .collect();
                assert_eq!(
                    painted.len(),
                    1,
                    "before={before}, anchor_lines={anchor_lines}: each anchored drawing paints on exactly one page ({kind}): {painted:?}"
                );
                assert_eq!(
                    painted, anchor_pages,
                    "before={before}, anchor_lines={anchor_lines}: {kind} anchor page"
                );
            }
            let cell_text: String = pages
                .iter()
                .map(|page| text(page["primitives"].as_array().unwrap()))
                .collect();
            for line in 0..20 {
                assert!(
                    cell_text.contains(&format!("Cell tail {line:02};")),
                    "before={before}, anchor_lines={anchor_lines}: missing in-flow line {line}"
                );
            }
            for line in 1..anchor_lines {
                assert!(
                    cell_text.contains(&format!("Anchor continuation {line:02};")),
                    "before={before}, anchor_lines={anchor_lines}: missing anchor continuation {line}"
                );
            }
        }
    }
}

#[test]
fn table_boundary_text_box_picture_and_text_controls() {
    assert_table_boundary_drawing(&document(PICTURE, true), &["shape", "image"]);
    assert_table_boundary_drawing(&document("<w:t>Plain text</w:t>", true), &["shape", "text"]);
}

#[test]
fn table_boundary_plain_anchored_picture_paints_once() {
    assert_table_boundary_drawing(&document(&anchored_picture(), false), &["image"]);
}

#[test]
fn table_boundary_anchored_shape_without_text_paints_once() {
    assert_table_boundary_drawing(&document(&anchored_shape(), false), &["shape"]);
}

#[test]
fn table_boundary_vertically_merged_picture_paints_once_on_anchor_page() {
    for leading_lines in [0, 12] {
        let picture = anchored_picture()
            .replace("cy=\"1371600\"", "cy=\"2743200\"")
            .replace("cy=\"914400\"", "cy=\"2743200\"");
        let bytes = rewrite_document(&document(&picture, false), |source| {
            let (prefix, body) = source.split_once("<w:body>").unwrap();
            let (paragraph, section) = body.split_once("<w:sectPr>").unwrap();
            let paragraph =
                paragraph.replace("</w:p>", "<w:r><w:t>Merged anchor</w:t></w:r></w:p>");
            let leading = r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>Merged lead;</w:t></w:r></w:p>"#.repeat(leading_lines);
            let tail = r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>Merged tail;</w:t></w:r></w:p>"#.repeat(12);
            let rows: String = (0..6)
                .map(|row| {
                    let (merge, content) = if row == 0 {
                        (r#"<w:vMerge w:val="restart"/>"#, format!("{leading}{paragraph}{tail}"))
                    } else {
                        ("<w:vMerge/>", "<w:p/>".to_owned())
                    };
                    format!(
                        r#"<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="1440" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/>{merge}</w:tcPr>{content}</w:tc><w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Merge row {row:02};</w:t></w:r></w:p></w:tc></w:tr>"#
                    )
                })
                .collect();
            format!(
                r#"{prefix}<w:body><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="10800" w:lineRule="exact"/></w:pPr><w:r><w:t>Before table</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>{rows}</w:tbl><w:sectPr>{section}"#
            )
        });
        let (display, layout) = display_with_layout(&bytes, true);
        let pages = display["pages"].as_array().unwrap();
        assert!(pages.len() > 1, "leading_lines={leading_lines}");
        assert!(text(pages[0]["primitives"].as_array().unwrap()).contains("Merge row 00;"));
        assert!(text(pages[1]["primitives"].as_array().unwrap()).contains("Merge row 05;"));
        let anchor_pages = shown_anchor_pages(&display, &layout, "Merged anchor");
        let image_pages: Vec<_> = pages
            .iter()
            .enumerate()
            .flat_map(|(page, value)| {
                images(value["primitives"].as_array().unwrap())
                    .into_iter()
                    .map(move |_| page)
            })
            .collect();
        assert_eq!(anchor_pages.len(), 1, "leading_lines={leading_lines}");
        assert_eq!(image_pages, anchor_pages, "leading_lines={leading_lines}");
        assert_eq!(anchor_pages, vec![usize::from(leading_lines > 0)]);
        let painted = images(pages[anchor_pages[0]]["primitives"].as_array().unwrap());
        assert_eq!(painted[0]["cell"]["rowSpan"], 6);
        assert_eq!(painted[0]["cell"]["row"], 0);
        if leading_lines > 0 {
            assert_eq!(painted[0]["cell"]["continuation"], true);
        }
    }
}

#[test]
fn table_boundary_in_flow_wrap_square_shape_is_not_lost() {
    let shape =
        anchored_shape().replace("<wp:wrapNone/>", r#"<wp:wrapSquare wrapText="bothSides"/>"#);
    let mut separates_shape_and_paragraph = false;
    for before in [10080, 10800, 11520, 12000] {
        let bytes = table_boundary_document(&document(&shape, false), before, 1);
        let (display, layout) = display_with_layout(&bytes, true);
        let pages = display["pages"].as_array().unwrap();
        assert!(pages.len() > 1, "before={before}");
        let shape_pages: Vec<_> = pages
            .iter()
            .enumerate()
            .flat_map(|(page, value)| {
                value["primitives"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter(|primitive| primitive["kind"] == "shape")
                    .map(move |_| page)
            })
            .collect();
        let anchor_pages = shown_anchor_pages(&display, &layout, "Anchor paragraph");
        assert!(!shape_pages.is_empty(), "before={before}");
        assert_eq!(anchor_pages.len(), 1, "before={before}");
        separates_shape_and_paragraph |= shape_pages[0] < anchor_pages[0];
        let cell_text: String = pages
            .iter()
            .map(|page| text(page["primitives"].as_array().unwrap()))
            .collect();
        for line in 0..20 {
            assert!(cell_text.contains(&format!("Cell tail {line:02};")));
        }
    }
    assert!(separates_shape_and_paragraph);
}

#[test]
fn table_boundary_empty_picture_anchor_after_overlay_and_nested_float() {
    for nested_float in [false, true] {
        let bytes = table_boundary_document(&document(&anchored_picture(), false), 10080, 1);
        let bytes = rewrite_document(&bytes, |source| {
            let shape = anchored_shape().replace("cy=\"1371600\"", "cy=\"2743200\"");
            let nested = if nested_float {
                r#"<w:tbl><w:tblPr><w:tblpPr w:vertAnchor="text" w:horzAnchor="margin" w:tblpXSpec="right" w:tblpY="0"/><w:tblW w:w="2000" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:trPr><w:trHeight w:val="2160" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl>"#
            } else {
                ""
            };
            source
                .replacen(
                    "</w:tcPr>",
                    &format!("</w:tcPr><w:p><w:r>{shape}</w:r></w:p>{nested}"),
                    1,
                )
                .replace("<w:r><w:t>Anchor paragraph</w:t></w:r>", "")
        });
        let display = display(&bytes, true);
        let pages = display["pages"].as_array().unwrap();
        assert_eq!(pages.len(), 2, "nested_float={nested_float}");
        assert!(text(pages[0]["primitives"].as_array().unwrap()).contains("Cell tail 00;"));
        assert!(text(pages[1]["primitives"].as_array().unwrap()).contains("Cell tail 19;"));
        let image_pages: Vec<_> = pages
            .iter()
            .enumerate()
            .flat_map(|(page, value)| {
                images(value["primitives"].as_array().unwrap())
                    .into_iter()
                    .map(move |_| page)
            })
            .collect();
        assert_eq!(image_pages, vec![0], "nested_float={nested_float}");
        let shape_pages: Vec<_> = pages
            .iter()
            .enumerate()
            .flat_map(|(page, value)| {
                value["primitives"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .filter(|primitive| primitive["kind"] == "shape")
                    .map(move |_| page)
            })
            .collect();
        assert_eq!(shape_pages, vec![0], "nested_float={nested_float}");
    }
}

#[test]
fn table_boundary_unsplit_exact_row_keeps_offset_picture() {
    let picture = anchored_picture().replace(
        "<wp:posOffset>0</wp:posOffset>",
        "<wp:posOffset>-914400</wp:posOffset>",
    );
    let bytes = rewrite_document(&document(&picture, false), |source| {
        let (prefix, body) = source.split_once("<w:body>").unwrap();
        let (paragraph, section) = body.split_once("<w:sectPr>").unwrap();
        let leading = r#"<w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr><w:r><w:t>Lead</w:t></w:r></w:p>"#.repeat(6);
        format!(
            r#"{prefix}<w:body><w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr><w:tblGrid><w:gridCol w:w="9360"/></w:tblGrid><w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="1440" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="9360" w:type="dxa"/></w:tcPr>{leading}{paragraph}</w:tc></w:tr></w:tbl><w:sectPr>{section}"#
        )
    });
    let display = display(&bytes, true);
    let pages = display["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    assert_eq!(images(pages[0]["primitives"].as_array().unwrap()).len(), 1);
}

fn rewrite_document(bytes: &[u8], rewrite: impl FnOnce(&str) -> String) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    *xml = rewrite(std::str::from_utf8(xml).unwrap()).into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn anchored_picture() -> String {
    PICTURE
        .replace("<wp:inline>", ANCHOR)
        .replace("</wp:inline>", "</wp:anchor>")
}

fn anchored_shape() -> String {
    format!(
        r#"<w:drawing>{ANCHOR}<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>"#
    )
}

#[test]
fn grouped_text_boxes_paint_once_at_composed_child_positions() {
    let paragraph = |text| format!("<w:p><w:r><w:t>{text}</w:t></w:r></w:p>");
    let text_box = |id, x, y, content: &str| {
        format!(
            r#"<wps:wsp><wps:cNvPr id="{id}"/><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/></wps:spPr><wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx><wps:bodyPr lIns="9525" tIns="19050" rIns="0" bIns="0"/></wps:wsp>"#
        )
    };
    let alpha = text_box(2, 190500, 95250, &paragraph("Alpha"));
    let beta_content = format!(
        "<w:sdt><w:sdtPr/><w:sdtContent><w:tbl><w:tr><w:tc>{}</w:tc></w:tr></w:tbl></w:sdtContent></w:sdt>",
        paragraph("Beta")
    );
    let beta = text_box(3, 95250, 47625, &beta_content);
    let origin = |primitive: &Value| {
        if primitive["kind"] == "glyphRun" {
            (
                primitive["glyphs"][0]["x"].as_f64().unwrap(),
                primitive["glyphs"][0]["y"].as_f64().unwrap(),
            )
        } else {
            (
                primitive["x"].as_f64().unwrap(),
                primitive["baselineY"].as_f64().unwrap(),
            )
        }
    };
    for (nested, picture) in [(false, false), (true, false), (true, true)] {
        let beta = if nested {
            format!(
                r#"<wpg:grpSp><wpg:grpSpPr><a:xfrm><a:off x="1047750" y="428625"/><a:ext cx="762000" cy="381000"/><a:chOff x="0" y="0"/><a:chExt cx="381000" cy="190500"/></a:xfrm></wpg:grpSpPr>{beta}</wpg:grpSp>"#
            )
        } else {
            text_box(3, 1143000, 476250, &beta_content)
        };
        let picture = if picture {
            &PICTURE[PICTURE.find("<pic:pic>").unwrap()
                ..PICTURE.find("</pic:pic>").unwrap() + "</pic:pic>".len()]
        } else {
            ""
        };
        let group = format!(
            r#"<w:drawing><wp:inline><wp:extent cx="3810000" cy="1905000"/><wp:docPr id="1" name="Group"/><a:graphic><a:graphicData><wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/><a:chOff x="95250" y="47625"/><a:chExt cx="1905000" cy="952500"/></a:xfrm></wpg:grpSpPr>{picture}{alpha}{beta}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>"#
        );
        let bytes = document(&group, false);
        let (display, _) = display_with_layout(&bytes, true);
        let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
            .unwrap()
            .document
            .package;
        let request = serde_json::from_value(json!({
            "determinism": {"seed": "0".repeat(64), "now": "2030-01-02T03:04:05.006Z"},
            "document": {"content": package.document.content},
            "relationshipEntries": package.relationship_entries,
            "selective": {"changedParaIds": []},
            "options": {"updateModifiedDate": false}
        }))
        .unwrap();
        let saved = docx_parse::write_docx_s13(request, &bytes).unwrap();
        assert_eq!(
            ooxml_opc::unzip_parts(&saved).unwrap(),
            ooxml_opc::unzip_parts(&bytes).unwrap()
        );
        let primitives = display["pages"][0]["primitives"].as_array().unwrap();
        let painted = text(primitives);
        let pictures: Vec<_> = primitives
            .iter()
            .filter(|p| p["fillPaint"]["kind"] == "picture")
            .collect();
        assert_eq!(pictures.len(), usize::from(!picture.is_empty()));
        if let Some(picture) = pictures.first() {
            assert_eq!(picture["fillPaint"]["pictureSrc"], "media:0");
        }
        assert_eq!(painted.matches("Alpha").count(), 1);
        assert_eq!(painted.matches("Beta").count(), 1);
        let shapes: Vec<_> = primitives.iter().filter(|p| p["kind"] == "shape").collect();
        let root = shapes[0];
        for (label, dx, dy) in [
            ("Alpha", 20.0, 10.0),
            (
                "Beta",
                if nested { 240.0 } else { 220.0 },
                if nested { 100.0 } else { 90.0 },
            ),
        ] {
            let label = first_glyph(primitives, label);
            let child_x = root["x"].as_f64().unwrap() + dx;
            let child_y = root["y"].as_f64().unwrap() + dy;
            let child = shapes
                .iter()
                .find(|p| {
                    (p["x"].as_f64().unwrap() - child_x).abs() < 0.01
                        && (p["y"].as_f64().unwrap() - child_y).abs() < 0.01
                })
                .unwrap();
            let (x, y) = origin(label);
            assert!((x - child_x - 1.0).abs() < 0.01);
            assert!(y >= child_y + 2.0 && y < child_y + child["h"].as_f64().unwrap());
        }
    }
    let control = primitives("<w:t>Standalone</w:t>", true, true);
    assert_eq!(text(&control), "Standalone");
    let shape = control.iter().find(|p| p["kind"] == "shape").unwrap();
    let label = first_glyph(&control, "Standalone");
    assert!((origin(label).0 - shape["x"].as_f64().unwrap()).abs() < 0.01);
}

#[test]
fn nested_group_text_boxes_with_fractional_transforms_paint_inside_child_boxes() {
    let groups = [
        (
            "Alpha", -19050, 45085, 641985, 2261, 2248, 1011, 2261, 2493, 2461,
        ),
        (
            "Beta", 6926580, 38735, 622935, 3387, 2109, 981, 3518, 2364, 3387,
        ),
        (
            "Gamma", 0, 3933825, 632460, 2366, 2237, 996, 2366, 2493, 2551,
        ),
        (
            "Delta", 6926580, 3933825, 615950, 3462, 2108, 970, 3582, 2364, 3462,
        ),
    ];
    let children: String = groups
        .iter()
        .enumerate()
        .map(|(index, group)| {
            let (label, x, y, width, child_x, child_y, child_width, line_x, line_y, text_x) = group;
            let id = index * 3 + 2;
            let line_id = id + 1;
            let text_id = id + 2;
            format!(
                r#"<wpg:grpSp><wpg:cNvPr id="{id}"/><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="184150"/><a:chOff x="{child_x}" y="{child_y}"/><a:chExt cx="{child_width}" cy="290"/></a:xfrm></wpg:grpSpPr><wps:wsp><wps:cNvPr id="{line_id}"/><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="{line_x}" y="{line_y}"/><a:ext cx="850" cy="0"/></a:xfrm><a:prstGeom prst="straightConnector1"/><a:noFill/><a:ln w="6350"><a:solidFill><a:srgbClr val="333333"/></a:solidFill><a:prstDash val="dash"/></a:ln></wps:spPr><wps:bodyPr/></wps:wsp><wps:wsp><wps:cNvPr id="{text_id}"/><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="{text_x}" y="{child_y}"/><a:ext cx="811" cy="290"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:sz w:val="18"/></w:rPr><w:t>{label}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t" anchorCtr="0" upright="1"><a:noAutofit/></wps:bodyPr></wps:wsp></wpg:grpSp>"#
            )
        })
        .collect();
    let group = format!(
        r#"<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>3287395</wp:posOffset></wp:positionV><wp:extent cx="7568640" cy="4078800"/><wp:effectExtent l="0" t="0" r="32385" b="17145"/><wp:wrapNone/><wp:docPr id="1" name="Group"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="7568640" cy="4078800"/><a:chOff x="-19050" y="38735"/><a:chExt cx="7568565" cy="4079240"/></a:xfrm></wpg:grpSpPr>{children}</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing>"#
    );
    let mut parts = ooxml_opc::unzip_parts(&document(&group, false)).unwrap();
    let (_, xml) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    *xml = std::str::from_utf8(xml)
        .unwrap()
        .replace(r#"w:w="12240" w:h="15840""#, r#"w:w="11906" w:h="16838""#)
        .replace(
            r#"w:top="1440" w:right="1440" w:bottom="1440" w:left="1440""#,
            r#"w:top="680" w:right="1134" w:bottom="680" w:left="1134""#,
        )
        .into_bytes();
    let shown = display(&ooxml_opc::rezip_parts(&parts).unwrap(), true);
    let pages = shown["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    let page = &pages[0];
    let primitives = page["primitives"].as_array().unwrap();
    let root = primitives.iter().find(|p| p["kind"] == "shape").unwrap();
    let root_x = root["x"].as_f64().unwrap();
    let root_y = root["y"].as_f64().unwrap();
    let scale_x = 7568640.0 / 7568565.0;
    let scale_y = 4078800.0 / 4079240.0;
    assert!(
        (root_x - (page["width"].as_f64().unwrap() - root["w"].as_f64().unwrap()) / 2.0).abs()
            < 0.01
    );
    assert!((root_y - 3287395.0 / 9525.0).abs() < 0.01);
    for (index, (label, x, y, width, child_x, _, child_width, _, _, text_x)) in
        groups.iter().enumerate()
    {
        assert_eq!(text(primitives).matches(*label).count(), 1);
        let key = format!(
            "{}:child:{index}:child:1",
            root["blockKey"].as_str().unwrap()
        );
        let child = primitives
            .iter()
            .find(|p| p["kind"] == "shape" && p["blockKey"] == key)
            .unwrap();
        let child_x = root_x
            + (f64::from(*x)
                + 19050.0
                + f64::from(text_x - child_x) * f64::from(*width) / f64::from(*child_width))
                * scale_x
                / 9525.0;
        let child_y = root_y + (f64::from(*y) - 38735.0) * scale_y / 9525.0;
        let child_width = 811.0 * f64::from(*width) / f64::from(*child_width) * scale_x / 9525.0;
        let child_height = 184150.0 * scale_y / 9525.0;
        assert!((child["x"].as_f64().unwrap() - child_x).abs() < 0.01);
        assert!((child["y"].as_f64().unwrap() - child_y).abs() < 0.01);
        assert!((child["w"].as_f64().unwrap() - child_width).abs() < 0.01);
        assert!((child["h"].as_f64().unwrap() - child_height).abs() < 0.01);
        assert!(child_x >= 0.0 && child_x + child_width <= page["width"].as_f64().unwrap());
        assert!(child_y >= 0.0 && child_y + child_height <= page["height"].as_f64().unwrap());
        let label = first_glyph(primitives, label);
        let glyphs: Vec<_> = if label["kind"] == "glyphRun" {
            label["glyphs"]
                .as_array()
                .unwrap()
                .iter()
                .map(|glyph| (glyph["x"].as_f64().unwrap(), glyph["y"].as_f64().unwrap()))
                .collect()
        } else {
            vec![(
                label["x"].as_f64().unwrap(),
                label["baselineY"].as_f64().unwrap(),
            )]
        };
        assert!(glyphs.iter().all(|(x, y)| {
            *x >= child_x
                && *x < child_x + child_width
                && *y >= child_y
                && *y < child_y + child_height
        }));
    }
}

fn glyph_origin(primitive: &Value) -> (f64, f64) {
    if primitive["kind"] == "glyphRun" {
        (
            primitive["glyphs"][0]["x"].as_f64().unwrap(),
            primitive["glyphs"][0]["y"].as_f64().unwrap(),
        )
    } else {
        (
            primitive["x"].as_f64().unwrap(),
            primitive["baselineY"].as_f64().unwrap(),
        )
    }
}

fn group_text_child(text: &str, x: i64, y: i64, width: i64, height: i64) -> String {
    format!(
        r#"<wps:wsp><wps:cNvPr id="2"/><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="{x}" y="{y}"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="20"/></w:rPr><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></wps:bodyPr></wps:wsp>"#
    )
}

fn text_group(children: &str, width: i64, height: i64, viewport: &str) -> String {
    format!(
        r#"<w:drawing><wp:inline><wp:extent cx="{width}" cy="{height}"/><wp:docPr id="1" name="Group"/><a:graphic><a:graphicData><wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{width}" cy="{height}"/>{viewport}</a:xfrm></wpg:grpSpPr>{children}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
}

fn force_group_fallback(engine: &EngineSession) {
    use crate::{EditCtx, RawOp, SegmentContent};

    let shape = engine
        .doc()
        .story_segments("body")
        .unwrap()
        .into_iter()
        .find_map(|segment| {
            let SegmentContent::OtherEmbed { kind, payload } = segment.content else {
                return None;
            };
            if kind != "shape" {
                return None;
            }
            let yrs::Any::String(source) = payload.get("shapeJson")? else {
                return None;
            };
            serde_json::from_str::<Value>(source).ok()
        })
        .unwrap();
    let sources = engine.doc().source_index().unwrap().group_picture_sources();
    assert!(crate::bridge::group_pictures::project(&shape, "body", &sources, true).is_some());
    let mut edited = shape;
    edited["scene"]["version"] = json!(2);
    assert!(crate::bridge::group_pictures::project(&edited, "body", &sources, true).is_none());
    engine
        .doc()
        .apply_raw_ops(
            "body",
            vec![
                RawOp::Delete { index: 0, len: 1 },
                RawOp::InsertEmbed {
                    index: 0,
                    kind: "shape".into(),
                    payload: vec![("shapeJson".into(), yrs::Any::from(edited.to_string()))],
                    attrs: yrs::types::Attrs::new(),
                },
            ],
            &EditCtx::local("", "2030-01-02T03:04:05.006Z"),
        )
        .unwrap();
}

#[test]
fn raw_nested_group_text_matches_projected_boxes_with_rotation_and_flips() {
    for (outer, inner) in [
        ("", ""),
        (r#"rot="5400000" flipH="1""#, r#"rot="16200000" flipV="1""#),
        (r#"flipV="1""#, r#"rot="1800000" flipH="1""#),
    ] {
        let child = group_text_child("Nested", 309, 1839, 3375, 363);
        let nested = format!(
            r#"<wpg:grpSp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:grpSpPr><a:xfrm {inner}><a:off x="266" y="0"/><a:ext cx="8803" cy="4570"/><a:chOff x="266" y="0"/><a:chExt cx="8803" cy="4570"/></a:xfrm></wpg:grpSpPr>{child}</wpg:grpSp>"#
        );
        let group = text_group(
            &nested,
            5779090,
            3074670,
            r#"<a:chOff x="0" y="0"/><a:chExt cx="9074" cy="4842"/>"#,
        )
        .replacen("<a:xfrm>", &format!("<a:xfrm {outer}>"), 1);
        let bytes = document(&group, false);
        let projected = display(&bytes, true);
        let raw = display_with_layout_edit(&bytes, true, force_group_fallback).0;
        let expected = projected["pages"][0]["primitives"].as_array().unwrap();
        let actual = raw["pages"][0]["primitives"].as_array().unwrap();
        assert_eq!(text(actual), "Nested");
        assert_eq!(actual.len(), expected.len());
        for (actual, expected) in actual.iter().zip(expected) {
            assert_eq!(actual["kind"], expected["kind"]);
            for field in ["x", "y", "w", "h", "baselineY"] {
                if let Some(expected) = expected[field].as_f64() {
                    assert!((actual[field].as_f64().unwrap() - expected).abs() < 0.01);
                }
            }
            if let Some(expected) = expected["rotationDeg"].as_f64() {
                assert!((actual["rotationDeg"].as_f64().unwrap() - expected).abs() < 0.01);
            }
            if let Some(glyphs) = actual["glyphs"].as_array() {
                for (actual, expected) in glyphs.iter().zip(expected["glyphs"].as_array().unwrap())
                {
                    for axis in ["x", "y"] {
                        assert!(
                            (actual[axis].as_f64().unwrap() - expected[axis].as_f64().unwrap())
                                .abs()
                                < 0.01
                        );
                    }
                }
            }
        }
    }
}

#[test]
fn grouped_paragraph_anchors_move_labels_onto_the_page() {
    for (width, height, viewport, x, y, child_width, child_height, offset) in [
        (
            1386272,
            1937650,
            r#"<a:chOff x="-4445" y="-373191"/><a:chExt cx="1386274" cy="1937651"/>"#,
            119449,
            326558,
            1262380,
            247650,
            49479,
        ),
        (
            1800225,
            362585,
            r#"<a:chOff x="0" y="0"/><a:chExt cx="1800225" cy="362585"/>"#,
            85725,
            0,
            1586230,
            362585,
            2239645,
        ),
    ] {
        let group = text_group(
            &group_text_child("Label", x, y, child_width, child_height),
            width,
            height,
            viewport,
        );
        let anchor = ANCHOR
            .replace(
                "<wp:posOffset>0</wp:posOffset>",
                &format!("<wp:posOffset>{offset}</wp:posOffset>"),
            )
            .replace(
                r#"cx="2743200" cy="1371600""#,
                &format!(r#"cx="{width}" cy="{height}""#),
            );
        let group = group
            .replace(&format!(r#"<wp:inline><wp:extent cx="{width}" cy="{height}"/><wp:docPr id="1" name="Group"/>"#), &anchor)
            .replace("</wp:inline>", "</wp:anchor>");
        let bytes = rewrite_document(&document(&group, false), |xml| {
            xml.replace("<w:body>", r#"<w:body><w:p><w:pPr><w:spacing w:line="12000" w:lineRule="exact"/></w:pPr><w:r><w:t>Spacer</w:t></w:r></w:p>"#)
        });
        let shown = display(&bytes, true);
        let pages = shown["pages"].as_array().unwrap();
        assert_eq!(pages.len(), 2);
        assert!(!text(pages[0]["primitives"].as_array().unwrap()).contains("Label"));
        let primitives = pages[1]["primitives"].as_array().unwrap();
        assert_eq!(text(primitives).matches("Label").count(), 1);
        let (_, y) = glyph_origin(first_glyph(primitives, "Label"));
        assert!(y >= 0.0 && y < pages[1]["height"].as_f64().unwrap());
    }
}

#[test]
fn grouped_fixed_text_body_does_not_paint_over_its_caption() {
    let content = (0..40)
        .map(|_| "<w:p><w:r><w:t>Inside</w:t></w:r></w:p>")
        .collect::<String>();
    let body = group_text_child("Inside", 0, -60305, 3552825, 1712595).replace(
        "<w:p><w:r><w:rPr><w:sz w:val=\"20\"/></w:rPr><w:t>Inside</w:t></w:r></w:p>",
        &content,
    );
    let caption = group_text_child("Caption", 34505, 1727703, 3444875, 266700);
    let group = text_group(
        &format!("{body}{caption}"),
        3552825,
        2052955,
        r#"<a:chOff x="0" y="-60305"/><a:chExt cx="3552825" cy="2054708"/>"#,
    );
    let shown = display(&document(&group, false), true);
    let primitives = shown["pages"][0]["primitives"].as_array().unwrap();
    let caption_y = glyph_origin(first_glyph(primitives, "Caption")).1;
    let mut inside: Vec<_> = primitives
        .iter()
        .filter(|p| {
            matches!(p["kind"].as_str(), Some("text" | "glyphRun"))
                && p["text"].as_str().is_some_and(|text| {
                    !text.is_empty() && ("Inside".starts_with(text) || text.starts_with("Inside"))
                })
        })
        .map(|p| glyph_origin(p).1)
        .collect();
    inside.sort_by(f64::total_cmp);
    inside.dedup_by(|a, b| (*a - *b).abs() < 1e-6);
    assert!(!inside.is_empty() && inside.len() < 40);
    assert!(inside.iter().all(|y| *y < caption_y));
    assert_eq!(text(primitives).matches("Caption").count(), 1);
}

#[test]
fn grouped_alternate_content_label_is_painted_once_with_charts() {
    let child = group_text_child("Q1 | Q2", -577, 3111, 9296, 3239)
        .replace(r#"w:val="20""#, r#"w:val="32""#)
        .replace(
            r#"lIns="0" tIns="0" rIns="0" bIns="0""#,
            r#"lIns="91440" tIns="45720" rIns="91440" bIns="45720""#,
        );
    let group = text_group(
        &child,
        997591,
        996950,
        r#"<a:chOff x="-626" y="-179"/><a:chExt cx="9360" cy="9360"/>"#,
    );
    let fallback = group.replace("Q1 | Q2", "Fallback");
    let alternate = format!(
        r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wpg">{group}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
    );
    let chart = r#"<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#;
    let mut parts =
        ooxml_opc::unzip_parts(&document(&format!("{alternate}{chart}"), false)).unwrap();
    for (path, bytes) in &mut parts {
        if path == "word/_rels/document.xml.rels" {
            *bytes = std::str::from_utf8(bytes).unwrap().replace("</Relationships>", r#"<Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>"#).into_bytes();
        } else if path == "[Content_Types].xml" {
            *bytes = std::str::from_utf8(bytes).unwrap().replace("</Types>", r#"<Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#).into_bytes();
        }
    }
    parts.push(("word/charts/chart1.xml".into(), br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/><c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>Quarter</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#.to_vec()));
    let shown = display(&ooxml_opc::rezip_parts(&parts).unwrap(), true);
    assert!(shown["pages"].as_array().unwrap().iter().any(|page| {
        page["primitives"]
            .as_array()
            .unwrap()
            .iter()
            .any(|primitive| !primitive["chart"].is_null())
    }));
    let painted: String = shown["pages"]
        .as_array()
        .unwrap()
        .iter()
        .map(|page| text(page["primitives"].as_array().unwrap()))
        .collect();
    assert_eq!(painted.matches("Q1 | Q2").count(), 1);
    assert!(!painted.contains("Fallback"));
}

#[test]
fn grouped_vertical_and_rotated_text_uses_the_child_frame() {
    for (vertical, rotation, expected) in [
        ("vert", 0, 90.0),
        ("vert270", 0, -90.0),
        ("horz", 5400000, 90.0),
    ] {
        let child = group_text_child("Turn", 952500, 476250, 952500, 476250)
            .replace("<a:xfrm>", &format!(r#"<a:xfrm rot="{rotation}">"#))
            .replace(
                "<wps:bodyPr ",
                &format!(r#"<wps:bodyPr vert="{vertical}" "#),
            );
        let group = text_group(
            &child,
            3810000,
            1905000,
            r#"<a:chOff x="0" y="0"/><a:chExt cx="3810000" cy="1905000"/>"#,
        );
        let shown = display(&document(&group, false), true);
        let primitives = shown["pages"][0]["primitives"].as_array().unwrap();
        assert_eq!(text(primitives), "Turn");
        let glyph = first_glyph(primitives, "Turn");
        assert!((glyph["rotationDeg"].as_f64().unwrap() - expected).abs() < 0.01);
        let clip = &glyph["clipGroup"]["clip"];
        let (x, y) = glyph_origin(glyph);
        assert!(x >= clip["x"].as_f64().unwrap());
        assert!(y >= clip["y"].as_f64().unwrap());
        assert!(x < clip["x"].as_f64().unwrap() + clip["w"].as_f64().unwrap());
        assert!(y < clip["y"].as_f64().unwrap() + clip["h"].as_f64().unwrap());
    }
}

#[test]
fn nested_group_no_wrap_text_keeps_one_line_inside_its_box() {
    let child = group_text_child("One line in the box", 309, 1839, 3375, 363)
        .replace("<wps:bodyPr ", r#"<wps:bodyPr wrap="none" "#);
    let nested = format!(
        r#"<wpg:grpSp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9069" cy="4570"/><a:chOff x="0" y="0"/><a:chExt cx="9069" cy="4570"/></a:xfrm></wpg:grpSpPr><wpg:grpSp><wpg:grpSpPr><a:xfrm><a:off x="266" y="0"/><a:ext cx="8803" cy="4570"/><a:chOff x="266" y="0"/><a:chExt cx="8803" cy="4570"/></a:xfrm></wpg:grpSpPr>{child}</wpg:grpSp></wpg:grpSp>"#
    );
    let group = text_group(
        &nested,
        5779090,
        3074670,
        r#"<a:chOff x="0" y="0"/><a:chExt cx="9074" cy="4842"/>"#,
    );
    let bytes = document(&group, false);
    for shown in [
        display(&bytes, true),
        display_with_layout_edit(&bytes, true, force_group_fallback).0,
    ] {
        let primitives = shown["pages"][0]["primitives"].as_array().unwrap();
        assert_eq!(text(primitives), "One line in the box");
        let glyph = first_glyph(primitives, "One line in the box");
        let clip = &glyph["clipGroup"]["clip"];
        let (x, y) = glyph_origin(glyph);
        assert!(
            x >= clip["x"].as_f64().unwrap()
                && x < clip["x"].as_f64().unwrap() + clip["w"].as_f64().unwrap()
        );
        assert!(
            y >= clip["y"].as_f64().unwrap()
                && y < clip["y"].as_f64().unwrap() + clip["h"].as_f64().unwrap()
        );
    }
}

fn first_glyph<'a>(primitives: &'a [Value], label: &str) -> &'a Value {
    primitives
        .iter()
        .find(|p| {
            p["text"]
                .as_str()
                .is_some_and(|text| !text.is_empty() && label.starts_with(text))
        })
        .unwrap()
}

fn grouped_drawing(picture_first: bool) -> String {
    let text_box = |text| {
        format!(
            r#"<wps:wsp><wps:cNvPr id="2"/><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="95250" y="47625"/><a:ext cx="952500" cy="476250"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>{text}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>"#
        )
    };
    let picture = &PICTURE[PICTURE.find("<pic:pic>").unwrap()
        ..PICTURE.find("</pic:pic>").unwrap() + "</pic:pic>".len()];
    let boxes = format!("{}{}", text_box("Alpha"), text_box("Beta"));
    let children = if picture_first {
        format!("{picture}{boxes}")
    } else {
        format!("{boxes}{picture}")
    };
    format!(
        r#"<w:drawing><wp:inline><wp:extent cx="3810000" cy="1905000"/><wp:docPr id="1" name="Group"/><a:graphic><a:graphicData><wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="3810000" cy="1905000"/><a:chOff x="95250" y="47625"/><a:chExt cx="1905000" cy="952500"/></a:xfrm></wpg:grpSpPr>{children}</wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
}

fn parsed_groups(value: &Value) -> Vec<&Value> {
    let mut groups = Vec::new();
    match value {
        Value::Object(object) => {
            if object.get("type").and_then(Value::as_str) == Some("shape")
                && object.contains_key("scene")
            {
                groups.push(value);
            } else {
                for child in object.values() {
                    groups.extend(parsed_groups(child));
                }
            }
        }
        Value::Array(array) => {
            for child in array {
                groups.extend(parsed_groups(child));
            }
        }
        _ => {}
    }
    groups
}

fn parsed_document(bytes: &[u8]) -> Value {
    serde_json::to_value(
        docx_parse::parse_docx_s9_wire(bytes, Default::default())
            .unwrap()
            .document
            .package
            .document,
    )
    .unwrap()
}

#[test]
fn identical_groups_in_one_run_each_keep_their_stories() {
    let group = grouped_drawing(true);
    let parsed = parsed_document(&document(&format!("{group}{group}"), false));
    let groups = parsed_groups(&parsed);
    assert_eq!(groups.len(), 2);
    for group in groups {
        for (index, label) in ["Alpha", "Beta"].iter().enumerate() {
            assert_eq!(
                group["children"][index]["textBody"]["content"][0]["content"][0]["content"][0]["text"],
                *label
            );
            assert_eq!(
                group["scene"]["root"]["children"][index + 1]["shape"]["textBody"],
                group["children"][index]["textBody"]
            );
        }
    }
}

#[test]
fn grouped_stories_follow_inline_containers_and_selected_alternate_branch() {
    let group = grouped_drawing(true);
    let run = format!("<w:r>{group}</w:r>");
    let wrappers = [
        format!(r#"<w:hyperlink w:anchor="target">{run}</w:hyperlink>"#),
        format!(r#"<w:fldSimple w:instr="REF target">{run}</w:fldSimple>"#),
        format!("<w:sdt><w:sdtPr/><w:sdtContent>{run}</w:sdtContent></w:sdt>"),
        format!(
            r#"<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>REF target</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>{run}<w:r><w:fldChar w:fldCharType="end"/></w:r>"#
        ),
    ];
    for wrapper in wrappers {
        let bytes = rewrite_document(&document(&group, false), |xml| xml.replace(&run, &wrapper));
        let parsed = parsed_document(&bytes);
        let groups = parsed_groups(&parsed);
        assert!(!groups.is_empty());
        for group in groups {
            let bodies = group["children"].to_string();
            assert!(bodies.contains("Alpha") && bodies.contains("Beta"));
        }
    }
    for choice in [group.clone(), "<w:drawing/>".into()] {
        let fallback = group.replace("Alpha", "FallbackAlpha");
        let alternate = format!(
            r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wpg">{choice}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
        );
        let parsed = parsed_document(&document(&alternate, false));
        let groups = parsed_groups(&parsed);
        assert_eq!(groups.len(), 1);
        let body = groups[0]["children"][0]["textBody"].to_string();
        assert!(body.contains(if choice == group {
            "\"Alpha\""
        } else {
            "FallbackAlpha"
        }));
    }
}

#[test]
fn grouped_picture_projection_preserves_enriched_stories_and_media_modes() {
    use crate::bridge::group_pictures::project;
    use crate::{EditingDoc, SegmentContent};

    for (picture_first, ids) in [(true, true), (false, true), (true, false)] {
        let mut group = grouped_drawing(picture_first);
        if !ids {
            group = group
                .replace(r#"<wp:docPr id="1" name="Group"/>"#, "")
                .replace(r#"<wps:cNvPr id="2"/>"#, "");
        }
        let bytes = document(&group, false);
        let doc = EditingDoc::new(1);
        crate::seed::seed_with_layout_tokens(&doc, &bytes).unwrap();
        let sources = doc.source_index().unwrap().group_picture_sources();
        let shape = doc
            .story_segments("body")
            .unwrap()
            .into_iter()
            .find_map(|segment| {
                let SegmentContent::OtherEmbed { kind, payload } = segment.content else {
                    return None;
                };
                if kind != "shape" {
                    return None;
                }
                let yrs::Any::String(json) = payload.get("shapeJson")? else {
                    return None;
                };
                serde_json::from_str::<Value>(json).ok()
            })
            .unwrap();
        for media_tokens in [true, false] {
            let projected = project(&shape, "body", &sources, media_tokens).unwrap();
            let children = projected["children"].as_array().unwrap();
            assert_eq!(children.len(), 3);
            let picture = children
                .iter()
                .find(|child| child["fillPaint"]["kind"] == "picture")
                .unwrap();
            let src = picture["fillPaint"]["picture"]["src"].as_str().unwrap();
            assert!(if media_tokens {
                src == "media:0"
            } else {
                src.starts_with("data:image/png;base64,")
            });
            let bodies: Vec<_> = children
                .iter()
                .filter_map(|child| child.get("textBody"))
                .collect();
            assert_eq!(bodies.len(), 2);
            assert!(bodies[0].to_string().contains("Alpha"));
            assert!(bodies[1].to_string().contains("Beta"));
            let lowered = super::lower_shape_json(
                &projected,
                1,
                &crate::bridge::RenderEnv {
                    media_tokens,
                    ..Default::default()
                },
            )
            .unwrap();
            for child in lowered.children {
                assert_eq!(child.pm_start, None);
                assert_eq!(child.doc_start, None);
                assert_eq!(child.doc_end, None);
            }
            let shown = display(&bytes, media_tokens);
            assert_eq!(
                text(shown["pages"][0]["primitives"].as_array().unwrap()),
                "AlphaBeta"
            );
        }
        for field in ["size", "transform", "children"] {
            let mut edited = shape.clone();
            edited[field] = json!({"edited": true});
            assert!(project(&edited, "body", &sources, false).is_none());
        }
    }
}

#[test]
fn skipped_rotated_group_picture_does_not_shift_text_bodies() {
    let group = grouped_drawing(true)
        .replace("<pic:blipFill>", r#"<pic:blipFill rotWithShape="0">"#)
        .replace("<pic:spPr><a:xfrm>", r#"<pic:spPr><a:xfrm rot="5400000">"#);
    let shown = display(&document(&group, false), false);
    let primitives = shown["pages"][0]["primitives"].as_array().unwrap();
    assert_eq!(text(primitives), "AlphaBeta");
    assert!(
        primitives
            .iter()
            .all(|primitive| primitive["fillPaint"]["kind"] != "picture")
    );
}
