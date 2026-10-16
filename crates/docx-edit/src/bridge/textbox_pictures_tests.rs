use serde_json::{Value, json};

use crate::EngineSession;

const FONT: &[u8] = include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#;
const PICTURE: &str = r#"<w:drawing><wp:inline><wp:extent cx="1828800" cy="914400"/><wp:docPr id="7" name="Picture 7" descr="Figure description"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rId9"/><a:srcRect t="10000" r="20000" b="30000" l="40000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#;

const ANCHOR: &str = r#"<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="1371600"/><wp:wrapNone/><wp:docPr id="1" name="Box"/>"#;

fn document(content: &str, boxed: bool) -> Vec<u8> {
    let paragraph = format!("<w:p><w:r>{content}</w:r></w:p>");
    let body = if boxed {
        format!(
            r#"<w:p><w:r><w:drawing>{ANCHOR}<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"/><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr><wps:txbx><w:txbxContent>{paragraph}</w:txbxContent></wps:txbx><wps:bodyPr lIns="0" tIns="0" rIns="0" bIns="0"/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p>"#
        )
    } else {
        paragraph
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

pub(super) fn display(bytes: &[u8], media_tokens: bool) -> Value {
    display_with_layout(bytes, media_tokens).0
}

fn display_with_layout(bytes: &[u8], media_tokens: bool) -> (Value, Value) {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75301);
    crate::seed::seed_with_layout_tokens(engine.doc(), bytes).unwrap();
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
