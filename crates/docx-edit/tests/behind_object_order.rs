use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

fn document(
    shape_rank: Option<u64>,
    image_rank: Option<u64>,
    shape_first: bool,
    image_behind: bool,
    shape_behind: bool,
) -> Vec<u8> {
    let fixture: Value =
        serde_json::from_str(include_str!("fixtures/behind_objects.json")).unwrap();
    let rank = |xml: &str, marker: &str, value: Option<u64>| match value {
        Some(value) => xml.replace(marker, &value.to_string()),
        None => xml.replace(&format!(r#" relativeHeight="{marker}""#), ""),
    };
    let shape = rank(fixture["shape"].as_str().unwrap(), "SHAPE_RANK", shape_rank)
        .replace("SHAPE_BEHIND", if shape_behind { "1" } else { "0" });
    let image = rank(fixture["image"].as_str().unwrap(), "IMAGE_RANK", image_rank)
        .replace("IMAGE_BEHIND", if image_behind { "1" } else { "0" });
    let body = if shape_first {
        format!("{shape}{image}")
    } else {
        format!("{image}{shape}")
    };
    let xml = format!(
        r#"<w:document {}><w:body>{body}{}</w:body></w:document>"#,
        fixture["namespaces"].as_str().unwrap(),
        fixture["body"].as_str().unwrap()
    );
    let parts = [
        ("[Content_Types].xml", br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec()),
        ("_rels/.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec()),
        ("word/_rels/document.xml.rels", br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.png"/></Relationships>"#.to_vec()),
        ("word/document.xml", xml.into_bytes()),
        ("word/media/image.png", fixture["png"].as_array().unwrap().iter().map(|byte| byte.as_u64().unwrap() as u8).collect()),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, bytes)| (name.to_owned(), bytes))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn layout(bytes: &[u8]) -> (Value, Vec<Value>) {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75101);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output: Value = serde_json::from_str(&engine.layout_document_with_regions_json(&json!({
        "bodyStory":"body", "renderEnv":{},
        "options":{"pageSize":{"w":816,"h":1056},"margins":{"top":96,"right":96,"bottom":96,"left":96}},
        "measurement":{"fontChains":{"calibri|0|0":[font]},"defaults":{"fontFamily":"Calibri","fontSize":12}}
    }).to_string()).unwrap()).unwrap();
    let display: Value = serde_json::from_str(
        &docx_layout::display_list::build_display_list_json(&output.to_string()).unwrap(),
    )
    .unwrap();
    (
        output,
        display["pages"][0]["primitives"]
            .as_array()
            .unwrap()
            .clone(),
    )
}

fn index(primitives: &[Value], label: &str) -> usize {
    primitives
        .iter()
        .position(|primitive| match label {
            "SHAPE" => primitive["kind"] == "shape",
            "IMAGE" => primitive["kind"] == "image",
            _ => {
                let Some(text) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
                    return false;
                };
                if !label.starts_with(text) {
                    return false;
                }
                let first = primitives
                    .iter()
                    .position(|candidate| std::ptr::eq(candidate, primitive))
                    .unwrap();
                let content = primitives[first..]
                    .iter()
                    .take_while(|candidate| candidate["blockKey"] == primitive["blockKey"])
                    .filter_map(|candidate| candidate["text"].as_str())
                    .collect::<String>();
                content == label
            }
        })
        .unwrap_or_else(|| panic!("missing {label}: {primitives:?}"))
}

#[test]
fn imported_picture_presets_survive_inline_and_anchored_layout() {
    for preset in ["rect", "ellipse", "roundRect"] {
        for inline in [false, true] {
            let mut parts =
                ooxml_opc::unzip_parts(&document(None, None, true, true, true)).unwrap();
            let (_, xml) = parts
                .iter_mut()
                .find(|(name, _)| name == "word/document.xml")
                .unwrap();
            let mut text = String::from_utf8(xml.clone()).unwrap().replace(
                r#"<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>"#,
                &format!(r#"<a:prstGeom prst="{preset}"><a:avLst/></a:prstGeom></pic:spPr>"#),
            );
            if inline {
                let picture = text.find(r#"<wp:docPr id="2""#).unwrap();
                let start = text[..picture].rfind("<wp:anchor").unwrap();
                let extent = start + text[start..].find("<wp:extent").unwrap();
                let end = picture + text[picture..].find("</wp:anchor>").unwrap();
                let content = text[extent..end].replace("<wp:wrapNone/>", "");
                text.replace_range(
                    start..end + "</wp:anchor>".len(),
                    &format!("<wp:inline>{content}</wp:inline>"),
                );
            }
            *xml = text.into_bytes();
            let (output, primitives) = layout(&ooxml_opc::rezip_parts(&parts).unwrap());
            let image = output["measured"]
                .as_array()
                .unwrap()
                .iter()
                .filter_map(|block| block["block"]["runs"].as_array())
                .flatten()
                .find(|run| run["kind"] == "image")
                .unwrap();
            let expected = (preset != "rect").then_some(preset);
            assert_eq!(image["shapeType"].as_str(), expected);
            assert_eq!(
                primitives[index(&primitives, "IMAGE")]["shapeType"].as_str(),
                expected
            );
            assert_eq!(image["width"], 192.0);
            assert_eq!(image["height"], 96.0);
        }
    }
}

#[test]
fn imported_behind_images_and_shapes_follow_relative_height() {
    for shape_first in [false, true] {
        for (shape_rank, image_rank) in [(10, 20), (20, 10), (0, 4_294_967_295)] {
            let (output, primitives) = layout(&document(
                Some(shape_rank),
                Some(image_rank),
                shape_first,
                true,
                true,
            ));
            let image = output["measured"]
                .as_array()
                .unwrap()
                .iter()
                .filter_map(|block| block["block"]["runs"].as_array())
                .flatten()
                .find(|run| run["kind"] == "image")
                .unwrap();
            assert_eq!(image["position"]["relativeHeight"], image_rank);
            assert_eq!(image["position"]["horizontal"]["posOffset"], 914400.0);
            let shape_index = index(&primitives, "SHAPE");
            let image_index = index(&primitives, "IMAGE");
            let body_index = index(&primitives, "BODY");
            let shape_text = index(&primitives, "SHAPE TEXT");
            assert_eq!(shape_index < image_index, shape_rank < image_rank);
            assert!(shape_index < shape_text);
            assert!(shape_text < body_index);
            assert!(image_index < body_index);
            assert_eq!(shape_text < image_index, shape_rank < image_rank);
        }
    }
}

#[test]
fn tied_and_missing_ranks_preserve_source_object_order() {
    for rank in [None, Some(20)] {
        for shape_first in [false, true] {
            let (_, primitives) = layout(&document(rank, rank, shape_first, true, true));
            assert_eq!(
                index(&primitives, "SHAPE") < index(&primitives, "IMAGE"),
                shape_first
            );
        }
    }
}

#[test]
fn front_objects_keep_their_body_pass_order() {
    let (_, front_image) = layout(&document(Some(10), Some(0), true, false, true));
    assert!(index(&front_image, "SHAPE") < index(&front_image, "BODY"));
    assert!(index(&front_image, "BODY") < index(&front_image, "IMAGE"));
    let (_, front_shape) = layout(&document(Some(0), Some(10), false, true, false));
    assert!(index(&front_shape, "IMAGE") < index(&front_shape, "SHAPE"));
    assert!(index(&front_shape, "SHAPE") < index(&front_shape, "BODY"));
}

#[test]
fn invalid_image_ranks_do_not_discard_valid_anchor_geometry() {
    for rank in ["-1", "1.5", "4294967296", "18446744073709551615"] {
        let mut parts =
            ooxml_opc::unzip_parts(&document(Some(10), Some(20), true, true, true)).unwrap();
        let (_, document) = parts
            .iter_mut()
            .find(|(name, _)| name == "word/document.xml")
            .unwrap();
        *document = String::from_utf8(document.clone())
            .unwrap()
            .replace(
                "relativeHeight=\"20\"",
                &format!("relativeHeight=\"{rank}\""),
            )
            .into_bytes();
        let (output, _) = layout(&ooxml_opc::rezip_parts(&parts).unwrap());
        let image = output["measured"]
            .as_array()
            .unwrap()
            .iter()
            .filter_map(|block| block["block"]["runs"].as_array())
            .flatten()
            .find(|run| run["kind"] == "image")
            .unwrap();
        assert!(image["position"]["relativeHeight"].is_null());
        for axis in ["horizontal", "vertical"] {
            assert_eq!(image["position"][axis]["relativeTo"], "page");
            assert_eq!(image["position"][axis]["posOffset"], 914400.0);
        }
    }
}

#[test]
fn equal_ranks_follow_source_order_inside_one_paragraph() {
    for rank in [None, Some(20)] {
        for shape_first in [false, true] {
            let mut parts =
                ooxml_opc::unzip_parts(&document(rank, rank, shape_first, true, true)).unwrap();
            let (_, document) = parts
                .iter_mut()
                .find(|(name, _)| name == "word/document.xml")
                .unwrap();
            *document = String::from_utf8(document.clone())
                .unwrap()
                .replacen("</w:p><w:p>", "", 1)
                .into_bytes();
            let (_, primitives) = layout(&ooxml_opc::rezip_parts(&parts).unwrap());
            assert_eq!(
                index(&primitives, "SHAPE") < index(&primitives, "IMAGE"),
                shape_first,
                "rank {rank:?}, shape first {shape_first}"
            );
        }
    }
}
