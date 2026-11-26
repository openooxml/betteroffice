#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EditingDoc, EngineSession, SegmentContent, seed_docx_preview, seed_from_docx};
use docx_layout::display_list::{Primitive, px};
use docx_layout::hit::{HoverTarget, hit_test_regions};
use serde_json::{Value, json};

fn connector_package() -> Vec<u8> {
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
            </Types>"#
                .to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
            </Relationships>"#
                .to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
          xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
          xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <w:body><w:p><w:r><w:drawing><wp:inline>
            <wp:extent cx="914400" cy="457200"/><wp:docPr id="1" name="Connector"/>
            <a:graphic><a:graphicData><wps:wsp><wps:spPr>
              <a:xfrm><a:ext cx="914400" cy="457200"/></a:xfrm>
              <a:prstGeom prst="bentConnector4"><a:avLst>
                <a:gd name="adj1" fmla="val 150000"/><a:gd name="adj2" fmla="val -25000"/>
              </a:avLst></a:prstGeom><a:noFill/>
              <a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>
            </wps:spPr></wps:wsp></a:graphicData></a:graphic>
          </wp:inline></w:drawing></w:r></w:p>
          <w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:left="1440"
            w:bottom="1440" w:right="1440"/></w:sectPr></w:body>
        </w:document>"#
                .to_vec(),
        ),
    ])
    .unwrap()
}

#[test]
fn connector_embed_keeps_the_released_path() {
    let doc = EditingDoc::new(1);
    seed_from_docx(&doc, &connector_package()).unwrap();
    let shape = doc
        .story_segments("body")
        .unwrap()
        .iter()
        .find_map(|segment| match &segment.content {
            SegmentContent::OtherEmbed { kind, payload } if kind == "shape" => {
                let yrs::Any::String(shape) = payload.get("shapeJson")? else {
                    return None;
                };
                Some(serde_json::from_str::<Value>(shape).unwrap())
            }
            _ => None,
        })
        .unwrap();
    assert_eq!(
        shape["geometryPath"],
        json!([
            {"type": "move", "x": 0, "y": 0},
            {"type": "line", "x": 1, "y": 0},
            {"type": "line", "x": 0.25, "y": 0.5},
            {"type": "line", "x": 0.75, "y": 0.5},
            {"type": "line", "x": 1, "y": 1}
        ])
    );
}

fn connector_engine(preview: bool) -> EngineSession {
    let bytes = connector_package();
    let engine = EngineSession::new(1);
    if preview {
        assert!(seed_docx_preview(engine.doc(), &bytes, 1).unwrap());
    } else {
        seed_from_docx(engine.doc(), &bytes).unwrap();
    }
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let request = fixture::region_request(&engine, &bytes, font);
    draw_connector(&engine, &request.to_string());
    engine
}

fn draw_connector(engine: &EngineSession, request: &str) {
    engine
        .layout_document_with_regions_retained_json(request)
        .unwrap();
    engine.build_display_list_frame("{}", 0).unwrap();
}

fn painted_connector(engine: &EngineSession) -> Value {
    let display = engine
        .with_display_list(|display| serde_json::to_value(display).unwrap())
        .unwrap();
    display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .find(|primitive| primitive["kind"] == "shape")
        .unwrap()
        .clone()
}

fn assert_authored_connector_path(engine: &EngineSession, context: &str) {
    let shape = painted_connector(engine);
    let [x, y, w, h] = ["x", "y", "w", "h"].map(|key| shape[key].as_f64().unwrap());
    assert_eq!(
        shape["geometryPath"],
        json!([
            {"type": "move", "x": px(x), "y": px(y)},
            {"type": "line", "x": px(x + w * 1.5), "y": px(y)},
            {"type": "line", "x": px(x + w * 1.5), "y": px(y - h * 0.25)},
            {"type": "line", "x": px(x + w), "y": px(y - h * 0.25)},
            {"type": "line", "x": px(x + w), "y": px(y + h)}
        ]),
        "{context}"
    );
}

#[test]
fn connector_display_list_uses_the_authored_guides() {
    for preview in [false, true] {
        let engine = connector_engine(preview);
        assert_authored_connector_path(&engine, &format!("preview={preview}"));
    }
}

#[test]
fn state_only_connector_replica_draws_the_released_path() {
    let bytes = connector_package();
    let source = EngineSession::new(1);
    seed_from_docx(source.doc(), &bytes).unwrap();
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let request = fixture::region_request(&source, &bytes, font).to_string();
    draw_connector(&source, &request);

    let replica = EngineSession::new(2);
    replica
        .doc()
        .apply_update_v1(&source.doc().encode_state_as_update_v1())
        .unwrap();
    draw_connector(&replica, &request);

    assert_authored_connector_path(&source, "source session");
    let shape = painted_connector(&replica);
    let [x, y, w, h] = ["x", "y", "w", "h"].map(|key| shape[key].as_f64().unwrap());
    assert_eq!(
        shape["geometryPath"],
        json!([
            {"type": "move", "x": px(x), "y": px(y)},
            {"type": "line", "x": px(x + w), "y": px(y)},
            {"type": "line", "x": px(x + w * 0.25), "y": px(y + h * 0.5)},
            {"type": "line", "x": px(x + w * 0.75), "y": px(y + h * 0.5)},
            {"type": "line", "x": px(x + w), "y": px(y + h)}
        ]),
        "state-only replica"
    );
}

#[test]
fn connector_hit_follows_the_painted_detour() {
    connector_engine(false)
        .with_display_list(|display| {
            let shape = display.pages[0]
                .primitives
                .iter()
                .find_map(|primitive| match primitive {
                    Primitive::Shape(shape) => Some(shape),
                    _ => None,
                })
                .unwrap();
            let x = shape.x.as_f64().unwrap();
            let y = shape.y.as_f64().unwrap();
            let w = shape.w.as_f64().unwrap();
            let h = shape.h.as_f64().unwrap();
            let hit = |x, y| hit_test_regions(display, 0, x, y).unwrap().target;
            assert_eq!(hit(x + w * 1.25, y), HoverTarget::Image);
            assert_ne!(hit(x + w * 0.5, y + h * 0.5), HoverTarget::Image);
        })
        .unwrap();
}

fn with_connector_document(edit: impl FnOnce(&str) -> String) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&connector_package()).unwrap();
    let (_, document) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    *document = edit(std::str::from_utf8(document).unwrap()).into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn state_only_connector_with_empty_guides_uses_the_released_default() {
    let bytes = with_connector_document(|xml| {
        xml.replace(
            r#"<a:gd name="adj1" fmla="val 150000"/><a:gd name="adj2" fmla="val -25000"/>"#,
            "",
        )
    });
    let source = EngineSession::new(1);
    seed_from_docx(source.doc(), &bytes).unwrap();
    let replica = EngineSession::new(2);
    replica
        .doc()
        .apply_update_v1(&source.doc().encode_state_as_update_v1())
        .unwrap();
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let request = fixture::region_request(&source, &bytes, font).to_string();
    draw_connector(&replica, &request);
    let shape = painted_connector(&replica);
    let [x, y, w, h] = ["x", "y", "w", "h"].map(|key| shape[key].as_f64().unwrap());
    assert_eq!(
        shape["geometryPath"],
        json!([
            {"type": "move", "x": px(x), "y": px(y)},
            {"type": "line", "x": px(x + w * 0.5), "y": px(y)},
            {"type": "line", "x": px(x + w * 0.25), "y": px(y + h * 0.5)},
            {"type": "line", "x": px(x + w * 0.75), "y": px(y + h * 0.5)},
            {"type": "line", "x": px(x + w), "y": px(y + h)}
        ])
    );
}

#[test]
fn grouped_connector_hits_its_drawing_atom() {
    for anchored in [false, true] {
        let bytes = with_connector_document(|xml| {
            let grouped = xml
                .replace("<a:graphicData>", concat!(
                    r#"<a:graphicData><wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">"#,
                    r#"<wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/>"#,
                    r#"<a:chOff x="0" y="0"/><a:chExt cx="914400" cy="457200"/></a:xfrm></wpg:grpSpPr>"#,
                ))
                .replace("</a:graphicData>", "</wpg:wgp></a:graphicData>");
            if anchored {
                grouped.replace("<wp:inline>", concat!(
                    r#"<wp:anchor simplePos="0" relativeHeight="1" behindDoc="0" layoutInCell="1" allowOverlap="1">"#,
                    r#"<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>"#,
                    r#"<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV><wp:wrapNone/>"#,
                )).replace("</wp:inline>", "</wp:anchor>")
            } else {
                grouped
            }
        });
        let engine = EngineSession::new(1);
        seed_from_docx(engine.doc(), &bytes).unwrap();
        let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
        draw_connector(
            &engine,
            &fixture::region_request(&engine, &bytes, font).to_string(),
        );
        engine
            .with_display_list(|display| {
                let mut shapes =
                    display.pages[0]
                        .primitives
                        .iter()
                        .filter_map(|primitive| match primitive {
                            Primitive::Shape(shape) => Some(shape),
                            _ => None,
                        });
                let frame = shapes.next().unwrap();
                let connector = shapes.next().unwrap();
                let [x, y, w, h] = [&connector.x, &connector.y, &connector.w, &connector.h]
                    .map(|value| value.as_f64().unwrap());
                let pos = frame.attrs.doc_start.unwrap();
                let hit = |x, y| hit_test_regions(display, 0, x, y).unwrap();
                let painted = hit(x + w * 1.25, y);
                assert_eq!(
                    (painted.target, painted.pos),
                    (HoverTarget::Image, Some(pos)),
                    "anchored={anchored}"
                );
                assert_eq!(
                    docx_layout::hit::hit_test(display, 0, x + w * 1.25, y),
                    Some(pos)
                );
                assert_ne!(
                    hit(x + w * 0.5, y + h * 0.5).target,
                    HoverTarget::Image,
                    "anchored={anchored}"
                );
            })
            .unwrap();
    }
}

#[test]
fn preview_connector_keeps_guides_when_the_uncut_story_exceeds_bookmark_limits() {
    let bytes = with_connector_document(|xml| {
        let attributes = (1..docx_parse::xml::ParseLimits::default().max_attributes_per_element)
            .map(|index| format!(r#" b{index}="""#))
            .collect::<String>();
        let tail = format!(
            r#"{}<w:bookmarkStart z:id="1"{attributes}/>"#,
            "<w:p/>".repeat(70)
        );
        xml.replace("<w:body>", r#"<w:body xmlns:z="urn:test">"#)
            .replace("<w:sectPr>", &format!("{tail}<w:sectPr>"))
    });
    let engine = EngineSession::new(1);
    assert!(seed_docx_preview(engine.doc(), &bytes, 1).unwrap());
    assert!(matches!(
        docx_parse::parse_docx_s9_wire(&bytes, Default::default()),
        Err(docx_parse::xml::ParseError::ResourceLimit {
            kind: "bookmarkNamespaceAliases",
            ..
        })
    ));
    let font = docx_layout::register_measure_font(fixture::FONT).unwrap();
    let request = fixture::region_request(&engine, &connector_package(), font);
    draw_connector(&engine, &request.to_string());
    assert_authored_connector_path(&engine, "bounded preview");
}
