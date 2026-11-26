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
fn state_only_connector_replica_draws_the_released_path_and_saves_identical_parts() {
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

    replica.doc().retain_source_docx(bytes.clone());
    let package = docx_parse::parse_docx_s9_wire(
        &bytes,
        docx_parse::S9ParseOptions {
            source_ordinals: true,
            ..Default::default()
        },
    )
    .unwrap()
    .document
    .package;
    let save = |engine: &EngineSession| {
        let plan = engine.doc().paragraph_save_plan();
        let request = serde_json::from_value(json!({
            "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
            "document": {"content": package.document.content, "comments": package.document.comments},
            "relationshipEntries": package.relationship_entries,
            "headerEntries": package.header_entries.as_deref().unwrap_or_default(),
            "options": {"updateModifiedDate": false},
            "paragraphIds": {
                "assignments": plan.assignments.iter().map(|(part, ordinal, para_id)| {
                    json!({"part": part, "ordinal": ordinal, "paraId": para_id})
                }).collect::<Vec<_>>(),
                "patchedParts": plan.patched_parts.iter().map(|(part, para_ids)| {
                    json!({"part": part, "paraIds": para_ids})
                }).collect::<Vec<_>>()
            }
        }))
        .unwrap();
        let saved = docx_parse::serializer::write_docx_s13(request, &bytes).unwrap();
        ooxml_opc::unzip_parts(&saved)
            .unwrap()
            .into_iter()
            .collect::<std::collections::BTreeMap<_, _>>()
    };
    assert_eq!(save(&source), save(&replica));
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
