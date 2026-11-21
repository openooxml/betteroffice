use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use docx_parse::xml::ParseError;
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas""#;

fn fixture(body: &str) -> Vec<u8> {
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned().into_bytes()),
        ("_rels/.rels".to_owned(), r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned().into_bytes()),
        ("word/document.xml".to_owned(), format!("<w:document {NS}><w:body><w:p><w:r><w:t>Before</w:t></w:r></w:p>{body}<w:sectPr/></w:body></w:document>").into_bytes()),
    ];
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn package(source: &[u8]) -> Value {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    let mut package: Value = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    package
}

fn find_mut<'a>(value: &'a mut Value, matches: &dyn Fn(&Value) -> bool) -> Option<&'a mut Value> {
    if matches(value) {
        return Some(value);
    }
    match value {
        Value::Object(map) => map.values_mut().find_map(|child| find_mut(child, matches)),
        Value::Array(items) => items.iter_mut().find_map(|child| find_mut(child, matches)),
        _ => None,
    }
}

fn save(source: &[u8], package: &Value) -> Result<String, ParseError> {
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    let saved = write_docx_s13(request, source)?;
    let (_, xml) = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap();
    Ok(String::from_utf8(xml).unwrap())
}

fn drawing_fixture(drawing: &str) -> Vec<u8> {
    fixture(&format!(
        "<w:p><w:r><w:t>Host</w:t></w:r><w:r>{drawing}</w:r></w:p>"
    ))
}

fn drawing(kind: &str, content: &str) -> String {
    format!(
        "<w:drawing><wp:inline><wp:extent cx=\"2743200\" cy=\"914400\"/>\
         <wp:docPr id=\"30\" name=\"Composite\"/><a:graphic>\
         <a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessing{kind}\">\
         {content}</a:graphicData></a:graphic></wp:inline></w:drawing>"
    )
}

fn group(tag: &str, children: &str) -> String {
    format!(
        "<{tag}><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/>\
         <a:ext cx=\"2743200\" cy=\"914400\"/><a:chOff x=\"0\" y=\"0\"/>\
         <a:chExt cx=\"2743200\" cy=\"914400\"/></a:xfrm></wpg:grpSpPr>{children}</{tag}>"
    )
}

fn shape(id: u32, name: &str, x: u32, geometry: &str, text_box: &str) -> String {
    format!(
        "<wps:wsp><wps:cNvPr id=\"{id}\" name=\"{name}\"/>\
         <wps:spPr><a:xfrm><a:off x=\"{x}\" y=\"0\"/><a:ext cx=\"914400\" cy=\"914400\"/>\
         </a:xfrm><a:prstGeom prst=\"{geometry}\"/></wps:spPr>{text_box}<wps:bodyPr/></wps:wsp>"
    )
}

fn group_shapes(bookmarked: bool, second: Option<&str>) -> String {
    let bookmark = if bookmarked {
        r#"<w:bookmarkStart w:id="0" w:name="_GoBack"/><w:bookmarkEnd w:id="0"/>"#
    } else {
        ""
    };
    let first = format!(
        "<wps:txbx><w:txbxContent><w:p><w:r><w:t>Inner</w:t></w:r>{bookmark}</w:p>\
         </w:txbxContent></wps:txbx>"
    );
    let mut shapes = shape(31, "Box1", 0, "rect", &first);
    if let Some(second) = second {
        let text_box = format!("<wps:txbx>{second}</wps:txbx>");
        shapes.push_str(&shape(33, "Box2", 914_400, "rect", &text_box));
    }
    let mark_x = if second.is_some() { 1_828_800 } else { 914_400 };
    shapes.push_str(&shape(32, "Mark", mark_x, "ellipse", ""));
    shapes
}

fn word_shaped_group(drawing: &str) -> String {
    let fallback = r#"<w:pict><v:group id="Composite" style="width:144pt;height:72pt" coordsize="1828800,914400"><v:rect style="width:72pt;height:72pt"><v:textbox><w:txbxContent><w:p><w:r><w:t>Inner</w:t></w:r></w:p></w:txbxContent></v:textbox></v:rect><v:oval style="left:72pt;width:72pt;height:72pt"/></v:group></w:pict>"#;
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wpg">{drawing}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
    )
}

fn assert_group<'a>(xml: &'a str, expected: &[(u32, &str)]) -> Vec<&'a str> {
    assert!(xml.contains("<wpg:wgp>"), "{xml}");
    let shapes: Vec<_> = xml
        .split("<wps:wsp>")
        .skip(1)
        .map(|shape| shape.split_once("</wps:wsp>").unwrap().0)
        .collect();
    assert_eq!(shapes.len(), expected.len(), "{xml}");
    for (shape, (id, name)) in shapes.iter().zip(expected) {
        assert!(
            shape.contains(&format!(r#"<wps:cNvPr id="{id}" name="{name}""#)),
            "{xml}"
        );
    }
    shapes
}

fn assert_bookmark(xml: &str) {
    assert!(xml.contains("_GoBack"), "{xml}");
    assert!(xml.contains("<w:bookmarkStart "), "{xml}");
    assert!(xml.contains("<w:bookmarkEnd "), "{xml}");
}

fn group_model(model: &mut Value) -> &mut Value {
    let host = &mut model["document"]["content"][1]["content"];
    find_mut(host, &|value| value["id"] == "30" && value["children"].is_array())
        .expect("group shape in the model")
}

fn edit_text(model: &mut Value, before: &str) {
    let children = &mut group_model(model)["children"];
    let text = find_mut(children, &|value| value["text"] == before)
        .unwrap_or_else(|| panic!("expected {before:?} in the group model"));
    text["text"] = json!("Changed");
}

fn check_group(word_shaped: bool) {
    let expected = [(31, "Box1"), (32, "Mark")];
    for bookmarked in [true, false] {
        let mut drawing = drawing("Group", &group("wpg:wgp", &group_shapes(bookmarked, None)));
        if word_shaped {
            drawing = word_shaped_group(&drawing);
        }
        let source = drawing_fixture(&drawing);
        let mut model = package(&source);
        if bookmarked {
            let xml = save(&source, &model).unwrap();
            let shapes = assert_group(&xml, &expected);
            assert!(shapes[0].contains(">Inner<"), "{xml}");
            assert_bookmark(shapes[0]);
            if word_shaped {
                assert!(xml.contains("<mc:Fallback>"), "{xml}");
            }
        }

        edit_text(&mut model, "Inner");
        let xml = save(&source, &model).unwrap();
        let shapes = assert_group(&xml, &expected);
        assert!(shapes[0].contains(">Changed<"), "{xml}");
        if bookmarked {
            assert!(!shapes[0].contains(">Inner<"), "{xml}");
            assert_bookmark(shapes[0]);
            if word_shaped {
                assert!(xml.contains("<mc:Fallback>"), "{xml}");
            }
        }
    }
}

#[test]
fn a_bookmarked_text_box_in_a_group_saves() {
    check_group(false);
}

#[test]
fn a_bookmarked_text_box_in_a_word_shaped_group_saves() {
    check_group(true);
}

#[test]
fn editing_either_group_text_box_preserves_the_other() {
    let second = "<w:txbxContent><w:p><w:r><w:t>Second</w:t></w:r></w:p></w:txbxContent>";
    let group = group("wpg:wgp", &group_shapes(true, Some(second)));
    let source = drawing_fixture(&drawing("Group", &group));
    for (before, first, second) in [
        ("Inner", ">Changed<", ">Second<"),
        ("Second", ">Inner<", ">Changed<"),
    ] {
        let mut model = package(&source);
        edit_text(&mut model, before);
        let xml = save(&source, &model).unwrap();
        let shapes = assert_group(&xml, &[(31, "Box1"), (33, "Box2"), (32, "Mark")]);
        assert!(shapes[0].contains(first), "{xml}");
        assert!(shapes[1].contains(second), "{xml}");
        assert_bookmark(shapes[0]);
    }
}

#[test]
fn a_bookmarked_group_keeps_an_empty_text_box_when_the_other_is_edited() {
    let group = group("wpg:wgp", &group_shapes(true, Some("<w:txbxContent/>")));
    let source = drawing_fixture(&drawing("Group", &group));
    let mut model = package(&source);
    edit_text(&mut model, "Inner");
    let xml = save(&source, &model).unwrap();
    let shapes = assert_group(&xml, &[(31, "Box1"), (33, "Box2"), (32, "Mark")]);
    assert!(shapes[0].contains(">Changed<"), "{xml}");
    assert!(shapes[1].contains("<w:txbxContent/>"), "{xml}");
    assert_bookmark(shapes[0]);
}

#[test]
fn a_group_shape_fill_edit_preserves_the_bookmarked_group() {
    let group = group("wpg:wgp", &group_shapes(true, None));
    let source = drawing_fixture(&drawing("Group", &group));
    let mut model = package(&source);
    let mark = &mut group_model(&mut model)["children"][1];
    assert_eq!(mark["name"], "Mark");
    mark["fill"] = json!({"type": "solid", "color": {"rgb": "FF0000"}});
    let xml = save(&source, &model).unwrap();
    let shapes = assert_group(&xml, &[(31, "Box1"), (32, "Mark")]);
    assert!(shapes[0].contains(">Inner<"), "{xml}");
    assert_bookmark(shapes[0]);
}

#[test]
fn a_bookmarked_group_in_a_group_or_canvas_saves() {
    let shapes = group_shapes(true, None);
    let nested = group("wpg:grpSp", &shapes);
    let canvas_group = group("wpg:wgp", &shapes);
    for (drawing, container) in [
        (drawing("Group", &group("wpg:wgp", &nested)), "wpg:grpSp"),
        (
            drawing("Canvas", &format!("<wpc:wpc>{canvas_group}</wpc:wpc>")),
            "wpc:wpc",
        ),
    ] {
        let source = drawing_fixture(&drawing);
        let mut model = package(&source);
        edit_text(&mut model, "Inner");
        let xml = save(&source, &model).unwrap();
        let shapes = assert_group(&xml, &[(31, "Box1"), (32, "Mark")]);
        let (_, rest) = xml.split_once(&format!("<{container}>")).unwrap();
        let (inner, _) = rest.split_once(&format!("</{container}>")).unwrap();
        assert!(shapes.iter().all(|&shape| inner.contains(shape)), "{xml}");
        assert!(shapes[0].contains(">Changed<"), "{xml}");
        assert_bookmark(shapes[0]);
    }
}

#[test]
fn a_bookmarked_group_fails_to_save_text_for_a_shape_without_a_source_text_box() {
    let group = group("wpg:wgp", &group_shapes(true, None));
    let plain = drawing("Group", &group);
    let word_shaped = word_shaped_group(&plain);
    for drawing in [plain, word_shaped] {
        let source = drawing_fixture(&drawing);
        let mut model = package(&source);
        let mark = &mut group_model(&mut model)["children"][1];
        mark["textBody"]["content"] = json!([{"type": "paragraph", "content": []}]);
        assert_eq!(
            save(&source, &model).unwrap_err(),
            ParseError::Canonical("bookmarked group source could not be replayed".to_owned())
        );
    }
}
