use docx_parse::serializer::{
    S13SaveRequest, SerializerContext, serialize_comments_part, try_serialize_comments_part,
    write_docx_s13,
};
use docx_parse::{S9ParseOptions, parse_docx_s9_wire};
use serde_json::{Value, json};

const START: &str =
    r#"<w:bookmarkStart w:name='Block&amp;Range' w:id = '&#55;' w:colFirst="0" w:colLast="2" />"#;
const END: &str = r#"<w:bookmarkEnd w:displacedByCustomXml='prev' w:id="7"></w:bookmarkEnd>"#;
const CONTAINERS: [&str; 8] = [
    "cell", "control", "header", "footer", "footnote", "endnote", "comment", "textbox",
];
const NS: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape""#;

fn p(id: &str, text: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}"><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

fn fixture(container: &str, blocks: &str) -> Vec<u8> {
    let mut body = p("00000010", "Unrelated");
    let mut section = String::new();
    let mut parts = Vec::new();
    let mut relationship = String::new();
    let mut content_type = String::new();
    match container {
        "cell" => body.push_str(&format!(
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{blocks}</w:tc></w:tr></w:tbl>"#
        )),
        "control" => body.push_str(&format!(
            r#"<w:sdt><w:sdtPr><w:id w:val="42"/></w:sdtPr><w:sdtContent>{blocks}</w:sdtContent></w:sdt>"#
        )),
        "textbox" => body.push_str(&format!(
            r#"<w:p w14:paraId="00000020"><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/><wp:docPr id="20" name="Text Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvPr id="20"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent>{blocks}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#
        )),
        _ => {
            let (name, kind, xml) = match container {
                "header" | "footer" => {
                    let root = if container == "header" { "hdr" } else { "ftr" };
                    section = format!(
                        r#"<w:{container}Reference w:type="default" r:id="story"/>"#
                    );
                    (
                        format!("{container}1"),
                        container.to_owned(),
                        format!("<w:{root} {NS}>{blocks}</w:{root}>"),
                    )
                }
                "footnote" | "endnote" => {
                    body.push_str(&format!(
                        r#"<w:p><w:r><w:{container}Reference w:id="1"/></w:r></w:p>"#
                    ));
                    (
                        format!("{container}s"),
                        format!("{container}s"),
                        format!(
                            r#"<w:{container}s {NS}><w:{container} w:id="1">{blocks}</w:{container}></w:{container}s>"#
                        ),
                    )
                }
                "comment" => (
                    "comments".to_owned(),
                    "comments".to_owned(),
                    format!(
                        r#"<w:comments {NS}><w:comment w:id="1" w:author="Reviewer">{blocks}</w:comment></w:comments>"#
                    ),
                ),
                _ => unreachable!(),
            };
            relationship = format!(
                r#"<Relationship Id="story" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{name}.xml"/>"#
            );
            content_type = format!(
                r#"<Override PartName="/word/{name}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/>"#
            );
            parts.push((format!("word/{name}.xml"), xml.into_bytes()));
        }
    }
    parts.extend([
        ("[Content_Types].xml".to_owned(), format!(r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>{content_type}</Types>"#).into_bytes()),
        ("word/_rels/document.xml.rels".to_owned(), format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relationship}</Relationships>"#).into_bytes()),
        ("word/document.xml".to_owned(), format!("<w:document {NS}><w:body>{body}<w:sectPr>{section}</w:sectPr></w:body></w:document>").into_bytes()),
    ]);
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn package(source: &[u8]) -> Value {
    let parsed = parse_docx_s9_wire(source, S9ParseOptions::default()).unwrap();
    let mut package = serde_json::to_value(parsed.document.package).unwrap();
    package["document"]
        .as_object_mut()
        .unwrap()
        .remove("sections");
    package
}

fn blocks<'a>(package: &'a mut Value, container: &str) -> &'a mut Vec<Value> {
    let path = match container {
        "cell" => "/document/content/1/rows/0/cells/0/content",
        "control" => "/document/content/1/content",
        "header" => "/headerEntries/0/1/content",
        "footer" => "/footerEntries/0/1/content",
        "footnote" => "/footnotes/0/content",
        "endnote" => "/endnotes/0/content",
        "comment" => "/document/comments/0/blockContent",
        "textbox" => "/document/content/1/content/0/content/0/shape/textBody/content",
        _ => unreachable!(),
    };
    package.pointer_mut(path).unwrap().as_array_mut().unwrap()
}

fn save_request(package: &Value) -> S13SaveRequest {
    let mut request = json!({
        "determinism": {"seed": "0".repeat(64), "now": "2000-01-01T00:00:00.000Z"},
        "document": package["document"],
        "options": {"updateModifiedDate": false}
    });
    for key in [
        "headerEntries",
        "footerEntries",
        "footnotes",
        "endnotes",
        "relationshipEntries",
    ] {
        if let Some(value) = package.get(key) {
            request[key] = value.clone();
        }
    }
    serde_json::from_value(request).unwrap()
}

fn save(source: &[u8], package: &Value) -> Vec<u8> {
    write_docx_s13(save_request(package), source).unwrap()
}

fn rewrite_text_box(package: &mut Value) {
    let shape = package
        .pointer_mut("/document/content/1/content/0/content/0/shape")
        .unwrap()
        .as_object_mut()
        .unwrap();
    for key in ["drawingXml", "drawingRunXml", "drawingModel"] {
        shape.remove(key);
    }
}

fn part(source: &[u8], container: &str) -> String {
    let path = match container {
        "header" => "word/header1.xml",
        "footer" => "word/footer1.xml",
        "footnote" => "word/footnotes.xml",
        "endnote" => "word/endnotes.xml",
        "comment" => "word/comments.xml",
        _ => "word/document.xml",
    };
    String::from_utf8(
        ooxml_opc::unzip_parts(source)
            .unwrap()
            .into_iter()
            .find(|(name, _)| name == path)
            .unwrap()
            .1,
    )
    .unwrap()
}

fn assert_pair(xml: &str) {
    assert_eq!(xml.matches("<w:bookmarkStart").count(), 1, "{xml}");
    assert_eq!(xml.matches("<w:bookmarkEnd").count(), 1, "{xml}");
    let mut offset = 0;
    for fragment in [">Before<", START, ">Anchor<", END, ">After<"] {
        offset += xml[offset..].find(fragment).unwrap() + fragment.len();
    }
}

#[test]
fn block_bookmarks_keep_their_source_xml_on_every_rewrite() {
    let content = format!(
        "{}{START}{}{END}{}",
        p("00000001", "Before"),
        p("00000002", "Anchor"),
        p("00000003", "After")
    );
    for container in CONTAINERS {
        let source = fixture(container, &content);
        let mut package = package(&source);
        if container == "footnote" || container == "endnote" {
            let verbatim = package[format!("{container}s")][0]["verbatimXml"]
                .as_str()
                .unwrap();
            assert!(verbatim.contains(START));
            assert!(verbatim.contains(END));
        }
        assert_eq!(
            blocks(&mut package, container)
                .iter()
                .map(|block| block["type"].as_str().unwrap())
                .collect::<Vec<_>>(),
            ["paragraph", "rawXml", "paragraph", "rawXml", "paragraph"],
            "{container}"
        );
        assert_pair(&part(&save(&source, &package), container));
        package["document"]["content"][0]["content"][0]["content"][0]["text"] = json!("Edited");
        assert_pair(&part(&save(&source, &package), container));
        blocks(&mut package, container).push(json!({"type": "paragraph", "content": []}));
        if container == "footnote" || container == "endnote" {
            package[format!("{container}s")][0]
                .as_object_mut()
                .unwrap()
                .remove("verbatimXml");
        }
        if container == "textbox" {
            rewrite_text_box(&mut package);
        }
        if container == "comment" {
            let paragraphs: Vec<_> = blocks(&mut package, container)
                .iter()
                .filter(|block| block["type"] == "paragraph")
                .cloned()
                .collect();
            package["document"]["comments"][0]["content"] = json!(paragraphs);
        }
        let xml = part(&save(&source, &package), container);
        assert_pair(&xml);
        assert!(part(&save(&source, &package), "cell").contains("Edited"));
    }
}

#[test]
fn block_bookmarks_moved_into_a_paragraph_are_written_once() {
    for container in CONTAINERS {
        let source = fixture(
            container,
            &format!("{START}{}{END}", p("00000002", "Anchor")),
        );
        let mut package = package(&source);
        let content = blocks(&mut package, container);
        content.retain(|block| block["type"] != "rawXml");
        content[0]["content"] = json!([
            {"type": "bookmarkStart", "id": 7, "name": "Block&Range"},
            {"type": "run", "content": [{"type": "text", "text": "Moved"}]},
            {"type": "bookmarkEnd", "id": 7}
        ]);
        if container == "comment" {
            let paragraphs = blocks(&mut package, container).clone();
            package["document"]["comments"][0]["content"] = json!(paragraphs);
        }
        if container == "textbox" {
            rewrite_text_box(&mut package);
        }
        if container == "footnote" || container == "endnote" {
            package[format!("{container}s")][0]
                .as_object_mut()
                .unwrap()
                .remove("verbatimXml");
        }
        let xml = part(&save(&source, &package), container);
        assert_eq!(
            xml.matches("<w:bookmarkStart").count(),
            1,
            "{container}: {xml}"
        );
        assert_eq!(
            xml.matches("<w:bookmarkEnd").count(),
            1,
            "{container}: {xml}"
        );
        assert!(xml.contains("Moved"), "{container}: {xml}");
    }
}

#[test]
fn untouched_bookmarked_notes_keep_revision_ids_and_field_runs() {
    for container in ["footnote", "endnote"] {
        let initial = fixture(container, &p("00000002", "Anchor"));
        let baseline = save(&initial, &package(&initial));
        let note = format!(
            r#"<w:{container} w:id="1"><w:bookmarkStart w:id="7" w:name="Note"/><w:p w14:paraId="00000002" w:rsidR="00112233" w:rsidRDefault="44556677" w:rsidP="8899AABB"><w:r w:rsidR="00112233"><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> HYPERLINK "https://example.test/note" </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>Anchor</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:bookmarkEnd w:id="7"/></w:{container}>"#
        );
        let mut parts = ooxml_opc::unzip_parts(&baseline).unwrap();
        let (_, bytes) = parts
            .iter_mut()
            .find(|(name, _)| name == &format!("word/{container}s.xml"))
            .unwrap();
        let mut xml = String::from_utf8(bytes.clone()).unwrap();
        let start = xml.find(&format!("<w:{container} w:id=\"1\">")).unwrap();
        let end =
            xml.find(&format!("</w:{container}>")).unwrap() + format!("</w:{container}>").len();
        xml.replace_range(start..end, &note);
        *bytes = xml.into_bytes();
        let source = ooxml_opc::rezip_parts(&parts).unwrap();
        let original = part(&source, container);
        let mut model = package(&source);
        assert_eq!(
            model[format!("{container}s")][0]["verbatimXml"].as_str(),
            Some(note.as_str())
        );
        for edited in [false, true] {
            if edited {
                model["document"]["content"][0]["content"][0]["content"][0]["text"] =
                    json!("Edited");
            }
            assert_eq!(part(&save(&source, &model), container), original);
        }
    }
}

#[test]
fn bookmark_only_control_keeps_its_container() {
    let source = fixture(
        "control",
        &format!("{START}{}{END}", p("00000002", "Anchor")),
    );
    let mut package = package(&source);
    blocks(&mut package, "control").retain(|block| block["type"] == "rawXml");
    let request = save_request(&package);
    assert!(matches!(
        &request.document.content[1],
        docx_parse::BlockContent::BlockSdt(_)
    ));
    let xml = part(&write_docx_s13(request, &source).unwrap(), "control");
    assert!(
        xml.contains(&format!("<w:sdtContent>{START}{END}</w:sdtContent>")),
        "{xml}"
    );
}

fn drawing_fixture(grouped: bool) -> Vec<u8> {
    let source = fixture(
        "textbox",
        &format!("{START}{}{END}", p("00000002", "Anchor")),
    );
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap();
    let mut xml = String::from_utf8(bytes.clone()).unwrap();
    let start = xml.find("<w:drawing>").unwrap();
    let end = xml.find("</w:drawing>").unwrap() + "</w:drawing>".len();
    let mut drawing = xml[start..end].to_owned();
    if grouped {
        let start = drawing.find("<wps:wsp>").unwrap();
        let end = drawing.find("</wps:wsp>").unwrap() + "</wps:wsp>".len();
        let shape = &drawing[start..end];
        let group = format!(
            r#"<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></wpg:grpSpPr>{shape}</wpg:wgp>"#
        );
        drawing.replace_range(start..end, &group);
        drawing = drawing.replace(
            "uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingShape\"",
            "uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\"",
        );
    }
    let start8 = START.replace("'&#55;'", "'8'");
    let end8 = END.replace("\"7\"", "\"8\"");
    let fallback = format!(
        r#"<w:pict><v:group><v:rect style="width:72pt;height:72pt"><v:textbox><w:txbxContent>{start8}{}{end8}</w:txbxContent></v:textbox></v:rect></v:group></w:pict>"#,
        p("00000008", "Fallback")
    );
    let requires = if grouped { "wpg" } else { "wps" };
    xml.replace_range(start..end, &format!(r#"<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><mc:Choice Requires="{requires}">{drawing}</mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#));
    *bytes = xml.into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn drawing_text_boxes_keep_choice_and_fallback_bookmarks_on_rewrite() {
    for grouped in [false, true] {
        let source = drawing_fixture(grouped);
        let mut package = package(&source);
        for edit in 0..3 {
            if edit == 1 {
                package["document"]["content"][0]["content"][0]["content"][0]["text"] =
                    json!("Edited");
                package["document"]["content"][1]["content"]
                    .as_array_mut()
                    .unwrap()
                    .push(json!({"type":"run","content":[{"type":"text","text":"Host edit"}]}));
            } else if edit == 2 {
                let mut shape =
                    &mut package["document"]["content"][1]["content"][0]["content"][0]["shape"];
                shape["size"]["width"] = json!(1828800);
                if grouped {
                    shape = &mut shape["children"][0];
                }
                shape["textBody"]["content"][1]["content"][0]["content"][0]["text"] =
                    json!("Edited box");
            }
            let xml = part(&save(&source, &package), "textbox");
            for marker in [
                START.to_owned(),
                END.to_owned(),
                START.replace("'&#55;'", "'8'"),
                END.replace("\"7\"", "\"8\""),
            ] {
                assert_eq!(xml.matches(&marker).count(), 1, "{grouped} {edit}: {xml}");
            }
            assert!(xml.contains("<mc:Fallback>"), "{xml}");
            if edit == 2 {
                assert!(xml.contains("Edited box"), "{xml}");
            }
        }
    }
}

#[test]
fn nested_block_controls_keep_bookmarks_inside_their_own_content() {
    let nested = format!(
        r#"<w:sdt><w:sdtPr><w:tag w:val="Nested"/></w:sdtPr><w:sdtContent>{START}{}{END}</w:sdtContent></w:sdt>"#,
        p("00000002", "Anchor")
    );
    for container in ["cell", "control"] {
        let source = fixture(container, &nested);
        let mut model = package(&source);
        let nested = &blocks(&mut model, container)[0];
        assert_eq!(nested["type"], "blockSdt");
        assert_eq!(nested["content"][0]["xml"], START);
        assert_eq!(nested["content"][2]["xml"], END);
        for edited in [false, true] {
            if edited {
                model["document"]["content"][0]["content"][0]["content"][0]["text"] =
                    json!("Edited");
            }
            let saved = save(&source, &model);
            let xml = part(&saved, container);
            assert_eq!(xml.matches(START).count(), 1, "{xml}");
            assert_eq!(xml.matches(END).count(), 1, "{xml}");
            let mut reopened = package(&saved);
            let nested = &blocks(&mut reopened, container)[0];
            assert_eq!(nested["content"][0]["xml"], START);
            assert_eq!(nested["content"][2]["xml"], END);
        }
    }
}

#[test]
fn a_cell_with_only_bookmark_boundaries_keeps_a_paragraph() {
    let source = fixture("cell", &format!("{START}{END}"));
    let mut package = package(&source);
    assert_eq!(blocks(&mut package, "cell").len(), 3);
    blocks(&mut package, "cell").retain(|block| block["type"] == "rawXml");
    let xml = part(&save(&source, &package), "cell");
    let cell = xml
        .split_once("<w:tc>")
        .unwrap()
        .1
        .split_once("</w:tc>")
        .unwrap()
        .0;
    assert!(
        cell.find(END).unwrap() < cell.rfind("<w:p").unwrap(),
        "{xml}"
    );
}

#[test]
fn a_cell_with_other_raw_content_does_not_gain_a_paragraph() {
    let raw = r#"<x:opaque xmlns:x="urn:cell"/>"#;
    let source = fixture("cell", raw);
    let mut package = package(&source);
    assert_eq!(blocks(&mut package, "cell").len(), 1);
    for raw in [raw, r#"<w:customXml w:element="cell"/>"#] {
        blocks(&mut package, "cell")[0]["xml"] = json!(raw);
        let xml = part(&save(&source, &package), "cell");
        let cell = xml
            .split_once("<w:tc>")
            .unwrap()
            .1
            .split_once("</w:tc>")
            .unwrap()
            .0;
        assert_eq!(cell, raw);
    }
}

#[test]
fn note_bookmark_attributes_keep_their_namespace_binding() {
    for container in ["footnote", "endnote"] {
        let source = fixture(
            container,
            &format!("{START}{}{END}", p("00000002", "Anchor")),
        );
        let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
        let (_, bytes) = parts
            .iter_mut()
            .find(|(name, _)| name == &format!("word/{container}s.xml"))
            .unwrap();
        let xml = String::from_utf8(bytes.clone())
            .unwrap()
            .replace(
                &format!("<w:{container} w:id=\"1\">"),
                &format!("<w:{container} xmlns:x=\"urn:bookmark\" w:id=\"1\">"),
            )
            .replace(START, &START.replace("w:name=", "x:keep=\"yes\" w:name="));
        *bytes = xml.into_bytes();
        let source = ooxml_opc::rezip_parts(&parts).unwrap();
        let model = package(&source);
        let xml = part(&save(&source, &model), container);
        assert!(xml.contains("xmlns:x=\"urn:bookmark\""), "{xml}");
        assert!(
            xml.contains(&START.replace("w:name=", "x:keep=\"yes\" w:name=")),
            "{xml}"
        );
        assert_eq!(
            attribute_namespace(&xml, "x:keep"),
            ["urn:bookmark"],
            "{xml}"
        );
    }
}

#[test]
fn invalid_comment_bookmark_fragments_fail_serialization_and_save() {
    let source = fixture(
        "comment",
        &format!("{START}{}{END}", p("00000002", "Anchor")),
    );
    for xml in [
        r#"<w:bookmarkStart w:id="7" w:name="Broken">"#,
        r#"<w:bookmarkStart w:id="7"/><w:bookmarkEnd w:id="7"/>"#,
        r#"<w:bookmarkStart w:name="Broken">"#,
    ] {
        let mut package = package(&source);
        blocks(&mut package, "comment")[0]["xml"] = json!(xml);
        let request = save_request(&package);
        let mut context = SerializerContext::new(&request.determinism).unwrap();
        let comments = request.document.comments.as_ref().unwrap();
        assert!(
            try_serialize_comments_part(comments, &mut context.clone()).is_err(),
            "{xml}"
        );
        assert!(
            !serialize_comments_part(comments, &mut context).contains("w:bookmark"),
            "{xml}"
        );
        assert!(write_docx_s13(request, &source).is_err(), "{xml}");
    }
}

#[test]
fn a_comment_bookmark_without_an_id_is_not_dropped() {
    let source = fixture(
        "comment",
        &format!("{START}{}{END}", p("00000002", "Anchor")),
    );
    let mut package = package(&source);
    let start = r#"<w:bookmarkStart w:name="Anonymous"/>"#;
    blocks(&mut package, "comment")[0]["xml"] = json!(start);
    let xml = part(&save(&source, &package), "comment");
    assert_eq!(xml.matches("<w:bookmarkStart").count(), 1, "{xml}");
    assert_eq!(xml.matches("<w:bookmarkEnd").count(), 1, "{xml}");
    assert!(xml.contains(start), "{xml}");
}

#[test]
fn duplicate_comment_bookmark_boundaries_fail_save() {
    let source = fixture(
        "comment",
        &format!("{START}{}{END}", p("00000002", "Anchor")),
    );
    let mut package = package(&source);
    let content = blocks(&mut package, "comment");
    content.insert(0, content[0].clone());
    assert!(write_docx_s13(save_request(&package), &source).is_err());
}

#[test]
fn a_legacy_comment_body_rewrite_restores_only_missing_boundaries() {
    let source = fixture(
        "comment",
        &format!(
            "{}{START}{}{END}{}",
            p("00000001", "Before"),
            p("00000002", "Anchor"),
            p("00000003", "After")
        ),
    );
    for scenario in ["delete", "merge", "move", "start", "end", "both"] {
        let mut package = package(&source);
        let comment = &mut package["document"]["comments"][0];
        comment.as_object_mut().unwrap().remove("blockContent");
        let paragraphs = comment["content"].as_array_mut().unwrap();
        match scenario {
            "delete" => {
                paragraphs.remove(1);
            }
            "merge" => {
                let anchor = paragraphs.remove(1);
                paragraphs[0]["content"]
                    .as_array_mut()
                    .unwrap()
                    .extend(anchor["content"].as_array().unwrap().iter().cloned());
            }
            "move" => {
                let anchor = paragraphs.remove(1);
                paragraphs.push(anchor);
            }
            _ => {
                let content = paragraphs[1]["content"].as_array_mut().unwrap();
                if scenario != "end" {
                    content.insert(
                        0,
                        json!({"type": "bookmarkStart", "id": 7, "name": "Block&Range"}),
                    );
                }
                if scenario != "start" {
                    content.push(json!({"type": "bookmarkEnd", "id": 7}));
                }
            }
        }
        let xml = part(&save(&source, &package), "comment");
        assert_eq!(
            xml.matches("<w:bookmarkStart").count(),
            1,
            "{scenario}: {xml}"
        );
        assert_eq!(
            xml.matches("<w:bookmarkEnd").count(),
            1,
            "{scenario}: {xml}"
        );
        assert!(xml.find("<w:bookmarkStart").unwrap() < xml.find("<w:bookmarkEnd").unwrap());
        if scenario == "move" {
            assert!(xml.find(">After<").unwrap() < xml.find(START).unwrap());
            assert!(xml.find(END).unwrap() < xml.find(">Anchor<").unwrap());
        }
    }
}

fn assert_bound_prefixes(xml: &str) {
    use quick_xml::NsReader;
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;

    let mut reader = NsReader::from_reader(xml.as_bytes());
    loop {
        match reader.read_resolved_event().unwrap() {
            (namespace, Event::Start(element) | Event::Empty(element)) => {
                assert!(!matches!(namespace, ResolveResult::Unknown(_)), "{xml}");
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
                    assert!(!matches!(namespace, ResolveResult::Unknown(_)), "{xml}");
                    if attribute.key.as_ref() == b"b:name" {
                        assert_eq!(
                            namespace,
                            ResolveResult::Bound(quick_xml::name::Namespace(
                                b"http://schemas.openxmlformats.org/wordprocessingml/2006/main"
                            ))
                        );
                    }
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }
}

#[test]
fn block_bookmarks_retain_inherited_namespace_aliases_in_every_container() {
    for container in CONTAINERS {
        let source = fixture(
            container,
            r#"<w:bookmarkStart w:id="7" b:name="B"/><w:p/><w:bookmarkEnd w:id="7"/>"#,
        );
        let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
        let tag = match container {
            "cell" => "<w:tc>",
            "control" => "<w:sdtContent>",
            "header" => "<w:hdr ",
            "footer" => "<w:ftr ",
            "footnote" => "<w:footnote ",
            "endnote" => "<w:endnote ",
            "comment" => "<w:comment ",
            "textbox" => "<w:txbxContent>",
            _ => unreachable!(),
        };
        for (_, bytes) in &mut parts {
            if let Ok(xml) = std::str::from_utf8(bytes)
                && xml.contains(tag)
            {
                let replacement = format!(
                    r#"{} xmlns:b="http://schemas.openxmlformats.org/wordprocessingml/2006/main"{}"#,
                    &tag[..tag.len() - 1],
                    &tag[tag.len() - 1..]
                );
                *bytes = xml.replace(tag, &replacement).into_bytes();
            }
        }
        let source = ooxml_opc::rezip_parts(&parts).unwrap();
        let mut model = package(&source);
        for rewritten in [false, true] {
            if rewritten {
                rewrite_container(&mut model, container);
            }
            let xml = part(&save(&source, &model), container);
            assert_eq!(
                xml.matches("<w:bookmarkStart").count(),
                1,
                "{container}: {xml}"
            );
            assert_eq!(
                xml.matches("<w:bookmarkEnd").count(),
                1,
                "{container}: {xml}"
            );
            assert!(xml.contains(r#"b:name="B""#), "{container}: {xml}");
            assert_bound_prefixes(&xml);
        }
    }
}

fn rewrite_container(model: &mut Value, container: &str) {
    blocks(model, container).push(json!({"type":"paragraph","content":[]}));
    if container == "footnote" || container == "endnote" {
        model[format!("{container}s")][0]
            .as_object_mut()
            .unwrap()
            .remove("verbatimXml");
    }
    if container == "textbox" {
        rewrite_text_box(model);
    }
    if container == "comment" {
        let paragraphs: Vec<_> = blocks(model, container)
            .iter()
            .filter(|block| block["type"] == "paragraph")
            .cloned()
            .collect();
        model["document"]["comments"][0]["content"] = json!(paragraphs);
    }
}

fn story_part(container: &str) -> &'static str {
    match container {
        "header" => "word/header1.xml",
        "footer" => "word/footer1.xml",
        "footnote" => "word/footnotes.xml",
        "endnote" => "word/endnotes.xml",
        "comment" => "word/comments.xml",
        _ => "word/document.xml",
    }
}

fn with_root_binding(source: &[u8], container: &str, binding: &str, start: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(source).unwrap();
    for (name, bytes) in &mut parts {
        let mut xml = String::from_utf8(bytes.clone()).unwrap();
        if name == story_part(container) {
            xml = xml.replacen(NS, &format!("{NS} {binding}"), 1);
        }
        *bytes = xml.replace(START, start).into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn attribute_namespace(xml: &str, key: &str) -> Vec<String> {
    use quick_xml::NsReader;
    use quick_xml::events::Event;
    use quick_xml::name::ResolveResult;

    let mut reader = NsReader::from_reader(xml.as_bytes());
    let mut found = Vec::new();
    loop {
        match reader.read_resolved_event().unwrap() {
            (_, Event::Start(element) | Event::Empty(element)) => {
                for attribute in element.attributes() {
                    let attribute = attribute.unwrap();
                    if attribute.key.as_ref() == key.as_bytes() {
                        let (namespace, _) = reader.resolver().resolve_attribute(attribute.key);
                        found.push(match namespace {
                            ResolveResult::Bound(namespace) => {
                                String::from_utf8(namespace.as_ref().to_vec()).unwrap()
                            }
                            _ => "unbound".to_owned(),
                        });
                    }
                }
            }
            (_, Event::Eof) => break,
            _ => {}
        }
    }
    found
}

#[test]
fn bookmark_prefixes_declared_on_the_part_root_stay_bound_in_every_container() {
    let start = START.replace("w:name=", "x:keep=\"yes\" w:name=");
    for container in CONTAINERS {
        let source = fixture(
            container,
            &format!("{START}{}{END}", p("00000002", "Anchor")),
        );
        let source = with_root_binding(&source, container, r#"xmlns:x="urn:bookmark""#, &start);
        let mut model = package(&source);
        for rewritten in [false, true] {
            if rewritten {
                rewrite_container(&mut model, container);
            }
            let xml = part(&save(&source, &model), container);
            assert_eq!(
                attribute_namespace(&xml, "x:keep"),
                ["urn:bookmark"],
                "{container} {rewritten}: {xml}"
            );
            assert_bound_prefixes(&xml);
        }
    }
}

#[test]
fn a_rebound_story_root_prefix_keeps_its_source_namespace_on_a_bookmark() {
    let start = START.replace("w:name=", "v:keep=\"yes\" w:name=");
    for container in ["cell", "header", "footnote", "comment", "textbox"] {
        let source = fixture(
            container,
            &format!("{START}{}{END}", p("00000002", "Anchor")),
        );
        let source = with_root_binding(&source, container, r#"xmlns:v="urn:custom""#, &start);
        let mut model = package(&source);
        for rewritten in [false, true] {
            if rewritten {
                rewrite_container(&mut model, container);
            }
            let xml = part(&save(&source, &model), container);
            assert_eq!(
                attribute_namespace(&xml, "v:keep"),
                ["urn:custom"],
                "{container} {rewritten}: {xml}"
            );
        }
    }
}

#[test]
fn a_note_bookmark_nested_in_a_cell_keeps_a_root_declared_prefix() {
    let start = START.replace("w:name=", "x:keep=\"yes\" w:name=");
    for container in ["footnote", "endnote"] {
        let cell = format!(
            r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc>{START}{}{END}</w:tc></w:tr></w:tbl>"#,
            p("00000002", "Anchor")
        );
        let source = fixture(container, &cell);
        let source = with_root_binding(&source, container, r#"xmlns:x="urn:bookmark""#, &start);
        let model = package(&source);
        let xml = part(&save(&source, &model), container);
        assert_eq!(
            attribute_namespace(&xml, "x:keep"),
            ["urn:bookmark"],
            "{container}: {xml}"
        );
        assert_bound_prefixes(&xml);
    }
}

#[test]
fn untouched_comments_with_nested_block_bookmarks_save_from_source() {
    let source = fixture(
        "comment",
        &format!(
            "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent>{START}{}{END}</w:sdtContent></w:sdt>{}",
            p("00000002", "Nested"),
            p("00000003", "After")
        ),
    );
    let original = ooxml_opc::unzip_parts(&source)
        .unwrap()
        .into_iter()
        .find(|(name, _)| name == "word/comments.xml")
        .unwrap()
        .1;
    let mut model = package(&source);
    for edited in [false, true] {
        if edited {
            model["document"]["content"][0]["content"][0]["content"][0]["text"] = json!("Edited");
        }
        let saved = save(&source, &model);
        assert_eq!(part(&saved, "comment").into_bytes(), original);
        if edited {
            assert!(part(&saved, "cell").contains("Edited"));
        }
    }
}

#[test]
fn a_legacy_comment_rewrite_rejects_bookmarks_only_in_nested_blocks() {
    let source = fixture(
        "comment",
        &format!(
            "<w:sdt><w:sdtPr><w:id w:val=\"5\"/></w:sdtPr><w:sdtContent>{START}{}{END}</w:sdtContent></w:sdt>{}",
            p("00000002", "Nested"),
            p("00000003", "After")
        ),
    );
    let mut model = package(&source);
    let comment = &mut model["document"]["comments"][0];
    comment.as_object_mut().unwrap().remove("blockContent");
    comment["author"] = json!("Editor");
    let error = write_docx_s13(save_request(&model), &source).unwrap_err();
    assert!(error.to_string().contains("nested blocks"), "{error}");
}

#[test]
fn wrapped_block_bookmarks_keep_a_root_declared_prefix() {
    let start = START.replace("w:name=", "x:keep=\"yes\" w:name=");
    for container in ["cell", "header", "footnote", "comment", "textbox"] {
        let source = fixture(
            container,
            &format!(
                r#"<w:customXml w:element="wrap">{START}{}{END}</w:customXml>"#,
                p("00000002", "Anchor")
            ),
        );
        let source = with_root_binding(&source, container, r#"xmlns:x="urn:bookmark""#, &start);
        let model = package(&source);
        let xml = part(&save(&source, &model), container);
        assert_eq!(
            attribute_namespace(&xml, "x:keep"),
            ["urn:bookmark"],
            "{container}: {xml}"
        );
        assert_bound_prefixes(&xml);
    }
}

#[test]
fn rewriting_comment_metadata_rejects_bookmarks_in_nested_blocks() {
    let source = fixture(
        "comment",
        &format!(
            r#"{START}<w:sdt><w:sdtContent>{}{END}</w:sdtContent></w:sdt>{}"#,
            p("00000002", "Nested"),
            p("00000003", "After")
        ),
    );
    let mut model = package(&source);
    assert_eq!(blocks(&mut model, "comment")[1]["type"], "blockSdt");
    model["document"]["comments"][0]["author"] = json!("Changed");
    let error = write_docx_s13(save_request(&model), &source).unwrap_err();
    assert!(
        error
            .to_string()
            .contains("comment bookmarks in nested blocks")
    );
}

#[test]
fn bookmarked_alternate_content_rejects_multiple_drawing_owners() {
    let source = drawing_fixture(false);
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(name, _)| name == "word/document.xml")
        .unwrap();
    let mut xml = String::from_utf8(bytes.clone()).unwrap();
    let start = xml.find("<w:drawing>").unwrap();
    let end = xml.find("</w:drawing>").unwrap() + "</w:drawing>".len();
    let drawing = xml[start..end]
        .replace("'&#55;'", "'9'")
        .replace(r#"w:id="7""#, r#"w:id="9""#);
    xml.insert_str(end, &drawing);
    *bytes = xml.into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let Err(error) = parse_docx_s9_wire(&source, S9ParseOptions::default()) else {
        panic!("multiple bookmarked drawing owners were accepted");
    };
    assert!(
        error
            .to_string()
            .contains("bookmarked alternate content has multiple drawings")
    );
}

#[test]
fn a_vml_image_with_a_text_box_keeps_its_block_bookmarks() {
    let source = fixture("textbox", "<w:p/>");
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    for (name, bytes) in &mut parts {
        let Ok(xml) = std::str::from_utf8(bytes) else {
            continue;
        };
        if name == "word/document.xml" {
            let mut xml = xml.to_owned();
            let start = xml.find("<w:drawing>").unwrap();
            let end = xml.find("</w:drawing>").unwrap() + "</w:drawing>".len();
            xml.replace_range(start..end, &format!(
                r#"<w:pict xmlns:v="urn:schemas-microsoft-com:vml"><v:shape id="TextBox" style="width:72pt;height:72pt"><v:imagedata r:id="image"/><v:textbox><w:txbxContent>{START}{}{END}</w:txbxContent></v:textbox></v:shape></w:pict>"#,
                p("00000002", "Anchor")
            ));
            *bytes = xml.into_bytes();
        } else if name == "word/_rels/document.xml.rels" {
            *bytes = xml.replace("</Relationships>", r#"<Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/pixel.png"/></Relationships>"#).into_bytes();
        } else if name == "[Content_Types].xml" {
            *bytes = xml
                .replace(
                    "</Types>",
                    r#"<Default Extension="png" ContentType="image/png"/></Types>"#,
                )
                .into_bytes();
        }
    }
    parts.push((
        "word/media/pixel.png".to_owned(),
        base64::Engine::decode(
            &base64::engine::general_purpose::STANDARD,
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
        ).unwrap(),
    ));
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut model = package(&source);
    assert_eq!(
        model["document"]["content"][1]["content"][0]["content"][0]["type"],
        "opaqueDrawing"
    );
    for edited in [false, true] {
        if edited {
            model["document"]["content"][0]["content"][0]["content"][0]["text"] = json!("Edited");
            model["document"]["content"][1]["content"]
                .as_array_mut()
                .unwrap()
                .push(json!({"type":"run","content":[{"type":"text","text":"Host edit"}]}));
        }
        let xml = part(&save(&source, &model), "textbox");
        assert_eq!(xml.matches(START).count(), 1, "{xml}");
        assert_eq!(xml.matches(END).count(), 1, "{xml}");
        assert!(xml.contains("<v:imagedata"), "{xml}");
        assert!(xml.contains("<v:textbox>"), "{xml}");
    }
}
