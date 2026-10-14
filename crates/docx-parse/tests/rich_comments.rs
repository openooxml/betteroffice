use docx_parse::document::DocumentBody;
use docx_parse::paragraph_identity::{
    paragraph_ids_by_part, paragraph_occurrences, parse_paragraph_id,
};
use docx_parse::relationships::{
    RelationshipTarget, parse_relationships, resolve_relationship_target,
};
use docx_parse::s9::{S9ParseOptions, parse_docx_s9_wire};
use docx_parse::serializer::{
    S13SaveOptions, S13SaveRequest, SerializerContext, SerializerDeterminism,
    serialize_comments_part, write_docx_s13,
};
use docx_parse::xml::{ParseBudget, ParseLimits, parse_xml};
use serde_json::json;
use sha2::{Digest, Sha256};

const RICH_COMMENT: &str = concat!(
    "<w:comment w:author='Rich reviewer' w:id=\"0\" w:initials=\"R\">\n",
    "  <w:p w14:paraId=\"10000001\"><w:hyperlink r:id=\"link\"><w:r><w:t>Linked</w:t></w:r></w:hyperlink>",
    "<w:r><w:rPr><w:color w:val=\"CC0000\"/><w:u w:val=\"single\"/><w:sz w:val=\"28\"/>",
    "<w:highlight w:val=\"yellow\"/></w:rPr><w:t xml:space=\"preserve\"> colorful</w:t></w:r></w:p>\n",
    "  <w:p w14:paraId=\"10000002\"><w:pPr><w:jc w:val=\"center\"/></w:pPr><w:r><w:t>Second</w:t></w:r></w:p>\n",
    "  <w:tbl><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr><w:tc>",
    "<w:p w14:paraId=\"10000003\"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>\n",
    "  <w:sdt><w:sdtPr><w:id w:val=\"7\"/></w:sdtPr><w:sdtContent>",
    "<w:p w14:paraId=\"10000005\"><w:fldSimple w:instr=\" PAGE \"><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>",
    "</w:sdtContent></w:sdt>\n",
    "  <w:p w14:paraId=\"10000006\"><w:r><w:pict xmlns:v=\"urn:schemas-microsoft-com:vml\">",
    "<v:shape id=\"syntheticPicture\" style=\"width:1pt;height:1pt\"><v:imagedata r:id=\"picture\"/></v:shape>",
    "</w:pict></w:r></w:p>\n",
    "</w:comment>"
);
const PLAIN_COMMENT: &str = "<w:comment w:id=\"1\" w:author=\"Plain reviewer\"><w:p w14:paraId=\"10000004\"><w:r><w:t>Plain</w:t></w:r></w:p></w:comment>";
const COMMENT_RELS: &str = concat!(
    "<?xml version='1.0' encoding='UTF-8'?>\n",
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\n",
    "  <Relationship TargetMode='External' Target='https://example.com/synthetic' Id='link' ",
    "Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'/>\n",
    "  <Relationship Id='picture' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/image' Target='media/pixel.png'/>\n",
    "</Relationships>"
);

fn fixture() -> Vec<u8> {
    let comments = format!(
        "<?xml version='1.0' encoding='UTF-8'?>\n<w:comments xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:w14=\"http://schemas.microsoft.com/office/word/2010/wordml\">\n{RICH_COMMENT}\n{PLAIN_COMMENT}\n</w:comments>"
    );
    ooxml_opc::rezip_parts(&[
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="office" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/_rels/document.xml.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="comments" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="00000001"><w:r><w:t>Synthetic body</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#.to_vec(),
        ),
        ("word/comments.xml".to_owned(), comments.into_bytes()),
        ("word/_rels/comments.xml.rels".to_owned(), COMMENT_RELS.as_bytes().to_vec()),
        (
            "word/media/pixel.png".to_owned(),
            base64::Engine::decode(
                &base64::engine::general_purpose::STANDARD,
                "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
            )
            .unwrap(),
        ),
    ])
    .unwrap()
}

fn save_request(bytes: &[u8]) -> S13SaveRequest {
    let package = parse_docx_s9_wire(bytes, S9ParseOptions::default())
        .unwrap()
        .document
        .package;
    let body = package.document;
    S13SaveRequest {
        determinism: SerializerDeterminism {
            seed: format!("{:x}", Sha256::digest(bytes)),
            now: "2026-01-01T00:00:00.000Z".to_owned(),
        },
        document: DocumentBody {
            content: body.content,
            sections: None,
            final_section_properties: body.final_section_properties,
            custom_root_bindings: body.custom_root_bindings,
            comments: body.comments,
        },
        header_entries: Vec::new(),
        footer_entries: Vec::new(),
        footnotes: Vec::new(),
        endnotes: Vec::new(),
        footnote_separators: Vec::new(),
        endnote_separators: Vec::new(),
        relationship_entries: package.relationship_entries,
        numbering: None,
        options: S13SaveOptions {
            update_modified_date: false,
            modified_by: None,
        },
        selective: None,
        paragraph_ids: None,
    }
}

fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == name)
        .unwrap()
        .1
}

fn assert_rich_comment(bytes: &[u8]) {
    let xml = String::from_utf8(part(bytes, "word/comments.xml")).unwrap();
    let start = xml.find("<w:comment w:author='Rich reviewer'").unwrap();
    let end = start + xml[start..].find("</w:comment>").unwrap() + "</w:comment>".len();
    assert_eq!(&xml.as_bytes()[start..end], RICH_COMMENT.as_bytes());
    assert_eq!(
        part(bytes, "word/_rels/comments.xml.rels"),
        COMMENT_RELS.as_bytes()
    );
}

fn assert_comment_text(bytes: &[u8], id: f64, text: &str) {
    let comments = save_request(bytes).document.comments.unwrap();
    let comment = comments.iter().find(|comment| comment.id == id).unwrap();
    assert!(
        serde_json::to_string(&comment.content)
            .unwrap()
            .contains(text)
    );
}

fn assert_comment_companion_ids(bytes: &[u8]) {
    let xml = String::from_utf8(part(bytes, "word/comments.xml")).unwrap();
    let ids: std::collections::HashSet<_> = paragraph_occurrences(&xml)
        .unwrap()
        .into_iter()
        .filter_map(|paragraph| paragraph.para_id.as_deref().and_then(parse_paragraph_id))
        .collect();
    let limits = ParseLimits::default();
    for (path, xml) in ooxml_opc::unzip_parts(bytes).unwrap() {
        if !matches!(
            path.as_str(),
            "word/commentsExtended.xml" | "word/commentsIds.xml"
        ) {
            continue;
        }
        let document = parse_xml(&xml, &path, &mut ParseBudget::new(&limits)).unwrap();
        for child in document.root().unwrap().child_elements() {
            for name in ["w15:paraId", "w15:paraIdParent", "w16cid:paraId"] {
                if let Some(id) = child.attribute_any(&[name]) {
                    assert!(
                        ids.contains(&parse_paragraph_id(id).unwrap()),
                        "{path}: {id}"
                    );
                }
            }
        }
    }
}

#[test]
fn parse_serialize_preserves_rich_comment_xml_and_relationship_bytes() {
    let source = fixture();
    let saved = write_docx_s13(save_request(&source), &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_rich_comment(&saved);
    assert_eq!(
        part(&saved, "word/comments.xml"),
        part(&source, "word/comments.xml")
    );
    assert_eq!(
        part(&saved, "word/media/pixel.png"),
        part(&source, "word/media/pixel.png")
    );
}

#[test]
fn assigning_a_comment_paragraph_id_saves_the_identity_without_collisions() {
    let mut parts = ooxml_opc::unzip_parts(&fixture()).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/document.xml")
        .unwrap();
    *bytes = std::str::from_utf8(bytes)
        .unwrap()
        .replace("00000001", "10000001")
        .into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut request = save_request(&source);
    request.document.comments = parse_docx_s9_wire(
        &source,
        S9ParseOptions {
            source_ordinals: true,
            ..S9ParseOptions::default()
        },
    )
    .unwrap()
    .document
    .package
    .document
    .comments;
    request.paragraph_ids = Some(
        serde_json::from_value(json!({
            "assignments": [{
                "part": "word/comments.xml", "ordinal": 0, "paraId": "10000007"
            }]
        }))
        .unwrap(),
    );
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Plain");
    let ids = paragraph_ids_by_part(&saved).unwrap();
    let body_ids = &ids["/word/document.xml"];
    let comment_ids = &ids["/word/comments.xml"];
    assert_eq!(body_ids, &["10000001".to_owned()]);
    assert!(comment_ids.contains(&"10000007".to_owned()));
    assert!(comment_ids.iter().all(|id| !body_ids.contains(id)));
}

#[test]
fn editing_a_plain_comment_saves_the_edit() {
    let source = fixture();
    let mut request = save_request(&source);
    let comment = &mut request.document.comments.as_mut().unwrap()[1];
    comment.content[0].content = vec![
        serde_json::from_value(json!({
            "type": "run", "content": [{ "type": "text", "text": "Edited plain comment" }]
        }))
        .unwrap(),
    ];
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Edited plain comment");
    let xml = String::from_utf8(part(&saved, "word/comments.xml")).unwrap();
    assert!(!xml.contains("<w:t>Plain</w:t>"));
}

#[test]
fn deleting_a_plain_comment_saves_the_deletion() {
    let source = fixture();
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap().remove(1);
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert!(
        save_request(&saved)
            .document
            .comments
            .unwrap()
            .iter()
            .all(|comment| comment.id != 1.0)
    );
}

#[test]
fn resolving_a_rich_comment_saves_its_done_flag() {
    let source = fixture();
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap()[0].done = Some(true);
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 1.0, "Plain");
    assert_comment_text(&saved, 0.0, " colorful");
    assert_eq!(
        save_request(&saved).document.comments.unwrap()[0].done,
        Some(true)
    );
}

#[test]
fn adding_a_comment_saves_the_new_comment() {
    let source = fixture();
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap().push(
        serde_json::from_value(json!({
            "id": 2, "author": "New reviewer", "content": [{
                "type": "paragraph", "content": [{
                    "type": "run", "content": [{ "type": "text", "text": "Added comment" }]
                }]
            }]
        }))
        .unwrap(),
    );
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Plain");
    assert_comment_text(&saved, 2.0, "Added comment");
}

fn without_rich_paragraph_ids(xml: &str) -> String {
    let mut xml = xml.to_owned();
    for id in ["10000001", "10000002", "10000003", "10000005", "10000006"] {
        xml = xml.replace(&format!(" w14:paraId=\"{id}\""), "");
    }
    xml
}

fn fixture_without_rich_paragraph_ids() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&fixture()).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/comments.xml")
        .unwrap();
    *bytes = without_rich_paragraph_ids(std::str::from_utf8(bytes).unwrap()).into_bytes();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn fixture_with_comment_parts(comments: &str, companions: &[(&str, &str)]) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&fixture()).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/comments.xml")
        .unwrap();
    *bytes = std::str::from_utf8(bytes)
        .unwrap()
        .replace(&format!("{RICH_COMMENT}\n{PLAIN_COMMENT}"), comments)
        .into_bytes();
    for (path, xml) in companions {
        parts.push(((*path).to_owned(), xml.as_bytes().to_vec()));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn comments_without_paragraph_ids_keep_their_source_bytes() {
    let source = fixture_without_rich_paragraph_ids();
    let saved = write_docx_s13(save_request(&source), &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_eq!(
        part(&saved, "word/comments.xml"),
        part(&source, "word/comments.xml")
    );
    assert_eq!(
        part(&saved, "word/_rels/comments.xml.rels"),
        part(&source, "word/_rels/comments.xml.rels")
    );
}

#[test]
fn resolving_and_reparenting_a_comment_without_source_ids_saves_its_metadata() {
    let source = fixture_without_rich_paragraph_ids();
    let mut request = save_request(&source);
    let comment = &mut request.document.comments.as_mut().unwrap()[0];
    comment.done = Some(true);
    comment.parent_id = Some(1.0);
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Plain");
    let comments = save_request(&saved).document.comments.unwrap();
    let comment = comments.iter().find(|comment| comment.id == 0.0).unwrap();
    assert_eq!(comment.done, Some(true));
    assert_eq!(comment.parent_id, Some(1.0));
    let extended = String::from_utf8(part(&saved, "word/commentsExtended.xml")).unwrap();
    assert!(extended.contains(&format!(
        "<w15:commentEx w15:paraId=\"{}\" w15:done=\"1\" w15:paraIdParent=\"10000004\" />",
        comment.para_id.as_deref().unwrap()
    )));
}

#[test]
fn resolving_a_plain_comment_without_source_ids_saves_its_done_flag() {
    let mut parts = ooxml_opc::unzip_parts(&fixture_without_rich_paragraph_ids()).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/comments.xml")
        .unwrap();
    *bytes = std::str::from_utf8(bytes)
        .unwrap()
        .replace(" w14:paraId=\"10000004\"", "")
        .into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap()[1].done = Some(true);
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Plain");
    assert_eq!(
        save_request(&saved).document.comments.unwrap()[1].done,
        Some(true)
    );
}

#[test]
fn clearing_inline_comment_metadata_saves_the_new_state() {
    let mut parts = ooxml_opc::unzip_parts(&fixture()).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/comments.xml")
        .unwrap();
    *bytes = std::str::from_utf8(bytes)
        .unwrap()
        .replace(
            "<w:comment w:author='Rich reviewer'",
            "<w:comment w:done='1' w:parentId='1' w:author='Rich reviewer'",
        )
        .into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut request = save_request(&source);
    let comment = &mut request.document.comments.as_mut().unwrap()[0];
    comment.done = Some(false);
    comment.parent_id = None;
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Plain");
    let xml = String::from_utf8(part(&saved, "word/comments.xml")).unwrap();
    assert!(!xml.contains("w:done="));
    assert!(!xml.contains("w:parentId="));
    let comments = save_request(&saved).document.comments.unwrap();
    let comment = comments.iter().find(|comment| comment.id == 0.0).unwrap();
    assert_ne!(comment.done, Some(true));
    assert_eq!(comment.parent_id, None);
    let extended = String::from_utf8(part(&saved, "word/commentsExtended.xml")).unwrap();
    assert!(extended.contains(&format!(
        "<w15:commentEx w15:paraId=\"{}\" w15:done=\"0\" />",
        comment.para_id.as_deref().unwrap()
    )));
}

#[test]
fn duplicate_source_comment_ids_fall_back_to_the_plain_writer() {
    let source = fixture();
    let request = save_request(&source);
    let expected = serialize_comments_part(
        request.document.comments.as_ref().unwrap(),
        &mut SerializerContext::new(&request.determinism).unwrap(),
    );
    let mut parts = ooxml_opc::unzip_parts(&source).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/comments.xml")
        .unwrap();
    *bytes = std::str::from_utf8(bytes)
        .unwrap()
        .replace(
            PLAIN_COMMENT,
            &PLAIN_COMMENT.replace("w:id=\"1\"", "w:id=\"0\""),
        )
        .into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_eq!(part(&saved, "word/comments.xml"), expected.as_bytes());
}

#[test]
fn adding_a_reply_to_a_comment_without_source_ids_saves_its_text_and_parent() {
    let source = fixture_without_rich_paragraph_ids();
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap().push(
        serde_json::from_value(json!({
            "id": 2, "author": "Reply reviewer", "parentId": 0, "content": [{
                "type": "paragraph", "content": [{
                    "type": "run", "content": [{ "type": "text", "text": "Reply" }]
                }]
            }]
        }))
        .unwrap(),
    );
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 2.0, "Reply");
    assert_comment_text(&saved, 1.0, "Plain");
    let comments = save_request(&saved).document.comments.unwrap();
    assert_eq!(
        comments
            .iter()
            .find(|comment| comment.id == 2.0)
            .unwrap()
            .parent_id,
        Some(0.0)
    );
    let parent = comments.iter().find(|comment| comment.id == 0.0).unwrap();
    let extended = String::from_utf8(part(&saved, "word/commentsExtended.xml")).unwrap();
    assert!(extended.contains(&format!(
        "w15:paraIdParent=\"{}\"",
        parent.para_id.as_deref().unwrap()
    )));
}

#[test]
fn resolving_a_comment_with_a_trailing_nested_paragraph_saves_its_done_flag() {
    for nested in [
        "<w:tbl><w:tblGrid><w:gridCol w:w=\"2400\"/></w:tblGrid><w:tr><w:tc><w:p w14:paraId=\"10000002\"><w:r><w:t>Cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>",
        "<w:sdt><w:sdtPr/><w:sdtContent><w:p w14:paraId=\"10000002\"><w:r><w:t>Control</w:t></w:r></w:p></w:sdtContent></w:sdt>",
    ] {
        let fragment = format!(
            "<w:comment w:id=\"0\" w:author=\"Reviewer\"><w:p w14:paraId=\"10000001\"><w:r><w:t>Direct</w:t></w:r></w:p>{nested}</w:comment>"
        );
        let source = fixture_with_comment_parts(&format!("{fragment}\n{PLAIN_COMMENT}"), &[]);
        let mut request = save_request(&source);
        request.document.comments.as_mut().unwrap()[0].done = Some(true);
        let saved = write_docx_s13(request, &source).unwrap();
        assert_comment_companion_ids(&saved);
        assert_comment_text(&saved, 1.0, "Plain");
        let ids = paragraph_ids_by_part(&saved).unwrap();
        assert!(ids["/word/comments.xml"].contains(&"10000001".to_owned()));
        assert_comment_text(&saved, 0.0, "Direct");
        let extended = String::from_utf8(part(&saved, "word/commentsExtended.xml")).unwrap();
        assert!(extended.contains("w15:paraId=\"10000001\" w15:done=\"1\""));
        assert_eq!(
            save_request(&saved).document.comments.unwrap()[0].done,
            Some(true)
        );
    }
}

#[test]
fn deleting_the_only_comment_with_source_ids_saves_the_remaining_comment() {
    let retained = PLAIN_COMMENT
        .replace("w:id=\"1\"", "w:id=\"0\"")
        .replace(" w14:paraId=\"10000004\"", "");
    for companion in [
        (
            "word/commentsExtended.xml",
            "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\"><w15:commentEx w15:paraId=\"10000004\" w15:done=\"1\"/></w15:commentsEx>",
        ),
        (
            "word/commentsIds.xml",
            "<w16cid:commentsIds xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\"><w16cid:commentId w16cid:paraId=\"10000004\" w16cid:durableId=\"20000001\"/></w16cid:commentsIds>",
        ),
        (
            "word/commentsExtensible.xml",
            "<w16cex:commentsExtensible xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\"><w16cex:commentExtensible w16cex:durableId=\"20000001\" w16cex:dateUtc=\"2026-01-01T00:00:00Z\"/></w16cex:commentsExtensible>",
        ),
    ] {
        let source =
            fixture_with_comment_parts(&format!("{retained}\n{PLAIN_COMMENT}"), &[companion]);
        let mut request = save_request(&source);
        request.document.comments.as_mut().unwrap().remove(1);
        let expected = serialize_comments_part(
            request.document.comments.as_ref().unwrap(),
            &mut SerializerContext::new(&request.determinism).unwrap(),
        );
        let saved = write_docx_s13(request, &source).unwrap();
        assert_eq!(part(&saved, "word/comments.xml"), expected.as_bytes());
        assert_comment_companion_ids(&saved);
        assert_comment_text(&saved, 0.0, "Plain");
        assert!(
            save_request(&saved)
                .document
                .comments
                .unwrap()
                .iter()
                .all(|comment| comment.id != 1.0)
        );
        let xml = String::from_utf8(part(&saved, companion.0)).unwrap();
        assert!(!xml.contains("10000004"));
        assert!(!xml.contains("20000001"));
        assert_eq!(save_request(&saved).document.comments.unwrap().len(), 1);
    }
}

#[test]
fn malformed_comment_text_uses_the_writer_part() {
    let malformed = PLAIN_COMMENT.replace("<w:t>Plain</w:t>", "<w:t>A & B</w:t>");
    let source = fixture_with_comment_parts(&format!("{RICH_COMMENT}\n{malformed}"), &[]);
    let request = save_request(&source);
    let expected = serialize_comments_part(
        request.document.comments.as_ref().unwrap(),
        &mut SerializerContext::new(&request.determinism).unwrap(),
    );
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_eq!(part(&saved, "word/comments.xml"), expected.as_bytes());
    let xml = String::from_utf8(part(&saved, "word/comments.xml")).unwrap();
    assert!(xml.contains("<w:t>A &amp; B</w:t>"));
    assert!(!xml.contains("<w:t>A & B</w:t>"));
}

#[test]
fn adding_a_comment_with_a_non_utf8_source_declaration_saves_its_text() {
    let mut parts =
        ooxml_opc::unzip_parts(&fixture_with_comment_parts(PLAIN_COMMENT, &[])).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/comments.xml")
        .unwrap();
    *bytes = std::str::from_utf8(bytes)
        .unwrap()
        .replace("encoding='UTF-8'", "encoding='windows-1252'")
        .into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap().push(
        serde_json::from_value(json!({
            "id": 2, "author": "New reviewer", "content": [{
                "type": "paragraph", "content": [{
                    "type": "run", "content": [{ "type": "text", "text": "café" }]
                }]
            }]
        }))
        .unwrap(),
    );
    let expected = serialize_comments_part(
        request.document.comments.as_ref().unwrap(),
        &mut SerializerContext::new(&request.determinism).unwrap(),
    );
    let saved = write_docx_s13(request, &source).unwrap();
    assert_eq!(part(&saved, "word/comments.xml"), expected.as_bytes());
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 1.0, "Plain");
    let comments = save_request(&saved).document.comments.unwrap();
    let added = comments.iter().find(|comment| comment.id == 2.0).unwrap();
    assert!(
        serde_json::to_string(&added.content)
            .unwrap()
            .contains("café")
    );
}

#[test]
fn invalid_source_comment_paragraph_id_uses_the_writer_part() {
    let invalid = PLAIN_COMMENT.replace("10000004", "1000000&quot;");
    let source = fixture_with_comment_parts(&invalid, &[]);
    let saved = write_docx_s13(save_request(&source), &source).unwrap();
    assert_comment_companion_ids(&saved);
    let limits = ParseLimits::default();
    for path in ["word/commentsExtended.xml", "word/commentsIds.xml"] {
        let document =
            parse_xml(&part(&saved, path), path, &mut ParseBudget::new(&limits)).unwrap();
        let root = document.root().unwrap();
        assert!(root.child_elements().next().is_some());
        for child in root.child_elements() {
            let id = child
                .attribute_any(&["w15:paraId", "w16cid:paraId"])
                .unwrap();
            assert!(parse_paragraph_id(id).is_some());
        }
    }
    let comments = save_request(&saved).document.comments.unwrap();
    assert!(parse_paragraph_id(comments[0].para_id.as_deref().unwrap()).is_some());
}

#[test]
fn comments_without_source_ids_register_only_existing_parts() {
    let comment = PLAIN_COMMENT.replace(" w14:paraId=\"10000004\"", "");
    let source = fixture_with_comment_parts(&comment, &[]);
    let saved = write_docx_s13(save_request(&source), &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_eq!(
        part(&saved, "word/comments.xml"),
        part(&source, "word/comments.xml")
    );
    assert_eq!(
        part(&saved, "word/_rels/comments.xml.rels"),
        part(&source, "word/_rels/comments.xml.rels")
    );
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let limits = ParseLimits::default();
    for (path, bytes) in parts.iter().filter(|(path, _)| path.ends_with(".rels")) {
        let relationships =
            parse_relationships(bytes, path, &mut ParseBudget::new(&limits)).unwrap();
        for relationship in relationships.values() {
            if let RelationshipTarget::Internal(target) =
                resolve_relationship_target(path, relationship).unwrap()
            {
                assert!(parts.iter().any(|(path, _)| path == &target), "{target}");
            }
        }
    }
    let content_types = parse_xml(
        &part(&saved, "[Content_Types].xml"),
        "[Content_Types].xml",
        &mut ParseBudget::new(&limits),
    )
    .unwrap();
    for entry in content_types
        .root()
        .unwrap()
        .children_by_local_name("Override")
    {
        let target = entry.attribute(None, "PartName").unwrap();
        assert!(
            parts
                .iter()
                .any(|(path, _)| path == target.trim_start_matches('/')),
            "{target}"
        );
    }
}

#[test]
fn malformed_comment_root_attribute_uses_the_writer_part() {
    let mut parts =
        ooxml_opc::unzip_parts(&fixture_with_comment_parts(PLAIN_COMMENT, &[])).unwrap();
    let (_, bytes) = parts
        .iter_mut()
        .find(|(path, _)| path == "word/comments.xml")
        .unwrap();
    *bytes = std::str::from_utf8(bytes)
        .unwrap()
        .replace("<w:comments ", "<w:comments xmlns:synthetic=\"urn:A & B\" ")
        .into_bytes();
    let source = ooxml_opc::rezip_parts(&parts).unwrap();
    let request = save_request(&source);
    let expected = serialize_comments_part(
        request.document.comments.as_ref().unwrap(),
        &mut SerializerContext::new(&request.determinism).unwrap(),
    );
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    let xml = part(&saved, "word/comments.xml");
    assert_eq!(xml, expected.as_bytes());
    assert!(docx_parse::xml::reads_as_written(&xml));
}

#[test]
fn setting_a_durable_id_without_source_paragraph_ids_saves_the_new_id() {
    let source = fixture_without_rich_paragraph_ids();
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap()[0].durable_id = Some("20000001".to_owned());
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Plain");
    let ids = String::from_utf8(part(&saved, "word/commentsIds.xml")).unwrap();
    assert!(ids.contains("w16cid:durableId=\"20000001\""));
    assert_eq!(
        save_request(&saved).document.comments.unwrap()[0]
            .durable_id
            .as_deref(),
        Some("20000001")
    );
}

#[test]
fn legacy_companion_date_is_written_on_the_comment_element() {
    let fragment = "<w:comment w:id=\"0\" w:author=\"Reviewer\" w14:paraId=\"10000001\"><w:p w14:paraId=\"10000001\"><w:r><w:t>Legacy date</w:t></w:r></w:p></w:comment>";
    let source = fixture_with_comment_parts(
        &format!("{fragment}\n{PLAIN_COMMENT}"),
        &[(
            "word/commentsExtensible.xml",
            "<w16cex:commentsExtensible xmlns:w16cex=\"http://schemas.microsoft.com/office/word/2018/wordml/cex\"><w16cex:comment w16cex:paraId=\"10000001\" w16cex:dateUtc=\"2026-01-01T00:00:00Z\"/></w16cex:commentsExtensible>",
        )],
    );
    let request = save_request(&source);
    assert_eq!(
        request.document.comments.as_ref().unwrap()[0]
            .date
            .as_deref(),
        Some("2026-01-01T00:00:00Z")
    );
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 1.0, "Plain");
    let xml = String::from_utf8(part(&saved, "word/comments.xml")).unwrap();
    assert!(xml.contains("w:date=\"2026-01-01T00:00:00Z\""));
    assert_eq!(
        save_request(&saved).document.comments.unwrap()[0]
            .date
            .as_deref(),
        Some("2026-01-01T00:00:00Z")
    );
}

#[test]
fn reparenting_a_comment_saves_its_paragraph_id_and_parent() {
    let source = fixture_with_comment_parts(
        &format!(
            "{PLAIN_COMMENT}\n{}",
            without_rich_paragraph_ids(RICH_COMMENT)
        ),
        &[],
    );
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap()[0].parent_id = Some(0.0);
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Plain");
    let comments = save_request(&saved).document.comments.unwrap();
    let reply = comments.iter().find(|comment| comment.id == 1.0).unwrap();
    assert_eq!(reply.para_id.as_deref(), Some("10000004"));
    assert_eq!(reply.parent_id, Some(0.0));
    let parent = comments.iter().find(|comment| comment.id == 0.0).unwrap();
    let extended = String::from_utf8(part(&saved, "word/commentsExtended.xml")).unwrap();
    assert!(extended.contains(&format!(
        "w15:paraId=\"10000004\" w15:done=\"0\" w15:paraIdParent=\"{}\"",
        parent.para_id.as_deref().unwrap()
    )));
}

#[test]
fn editing_an_inline_durable_id_saves_the_new_id() {
    let plain = PLAIN_COMMENT.replace(
        "<w:comment ",
        "<w:comment xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" w16cid:durableId=\"20000001\" ",
    );
    let source = fixture_with_comment_parts(&format!("{RICH_COMMENT}\n{plain}"), &[]);
    let mut request = save_request(&source);
    let comment = &mut request.document.comments.as_mut().unwrap()[1];
    assert_eq!(comment.durable_id.as_deref(), Some("20000001"));
    comment.durable_id = Some("20000002".to_owned());
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 1.0, "Plain");
    assert_comment_text(&saved, 0.0, " colorful");
    let ids = String::from_utf8(part(&saved, "word/commentsIds.xml")).unwrap();
    assert!(ids.contains("w16cid:paraId=\"10000004\" w16cid:durableId=\"20000002\""));
    let comments = save_request(&saved).document.comments.unwrap();
    assert_eq!(comments.len(), 2);
    let comment = comments.iter().find(|comment| comment.id == 1.0).unwrap();
    assert_eq!(comment.durable_id.as_deref(), Some("20000002"));
    assert_eq!(comment.para_id.as_deref(), Some("10000004"));
}

#[test]
fn resolving_a_comment_with_a_trailing_unidentified_paragraph_saves_its_done_flag() {
    let plain = PLAIN_COMMENT.replace("</w:comment>", "<w:p/></w:comment>");
    let source = fixture_with_comment_parts(&format!("{RICH_COMMENT}\n{plain}"), &[]);
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap()[1].done = Some(true);
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 1.0, "Plain");
    assert_comment_text(&saved, 0.0, " colorful");
    let xml = String::from_utf8(part(&saved, "word/comments.xml")).unwrap();
    let paragraphs: Vec<_> = paragraph_occurrences(&xml)
        .unwrap()
        .into_iter()
        .filter(|paragraph| paragraph.item_id.as_deref() == Some("1"))
        .collect();
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(paragraphs[1].para_id.as_deref(), Some("10000004"));
    let extended = String::from_utf8(part(&saved, "word/commentsExtended.xml")).unwrap();
    assert!(extended.contains("w15:paraId=\"10000004\" w15:done=\"1\""));
    let comments = save_request(&saved).document.comments.unwrap();
    assert_eq!(comments.len(), 2);
    let comment = comments.iter().find(|comment| comment.id == 1.0).unwrap();
    assert_eq!(comment.done, Some(true));
    assert_eq!(comment.content.len(), 2);
    assert_eq!(comment.content[1].para_id.as_deref(), Some("10000004"));
    assert_eq!(comment.para_id.as_deref(), Some("10000004"));
}

#[test]
fn editing_a_model_comment_paragraph_id_updates_all_comment_parts() {
    let source = fixture();
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap()[1].content[0].para_id = Some("10000009".to_owned());
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Plain");
    for path in [
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
    ] {
        let xml = String::from_utf8(part(&saved, path)).unwrap();
        assert!(xml.contains("paraId=\"10000009\""), "{path}");
        assert!(!xml.contains("10000004"), "{path}");
    }
    let comments = save_request(&saved).document.comments.unwrap();
    assert_eq!(comments.len(), 2);
    let comment = comments.iter().find(|comment| comment.id == 1.0).unwrap();
    assert_eq!(comment.content[0].para_id.as_deref(), Some("10000009"));
    assert_eq!(comment.para_id.as_deref(), Some("10000009"));
}

#[test]
fn invalid_earlier_comment_paragraph_ids_use_the_writer_part() {
    for id in ["10000001", "10000003"] {
        let rich = RICH_COMMENT.replace(id, "1000000&quot;");
        let source = fixture_with_comment_parts(&format!("{rich}\n{PLAIN_COMMENT}"), &[]);
        let request = save_request(&source);
        let written = serialize_comments_part(
            request.document.comments.as_ref().unwrap(),
            &mut SerializerContext::new(&request.determinism).unwrap(),
        );
        let saved = write_docx_s13(request, &source).unwrap();
        assert_comment_companion_ids(&saved);
        let xml = String::from_utf8(part(&saved, "word/comments.xml")).unwrap();
        assert_eq!(xml, written);
        assert_comment_text(&saved, 1.0, "Plain");
        assert!(!xml.contains("1000000&quot;"));
        for paragraph in paragraph_occurrences(&xml).unwrap() {
            if let Some(id) = paragraph.para_id {
                assert!(parse_paragraph_id(&id).is_some());
            }
        }
        let comments = save_request(&saved).document.comments.unwrap();
        assert_eq!(comments.len(), 2);
        let comment = comments.iter().find(|comment| comment.id == 0.0).unwrap();
        assert_eq!(comment.content.len(), 3);
        assert_eq!(comment.para_id.as_deref(), Some("10000006"));
        for paragraph in &comment.content {
            if let Some(id) = &paragraph.para_id {
                assert!(parse_paragraph_id(id).is_some());
            }
        }
        let content = serde_json::to_string(&comment.content).unwrap();
        assert!(content.contains(" colorful"));
        assert!(content.contains("Second"));
    }
}

#[test]
fn noncanonical_comment_id_saves_one_canonical_comment() {
    let plain = PLAIN_COMMENT.replace("w:id=\"1\"", "w:id=\"1tail\"");
    let source = fixture_with_comment_parts(&format!("{RICH_COMMENT}\n{plain}"), &[]);
    let saved = write_docx_s13(save_request(&source), &source).unwrap();
    assert_comment_companion_ids(&saved);
    let xml = part(&saved, "word/comments.xml");
    assert!(!std::str::from_utf8(&xml).unwrap().contains("1tail"));
    let limits = ParseLimits::default();
    let document = parse_xml(&xml, "word/comments.xml", &mut ParseBudget::new(&limits)).unwrap();
    let root = document.root().unwrap();
    assert_eq!(root.children_by_local_name("comment").count(), 2);
    assert_eq!(
        root.children_by_local_name("comment")
            .filter(|comment| comment.attribute(Some("w"), "id") == Some("1"))
            .count(),
        1
    );
    let comments = save_request(&saved).document.comments.unwrap();
    assert_eq!(comments.len(), 2);
    let comment = comments.iter().find(|comment| comment.id == 1.0).unwrap();
    assert_eq!(comment.para_id.as_deref(), Some("10000004"));
    assert!(
        serde_json::to_string(&comment.content)
            .unwrap()
            .contains("Plain")
    );
}

#[test]
fn setting_a_comment_paragraph_id_without_source_ids_saves_it_in_all_comment_parts() {
    let plain = PLAIN_COMMENT.replace(" w14:paraId=\"10000004\"", "");
    let source = fixture_with_comment_parts(&format!("{RICH_COMMENT}\n{plain}"), &[]);
    let mut request = save_request(&source);
    request.document.comments.as_mut().unwrap()[1].para_id = Some("10000009".to_owned());
    let saved = write_docx_s13(request, &source).unwrap();
    assert_comment_companion_ids(&saved);
    for path in [
        "word/comments.xml",
        "word/commentsExtended.xml",
        "word/commentsIds.xml",
    ] {
        let xml = String::from_utf8(part(&saved, path)).unwrap();
        assert!(xml.contains("paraId=\"10000009\""), "{path}");
    }
    let comments = save_request(&saved).document.comments.unwrap();
    let comment = comments.iter().find(|comment| comment.id == 1.0).unwrap();
    assert_eq!(comment.para_id.as_deref(), Some("10000009"));
    assert_comment_text(&saved, 0.0, " colorful");
    assert_comment_text(&saved, 1.0, "Plain");
}

#[test]
fn changing_only_comment_metadata_writes_the_whole_comments_part() {
    for change in ["done", "durable_id", "para_id", "explicit_done"] {
        let source = fixture();
        let mut request = save_request(&source);
        let comment = &mut request.document.comments.as_mut().unwrap()[1];
        match change {
            "done" => comment.done = Some(true),
            "durable_id" => comment.durable_id = Some("20000009".to_owned()),
            "para_id" => comment.content[0].para_id = Some("10000009".to_owned()),
            "explicit_done" => comment.done = Some(false),
            _ => unreachable!(),
        }
        let saved = write_docx_s13(request, &source).unwrap();
        assert_ne!(
            part(&saved, "word/comments.xml"),
            part(&source, "word/comments.xml"),
            "{change}"
        );
        assert_comment_companion_ids(&saved);
        assert_comment_text(&saved, 0.0, " colorful");
        assert_comment_text(&saved, 1.0, "Plain");
        let comments = save_request(&saved).document.comments.unwrap();
        let comment = comments.iter().find(|comment| comment.id == 1.0).unwrap();
        match change {
            "done" => assert_eq!(comment.done, Some(true)),
            "durable_id" => assert_eq!(comment.durable_id.as_deref(), Some("20000009")),
            "para_id" => assert_eq!(comment.para_id.as_deref(), Some("10000009")),
            "explicit_done" => assert_ne!(comment.done, Some(true)),
            _ => unreachable!(),
        }
    }
}
