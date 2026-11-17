use serde_json::{Value, json};

use crate::EngineSession;

const NAMESPACES: &str = r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#;

#[derive(Clone, Copy)]
enum TextBox {
    Drawing,
    Vml,
    Header,
}

fn run(text: &str, font: &str) -> String {
    format!(
        r#"<w:r><w:rPr><w:rFonts w:ascii="{font}" w:eastAsia="{font}" w:hAnsi="{font}"/></w:rPr><w:t>{text}</w:t></w:r>"#
    )
}

fn drawing_text_box(story: &str) -> String {
    format!(
        r#"<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>457200</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>228600</wp:posOffset></wp:positionV><wp:extent cx="2743200" cy="1371600"/><wp:wrapNone/><wp:docPr id="1" name="Box"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"/></wps:spPr><wps:txbx><w:txbxContent><w:p>{story}</w:p></w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:t>Fallback</w:t></mc:Fallback></mc:AlternateContent></w:r></w:p>"#
    )
}

fn vml_text_box(story: &str) -> String {
    format!(
        r#"<w:p><w:r><w:pict><v:shape id="Text Box 1" o:spid="_x0000_s1026" style="position:absolute;margin-left:36pt;margin-top:72pt;width:144pt;height:72pt;z-index:1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" filled="f" stroked="f"><v:textbox inset="0,0,0,0"><w:txbxContent><w:p>{story}</w:p></w:txbxContent></v:textbox><w10:wrap type="none"/></v:shape></w:pict></w:r></w:p>"#
    )
}

/// Body paragraphs in MS Mincho, each with the East Asian language of its
/// paragraph style if it names one, and a text box of pure Han in MS Gothic, in
/// a document whose default East Asian language is `language`.
fn document(language: Option<&str>, body: &[(&str, Option<&str>)], text_box: TextBox) -> Vec<u8> {
    let default_language = language.map_or_else(String::new, |tag| {
        format!(r#"<w:lang w:val="en-US" w:eastAsia="{tag}"/>"#)
    });
    let mut styles = String::new();
    let mut paragraphs = String::new();
    for (index, (text, language)) in body.iter().enumerate() {
        let properties = language.map_or_else(String::new, |tag| {
            styles.push_str(&format!(
                r#"<w:style w:type="paragraph" w:styleId="Language{index}"><w:name w:val="Language {index}"/><w:rPr><w:lang w:eastAsia="{tag}"/></w:rPr></w:style>"#
            ));
            format!(r#"<w:pPr><w:pStyle w:val="Language{index}"/></w:pPr>"#)
        });
        paragraphs.push_str(&format!(
            "<w:p>{properties}{}</w:p>",
            run(text, "MS Mincho")
        ));
    }
    let story = run("骨直角", "MS Gothic");
    let (body_box, header, header_reference) = match text_box {
        TextBox::Drawing => (drawing_text_box(&story), None, ""),
        TextBox::Vml => (vml_text_box(&story), None, ""),
        TextBox::Header => (
            String::new(),
            Some(format!(
                "<w:hdr {NAMESPACES}>{}</w:hdr>",
                drawing_text_box(&story)
            )),
            r#"<w:headerReference w:type="default" r:id="rId2"/>"#,
        ),
    };
    let mut parts = vec![
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_owned()),
        ("word/styles.xml", format!(r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="21"/>{default_language}</w:rPr></w:rPrDefault></w:docDefaults>{styles}</w:styles>"#)),
        ("word/document.xml", format!(r#"<w:document {NAMESPACES}><w:body>{paragraphs}{body_box}<w:sectPr>{header_reference}<w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/><w:cols w:space="720"/></w:sectPr></w:body></w:document>"#)),
    ];
    parts.extend(header.map(|header| ("word/header1.xml", header)));
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The scripts `family` asks for, through the request both renderers send, or
/// null when nothing in the document asks for it.
fn scripts(bytes: &[u8], family: &str) -> Value {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let mut request =
        crate::layout_request::region_layout_request(&serde_json::to_value(package).unwrap(), 24.0);
    request["measurement"] = json!({"defaults": {"fontFamily": "Calibri", "fontSize": 11}});
    let engine = EngineSession::new(75303);
    crate::seed_from_docx(engine.doc(), bytes).unwrap();
    let requirements: Vec<Value> = serde_json::from_str(
        &engine
            .layout_font_requirements_json(&request.to_string())
            .unwrap(),
    )
    .unwrap();
    requirements
        .into_iter()
        .find(|requirement| requirement["key"] == format!("{family}|0|0"))
        .map(|requirement| requirement["scripts"].clone())
        .unwrap_or(Value::Null)
}

fn text_box_scripts(language: Option<&str>, text_box: TextBox) -> Value {
    scripts(
        &document(language, &[("本文", None)], text_box),
        "ms gothic",
    )
}

fn body_scripts(language: Option<&str>, body: &[(&str, Option<&str>)]) -> Value {
    scripts(&document(language, body, TextBox::Drawing), "ms mincho")
}

#[test]
fn a_text_box_inherits_the_documents_default_east_asian_language() {
    assert_eq!(
        text_box_scripts(Some("ja-JP"), TextBox::Drawing),
        json!(["cjk-jp"])
    );
    assert_eq!(
        text_box_scripts(Some("ko-KR"), TextBox::Drawing),
        json!(["cjk-kr"])
    );
    assert_eq!(text_box_scripts(None, TextBox::Drawing), json!(["cjk-sc"]));
}

#[test]
fn a_vml_text_box_inherits_the_documents_default_east_asian_language() {
    assert_eq!(
        text_box_scripts(Some("ja-JP"), TextBox::Vml),
        json!(["cjk-jp"])
    );
}

#[test]
fn a_header_text_box_inherits_the_documents_default_east_asian_language() {
    assert_eq!(
        text_box_scripts(Some("ja-JP"), TextBox::Header),
        json!(["cjk-jp"])
    );
}

#[test]
fn a_japanese_document_keeps_its_forms_past_a_stray_chinese_paragraph() {
    assert_eq!(
        body_scripts(Some("ja-JP"), &[("汉字", Some("zh-CN")), ("日本語", None)]),
        json!(["cjk-jp", "cjk-sc"])
    );
}

#[test]
fn a_chinese_document_keeps_its_forms_past_a_stray_japanese_paragraph() {
    assert_eq!(
        body_scripts(Some("zh-CN"), &[("日本語", Some("ja-JP")), ("汉字", None)]),
        json!(["cjk-sc", "cjk-jp"])
    );
}

#[test]
fn a_document_naming_no_language_prefers_simplified_forms() {
    assert_eq!(
        body_scripts(None, &[("日本語", Some("ja-JP")), ("汉字", None)]),
        json!(["cjk-sc", "cjk-jp"])
    );
}
