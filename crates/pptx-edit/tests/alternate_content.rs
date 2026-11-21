use pptx_edit::{DeckSession, EditCtx};

const FIXTURE: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/alternate-content.pptx");

fn slide_xml(bytes: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let part = parts
        .iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    String::from_utf8(part.1.clone()).unwrap()
}

fn wrapper(xml: &str) -> &str {
    let start = xml.find("<mc:AlternateContent").unwrap();
    let end = xml.find("</mc:AlternateContent>").unwrap() + "</mc:AlternateContent>".len();
    &xml[start..end]
}

#[test]
fn a_fill_edit_lands_on_the_shape_the_fallback_holds() {
    let session = DeckSession::open(FIXTURE, 332).unwrap();
    let context = EditCtx::local("test");
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let names: Vec<&str> = slide
        .shapes
        .iter()
        .map(|shape| shape.name.as_str())
        .collect();
    assert_eq!(names, ["title", "control", "fallback", "caption"]);

    session
        .set_shape_fill(&context, &slide.id, &slide.shapes[2].id, Some("#DC2626"))
        .unwrap();
    session
        .set_shape_fill(&context, &slide.id, &slide.shapes[3].id, Some("#12B76A"))
        .unwrap();
    let xml = slide_xml(&session.save().unwrap());

    let fallback = xml.split("<mc:Fallback>").nth(1).unwrap();
    assert!(fallback.contains(r#"name="fallback""#));
    assert!(fallback.contains(r#"<a:srgbClr val="DC2626"/>"#));
    assert!(
        xml.contains(r#"<mc:Choice xmlns:bo="urn:betteroffice:fixture-extension" Requires="bo">"#)
    );
    let caption = xml.split(r#"name="caption""#).nth(1).unwrap();
    assert!(caption.contains(r#"<a:srgbClr val="12B76A"/>"#));
}

#[test]
fn a_save_without_edits_leaves_the_alternate_content_untouched() {
    let session = DeckSession::open(FIXTURE, 333).unwrap();
    assert_eq!(slide_xml(&session.save().unwrap()), slide_xml(FIXTURE));
}

#[test]
fn an_edit_elsewhere_on_the_slide_leaves_the_branch_intact() {
    let session = DeckSession::open(FIXTURE, 335).unwrap();
    let context = EditCtx::local("test");
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    session
        .set_shape_fill(&context, &slide.id, &slide.shapes[3].id, Some("#12B76A"))
        .unwrap();
    let xml = slide_xml(&session.save().unwrap());

    let wrapper = wrapper(&xml);
    assert!(wrapper.contains(r#"<bo:extensionShape name="unsupported"/>"#));
    assert!(wrapper.contains(r#"name="fallback""#));
    assert!(wrapper.contains(r#"<a:srgbClr val="315EFB"/>"#));
    assert!(!wrapper.contains("12B76A"), "the edit landed in the branch");
    let caption = xml.split(r#"name="caption""#).nth(1).unwrap();
    assert!(caption.contains(r#"<a:srgbClr val="12B76A"/>"#));
}

#[test]
fn deleting_the_only_shape_in_a_branch_removes_the_wrapper() {
    let session = DeckSession::open(FIXTURE, 334).unwrap();
    let context = EditCtx::local("test");
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    assert_eq!(slide.shapes[2].name, "fallback");
    session
        .remove_shape(&context, &slide.id, &slide.shapes[2].id)
        .unwrap();
    let xml = slide_xml(&session.save().unwrap());
    assert!(
        !xml.contains(r#"name="fallback""#),
        "the deleted shape came back with its wrapper"
    );
    assert!(!xml.contains("<mc:AlternateContent"), "{xml}");
}

const PNG: &[u8] = &[
    0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
    0x89, 0x00, 0x00, 0x00, 0x0D, 0x49, 0x44, 0x41, 0x54, 0x78, 0xDA, 0x63, 0xFC, 0xCF, 0xC0, 0xF0,
    0x1F, 0x00, 0x05, 0x05, 0x02, 0x00, 0x5F, 0xC8, 0xF1, 0xD2, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45,
    0x4E, 0x44, 0xAE, 0x42, 0x60, 0x82,
];

fn with_alternate_picture() -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
    for (path, bytes) in &mut parts {
        let text = String::from_utf8(bytes.clone()).unwrap_or_default();
        let replaced = match path.as_str() {
            "ppt/slides/slide1.xml" => text.replace(
                "</p:spTree>",
                r#"<p:pic><p:nvPicPr><p:cNvPr id="90" name="alternate picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:ma="http://schemas.microsoft.com/office/mac/drawingml/2008/main" Requires="ma"><p:blipFill><a:blip r:embed="rIdAltPdf"/><a:stretch><a:fillRect/></a:stretch></p:blipFill></mc:Choice><mc:Fallback><p:blipFill><a:blip r:embed="rIdAltPng"/><a:stretch><a:fillRect/></a:stretch></p:blipFill></mc:Fallback></mc:AlternateContent><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="900000" cy="900000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:spTree>"#,
            ),
            "ppt/slides/_rels/slide1.xml.rels" => text.replace(
                "</Relationships>",
                r#"<Relationship Id="rIdAltPdf" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/alternate.pdf"/><Relationship Id="rIdAltPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/alternate.png"/></Relationships>"#,
            ),
            "[Content_Types].xml" => text.replace(
                "<Default Extension=\"xml\"",
                r#"<Default Extension="png" ContentType="image/png"/><Default Extension="pdf" ContentType="application/pdf"/><Default Extension="xml""#,
            ),
            _ => continue,
        };
        *bytes = replaced.into_bytes();
    }
    parts.push(("ppt/media/alternate.png".to_owned(), PNG.to_vec()));
    parts.push((
        "ppt/media/alternate.pdf".to_owned(),
        b"%PDF-1.4\n%%EOF\n".to_vec(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn a_moved_and_resized_alternate_content_picture_keeps_both_branches() {
    let deck = with_alternate_picture();
    let session = DeckSession::open(&deck, 336).unwrap();
    let context = EditCtx::local("test");
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let picture = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "alternate picture")
        .unwrap();
    assert_eq!(
        picture.media_part_path.as_deref(),
        Some("ppt/media/alternate.png")
    );

    session
        .move_shape(&context, &slide.id, &picture.id, 1_000_000, 2_000_000)
        .unwrap();
    session
        .resize_shape(&context, &slide.id, &picture.id, 2_500_000, 1_500_000)
        .unwrap();
    let saved = session.save().unwrap();
    let xml = slide_xml(&saved);
    let pic = xml.split(r#"name="alternate picture""#).nth(1).unwrap();
    let pic = &pic[..pic.find("</p:pic>").unwrap()];
    let wrapper = wrapper(pic);
    assert!(wrapper.contains(r#"Requires="ma""#));
    assert!(wrapper.contains(r#"r:embed="rIdAltPdf""#));
    assert!(wrapper.contains(r#"r:embed="rIdAltPng""#));
    let properties = pic.split("<p:spPr>").nth(1).unwrap();
    assert!(properties.contains(r#"<a:off x="1000000" y="2000000"/>"#));
    assert!(properties.contains(r#"<a:ext cx="2500000" cy="1500000"/>"#));

    let reopened = DeckSession::open(&saved, 337).unwrap().snapshot().unwrap();
    let again = reopened.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "alternate picture")
        .unwrap();
    assert_eq!(
        again.media_part_path.as_deref(),
        Some("ppt/media/alternate.png")
    );
    assert_eq!(
        (again.x, again.y, again.width, again.height),
        (1_000_000, 2_000_000, 2_500_000, 1_500_000)
    );
}
