use std::collections::BTreeMap;
use std::sync::{Arc, Mutex};

use pptx_edit::{DeckSession, EditCtx, ShapeKind, TextStyle, TextStylePatch, UpdateOrigin};

const PNG: &[u8] = b"\x89\x50\x4e\x47\x0d\x0a\x1a\x0a\x00\x00\x00\x0d\x49\x48\x44\x52\x00\x00\x00\x01\x00\x00\x00\x01\x08\x04\x00\x00\x00\xb5\x1c\x0c\x02\x00\x00\x00\x0b\x49\x44\x41\x54\x78\xda\x63\x64\xf8\x0f\x00\x01\x05\x01\x01\x27\x18\xe3\x66\x00\x00\x00\x00\x49\x45\x4e\x44\xae\x42\x60\x82";

fn synthetic_deck() -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/authors.xml" ContentType="application/vnd.ms-powerpoint.authors+xml"/><Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.ms-powerpoint.comments+xml"/></Types>"#,
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#,
        ),
        (
            "ppt/presentation.xml",
            r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#,
        ),
        (
            "ppt/_rels/presentation.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" Target="authors.xml"/></Relationships>"#,
        ),
        (
            "ppt/slides/slide1.xml",
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Styled"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="4000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1" sz="2400"><a:solidFill><a:srgbClr val="123456"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>alpha </a:t></a:r><a:r><a:rPr i="1" u="sng" sz="1800"/><a:t>beta</a:t></a:r></a:p></p:txBody></p:sp>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="100000" y="2000000"/><a:ext cx="2000000" cy="300000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="2000000"/></a:tblGrid><a:tr h="300000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>Cell</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
<p:pic><p:nvPicPr><p:cNvPr id="4" name="Image"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="5000000" y="100000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>
</p:spTree></p:cSld></p:sld>"#,
        ),
        (
            "ppt/slides/_rels/slide1.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../comments/comment1.xml"/></Relationships>"#,
        ),
        (
            "ppt/authors.xml",
            r#"<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:author id="{11111111-1111-1111-1111-111111111111}" name="Author" initials="A" userId="" providerId=""/></p188:authorLst>"#,
        ),
        (
            "ppt/comments/comment1.xml",
            r#"<p188:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command" xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main"><p188:cm id="{22222222-2222-2222-2222-222222222222}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-01-01T00:00:00Z"><pc:sldMkLst><pc:docMk/><pc:sldMk cId="0" sldId="256"/></pc:sldMkLst><p188:pos x="100" y="200"/><p188:replyLst><p188:reply id="{33333333-3333-3333-3333-333333333333}" authorId="{11111111-1111-1111-1111-111111111111}" created="2026-01-01T00:01:00Z"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Reply</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Root</a:t></a:r></a:p></p188:txBody></p188:cm></p188:cmLst>"#,
        ),
    ];
    let mut parts: Vec<(String, Vec<u8>)> = parts
        .into_iter()
        .map(|(path, xml)| (path.to_owned(), xml.as_bytes().to_vec()))
        .collect();
    parts.push(("ppt/media/image1.png".to_owned(), PNG.to_vec()));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

#[test]
fn source_seed_and_update_baselines_replay_identically_across_live_clients() {
    let bytes = synthetic_deck();
    let author = DeckSession::open(&bytes, 101).unwrap();
    let source = DeckSession::open(&bytes, 202).unwrap();
    let baseline = author.encode_state_as_update_v1();
    assert_eq!(source.encode_state_as_update_v1(), baseline);
    assert_eq!(
        source.encode_state_vector_v1(),
        author.encode_state_vector_v1()
    );
    let hydrated = DeckSession::open_from_update_with_source(&baseline, &bytes, 303).unwrap();
    assert_eq!(hydrated.encode_state_as_update_v1(), baseline);
    let snapshot = author.snapshot().unwrap();
    assert_eq!(snapshot.comments.len(), 2);
    assert!(
        snapshot
            .comments
            .iter()
            .any(|comment| comment.parent_id.is_some())
    );
    assert_eq!(author.package().media.len(), 1);
    assert_eq!(snapshot.slides[0].shapes.len(), 3);
    assert!(
        snapshot.slides[0]
            .shapes
            .iter()
            .any(|shape| shape.kind == ShapeKind::GraphicFrame)
    );
    let slide = &snapshot.slides[0];
    let shape = &slide.shapes[0];
    let story = &shape.text_stories[0];
    assert_eq!(story.paragraphs[0].runs[0].style.bold, Some(true));
    assert_eq!(story.paragraphs[0].runs[1].style.italic, Some(true));
    assert!(
        slide
            .shapes
            .iter()
            .any(|shape| shape.kind == ShapeKind::Picture)
    );
    let story = &story.id;
    let updates = Arc::new(Mutex::new(Vec::new()));
    let captured = Arc::clone(&updates);
    let _subscription = author
        .observe_update_v1(move |event| {
            if event.origin == UpdateOrigin::Local {
                captured.lock().unwrap().push(event.update);
            }
        })
        .unwrap();
    let context = EditCtx::local("test");
    author
        .insert_text(&context, story, 0, "typed ", &TextStyle::default())
        .unwrap();
    author.delete_text(&context, story, 1, 3).unwrap();
    author
        .format_text(
            &context,
            story,
            0,
            4,
            &TextStylePatch {
                bold: Some(true),
                italic: Some(true),
                ..TextStylePatch::default()
            },
        )
        .unwrap();
    author
        .move_shape(&context, &slide.id, &shape.id, 300_000, 400_000)
        .unwrap();
    assert!(updates.lock().unwrap().len() >= 4);
    for update in updates.lock().unwrap().iter() {
        source.apply_update_v1(update).unwrap();
        hydrated.apply_update_v1(update).unwrap();
    }
    let expected = parts(&author.save().unwrap());
    for recovered in [&source, &hydrated] {
        assert_eq!(
            recovered.encode_state_as_update_v1(),
            author.encode_state_as_update_v1()
        );
        assert_eq!(
            recovered.encode_state_vector_v1(),
            author.encode_state_vector_v1()
        );
        assert_eq!(recovered.snapshot().unwrap(), author.snapshot().unwrap());
        assert_eq!(parts(&recovered.save().unwrap()), expected);
    }
}
