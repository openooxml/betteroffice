use pptx_edit::{
    CommentFlavor, DeckSession, EditCtx, ProposalEdit, ProposalError, ProposalRequest,
};

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/custom.xml" ContentType="application/vnd.openxmlformats-officedocument.custom-properties+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" Target="docProps/custom.xml"/>
</Relationships>"#;

const CUSTOM: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/custom-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><property fmtid="{D5CDD505-2E9C-101B-9397-08002B2CF9AE}" pid="2" name="Classification"><vt:lpwstr>Internal</vt:lpwstr></property></Properties>"#;

const PRESENTATION: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"#;

const PRESENTATION_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>"#;

const SLIDE1: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="6000000" cy="1200000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:r><a:rPr lang="en-US" sz="3200"/><a:t>Quarterly review</a:t></a:r></a:p>
<a:p><a:r><a:rPr lang="en-US"/><a:t>Revenue</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br><a:r><a:rPr lang="en-US"/><a:t>grew 😀 fast</a:t></a:r></a:p>
</p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Footer"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="100000" y="6000000"/><a:ext cx="4000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Slide </a:t></a:r><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>12</a:t></a:fld><a:r><a:rPr lang="en-US"/><a:t> of deck</a:t></a:r></a:p></p:txBody></p:sp>
<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="100000" y="4000000"/><a:ext cx="2000000" cy="500000"/><a:chOff x="100000" y="4000000"/><a:chExt cx="2000000" cy="500000"/></a:xfrm></p:grpSpPr>
<p:sp><p:nvSpPr><p:cNvPr id="11" name="Nested"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="100000" y="4000000"/><a:ext cx="2000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Nested text</a:t></a:r></a:p></p:txBody></p:sp>
</p:grpSp>
<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="20" name="Table"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>
<p:xfrm><a:off x="3000000" y="3000000"/><a:ext cx="4000000" cy="600000"/></p:xfrm>
<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
<a:tr h="300000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Cell A</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Cell B</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>
</a:tbl></a:graphicData></a:graphic></p:graphicFrame>
<p:sp><p:nvSpPr><p:cNvPr id="30" name="Card"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="7000000" y="1000000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Card</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="40" name="Counter"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="7000000" y="3000000"/><a:ext cx="2000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>X12</a:t></a:r><a:fld id="{7A1E6B2C-3D4F-4E5A-8B9C-0D1E2F3A4B5C}" type="slidenum"><a:rPr lang="en-US"/><a:t>12</a:t></a:fld><a:r><a:rPr lang="en-US"/><a:t>12Y</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="50" name="Stamp"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="7000000" y="4000000"/><a:ext cx="2000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>A</a:t></a:r><a:fld id="{0F1E2D3C-4B5A-4978-8695-A4B3C2D1E0F9}" type="datetime1"><a:rPr lang="en-US"/><a:t></a:t></a:fld><a:r><a:rPr lang="en-US"/><a:t>B</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>"#;

const SLIDE_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>"#;

const SLIDE2: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="3000000" cy="400000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Second slide</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>"#;

const SLIDE2_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>"#;

const NOTES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Speaker notes</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:notes>"#;

const NOTES_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide2.xml"/>
</Relationships>"#;

const LAYOUT: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sldLayout>"#;

const LAYOUT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>"#;

const MASTER: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst></p:sldMaster>"#;

const MASTER_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>"#;

const THEME: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Fixture">
<a:themeElements><a:clrScheme name="Fixture">
<a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="222222"/></a:dk2><a:lt2><a:srgbClr val="EEEEEE"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme></a:themeElements></a:theme>"#;

fn fixture() -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = [
        ("[Content_Types].xml", CONTENT_TYPES),
        ("_rels/.rels", ROOT_RELS),
        ("docProps/custom.xml", CUSTOM),
        ("ppt/presentation.xml", PRESENTATION),
        ("ppt/_rels/presentation.xml.rels", PRESENTATION_RELS),
        ("ppt/slides/slide1.xml", SLIDE1),
        ("ppt/slides/_rels/slide1.xml.rels", SLIDE_RELS),
        ("ppt/slides/slide2.xml", SLIDE2),
        ("ppt/slides/_rels/slide2.xml.rels", SLIDE2_RELS),
        ("ppt/notesSlides/notesSlide1.xml", NOTES),
        ("ppt/notesSlides/_rels/notesSlide1.xml.rels", NOTES_RELS),
        ("ppt/slideLayouts/slideLayout1.xml", LAYOUT),
        ("ppt/slideLayouts/_rels/slideLayout1.xml.rels", LAYOUT_RELS),
        ("ppt/slideMasters/slideMaster1.xml", MASTER),
        ("ppt/slideMasters/_rels/slideMaster1.xml.rels", MASTER_RELS),
        ("ppt/theme/theme1.xml", THEME),
    ]
    .into_iter()
    .map(|(path, body)| (path.to_owned(), body.as_bytes().to_vec()))
    .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open() -> DeckSession {
    DeckSession::open(&fixture(), 71).unwrap()
}

fn request(edits: Vec<ProposalEdit>) -> ProposalRequest {
    ProposalRequest {
        agent_id: "editor-agent".into(),
        note: Some("Tighten the title and align it".into()),
        edits,
    }
}

fn stage_delete(session: &DeckSession, slide_id: &str) -> String {
    let edit = serde_json::from_value(serde_json::json!({
        "type": "deleteSlide",
        "slideId": slide_id,
    }))
    .unwrap();
    session.propose(request(vec![edit])).unwrap().id
}

#[test]
fn text_edit_after_staging_makes_delete_slide_stale() {
    let session = open();
    let slide = &session.snapshot().unwrap().slides[1];
    let story = slide
        .shapes
        .iter()
        .find_map(|shape| shape.text_stories.first())
        .unwrap();
    let proposal_id = stage_delete(&session, &slide.id);
    session
        .insert_text(
            &EditCtx::local("human"),
            &story.id,
            0,
            "Revised: ",
            &Default::default(),
        )
        .unwrap();

    let result = session.accept_proposal(&proposal_id, false);
    let after = session.snapshot().unwrap();
    let kept = after
        .slides
        .get(1)
        .expect("deleteSlide removed the edited slide");
    assert_eq!(kept.id, slide.id);
    let kept_story = kept
        .shapes
        .iter()
        .flat_map(|shape| &shape.text_stories)
        .find(|candidate| candidate.id == story.id)
        .expect("the edited shape's story should survive");
    assert_eq!(kept_story.plain_text(), "Revised: Second slide");
    assert!(matches!(result, Err(ProposalError::Stale(ids)) if ids == [slide.id.clone()]));
}

#[test]
fn comment_added_after_staging_blocks_delete_slide_accept() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[1].id.clone();
    let proposal_id = stage_delete(&session, &slide_id);
    let text = "Please keep this slide.";
    let added = session
        .add_comment(
            &EditCtx::local("human"),
            &slide_id,
            "Ada",
            "AL",
            text,
            "2026-09-05T10:00:00Z",
            0,
            0,
        )
        .unwrap();

    let result = session.accept_proposal(&proposal_id, false);
    let after = session.snapshot().unwrap();
    let kept = after
        .slides
        .get(1)
        .expect("deleteSlide removed the commented slide");
    assert_eq!(kept.id, slide_id);
    let comment = session
        .comments()
        .unwrap()
        .into_iter()
        .find(|comment| comment.id == added.comment_id)
        .expect("the added comment should survive");
    assert_eq!(comment.slide_id, slide_id);
    assert_eq!(comment.text, text);
    assert!(matches!(result, Err(ProposalError::Stale(ids)) if ids == [slide_id.clone()]));
}

#[test]
fn notes_edit_after_staging_blocks_delete_slide_accept() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[1].id.clone();
    let proposal_id = stage_delete(&session, &slide_id);
    let notes = "Keep the revised speaker notes.";
    session
        .set_slide_notes(&EditCtx::local("human"), &slide_id, notes)
        .unwrap();

    let result = session.accept_proposal(&proposal_id, false);
    let after = session.snapshot().unwrap();
    let kept = after
        .slides
        .get(1)
        .expect("deleteSlide removed the slide with edited notes");
    assert_eq!(kept.id, slide_id);
    assert_eq!(kept.notes, notes);
    assert!(matches!(result, Err(ProposalError::Stale(ids)) if ids == [slide_id.clone()]));
}

fn commented(flavor: CommentFlavor) -> (DeckSession, String, String) {
    let session = open();
    let ctx = EditCtx::local("human");
    session.set_comment_flavor(&ctx, flavor).unwrap();
    let slide_id = session.snapshot().unwrap().slides[1].id.clone();
    let comment = session
        .add_comment(
            &ctx,
            &slide_id,
            "Ada",
            "AL",
            "Check the figures.",
            "2026-09-05T10:00:00Z",
            0,
            0,
        )
        .unwrap();
    (session, slide_id, comment.comment_id)
}

#[test]
fn reply_added_after_staging_blocks_delete_slide_accept() {
    let (session, slide_id, comment_id) = commented(CommentFlavor::Modern);
    let proposal_id = stage_delete(&session, &slide_id);
    let reply = session
        .reply_to_comment(
            &EditCtx::local("human"),
            &comment_id,
            "Bo",
            "BO",
            "Figures checked.",
            "2026-09-05T11:00:00Z",
        )
        .unwrap();

    let result = session.accept_proposal(&proposal_id, false);
    assert!(matches!(result, Err(ProposalError::Stale(ids)) if ids == [slide_id.clone()]));
    assert_eq!(session.snapshot().unwrap().slides[1].id, slide_id);
    assert!(
        session
            .comments()
            .unwrap()
            .iter()
            .any(|comment| comment.id == reply.comment_id)
    );
}

#[test]
fn comment_resolved_after_staging_blocks_delete_slide_accept() {
    let (session, slide_id, comment_id) = commented(CommentFlavor::Modern);
    let proposal_id = stage_delete(&session, &slide_id);
    session
        .set_comment_status(&EditCtx::local("human"), &comment_id, true)
        .unwrap();

    let result = session.accept_proposal(&proposal_id, false);
    assert!(matches!(result, Err(ProposalError::Stale(ids)) if ids == [slide_id.clone()]));
    assert_eq!(session.snapshot().unwrap().slides[1].id, slide_id);
    assert!(
        session
            .comments()
            .unwrap()
            .iter()
            .any(|comment| comment.id == comment_id && comment.resolved)
    );
}

#[test]
fn unchanged_comments_keep_delete_slide_acceptable() {
    let (session, slide_id, _) = commented(CommentFlavor::Legacy);
    let proposal_id = stage_delete(&session, &slide_id);
    let first = session.snapshot().unwrap().slides[0].id.clone();
    session
        .add_comment(
            &EditCtx::local("human"),
            &first,
            "Ada",
            "AL",
            "Other slide.",
            "2026-09-05T10:00:00Z",
            0,
            0,
        )
        .unwrap();

    session.accept_proposal(&proposal_id, false).unwrap();
    assert!(
        session
            .snapshot()
            .unwrap()
            .slides
            .iter()
            .all(|slide| slide.id != slide_id)
    );
}

#[test]
fn force_accept_deletes_a_slide_commented_after_staging() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[1].id.clone();
    let proposal_id = stage_delete(&session, &slide_id);
    let late = session
        .add_comment(
            &EditCtx::local("human"),
            &slide_id,
            "Ada",
            "AL",
            "Late note.",
            "2026-09-05T10:00:00Z",
            0,
            0,
        )
        .unwrap();

    assert!(matches!(
        session.accept_proposal(&proposal_id, false),
        Err(ProposalError::Stale(ids)) if ids == [slide_id.clone()]
    ));
    assert!(session.accept_proposal(&proposal_id, true).unwrap().applied);
    assert!(
        session
            .snapshot()
            .unwrap()
            .slides
            .iter()
            .all(|slide| slide.id != slide_id)
    );
    assert!(
        session
            .comments()
            .unwrap()
            .iter()
            .all(|comment| comment.id != late.comment_id)
    );
}

#[test]
fn a_live_slide_sharing_a_proposed_slide_id_does_not_make_the_proposal_stale() {
    let session = open();
    let edits = serde_json::from_value(serde_json::json!([
        {"type": "insertSlide", "ref": "draft", "afterSlideId": null},
        {"type": "deleteSlide", "slideId": "@draft"}
    ]))
    .unwrap();
    let proposal = session.propose(request(edits)).unwrap();
    let ctx = EditCtx::local("human");
    let live = session.insert_slide(&ctx, 0, None).unwrap().slide_id;
    assert_eq!(live, proposal.refs["draft"]);
    let comment = session
        .add_comment(
            &ctx,
            &live,
            "Ada",
            "AL",
            "On the live slide.",
            "2026-09-05T10:00:00Z",
            0,
            0,
        )
        .unwrap();

    session.accept_proposal(&proposal.id, false).unwrap();
    assert!(
        session
            .snapshot()
            .unwrap()
            .slides
            .iter()
            .any(|slide| slide.id == live)
    );
    assert!(
        session
            .comments()
            .unwrap()
            .iter()
            .any(|candidate| candidate.id == comment.comment_id)
    );
}
