//! Version-checked edit batches against a synthetic deck carrying a field, a soft line break, a
//! group, a table, speaker notes and an unrelated custom part.

use std::cell::RefCell;
use std::collections::BTreeMap;
use std::rc::Rc;
use std::sync::Arc;
use std::sync::atomic::{AtomicUsize, Ordering};

use pptx_edit::{
    DeckSession, DocumentVersion, EditCtx, EditFailure, EditFailureCode, EditHistory, EditOutcome,
    EditRequest, EditSource, EditTarget, FindRequest, FindScope, ProposalEdit, ProposalError,
    ProposalRequest, ReadRequest, ReadResponse, ShapeDraft, ShapeRect, ShapeSnapshot, StoryText,
    TextField, TextRange, TextStyle, UndoCaptureMode,
};
use serde_json::{Value, json};

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

fn read(session: &DeckSession) -> ReadResponse {
    session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap()
}

fn named<'a>(shapes: &'a [ShapeSnapshot], name: &str) -> Option<&'a ShapeSnapshot> {
    shapes.iter().find_map(|shape| {
        if shape.name == name {
            Some(shape)
        } else {
            named(&shape.children, name)
        }
    })
}

fn shape<'a>(read: &'a ReadResponse, name: &str) -> &'a ShapeSnapshot {
    read.slides
        .iter()
        .find_map(|slide| named(&slide.shapes, name))
        .unwrap()
}

fn story(read: &ReadResponse, name: &str) -> StoryText {
    let shape_id = &shape(read, name).id;
    read.stories
        .iter()
        .find(|story| &story.shape_id == shape_id)
        .unwrap()
        .clone()
}

fn range(story: &StoryText, start: u32, end: u32) -> Value {
    json!({
        "kind": "range", "slideId": story.slide_id, "shapeId": story.shape_id,
        "storyId": story.story_id, "start": start, "end": end,
    })
}

fn search(story: &StoryText, text: &str) -> Value {
    json!({
        "kind": "search", "text": text,
        "within": {"slideId": story.slide_id, "shapeId": story.shape_id, "storyId": story.story_id},
    })
}

fn shape_target(read: &ReadResponse, name: &str) -> Value {
    let slide = read
        .slides
        .iter()
        .find(|slide| named(&slide.shapes, name).is_some())
        .unwrap();
    json!({"slideId": slide.id, "shapeId": shape(read, name).id})
}

fn request(version: &DocumentVersion, body: Value) -> EditRequest {
    let mut value = json!({"expectVersion": version});
    value
        .as_object_mut()
        .unwrap()
        .extend(body.as_object().unwrap().clone());
    serde_json::from_value(value).unwrap()
}

/// Applies `steps` against the session's current version.
fn edit(session: &DeckSession, steps: Value) -> EditOutcome {
    session
        .apply_edits(&request(&session.version(), json!({"steps": steps})))
        .unwrap()
}

fn refused(outcome: EditOutcome) -> EditFailure {
    outcome.unwrap_err().failure
}

fn text_of(session: &DeckSession, story: &StoryText) -> String {
    session.story(&story.story_id).unwrap().plain_text()
}

fn range_of(target: &EditTarget) -> (u32, u32) {
    match target {
        EditTarget::Range(TextRange { start, end, .. }) => (*start, *end),
        other => panic!("not a range: {other:?}"),
    }
}

#[test]
fn reads_project_stories_with_separators_breaks_and_fields() {
    let session = open();
    let read = read(&session);
    assert_eq!(read.version, session.version());
    let title = story(&read, "Title");
    assert_eq!(title.text, "Quarterly review\nRevenue\ngrew 😀 fast");
    assert_eq!(
        title
            .paragraphs
            .iter()
            .map(|paragraph| (paragraph.start, paragraph.end))
            .collect::<Vec<_>>(),
        [(0, 16), (17, 37)]
    );
    assert_eq!(title.paragraphs[1].line_breaks, [24]);
    let footer = story(&read, "Footer");
    assert_eq!(
        footer.paragraphs[0].fields,
        [TextField {
            start: 6,
            end: 8,
            field_type: Some("slidenum".to_owned()),
        }]
    );
    assert!(footer.paragraphs[0].editable);
    assert_eq!(story(&read, "Nested").text, "Nested text");
    let cells: Vec<&str> = read
        .stories
        .iter()
        .filter(|story| story.story_id.contains(":table:"))
        .map(|story| story.text.as_str())
        .collect();
    assert_eq!(cells, ["Cell A", "Cell B"]);

    let second = read.slides[1].id.clone();
    let scoped = session
        .read_content(&ReadRequest {
            slide_ids: Some(vec![second.clone()]),
        })
        .unwrap()
        .unwrap();
    assert_eq!(scoped.slides.len(), 1);
    assert!(scoped.stories.iter().all(|story| story.slide_id == second));
    let missing = session
        .read_content(&ReadRequest {
            slide_ids: Some(vec!["slide:9:999".to_owned()]),
        })
        .unwrap()
        .unwrap_err();
    assert_eq!(missing.failure.code, EditFailureCode::MissingTarget);
    assert_eq!(missing.version, session.version());
}

#[test]
fn a_batch_commits_once_or_refuses_without_changing_anything() {
    let session = open();
    let events = Arc::new(AtomicUsize::new(0));
    let counter = Arc::clone(&events);
    let _observer = session
        .observe_update_v1(move |_| {
            counter.fetch_add(1, Ordering::Relaxed);
        })
        .unwrap();
    let read = read(&session);
    let title = story(&read, "Title");
    let before = session.encode_state_as_update_v1();
    let refusal = session
        .apply_edits(&request(
            &read.version,
            json!({"steps": [
                {"op": "replaceText", "target": search(&title, "Quarterly"), "text": "Annual"},
                {"op": "deleteText", "target": search(&title, "absent")},
            ]}),
        ))
        .unwrap()
        .unwrap_err();
    assert_eq!(refusal.failure.code, EditFailureCode::MissingTarget);
    assert_eq!(refusal.failure.step_index, Some(1));
    assert_eq!(refusal.version, read.version);
    assert_eq!(session.encode_state_as_update_v1(), before);
    assert_eq!(events.load(Ordering::Relaxed), 0);
    assert!(!session.can_undo());

    let applied = session
        .apply_edits(&request(
            &read.version,
            json!({"steps": [
                {"op": "replaceText", "target": search(&title, "Quarterly"), "text": "Annual",
                 "expect": {"text": "Quarterly"}},
                {"op": "insertText", "target": range(&title, 37, 37), "at": "end", "text": " (draft)"},
            ]}),
        ))
        .unwrap()
        .unwrap();
    assert!(applied.applied);
    assert_eq!(applied.base_version, read.version);
    assert_eq!(applied.version, session.version());
    assert_ne!(applied.version, read.version);
    assert_eq!(applied.changed_stories, [title.story_id.as_str()]);
    assert_eq!(applied.changed_slides, [title.slide_id.as_str()]);
    assert_eq!(range_of(&applied.receipts[0].target), (0, 6));
    assert_eq!(range_of(&applied.receipts[1].target), (34, 42));
    assert_eq!(events.load(Ordering::Relaxed), 1);
    assert_eq!(
        text_of(&session, &title),
        "Annual review\nRevenue\ngrew 😀 fast (draft)"
    );
    assert!(session.undo());
    assert_eq!(text_of(&session, &title), title.text);
    assert!(!session.can_undo());
}

#[test]
fn every_target_resolves_against_the_pre_batch_state() {
    let session = open();
    let title = story(&read(&session), "Title");
    let applied = edit(
        &session,
        json!([
            {"op": "replaceText", "target": range(&title, 10, 16), "text": "recap"},
            {"op": "replaceText", "target": range(&title, 0, 9), "text": "Q"},
            {"op": "deleteText", "target": range(&title, 17, 24)},
            {"op": "formatText", "target": range(&title, 25, 29), "patch": {"bold": true}},
        ]),
    )
    .unwrap();
    let text = text_of(&session, &title);
    assert_eq!(text, "Q recap\n\ngrew 😀 fast");
    let spans: Vec<(u32, u32)> = applied
        .receipts
        .iter()
        .map(|receipt| range_of(&receipt.target))
        .collect();
    assert_eq!(spans, [(2, 7), (0, 1), (8, 8), (9, 13)]);
    let units: Vec<u16> = text.encode_utf16().collect();
    assert_eq!(String::from_utf16(&units[9..13]).unwrap(), "grew");
    let runs = &session.story(&title.story_id).unwrap().paragraphs[1].runs;
    assert!(
        runs.iter()
            .any(|run| run.text == "grew" && run.style.bold == Some(true))
    );
}

#[test]
fn surrogates_and_overlapping_occurrences_are_never_split_or_guessed() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    session
        .add_text_box(
            &EditCtx::local("test"),
            &slide_id,
            &ShapeDraft {
                name: "Echo".into(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 1_000_000,
                    height: 500_000,
                },
                text: "aaa".into(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    let read = read(&session);
    let (title, echo) = (story(&read, "Title"), story(&read, "Echo"));
    let split = refused(edit(
        &session,
        json!([{"op": "deleteText", "target": range(&title, 31, 32)}]),
    ));
    assert_eq!(split.code, EditFailureCode::InvalidStep);
    let ambiguous = refused(edit(
        &session,
        json!([{"op": "replaceText", "target": search(&echo, "aa"), "text": "b"}]),
    ));
    assert_eq!(ambiguous.code, EditFailureCode::AmbiguousTarget);
    assert_eq!(session.search_text("aa", true, None).unwrap().len(), 1);

    let find = |within: Option<FindScope>, limit| {
        session
            .find_text(&FindRequest {
                text: "aa".into(),
                within,
                limit,
            })
            .unwrap()
            .unwrap()
    };
    let first = find(None, Some(1));
    assert!(first.truncated);
    assert_eq!(first.matches.len(), 1);
    let scoped = find(
        Some(FindScope {
            slide_id: slide_id.clone(),
            shape_id: Some(echo.shape_id.clone()),
            story_id: None,
        }),
        None,
    );
    assert!(!scoped.truncated);
    assert_eq!(
        scoped
            .matches
            .iter()
            .map(|found| (found.range.start, found.range.end))
            .collect::<Vec<_>>(),
        [(0, 2), (1, 3)]
    );
    let across = session
        .find_text(&FindRequest {
            text: "review\nRevenue".into(),
            within: None,
            limit: None,
        })
        .unwrap()
        .unwrap();
    assert!(across.matches.is_empty());
    let orphan = session
        .find_text(&FindRequest {
            text: "a".into(),
            within: Some(FindScope {
                slide_id,
                shape_id: None,
                story_id: Some(echo.story_id.clone()),
            }),
            limit: None,
        })
        .unwrap()
        .unwrap_err();
    assert_eq!(orphan.failure.code, EditFailureCode::InvalidStep);

    edit(
        &session,
        json!([{"op": "replaceText", "target": search(&title, "😀"), "text": "🚀"}]),
    )
    .unwrap();
    assert_eq!(
        text_of(&session, &title),
        "Quarterly review\nRevenue\ngrew 🚀 fast"
    );
}

#[test]
fn ownership_chains_resolve_exactly() {
    let session = open();
    let read = read(&session);
    let (title, footer, nested) = (
        story(&read, "Title"),
        story(&read, "Footer"),
        story(&read, "Nested"),
    );
    let cell = read
        .stories
        .iter()
        .find(|story| story.text == "Cell B")
        .unwrap()
        .clone();
    let point = |slide: &str, shape: &str, story: &str| json!({"kind": "range", "slideId": slide, "shapeId": shape, "storyId": story, "start": 0, "end": 0});
    let second = read.slides[1].id.as_str();
    for (target, code) in [
        (
            point(&title.slide_id, &footer.shape_id, &title.story_id),
            EditFailureCode::InvalidStep,
        ),
        (
            point(second, &title.shape_id, &title.story_id),
            EditFailureCode::InvalidStep,
        ),
        (
            point(&title.slide_id, &title.shape_id, "story:missing"),
            EditFailureCode::MissingTarget,
        ),
        (
            point(&title.slide_id, "shape:missing", &title.story_id),
            EditFailureCode::MissingTarget,
        ),
        (
            point("slide:missing", &title.shape_id, &title.story_id),
            EditFailureCode::MissingTarget,
        ),
    ] {
        let failure = refused(edit(
            &session,
            json!([{"op": "insertText", "target": target, "at": "start", "text": "x"}]),
        ));
        assert_eq!(failure.code, code, "{failure:?}");
        assert!(failure.target.is_some());
    }

    let applied = edit(
        &session,
        json!([
            {"op": "insertText", "target": search(&nested, "text"), "at": "start", "text": "inner "},
            {"op": "replaceText", "target": search(&cell, "B"), "text": "Beta"},
        ]),
    )
    .unwrap();
    assert_eq!(text_of(&session, &nested), "Nested inner text");
    assert_eq!(text_of(&session, &cell), "Cell Beta");
    assert_eq!(applied.changed_stories.len(), 2);
    assert!(applied.changed_stories.contains(&cell.story_id));

    let nested_rect = refused(edit(
        &session,
        json!([{"op": "setShapeRect", "target": {"slideId": nested.slide_id, "shapeId": nested.shape_id},
                "rect": {"x": 0, "y": 0, "width": 10, "height": 10}}]),
    ));
    assert_eq!(nested_rect.code, EditFailureCode::Unsupported);
}

#[test]
fn fields_and_line_breaks_refuse_text_steps_that_would_lose_them() {
    let session = open();
    let read = read(&session);
    let (title, footer) = (story(&read, "Title"), story(&read, "Footer"));
    for step in [
        json!({"op": "insertText", "target": range(&footer, 7, 7), "at": "start", "text": "x"}),
        json!({"op": "replaceText", "target": range(&footer, 5, 7), "text": "x"}),
        json!({"op": "deleteText", "target": range(&footer, 7, 10)}),
        json!({"op": "deleteText", "target": range(&title, 23, 26)}),
        json!({"op": "formatText", "target": range(&footer, 7, 10), "patch": {"bold": true}}),
        json!({"op": "replaceText", "target": range(&title, 15, 18), "text": "x"}),
    ] {
        let failure = refused(edit(&session, json!([step])));
        assert_eq!(failure.code, EditFailureCode::Unsupported, "{failure:?}");
    }
    for text in ["a\nb", "a\rb", "a\u{2028}b"] {
        let failure = refused(edit(
            &session,
            json!([{"op": "insertText", "target": range(&title, 0, 0), "at": "start", "text": text}]),
        ));
        assert_eq!(failure.code, EditFailureCode::InvalidStep);
    }

    edit(
        &session,
        json!([{"op": "formatText", "target": range(&footer, 6, 8), "patch": {"italic": true}}]),
    )
    .unwrap();
    edit(
        &session,
        json!([{"op": "insertText", "target": range(&footer, 8, 8), "at": "start", "text": "!"}]),
    )
    .unwrap();
    let footer_now = story(&self::read(&session), "Footer");
    assert_eq!(footer_now.text, "Slide 12! of deck");
    assert!(footer_now.paragraphs[0].editable);
    assert_eq!(
        (
            footer_now.paragraphs[0].fields[0].start,
            footer_now.paragraphs[0].fields[0].end
        ),
        (6, 8)
    );

    let saved = session.save().unwrap();
    let parts: BTreeMap<String, Vec<u8>> = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .collect();
    let slide = String::from_utf8(parts["ppt/slides/slide1.xml"].clone()).unwrap();
    assert!(slide.contains(r#"type="slidenum""#), "{slide}");
    assert!(slide.contains("<a:t>12</a:t></a:fld>"), "{slide}");
    let reopened = DeckSession::open(&saved, 72).unwrap();
    assert_eq!(
        story(&self::read(&reopened), "Footer").text,
        "Slide 12! of deck"
    );
}

#[test]
fn fields_stay_fields_through_save_or_the_batch_refuses() {
    let session = open();
    let counter = story(&read(&session), "Counter");
    assert_eq!(counter.text, "X121212Y");
    assert_eq!(
        (
            counter.paragraphs[0].fields[0].start,
            counter.paragraphs[0].fields[0].end
        ),
        (3, 5)
    );
    let before = session.encode_state_as_update_v1();
    let lost = refused(edit(
        &session,
        json!([
            {"op": "deleteText", "target": range(&counter, 0, 3)},
            {"op": "deleteText", "target": range(&counter, 5, 8)},
        ]),
    ));
    assert_eq!(lost.code, EditFailureCode::Unsupported);
    assert!(lost.message.contains("plain text"), "{lost:?}");
    assert_eq!(lost.step_index, Some(0));
    assert_eq!(session.encode_state_as_update_v1(), before);

    let kept = edit(
        &session,
        json!([{"op": "deleteText", "target": range(&counter, 0, 1)}]),
    )
    .unwrap();
    assert!(kept.applied);
    let saved = session.save().unwrap();
    let parts: BTreeMap<String, Vec<u8>> = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .collect();
    let slide = String::from_utf8(parts["ppt/slides/slide1.xml"].clone()).unwrap();
    assert!(slide.contains("<a:t>12</a:t></a:fld><a:r>"), "{slide}");
    let reopened = DeckSession::open(&saved, 77).unwrap();
    let counter = story(&self::read(&reopened), "Counter");
    assert_eq!(counter.text, "121212Y");
    assert_eq!(
        (
            counter.paragraphs[0].fields[0].start,
            counter.paragraphs[0].fields[0].end
        ),
        (2, 4)
    );

    let context = EditCtx::local("human");
    let style = TextStyle::default();
    session
        .insert_text(&context, &counter.story_id, 0, "A", &style)
        .unwrap();
    session
        .insert_text(&context, &counter.story_id, 8, "B", &style)
        .unwrap();
    let stranded = story(&self::read(&session), "Counter");
    assert!(!stranded.paragraphs[0].editable);
    assert!(stranded.paragraphs[0].fields.is_empty());
    let refusal = refused(edit(
        &session,
        json!([{"op": "insertText", "target": range(&stranded, 1, 1), "at": "start", "text": "x"}]),
    ));
    assert_eq!(refusal.code, EditFailureCode::Unsupported);
}

#[test]
fn paragraphs_with_empty_field_results_refuse_changes_and_keep_the_field() {
    let session = open();
    let read = read(&session);
    let stamp = story(&read, "Stamp");
    assert_eq!(stamp.text, "AB");
    assert_eq!(
        stamp.paragraphs[0].fields,
        [TextField {
            start: 1,
            end: 1,
            field_type: Some("datetime1".to_owned()),
        }]
    );
    assert!(!stamp.paragraphs[0].editable);
    let before = session.encode_state_as_update_v1();
    for step in [
        json!({"op": "setParagraphAlignment", "target": range(&stamp, 0, 0), "alignment": "ctr"}),
        json!({"op": "insertText", "target": range(&stamp, 2, 2), "at": "end", "text": "C"}),
        json!({"op": "formatText", "target": range(&stamp, 0, 1), "patch": {"bold": true}}),
    ] {
        let failure = refused(edit(&session, json!([step])));
        assert_eq!(failure.code, EditFailureCode::Unsupported, "{failure:?}");
    }
    assert_eq!(session.encode_state_as_update_v1(), before);

    let title = story(&read, "Title");
    edit(
        &session,
        json!([{"op": "replaceText", "target": search(&title, "Quarterly"), "text": "Annual"}]),
    )
    .unwrap();
    let saved = session.save().unwrap();
    let parts: BTreeMap<String, Vec<u8>> = ooxml_opc::unzip_parts(&saved)
        .unwrap()
        .into_iter()
        .collect();
    let slide = String::from_utf8(parts["ppt/slides/slide1.xml"].clone()).unwrap();
    assert!(slide.contains(r#"type="datetime1""#), "{slide}");
    let reopened = DeckSession::open(&saved, 78).unwrap();
    let stamp = story(&self::read(&reopened), "Stamp");
    assert_eq!(stamp.text, "AB");
    assert_eq!(stamp.paragraphs[0].fields.len(), 1);
}

#[test]
fn refusals_stay_small_whatever_the_text_they_describe() {
    use pptx_edit::{MAX_REQUEST_BYTES, outcome_json};

    let session = open();
    let title = story(&read(&session), "Title");
    let quotes = "\"".repeat(20 * 1024 * 1024);
    session
        .insert_text(
            &EditCtx::local("human"),
            &title.story_id,
            0,
            &quotes,
            &TextStyle::default(),
        )
        .unwrap();
    let end = quotes.len() as u32;
    let mismatch = session
        .apply_edits(&request(
            &session.version(),
            json!({"steps": [{"op": "deleteText", "target": range(&title, 0, end), "expect": {"text": "x"}}]}),
        ))
        .unwrap();
    assert_eq!(
        mismatch.as_ref().unwrap_err().failure.code,
        EditFailureCode::ContentMismatch
    );
    assert!(outcome_json(&mismatch).unwrap().len() < 4096);

    let search = session
        .find_text(&FindRequest {
            text: quotes.clone(),
            within: None,
            limit: None,
        })
        .unwrap();
    assert_eq!(
        search.as_ref().unwrap_err().failure.code,
        EditFailureCode::LimitExceeded
    );
    assert!(outcome_json(&search).unwrap().len() < 4096);
    let read = session
        .read_content(&ReadRequest {
            slide_ids: Some(vec!["s".repeat(MAX_REQUEST_BYTES)]),
        })
        .unwrap();
    assert_eq!(
        read.as_ref().unwrap_err().failure.code,
        EditFailureCode::LimitExceeded
    );

    let fixture_read = self::read(&open());
    let card = shape_target(&fixture_read, "Card");
    let junk = "\u{7f}".repeat(10 * 1024 * 1024);
    let long_id = "\u{1}".repeat(1024 * 1024);
    for step in [
        json!({"op": "setShapeFill", "target": card, "color": junk}),
        json!({"op": "setShapeRect", "target": {"slideId": title.slide_id, "shapeId": long_id},
               "rect": {"x": 0, "y": 0, "width": 1, "height": 1}}),
    ] {
        let fresh = open();
        let outcome = edit(&fresh, json!([step]));
        assert!(outcome.is_err());
        assert!(outcome_json(&outcome).unwrap().len() < 4096);
    }
}

#[test]
fn shape_steps_follow_the_existing_shape_policy() {
    let session = open();
    let read = read(&session);
    let card = shape(&read, "Card").clone();
    let card_target = shape_target(&read, "Card");
    let current = json!({"x": card.x, "y": card.y, "width": card.width, "height": card.height});
    let moved =
        json!({"x": card.x + 1000, "y": card.y, "width": card.width, "height": card.height});
    let applied = edit(
        &session,
        json!([
            {"op": "setShapeRect", "target": card_target, "rect": moved, "expect": {"rect": current}},
            {"op": "setShapeFill", "target": card_target, "color": "#112233", "expect": {"fill": card.fill}},
            {"op": "setShapeStroke", "target": card_target, "stroke": {"color": "#445566", "widthPt": 2.0},
             "expect": {"outline": card.outline}},
            {"op": "replaceText", "target": search(&story(&read, "Card"), "Card"), "text": "Metric"},
        ]),
    )
    .unwrap();
    assert!(applied.receipts.iter().all(|receipt| receipt.changed));
    assert!(matches!(applied.receipts[0].target, EditTarget::Shape(_)));
    let snapshot = session.snapshot().unwrap();
    let after = named(&snapshot.slides[0].shapes, "Card").unwrap();
    assert_eq!(after.x, card.x + 1000);
    assert_eq!(after.resolved_fill_color.as_deref(), Some("#112233"));
    assert_eq!(after.resolved_outline_color.as_deref(), Some("#445566"));

    let conflict = refused(edit(
        &session,
        json!([
            {"op": "setShapeRect", "target": card_target, "rect": current},
            {"op": "setShapeRect", "target": card_target, "rect": moved},
        ]),
    ));
    assert_eq!(
        (
            conflict.code,
            conflict.step_index,
            conflict.conflicting_step_index
        ),
        (EditFailureCode::OverlappingSteps, Some(1), Some(0))
    );
    let table = refused(edit(
        &session,
        json!([{"op": "setShapeFill", "target": shape_target(&read, "Table"), "color": "#000000"}]),
    ));
    assert_eq!(table.code, EditFailureCode::Unsupported);
    let stale = refused(edit(
        &session,
        json!([{"op": "setShapeFill", "target": card_target, "color": "#000000",
                "expect": {"fill": card.fill}}]),
    ));
    assert_eq!(stale.code, EditFailureCode::ContentMismatch);
    for step in [
        json!({"op": "setShapeFill", "target": card_target, "color": "blue"}),
        json!({"op": "setShapeRect", "target": card_target, "rect": {"x": 0, "y": 0, "width": 0, "height": 5}}),
        json!({"op": "setShapeStroke", "target": card_target, "stroke": {"widthPt": -1.0}}),
    ] {
        assert_eq!(
            refused(edit(&session, json!([step]))).code,
            EditFailureCode::InvalidStep
        );
    }
    let repeat = edit(
        &session,
        json!([{"op": "setShapeFill", "target": card_target, "color": "#112233"}]),
    )
    .unwrap();
    assert!(!repeat.applied);
}

#[test]
fn overlapping_text_writes_refuse_with_both_step_indices() {
    let session = open();
    let title = story(&read(&session), "Title");
    for steps in [
        json!([
            {"op": "replaceText", "target": range(&title, 0, 9), "text": "Q"},
            {"op": "insertText", "target": range(&title, 9, 9), "at": "start", "text": "x"},
        ]),
        json!([
            {"op": "insertText", "target": range(&title, 5, 5), "at": "start", "text": "x"},
            {"op": "insertText", "target": range(&title, 5, 5), "at": "end", "text": "y"},
        ]),
        json!([
            {"op": "formatText", "target": range(&title, 0, 5), "patch": {"bold": true}},
            {"op": "formatText", "target": range(&title, 3, 8), "patch": {"italic": true}},
        ]),
        json!([
            {"op": "setParagraphAlignment", "target": range(&title, 0, 0), "alignment": "ctr"},
            {"op": "insertText", "target": range(&title, 3, 3), "at": "start", "text": "x"},
        ]),
    ] {
        let failure = refused(edit(&session, steps));
        assert_eq!(failure.code, EditFailureCode::OverlappingSteps);
        assert_eq!(
            (failure.step_index, failure.conflicting_step_index),
            (Some(1), Some(0))
        );
    }
    edit(
        &session,
        json!([
            {"op": "replaceText", "target": range(&title, 0, 9), "text": "Q"},
            {"op": "replaceText", "target": range(&title, 9, 16), "text": " recap"},
            {"op": "setParagraphAlignment", "target": range(&title, 17, 17), "alignment": "r"},
        ]),
    )
    .unwrap();
    let story_now = session.story(&title.story_id).unwrap();
    assert_eq!(story_now.plain_text(), "Q recap\nRevenue\ngrew 😀 fast");
    assert_eq!(story_now.paragraphs[1].alignment.as_deref(), Some("r"));
    assert_eq!(story_now.paragraphs[0].alignment, None);
}

#[test]
fn no_op_batches_keep_the_version_history_and_redo() {
    let session = open();
    let title = story(&read(&session), "Title");
    session
        .insert_text(
            &EditCtx::local("human"),
            &title.story_id,
            0,
            "X",
            &TextStyle::default(),
        )
        .unwrap();
    assert!(session.undo());
    assert!(session.can_redo());
    let events = Arc::new(AtomicUsize::new(0));
    let counter = Arc::clone(&events);
    let _observer = session
        .observe_update_v1(move |_| {
            counter.fetch_add(1, Ordering::Relaxed);
        })
        .unwrap();
    let read = read(&session);
    let card = shape(&read, "Card");
    let card_rect = json!({"x": card.x, "y": card.y, "width": card.width, "height": card.height});
    let request = request(
        &read.version,
        json!({"steps": [
            {"op": "replaceText", "target": search(&title, "review"), "text": "review"},
            {"op": "insertText", "target": range(&title, 0, 0), "at": "start", "text": ""},
            {"op": "deleteText", "target": range(&title, 3, 3)},
            {"op": "formatText", "target": range(&title, 5, 9), "patch": {}},
            {"op": "setParagraphAlignment", "target": range(&title, 20, 20), "alignment": null},
            {"op": "setShapeRect", "target": shape_target(&read, "Card"), "rect": card_rect},
            {"op": "setSlideNotes", "target": {"slideId": read.slides[1].id}, "text": "Speaker notes"},
        ]}),
    );
    let validation = session.validate_edits(&request).unwrap().unwrap();
    assert!(!validation.would_apply);
    let outcome = session.apply_edits(&request).unwrap().unwrap();
    assert!(!outcome.applied);
    assert_eq!(outcome.version, read.version);
    assert!(outcome.receipts.iter().all(|receipt| !receipt.changed));
    assert!(outcome.changed_slides.is_empty());
    assert_eq!(range_of(&outcome.receipts[0].target), (10, 16));
    assert_eq!(session.version(), read.version);
    assert_eq!(events.load(Ordering::Relaxed), 0);
    assert!(session.can_redo());
    assert!(!session.can_undo());
}

#[test]
fn deletions_history_peers_and_reopening_invalidate_versions() {
    let session = open();
    let title = story(&read(&session), "Title");
    let context = EditCtx::local("human");
    let (version, state) = (session.version(), session.encode_state_vector_v1());
    session
        .delete_text(&context, &title.story_id, 0, 1)
        .unwrap();
    assert_eq!(session.encode_state_vector_v1(), state);
    assert_ne!(session.version(), version);
    let stale = session
        .apply_edits(&request(
            &version,
            json!({"steps": [{"op": "insertText", "target": range(&title, 0, 0), "at": "start", "text": "Q"}]}),
        ))
        .unwrap()
        .unwrap_err();
    assert_eq!(stale.failure.code, EditFailureCode::StaleVersion);
    assert_eq!(stale.version, session.version());

    let before_undo = session.version();
    assert!(session.undo());
    assert_ne!(session.version(), before_undo);
    let before_redo = session.version();
    assert!(session.redo());
    assert_ne!(session.version(), before_redo);

    let peer = DeckSession::open(&fixture(), 73).unwrap();
    peer.set_slide_notes(&context, &title.slide_id, "Peer notes")
        .unwrap();
    let before_remote = session.version();
    session
        .apply_update_v1(
            &peer
                .encode_diff_v1(&session.encode_state_vector_v1())
                .unwrap(),
        )
        .unwrap();
    assert_ne!(session.version(), before_remote);

    let bytes = fixture();
    assert_ne!(
        DeckSession::open(&bytes, 71).unwrap().version(),
        DeckSession::open(&bytes, 71).unwrap().version()
    );
}

#[test]
fn separate_history_is_one_undo_step_in_both_capture_modes() {
    for mode in [UndoCaptureMode::Auto, UndoCaptureMode::Manual] {
        let session = open();
        session.set_undo_capture_mode(mode);
        let title = story(&read(&session), "Title");
        let context = EditCtx::local("human");
        let style = TextStyle::default();
        session
            .insert_text(&context, &title.story_id, 0, "A", &style)
            .unwrap();
        let applied = session
            .apply_edits(&request(
                &session.version(),
                json!({"source": "agent", "steps": [
                    {"op": "replaceText", "target": search(&title, "review"), "text": "recap"},
                ]}),
            ))
            .unwrap()
            .unwrap();
        assert_eq!(applied.source, EditSource::Agent);
        session
            .insert_text(&context, &title.story_id, 0, "B", &style)
            .unwrap();
        let first_line = |session: &DeckSession| {
            text_of(session, &title)
                .split('\n')
                .next()
                .unwrap()
                .to_owned()
        };
        assert_eq!(first_line(&session), "BAQuarterly recap");
        assert!(session.undo());
        assert_eq!(first_line(&session), "AQuarterly recap", "{mode:?}");
        assert!(session.undo());
        assert_eq!(first_line(&session), "AQuarterly review", "{mode:?}");
        assert!(session.undo());
        assert_eq!(first_line(&session), "Quarterly review", "{mode:?}");
        assert!(!session.can_undo());
    }
}

#[test]
fn source_is_provenance_and_history_alone_decides_undo() {
    for (source, history, undoable) in [
        ("host", "separate", true),
        ("agent", "separate", true),
        ("host", "none", false),
        ("agent", "none", false),
    ] {
        let session = open();
        let title = story(&read(&session), "Title");
        let applied = session
            .apply_edits(&request(
                &session.version(),
                json!({"source": source, "history": history, "steps": [
                    {"op": "replaceText", "target": search(&title, "review"), "text": "recap"},
                ]}),
            ))
            .unwrap()
            .unwrap();
        assert_eq!(serde_json::to_value(applied.source).unwrap(), source);
        assert_eq!(session.can_undo(), undoable, "{source} {history}");
    }
}

#[test]
fn untracked_batches_keep_existing_undo_and_redo_entries() {
    let session = open();
    let title = story(&read(&session), "Title");
    let context = EditCtx::local("human");
    let style = TextStyle::default();
    session
        .insert_text(&context, &title.story_id, 0, "A", &style)
        .unwrap();
    session.add_undo_barrier();
    session
        .insert_text(&context, &title.story_id, 0, "B", &style)
        .unwrap();
    assert!(session.undo());
    let applied = session
        .apply_edits(&request(
            &session.version(),
            json!({"history": "none", "steps": [
                {"op": "replaceText", "target": search(&title, "review"), "text": "recap"},
            ]}),
        ))
        .unwrap()
        .unwrap();
    assert!(applied.applied);
    assert!(session.can_redo());
    assert!(session.undo());
    assert!(text_of(&session, &title).starts_with("Quarterly recap"));
    assert!(!session.can_undo());
}

#[test]
fn proposals_keep_their_contract_next_to_batches() {
    let session = open();
    let title = story(&read(&session), "Title");
    let proposal = session
        .propose(ProposalRequest {
            agent_id: "agent".into(),
            note: None,
            edits: vec![ProposalEdit::ReplaceText {
                story_id: title.story_id.clone(),
                start: 0,
                end: 9,
                text: "Yearly".into(),
                style: None,
            }],
        })
        .unwrap();
    let version = session.version();
    edit(
        &session,
        json!([{"op": "replaceText", "target": search(&title, "review"), "text": "recap"}]),
    )
    .unwrap();
    assert_ne!(session.version(), version);
    assert_eq!(
        session.proposals().unwrap()[0].stale_targets,
        [title.shape_id.as_str()]
    );
    assert!(matches!(
        session.accept_proposal(&proposal.id, false),
        Err(ProposalError::Stale(_))
    ));
    assert!(session.accept_proposal(&proposal.id, true).unwrap().applied);
    assert!(text_of(&session, &title).starts_with("Yearly recap"));
}

#[test]
fn update_observers_see_the_committed_version() {
    let session = Rc::new(open());
    let seen = Rc::new(RefCell::new(Vec::new()));
    let (weak, log) = (Rc::downgrade(&session), Rc::clone(&seen));
    let _observer = session
        .observe_update_v1(move |_| {
            if let Some(session) = weak.upgrade() {
                log.borrow_mut().push(session.version());
            }
        })
        .unwrap();
    let title = story(&read(&session), "Title");
    let applied = edit(
        &session,
        json!([{"op": "replaceText", "target": search(&title, "review"), "text": "recap"}]),
    )
    .unwrap();
    assert_eq!(*seen.borrow(), [applied.version]);
}

#[test]
fn updates_waiting_on_missing_structure_refuse_batches() {
    let origin = open();
    let base = origin.encode_state_as_update_v1();
    let title = story(&read(&origin), "Title");
    let (context, style) = (EditCtx::local("human"), TextStyle::default());
    origin
        .insert_text(&context, &title.story_id, 0, "A", &style)
        .unwrap();
    let after_first = origin.encode_state_vector_v1();
    origin
        .insert_text(&context, &title.story_id, 1, "B", &style)
        .unwrap();
    let dependent = origin.encode_diff_v1(&after_first).unwrap();
    let merged = yrs::merge_updates_v1([base.as_slice(), dependent.as_slice()]).unwrap();
    let joined = DeckSession::open_from_update(&merged, 74).unwrap();
    let failure = refused(edit(
        &joined,
        json!([{"op": "replaceText", "target": search(&story(&read(&joined), "Title"), "review"), "text": "recap"}]),
    ));
    assert_eq!(failure.code, EditFailureCode::Unsupported);
    assert_eq!(failure.step_index, None);
    assert!(failure.message.contains("not integrated"), "{failure:?}");
}

#[test]
fn oversized_batches_refuse_before_planning() {
    let session = open();
    let title = story(&read(&session), "Title");
    let empty =
        json!({"op": "insertText", "target": range(&title, 0, 0), "at": "start", "text": ""});
    let failure = refused(edit(&session, Value::Array(vec![empty; 129])));
    assert_eq!(failure.code, EditFailureCode::LimitExceeded);
    assert_eq!(failure.step_index, None);
    let huge = "x".repeat(1_048_577);
    let failure = refused(edit(
        &session,
        json!([{"op": "insertText", "target": range(&title, 0, 0), "at": "start", "text": huge}]),
    ));
    assert_eq!(failure.code, EditFailureCode::LimitExceeded);
    let guard = "x".repeat(pptx_edit::MAX_REQUEST_BYTES);
    let failure = refused(edit(
        &session,
        json!([{"op": "deleteText", "target": range(&title, 0, 0), "expect": {"text": guard}}]),
    ));
    assert_eq!(failure.code, EditFailureCode::LimitExceeded);
    let oversized = pptx_edit::oversized_request(session.version());
    assert_eq!(oversized.failure.code, EditFailureCode::LimitExceeded);
}

#[test]
fn validation_previews_pre_batch_targets_and_reserves_nothing() {
    let session = open();
    let read = read(&session);
    let title = story(&read, "Title");
    let request = request(
        &read.version,
        json!({"steps": [
            {"op": "replaceText", "target": range(&title, 0, 9), "text": "Q"},
            {"op": "insertText", "target": search(&title, "fast"), "at": "end", "text": "!"},
        ]}),
    );
    let before = session.encode_state_as_update_v1();
    let validation = session.validate_edits(&request).unwrap().unwrap();
    assert!(validation.would_apply);
    assert_eq!(validation.base_version, read.version);
    assert_eq!(range_of(&validation.previews[1].target), (33, 37));
    assert!(
        validation
            .previews
            .iter()
            .all(|preview| preview.would_change)
    );
    assert_eq!(session.encode_state_as_update_v1(), before);
    assert_eq!(session.version(), read.version);
    let applied = session.apply_edits(&request).unwrap().unwrap();
    assert_eq!(range_of(&applied.receipts[1].target), (29, 30));
    let stale = session.validate_edits(&request).unwrap().unwrap_err();
    assert_eq!(stale.failure.code, EditFailureCode::StaleVersion);
}

#[test]
fn malformed_requests_fail_to_decode() {
    let target = r#"{"kind":"range","slideId":"s","shapeId":"h","storyId":"t","start":0,"end":0}"#;
    for body in [
        r#"{"steps":[]}"#.to_owned(),
        r#"{"expectVersion":"v","steps":[{"op":"explode"}]}"#.to_owned(),
        r#"{"expectVersion":"v","steps":[],"extra":1}"#.to_owned(),
        format!(r#"{{"expectVersion":"v","steps":[{{"op":"setParagraphAlignment","target":{target}}}]}}"#),
        r#"{"expectVersion":"v","steps":[{"op":"setShapeFill","target":{"slideId":"s","shapeId":"h"}}]}"#.to_owned(),
        r#"{"expectVersion":"v","steps":[{"op":"deleteText","target":{"kind":"range","slideId":"s","shapeId":"h","storyId":"t","start":-1,"end":0}}]}"#.to_owned(),
        r#"{"expectVersion":"v","steps":[{"op":"setShapeRect","target":{"slideId":"s","shapeId":"h"},"rect":{"x":1.5,"y":0,"width":1,"height":1}}]}"#.to_owned(),
        format!(r#"{{"expectVersion":"v","steps":[{{"op":"deleteText","target":{target},"expect":{{"text":"x","extra":true}}}}]}}"#),
    ] {
        assert!(serde_json::from_str::<EditRequest>(&body).is_err(), "{body}");
    }
    let empty: EditRequest = serde_json::from_str(r#"{"expectVersion":"v","steps":[]}"#).unwrap();
    assert_eq!(
        (empty.source, empty.history),
        (EditSource::Host, EditHistory::Separate)
    );
    let session = open();
    let outcome = session
        .apply_edits(&EditRequest {
            expect_version: session.version(),
            ..empty
        })
        .unwrap()
        .unwrap();
    assert!(!outcome.applied);
    assert!(outcome.receipts.is_empty());
}

#[test]
fn an_applied_batch_saves_and_reopens_with_unrelated_parts_intact() {
    let bytes = fixture();
    let session = DeckSession::open(&bytes, 75).unwrap();
    let read = read(&session);
    let title = story(&read, "Title");
    let cell = read
        .stories
        .iter()
        .find(|story| story.text == "Cell B")
        .unwrap()
        .clone();
    let card = shape(&read, "Card");
    let rect = json!({"x": 6_500_000, "y": card.y, "width": card.width, "height": card.height});
    let applied = edit(
        &session,
        json!([
            {"op": "replaceText", "target": search(&title, "Quarterly"), "text": "Annual"},
            {"op": "setParagraphAlignment", "target": range(&title, 17, 17), "alignment": "ctr"},
            {"op": "replaceText", "target": search(&cell, "Cell B"), "text": "Cell Beta"},
            {"op": "setShapeRect", "target": shape_target(&read, "Card"), "rect": rect},
            {"op": "setSlideNotes", "target": {"slideId": title.slide_id}, "text": "Opening notes"},
        ]),
    )
    .unwrap();
    assert_eq!(applied.changed_slides, [title.slide_id.as_str()]);
    let saved = session.save().unwrap();
    let parts = |bytes: &[u8]| -> BTreeMap<String, Vec<u8>> {
        ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
    };
    let (source, written) = (parts(&bytes), parts(&saved));
    for untouched in [
        "docProps/custom.xml",
        "ppt/slides/slide2.xml",
        "ppt/notesSlides/notesSlide1.xml",
        "ppt/theme/theme1.xml",
    ] {
        assert_eq!(source[untouched], written[untouched], "{untouched}");
    }
    let slide = String::from_utf8(written["ppt/slides/slide1.xml"].clone()).unwrap();
    for kept in [r#"type="slidenum""#, r#"firstRow="1""#, r#"name="Nested""#] {
        assert!(slide.contains(kept), "{kept}");
    }

    let reopened = DeckSession::open(&saved, 76).unwrap();
    let again = self::read(&reopened);
    let title_again = story(&again, "Title");
    assert_eq!(title_again.text, "Annual review\nRevenue\ngrew 😀 fast");
    assert_eq!(
        reopened.story(&title_again.story_id).unwrap().paragraphs[1]
            .alignment
            .as_deref(),
        Some("ctr")
    );
    assert!(again.stories.iter().any(|story| story.text == "Cell Beta"));
    assert_eq!(shape(&again, "Card").x, 6_500_000);
    assert_eq!(again.slides[0].notes, "Opening notes");
    assert_eq!(again.slides[1].notes, "Speaker notes");
    let carried = reopened
        .apply_edits(&request(
            &applied.version,
            json!({"steps": [{"op": "deleteText", "target": range(&title_again, 0, 1)}]}),
        ))
        .unwrap()
        .unwrap_err();
    assert_eq!(carried.failure.code, EditFailureCode::StaleVersion);
}
