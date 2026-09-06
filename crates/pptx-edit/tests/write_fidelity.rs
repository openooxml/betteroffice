//! Write-back fidelity against a deck exercising markup the model does not
//! carry: connectors, inherited placeholder rects, hyperlinks, fields,
//! theme colours, notes slides, custom shows, and hostile inputs.

use std::collections::BTreeMap;

use pptx_edit::{CommentFlavor, DeckSession, EditCtx, EditError, TextStyle};

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
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#;

const PRESENTATION_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
</Relationships>"#;

const SLIDE1: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr/>
<p:txBody><a:bodyPr/><a:lstStyle/>
<a:p><a:r><a:rPr b="1"/><a:t>Hello </a:t></a:r><a:r><a:rPr strike="sngStrike"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>link</a:t></a:r></a:p>
<a:p><a:fld id="{D038279B-FC19-497E-A7D1-5ADD9CAF016F}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld><a:r><a:rPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:rPr><a:t>Accent</a:t></a:r></a:p>
</p:txBody></p:sp>
<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="300" cy="400"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000" y="2000"/><a:ext cx="3000" cy="4000"/></a:xfrm>
<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 20000"/><a:gd name="adj2" fmla="*/ missing 2 3"/></a:avLst></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:noFill/></a:rPr><a:t>Box</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="4" name="Halfway"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm rot="1200000"><a:off x="123456" y="654321"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Halfway</a:t></a:r></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sld>"#;

const SLIDE1_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/" TargetMode="External"/>
</Relationships>"#;

const SLIDE2: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Second"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second</a:t></a:r></a:p></p:txBody></p:sp>
<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Trailing Connector"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
<p:spPr><a:xfrm><a:off x="1" y="2"/><a:ext cx="3" cy="4"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>
</p:spTree></p:cSld></p:sld>"#;

const SLIDE2_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
</Relationships>"#;

const NOTES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:notes>"#;

const NOTES_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide2.xml"/>
</Relationships>"#;

const LAYOUT: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title">
<p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr/>
<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm flipH="1" rot="2700000"><a:off x="762000" y="457200"/><a:ext cx="10668000" cy="1143000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="body"/></p:nvPr></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="8000000" cy="3000000"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>
</p:spTree></p:cSld></p:sldLayout>"#;

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

fn presentation_xml(first_slide_id: u32) -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="{first_slide_id}" r:id="rId2"/><p:sldId id="257" r:id="rId3"/></p:sldIdLst>
<p:custShowLst><p:custShow name="Short" id="0"><p:sldLst><p:sld r:id="rId2"/><p:sld r:id="rId3"/></p:sldLst></p:custShow></p:custShowLst>
<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"#
    )
}

const LEGACY_AUTHORS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmAuthorLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cmAuthor id="0" name="Mary Smith" initials="mas" lastIdx="1" clrIdx="0"/>
</p:cmAuthorLst>"#;

const LEGACY_COMMENTS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmLst xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cm authorId="0" dt="2005-11-13T17:00:22.071" idx="1"><p:pos x="576" y="288"/><p:text>Needs a source.</p:text></p:cm>
</p:cmLst>"#;

const MODERN_AUTHORS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:authorLst xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">
<p188:author id="{CD37207E-7903-4ED4-8AE8-017538D2DF7E}" name="Mary Smith" initials="mas" userId="" providerId=""/>
</p188:authorLst>"#;

const MODERN_COMMENTS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p188:cmLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pc="http://schemas.microsoft.com/office/powerpoint/2013/main/command" xmlns:p188="http://schemas.microsoft.com/office/powerpoint/2018/8/main">
<p188:cm id="{62A8A96D-E5A8-4BFC-B993-A6EAE3907CAD}" authorId="{CD37207E-7903-4ED4-8AE8-017538D2DF7E}" created="2024-12-30T20:26:06.503">
<pc:sldMkLst><pc:docMk/><pc:sldMk cId="0" sldId="256"/></pc:sldMkLst>
<p188:pos x="576" y="288"/>
<p188:replyLst><p188:reply id="{9F5B1E2C-4A7D-4E93-9B1A-2C3D4E5F6A7B}" authorId="{CD37207E-7903-4ED4-8AE8-017538D2DF7E}" created="2024-12-30T20:31:12.117"><p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Added one.</a:t></a:r></a:p></p188:txBody></p188:reply></p188:replyLst>
<p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Needs a source.</a:t></a:r></a:p></p188:txBody>
</p188:cm>
</p188:cmLst>"#;

const LEGACY_COMMENTS_REL: &str = r#"<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments/comment1.xml"/>"#;
const LEGACY_AUTHORS_REL: &str = r#"<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" Target="commentAuthors.xml"/>"#;
const MODERN_COMMENTS_REL: &str = r#"<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2018/10/relationships/comments" Target="../comments/modernComment1.xml"/>"#;
const MODERN_AUTHORS_REL: &str = r#"<Relationship Id="rId9" Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" Target="authors.xml"/>"#;

fn commented_fixture(flavor: CommentFlavor) -> Vec<u8> {
    let (
        authors_part,
        comments_part,
        authors,
        comments,
        authors_rel,
        comments_rel,
        authors_type,
        comments_type,
    ) = match flavor {
        CommentFlavor::Legacy => (
            "ppt/commentAuthors.xml",
            "ppt/comments/comment1.xml",
            LEGACY_AUTHORS,
            LEGACY_COMMENTS,
            LEGACY_AUTHORS_REL,
            LEGACY_COMMENTS_REL,
            "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml",
            "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
        ),
        CommentFlavor::Modern => (
            "ppt/authors.xml",
            "ppt/comments/modernComment1.xml",
            MODERN_AUTHORS,
            MODERN_COMMENTS,
            MODERN_AUTHORS_REL,
            MODERN_COMMENTS_REL,
            "application/vnd.ms-powerpoint.authors+xml",
            "application/vnd.ms-powerpoint.comments+xml",
        ),
    };
    let mut parts = fixture_parts(256);
    for (path, body) in parts.iter_mut() {
        match path.as_str() {
            "[Content_Types].xml" => {
                *body = body.replace(
                    "</Types>",
                    &format!(
                        r#"<Override PartName="/{authors_part}" ContentType="{authors_type}"/><Override PartName="/{comments_part}" ContentType="{comments_type}"/></Types>"#
                    ),
                );
            }
            "ppt/_rels/presentation.xml.rels" => {
                *body = body.replace(
                    "</Relationships>",
                    &format!("{authors_rel}</Relationships>"),
                );
            }
            "ppt/slides/_rels/slide1.xml.rels" => {
                *body = body.replace(
                    "</Relationships>",
                    &format!("{comments_rel}</Relationships>"),
                );
            }
            _ => {}
        }
    }
    parts.push((authors_part.to_owned(), authors.to_owned()));
    parts.push((comments_part.to_owned(), comments.to_owned()));
    zip(parts)
}

fn zip(parts: Vec<(String, String)>) -> Vec<u8> {
    let parts: Vec<(String, Vec<u8>)> = parts
        .into_iter()
        .map(|(path, body)| (path, body.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn fixture(first_slide_id: u32) -> Vec<u8> {
    zip(fixture_parts(first_slide_id))
}

fn fixture_parts(first_slide_id: u32) -> Vec<(String, String)> {
    [
        ("[Content_Types].xml", CONTENT_TYPES.to_owned()),
        ("_rels/.rels", ROOT_RELS.to_owned()),
        ("ppt/presentation.xml", presentation_xml(first_slide_id)),
        (
            "ppt/_rels/presentation.xml.rels",
            PRESENTATION_RELS.to_owned(),
        ),
        ("ppt/slides/slide1.xml", SLIDE1.to_owned()),
        ("ppt/slides/_rels/slide1.xml.rels", SLIDE1_RELS.to_owned()),
        ("ppt/slides/slide2.xml", SLIDE2.to_owned()),
        ("ppt/slides/_rels/slide2.xml.rels", SLIDE2_RELS.to_owned()),
        ("ppt/notesSlides/notesSlide1.xml", NOTES.to_owned()),
        (
            "ppt/notesSlides/_rels/notesSlide1.xml.rels",
            NOTES_RELS.to_owned(),
        ),
        ("ppt/slideLayouts/slideLayout1.xml", LAYOUT.to_owned()),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            LAYOUT_RELS.to_owned(),
        ),
        ("ppt/slideMasters/slideMaster1.xml", MASTER.to_owned()),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            MASTER_RELS.to_owned(),
        ),
        ("ppt/theme/theme1.xml", THEME.to_owned()),
    ]
    .into_iter()
    .map(|(path, body)| (path.to_owned(), body))
    .collect()
}

fn open() -> DeckSession {
    DeckSession::open(&fixture(256), 11).unwrap()
}

fn context() -> EditCtx {
    EditCtx::local("fidelity")
}

fn parts(bytes: &[u8]) -> BTreeMap<String, Vec<u8>> {
    ooxml_opc::unzip_parts(bytes).unwrap().into_iter().collect()
}

fn part_text(parts: &BTreeMap<String, Vec<u8>>, path: &str) -> String {
    String::from_utf8(
        parts
            .get(path)
            .unwrap_or_else(|| panic!("missing {path}"))
            .clone(),
    )
    .unwrap()
}

/// Every internal relationship in every .rels part must resolve to a part.
fn assert_relationships_resolve(parts: &BTreeMap<String, Vec<u8>>) {
    for (path, bytes) in parts {
        let Some((directory, name)) = path.rsplit_once("/_rels/") else {
            continue;
        };
        let source_directory = directory;
        let text = String::from_utf8(bytes.clone()).unwrap();
        for chunk in text.split("<Relationship ").skip(1) {
            if chunk.contains("TargetMode=\"External\"") {
                continue;
            }
            let target = chunk
                .split("Target=\"")
                .nth(1)
                .and_then(|rest| rest.split('"').next())
                .unwrap_or_else(|| panic!("{path}: relationship without a target"));
            let mut segments: Vec<&str> = source_directory.split('/').collect();
            for segment in target.split('/') {
                match segment {
                    "." | "" => {}
                    ".." => {
                        segments.pop();
                    }
                    segment => segments.push(segment),
                }
            }
            let resolved = segments.join("/");
            assert!(
                parts.contains_key(&resolved),
                "{path} ({name}) references missing part {resolved}"
            );
        }
    }
}

fn box_shape(session: &DeckSession) -> (String, String) {
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Box")
        .unwrap();
    (slide.id.clone(), shape.id.clone())
}

#[test]
fn a_connector_between_shapes_keeps_its_slot() {
    let session = open();
    let (slide_id, shape_id) = box_shape(&session);
    session
        .move_shape(&context(), &slide_id, &shape_id, 5000, 6000)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    let title = slide.find("Title").unwrap();
    let connector = slide.find("<p:cxnSp>").unwrap();
    let moved = slide.find("\"Box\"").unwrap();
    assert!(title < connector && connector < moved);
    assert!(slide.contains(r#"<a:off x="5000" y="6000"/>"#));
    assert_relationships_resolve(&saved);
}

#[test]
fn a_moved_inherited_placeholder_keeps_its_rect_rotation_and_flip() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let title = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap();
    session
        .move_shape(&context(), &slide.id, &title.id, 900_000, 1_000_000)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<a:off x="900000" y="1000000"/>"#));
    assert!(slide.contains(r#"<a:ext cx="10668000" cy="1143000"/>"#));
    assert!(slide.contains(r#"rot="2700000""#));
    assert!(slide.contains(r#"flipH="1""#));

    let reopened = DeckSession::open(&session.save().unwrap(), 12).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    let title = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap();
    assert_eq!((title.x, title.y), (900_000, 1_000_000));
    assert_eq!((title.width, title.height), (10_668_000, 1_143_000));
    assert_eq!(title.rotation_deg, 45.0);
    assert!(title.flip_h);
}

#[test]
fn a_resized_inherited_placeholder_keeps_its_inherited_offset() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let title = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap();
    session
        .resize_shape(&context(), &slide.id, &title.id, 5_000_000, 900_000)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<a:off x="762000" y="457200"/>"#));
    assert!(slide.contains(r#"<a:ext cx="5000000" cy="900000"/>"#));
}

#[test]
fn a_shape_added_to_an_emptied_slide_lands_after_the_group_properties() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = snapshot.slides[1].clone();
    session
        .remove_shape(&context(), &slide.id, &slide.shapes[0].id)
        .unwrap();
    session
        .add_text_box(
            &context(),
            &slide.id,
            &pptx_edit::ShapeDraft {
                name: "Fresh".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 100,
                    y: 200,
                    width: 3_000,
                    height: 4_000,
                },
                text: "on an empty slide".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide2.xml");
    let non_visual = slide.find("<p:nvGrpSpPr>").unwrap();
    let group_properties = slide.find("<p:grpSpPr").unwrap();
    let added = slide.find("<p:sp>").unwrap();
    assert!(non_visual < group_properties && group_properties < added);
}

#[test]
fn junk_style_values_are_rejected_at_the_edit() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0].shapes[0].text_stories[0].id.clone();
    let styles = [
        TextStyle {
            color: Some("rebeccapurple".to_owned()),
            ..TextStyle::default()
        },
        TextStyle {
            underline: Some("wobbly".to_owned()),
            ..TextStyle::default()
        },
        TextStyle {
            font_size_pt: Some(0.001),
            ..TextStyle::default()
        },
    ];
    for style in styles {
        let error = session
            .insert_text(&context(), &story_id, 0, "X", &style)
            .unwrap_err();
        assert!(matches!(error, EditError::InvalidText(_)));
    }
    assert_eq!(parts(&session.save().unwrap()), parts(&fixture(256)));
}

#[test]
fn an_edit_inside_a_linked_run_keeps_the_link() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session
        .insert_text(&context(), &story_id, 8, "X", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let slide = part_text(&parts(&saved), "ppt/slides/slide1.xml");
    assert!(slide.contains("<a:t>liXnk</a:t>"));
    assert!(slide.contains("hlinkClick"));
    assert!(slide.contains(r#"strike="sngStrike""#));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert!(
        reopened
            .story(&story_id)
            .unwrap()
            .plain_text()
            .starts_with("Hello liXnk")
    );
}

#[test]
fn an_edit_before_a_hyperlink_keeps_links_fields_and_theme_colours() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session
        .insert_text(&context(), &story_id, 0, "Hi ", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let slide = part_text(&parts(&saved), "ppt/slides/slide1.xml");
    assert!(slide.contains("hlinkClick"));
    assert!(slide.contains(r#"strike="sngStrike""#));
    assert!(slide.contains("<a:fld "));
    assert!(slide.contains(r#"schemeClr val="accent1""#));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    let text = reopened.story(&story_id).unwrap().plain_text();
    assert_eq!(text, "Hi Hello link\n1Accent");
}

#[test]
fn deleting_a_slide_prunes_its_notes_and_custom_show_entry() {
    let session = open();
    let second = session.snapshot().unwrap().slides[1].id.clone();
    session.delete_slide(&context(), &second).unwrap();

    let saved = parts(&session.save().unwrap());
    assert!(!saved.contains_key("ppt/slides/slide2.xml"));
    assert!(!saved.contains_key("ppt/notesSlides/notesSlide1.xml"));
    assert!(!saved.contains_key("ppt/notesSlides/_rels/notesSlide1.xml.rels"));
    let content_types = part_text(&saved, "[Content_Types].xml");
    assert!(!content_types.contains("notesSlide1.xml"));
    assert!(!content_types.contains("slide2.xml"));
    let presentation = part_text(&saved, "ppt/presentation.xml");
    assert_eq!(presentation.matches("rId3").count(), 0);
    assert!(presentation.contains("custShowLst"));
    assert_relationships_resolve(&saved);
}

#[test]
fn illegal_control_characters_are_rejected_at_the_edit() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0].shapes[0].text_stories[0].id.clone();
    let error = session
        .insert_text(&context(), &story_id, 0, "a\u{1}b", &TextStyle::default())
        .unwrap_err();
    assert!(matches!(error, EditError::InvalidText(_)));
    assert_eq!(parts(&session.save().unwrap()), parts(&fixture(256)));
}

#[test]
fn a_carriage_return_survives_the_round_trip() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0].shapes[0].text_stories[0].id.clone();
    session
        .insert_text(&context(), &story_id, 0, "a\rb ", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert!(
        reopened
            .story(&story_id)
            .unwrap()
            .plain_text()
            .starts_with("a\rb ")
    );
}

#[test]
fn unevaluated_adjustment_guides_survive_a_shape_adjust() {
    let session = open();
    let (slide_id, shape_id) = box_shape(&session);
    let mut adjustments = BTreeMap::new();
    adjustments.insert("adj".to_owned(), 0.25);
    session
        .set_shape_adjust(&context(), &slide_id, &shape_id, &adjustments)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"fmla="val 25000" name="adj""#));
    assert!(slide.contains(r#"fmla="*/ missing 2 3" name="adj2""#));
}

#[test]
fn an_exhausted_slide_id_space_errors_instead_of_panicking() {
    let session = DeckSession::open(&fixture(4_294_967_295), 11).unwrap();
    session.insert_slide(&context(), 2, None).unwrap();
    let error = session.save().unwrap_err();
    assert!(matches!(error, EditError::Write(message) if message.contains("slide id")));
}

#[test]
fn a_colour_write_replaces_an_existing_no_fill() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Box")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session
        .format_text(
            &context(),
            &story_id,
            0,
            3,
            &pptx_edit::TextStylePatch {
                color: Some("#FF0000".to_owned()),
                ..pptx_edit::TextStylePatch::default()
            },
        )
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"srgbClr val="FF0000""#));
    assert!(!slide.contains("<a:noFill/></a:rPr>"));
}

#[test]
fn an_alignment_write_reaches_the_paragraph_properties() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let story_id = snapshot.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Title")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session
        .set_paragraph_alignment(&context(), &story_id, 0, 0, Some("ctr"))
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert_eq!(slide.matches(r#"algn="ctr""#).count(), 1);
    assert!(slide.contains("Accent"));
}

#[test]
fn a_resize_keeps_a_partial_transforms_own_offset_and_rotation() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let halfway = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Halfway")
        .unwrap();
    session
        .resize_shape(&context(), &slide.id, &halfway.id, 5_000_000, 900_000)
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<a:off x="123456" y="654321"/>"#));
    assert!(slide.contains(r#"rot="1200000""#));
    assert!(slide.contains(r#"<a:ext cx="5000000" cy="900000"/>"#));
    assert!(!slide.contains(r#"<a:off x="1000000" y="2000000"/>"#));
}

#[test]
fn a_shape_added_after_a_trailing_connector_lands_on_top() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[1].id.clone();
    session
        .add_text_box(
            &context(),
            &slide_id,
            &pptx_edit::ShapeDraft {
                name: "OnTop".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 1,
                    y: 2,
                    width: 300,
                    height: 400,
                },
                text: "above the connector".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide2.xml");
    let second = slide.find("\"Second\"").unwrap();
    let connector = slide.find("<p:cxnSp>").unwrap();
    let added = slide.find("\"OnTop\"").unwrap();
    assert!(second < connector && connector < added);
}

#[test]
fn an_unknown_layout_is_rejected_at_insert_slide() {
    let session = open();
    let error = session
        .insert_slide(&context(), 2, Some("ppt/slideLayouts/nope.xml"))
        .unwrap_err();
    assert!(matches!(error, EditError::InvalidState(_)));
    assert_eq!(parts(&session.save().unwrap()), parts(&fixture(256)));
}

#[test]
fn an_update_seeded_by_a_different_parse_refuses_to_save_through_stale_ordinals() {
    let source = fixture(256);

    let mut stale = pptx_parse::parse_pptx(&source).unwrap();
    let slide = stale
        .slides
        .iter_mut()
        .find(|slide| slide.shapes.len() > 1)
        .expect("the fidelity deck has a slide with several shapes");
    slide.shapes.remove(0);

    let seeded = DeckSession::from_package_with_source(stale, &source, 21).unwrap();
    let update = seeded.encode_state_as_update_v1();

    let reattached = DeckSession::open_from_update_with_source(&update, &source, 22).unwrap();
    let snapshot = reattached.snapshot().unwrap();
    let (slide_id, shape_id) = snapshot
        .slides
        .iter()
        .find_map(|slide| {
            slide
                .shapes
                .first()
                .map(|shape| (slide.id.clone(), shape.id.clone()))
        })
        .expect("the reattached session has a shape");
    reattached
        .move_shape(&context(), &slide_id, &shape_id, 1_000, 2_000)
        .unwrap();

    assert!(matches!(
        reattached.save(),
        Err(EditError::Write(message)) if message.contains("no longer addresses source shape")
    ));
}

#[test]
fn a_session_opened_from_an_update_alone_refuses_to_save() {
    let source = fixture(256);
    let seeded = DeckSession::open(&source, 11).unwrap();
    let update = seeded.encode_state_as_update_v1();

    let bare = DeckSession::open_from_update(&update, 12).unwrap();
    assert!(matches!(bare.save(), Err(EditError::Write(_))));

    let attached = DeckSession::open_from_update_with_source(&update, &source, 13).unwrap();
    let saved = attached.save().unwrap();
    assert_eq!(parts(&saved), parts(&source));

    let mismatched = DeckSession::open_from_update_with_source(&update, b"garbage", 14);
    assert!(mismatched.is_err());
}

#[test]
fn a_deck_reads_the_comments_it_was_opened_with() {
    let session = DeckSession::open(&commented_fixture(CommentFlavor::Legacy), 21).unwrap();
    let snapshot = session.snapshot().unwrap();
    assert_eq!(snapshot.comment_flavor, CommentFlavor::Legacy);
    assert_eq!(snapshot.comments.len(), 1);
    let comment = &snapshot.comments[0];
    assert_eq!(comment.text, "Needs a source.");
    assert_eq!(comment.author, "Mary Smith");
    assert_eq!(comment.initials, "mas");
    assert_eq!(comment.slide_id, snapshot.slides[0].id);
    assert_eq!(comment.x_emu, 914_400);
    assert_eq!(comment.y_emu, 457_200);
}

#[test]
fn a_modern_deck_reads_its_threads_and_flavour() {
    let session = DeckSession::open(&commented_fixture(CommentFlavor::Modern), 22).unwrap();
    let snapshot = session.snapshot().unwrap();
    assert_eq!(snapshot.comment_flavor, CommentFlavor::Modern);
    assert_eq!(snapshot.comments.len(), 2);
    let root = snapshot
        .comments
        .iter()
        .find(|comment| comment.parent_id.is_none())
        .unwrap();
    let reply = snapshot
        .comments
        .iter()
        .find(|comment| comment.parent_id.is_some())
        .unwrap();
    assert_eq!(root.text, "Needs a source.");
    assert_eq!(reply.text, "Added one.");
    assert_eq!(reply.parent_id.as_deref(), Some(root.id.as_str()));
}

#[test]
fn an_unrelated_edit_leaves_existing_comment_parts_byte_for_byte() {
    let source = commented_fixture(CommentFlavor::Legacy);
    let session = DeckSession::open(&source, 23).unwrap();
    let story = session.snapshot().unwrap().slides[1].shapes[0].text_stories[0]
        .id
        .clone();
    session
        .insert_text(&context(), &story, 0, "x", &TextStyle::default())
        .unwrap();

    let before = parts(&source);
    let after = parts(&session.save().unwrap());
    for path in ["ppt/commentAuthors.xml", "ppt/comments/comment1.xml"] {
        assert_eq!(before[path], after[path], "{path} was rewritten");
    }
    assert_relationships_resolve(&after);
}

#[test]
fn deleting_a_commented_slide_prunes_its_comment_part() {
    let session = DeckSession::open(&commented_fixture(CommentFlavor::Legacy), 24).unwrap();
    let first = session.snapshot().unwrap().slides[0].id.clone();
    session.delete_slide(&context(), &first).unwrap();

    let saved = parts(&session.save().unwrap());
    assert!(!saved.contains_key("ppt/comments/comment1.xml"));
    let content_types = part_text(&saved, "[Content_Types].xml");
    assert!(!content_types.contains("comments/comment1.xml"));
    assert_relationships_resolve(&saved);
}

#[test]
fn adding_a_comment_mints_the_part_relationship_and_override() {
    let session = open();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    session
        .add_comment(
            &context(),
            &slide,
            "Ada Lovelace",
            "AL",
            "Check this figure.",
            "2026-09-01T10:00:00.000",
            914_400,
            457_200,
        )
        .unwrap();

    let saved = parts(&session.save().unwrap());
    let comments = part_text(&saved, "ppt/comments/comment1.xml");
    assert!(comments.contains("<p:text>Check this figure.</p:text>"));
    assert!(comments.contains(r#"<p:pos x="576" y="288"/>"#));
    assert!(comments.contains(r#"authorId="0""#));
    assert!(comments.contains(r#"idx="1""#));

    let authors = part_text(&saved, "ppt/commentAuthors.xml");
    assert!(authors.contains(r#"name="Ada Lovelace""#));
    assert!(authors.contains(r#"initials="AL""#));
    assert!(authors.contains(r#"lastIdx="1""#));

    let content_types = part_text(&saved, "[Content_Types].xml");
    assert!(content_types.contains(
        "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"
    ));
    assert!(
        content_types
            .contains("application/vnd.openxmlformats-officedocument.presentationml.comments+xml")
    );
    let slide_rels = part_text(&saved, "ppt/slides/_rels/slide1.xml.rels");
    assert!(slide_rels.contains("Target=\"../comments/comment1.xml\""));
    assert_relationships_resolve(&saved);
}

#[test]
fn removing_the_last_comment_drops_the_part_and_its_bookkeeping() {
    let session = DeckSession::open(&commented_fixture(CommentFlavor::Legacy), 25).unwrap();
    let comment = session.snapshot().unwrap().comments[0].id.clone();
    session.remove_comment(&context(), &comment).unwrap();

    let saved = parts(&session.save().unwrap());
    assert!(!saved.contains_key("ppt/comments/comment1.xml"));
    assert!(!saved.contains_key("ppt/commentAuthors.xml"));
    let content_types = part_text(&saved, "[Content_Types].xml");
    assert!(!content_types.contains("commentAuthors.xml"));
    assert!(!content_types.contains("comments/comment1.xml"));
    assert!(!part_text(&saved, "ppt/slides/_rels/slide1.xml.rels").contains("comments"));
    assert_relationships_resolve(&saved);
}

#[test]
fn a_modern_thread_round_trips_through_a_save() {
    let session = DeckSession::open(&commented_fixture(CommentFlavor::Modern), 26).unwrap();
    let root = session
        .snapshot()
        .unwrap()
        .comments
        .iter()
        .find(|comment| comment.parent_id.is_none())
        .unwrap()
        .id
        .clone();
    session.set_comment_status(&context(), &root, true).unwrap();

    let saved = session.save().unwrap();
    let text = part_text(&parts(&saved), "ppt/comments/modernComment1.xml");
    assert!(text.contains(r#"status="resolved""#));
    assert!(text.contains("<p188:replyLst>"));
    assert!(text.contains("<a:t>Added one.</a:t>"));
    assert!(text.contains(r#"sldId="256""#));

    let reopened = DeckSession::open(&saved, 27).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.comment_flavor, CommentFlavor::Modern);
    assert_eq!(snapshot.comments.len(), 2);
    assert!(
        snapshot
            .comments
            .iter()
            .any(|comment| comment.resolved && comment.parent_id.is_none())
    );
    assert_relationships_resolve(&parts(&saved));
}

#[test]
fn replies_and_status_are_rejected_on_a_legacy_deck() {
    let session = DeckSession::open(&commented_fixture(CommentFlavor::Legacy), 28).unwrap();
    let comment = session.snapshot().unwrap().comments[0].id.clone();
    assert!(matches!(
        session.reply_to_comment(
            &context(),
            &comment,
            "Ada",
            "AL",
            "no",
            "2026-09-01T10:00:00"
        ),
        Err(EditError::InvalidComment(_))
    ));
    assert!(matches!(
        session.set_comment_status(&context(), &comment, true),
        Err(EditError::InvalidComment(_))
    ));
}

#[test]
fn the_comment_flavour_is_fixed_once_a_deck_has_comments() {
    let session = DeckSession::open(&commented_fixture(CommentFlavor::Legacy), 29).unwrap();
    assert!(matches!(
        session.set_comment_flavor(&context(), CommentFlavor::Modern),
        Err(EditError::InvalidComment(_))
    ));
}

#[test]
fn switching_an_emptied_deck_to_modern_drops_the_legacy_parts() {
    let session = DeckSession::open(&commented_fixture(CommentFlavor::Legacy), 30).unwrap();
    let comment = session.snapshot().unwrap().comments[0].id.clone();
    session.remove_comment(&context(), &comment).unwrap();
    session
        .set_comment_flavor(&context(), CommentFlavor::Modern)
        .unwrap();
    let slide = session.snapshot().unwrap().slides[0].id.clone();
    let root = session
        .add_comment(
            &context(),
            &slide,
            "Ada Lovelace",
            "AL",
            "Threaded now.",
            "2026-09-01T10:00:00.000",
            0,
            0,
        )
        .unwrap();
    session
        .reply_to_comment(
            &context(),
            &root.comment_id,
            "Grace Hopper",
            "GH",
            "Agreed.",
            "2026-09-01T10:05:00.000",
        )
        .unwrap();

    let saved = parts(&session.save().unwrap());
    assert!(!saved.contains_key("ppt/comments/comment1.xml"));
    assert!(!saved.contains_key("ppt/commentAuthors.xml"));
    let modern = part_text(&saved, "ppt/comments/modernComment1.xml");
    assert!(modern.contains("<a:t>Threaded now.</a:t>"));
    assert!(modern.contains("<a:t>Agreed.</a:t>"));
    let content_types = part_text(&saved, "[Content_Types].xml");
    assert!(content_types.contains("application/vnd.ms-powerpoint.comments+xml"));
    assert!(content_types.contains("application/vnd.ms-powerpoint.authors+xml"));
    assert!(!content_types.contains("presentationml.comments+xml"));
    assert!(!content_types.contains("presentationml.commentAuthors+xml"));
    assert_relationships_resolve(&saved);
}

#[test]
fn a_minted_comment_part_never_overwrites_another_slides() {
    let mut sources = fixture_parts(256);
    for (path, body) in sources.iter_mut() {
        match path.as_str() {
            "[Content_Types].xml" => {
                *body = body.replace(
                    "</Types>",
                    r#"<Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/><Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/></Types>"#,
                );
            }
            "ppt/_rels/presentation.xml.rels" => {
                *body = body.replace(
                    "</Relationships>",
                    &format!("{LEGACY_AUTHORS_REL}</Relationships>"),
                );
            }
            "ppt/slides/_rels/slide2.xml.rels" => {
                *body = body.replace(
                    "</Relationships>",
                    &format!("{LEGACY_COMMENTS_REL}</Relationships>"),
                );
            }
            _ => {}
        }
    }
    sources.push((
        "ppt/commentAuthors.xml".to_owned(),
        LEGACY_AUTHORS.to_owned(),
    ));
    sources.push((
        "ppt/comments/comment1.xml".to_owned(),
        LEGACY_COMMENTS.to_owned(),
    ));

    let session = DeckSession::open(&zip(sources), 31).unwrap();
    let snapshot = session.snapshot().unwrap();
    assert_eq!(
        snapshot.comments.len(),
        1,
        "slide 2 starts with the comment"
    );
    assert_eq!(snapshot.comments[0].slide_id, snapshot.slides[1].id);

    let first = snapshot.slides[0].id.clone();
    session
        .add_comment(
            &context(),
            &first,
            "Ada Lovelace",
            "AL",
            "First slide, new comment.",
            "2026-09-02T10:00:00.000",
            0,
            0,
        )
        .unwrap();

    let saved = session.save().unwrap();
    let back = DeckSession::open(&saved, 32).unwrap();
    let comments = back.snapshot().unwrap();
    assert_eq!(
        comments.comments.len(),
        2,
        "both slides keep their own comment"
    );
    let texts: Vec<&str> = comments
        .comments
        .iter()
        .map(|comment| comment.text.as_str())
        .collect();
    assert!(
        texts.contains(&"Needs a source."),
        "slide 2's comment was lost: {texts:?}"
    );
    assert!(texts.contains(&"First slide, new comment."));
    assert_relationships_resolve(&parts(&saved));
}
