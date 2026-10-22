//! Write-back fidelity against a deck exercising markup the model does not
//! carry: connectors, inherited placeholder rects, hyperlinks, fields,
//! theme colours, notes slides, custom shows, and hostile inputs.

use std::collections::BTreeMap;

use pptx_edit::{
    CommentFlavor, DeckSession, EditCtx, EditError, PresetShapeDraft, ShapeRect, TextStyle,
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
<p:sp><p:nvSpPr><p:cNvPr id="6" name="Script"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="3000" cy="400"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1200"/><a:t>1</a:t></a:r><a:r><a:rPr sz="1200" baseline="30000"/><a:t>st</a:t></a:r><a:r><a:rPr sz="1200"/><a:t> place, H</a:t></a:r><a:r><a:rPr sz="1200" baseline="-25000"/><a:t>2</a:t></a:r><a:r><a:rPr sz="1200"/><a:t>O</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="7" name="Tracked"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr spc="300"/><a:t>Wide</a:t></a:r><a:r><a:rPr spc="300" b="1"/><a:t>Caps</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="8" name="Linked"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="3000" cy="400"/></a:xfrm></p:spPr>
<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>See </a:t></a:r><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>the docs</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> today </a:t></a:r><a:fld id="{5C2A3F1E-8B7D-4C6A-9E0F-1A2B3C4D5E6F}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld></a:p></p:txBody></p:sp>
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

#[test]
fn replacing_paragraphs_preserves_properties_and_end_run_markup() {
    let properties = r#"<a:pPr algn="r" lvl="2" marL="3000" indent="-1000"><a:lnSpc><a:spcPct val="120000"/></a:lnSpc><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buNone/></a:pPr>"#;
    let end = r#"<a:endParaRPr lang="de-DE" sz="1700"/>"#;
    let mut source = fixture_parts(256);
    let slide = source
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide2.xml")
        .unwrap();
    slide.1 = slide.1.replace("<a:p><a:r><a:t>Second</a:t></a:r></a:p>", &format!("<a:p>{properties}<a:r><a:rPr b=\"1\" lang=\"en-US\" strike=\"sngStrike\"/><a:t>Second</a:t></a:r>{end}</a:p>"));
    let deck = DeckSession::open(&zip(source), 11).unwrap();
    let story = deck.snapshot().unwrap().slides[1].shapes[0].text_stories[0].clone();
    deck.set_story_paragraphs(
        &context(),
        &story.id,
        &[pptx_edit::TextParagraphDraft {
            bullet: None,
            alignment: None,
            runs: vec![pptx_edit::TextRunDraft {
                text: "Replacement".into(),
                style: story.paragraphs[0].runs[0].style.clone(),
            }],
        }],
    )
    .unwrap();
    assert_eq!(
        deck.story(&story.id).unwrap().paragraphs[0].id,
        story.paragraphs[0].id
    );
    let xml = part_text(&parts(&deck.save().unwrap()), "ppt/slides/slide2.xml");
    for markup in [
        "algn=\"r\"",
        "lvl=\"2\"",
        "marL=\"3000\"",
        "indent=\"-1000\"",
        "<a:lnSpc><a:spcPct val=\"120000\"/></a:lnSpc>",
        "<a:spcBef><a:spcPts val=\"500\"/></a:spcBef>",
        "<a:buNone/>",
    ] {
        assert!(xml.contains(markup), "{markup}");
    }
    assert!(xml.contains("<a:endParaRPr lang=\"de-DE\" sz=\"1700\"/>"));
    assert!(xml.contains("lang=\"en-US\""));
    assert!(xml.contains("strike=\"sngStrike\""));
    assert!(xml.contains("Replacement"));
    assert!(deck.undo());
    assert_eq!(deck.story(&story.id).unwrap(), story);
}

fn replacement_source(texts: &[&str]) -> Vec<u8> {
    let mut source = fixture_parts(256);
    let slide = source
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide2.xml")
        .unwrap();
    let paragraphs = texts
        .iter()
        .map(|text| format!("<a:p><a:r><a:t>{text}</a:t></a:r></a:p>"))
        .collect::<String>();
    slide.1 = slide
        .1
        .replace("<a:p><a:r><a:t>Second</a:t></a:r></a:p>", &paragraphs);
    zip(source)
}

fn replacement_story(deck: &DeckSession) -> pptx_edit::StorySnapshot {
    deck.snapshot().unwrap().slides[1].shapes[0].text_stories[0].clone()
}

#[test]
fn joining_paragraphs_preserves_each_source_run_and_last_end_properties() {
    let paragraphs = [
        r#"<a:p><a:pPr algn="ctr" marL="1234"/><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>First link</a:t></a:r><a:r><a:rPr strike="sngStrike"/><a:t> first tail</a:t></a:r><a:endParaRPr lang="en-US" sz="1100"/></a:p>"#,
        r#"<a:p><a:pPr algn="r"/><a:r><a:rPr lang="fr-FR"><a:hlinkClick r:id="rId3"/></a:rPr><a:t>Second link</a:t></a:r><a:r><a:rPr dirty="0"/><a:t> second tail</a:t></a:r><a:endParaRPr lang="de-DE" sz="1700"/></a:p>"#,
    ];
    let mut source = fixture_parts(256);
    let slide = source
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide2.xml")
        .unwrap();
    slide.1 = slide.1.replace(
        "<a:p><a:r><a:t>Second</a:t></a:r></a:p>",
        &paragraphs.concat(),
    );
    let deck = DeckSession::open(&zip(source), 13).unwrap();
    let story = replacement_story(&deck);
    let boundary = "First link first tail".len() as u32;
    deck.delete_text_joining_paragraphs(&context(), &story.id, boundary, boundary + 1)
        .unwrap();
    let saved = deck.save().unwrap();
    let xml = part_text(&parts(&saved), "ppt/slides/slide2.xml");
    let body = xml
        .split("<p:txBody>")
        .nth(1)
        .unwrap()
        .split("</p:txBody>")
        .next()
        .unwrap();
    assert_eq!(body.matches("<a:p>").count(), 1);
    let runs: Vec<_> = body
        .split("<a:r>")
        .skip(1)
        .map(|run| run.split("</a:r>").next().unwrap())
        .collect();
    assert_eq!(runs.len(), 4);
    for (run, (text, markup, linked)) in runs.iter().zip([
        ("First link", r#"lang="en-US""#, true),
        (" first tail", r#"strike="sngStrike""#, false),
        ("Second link", r#"lang="fr-FR""#, true),
        (" second tail", r#"dirty="0""#, false),
    ]) {
        assert!(run.contains(&format!("<a:t>{text}</a:t>")), "{run}");
        assert!(run.contains(markup), "{run}");
        assert_eq!(run.contains("hlinkClick"), linked, "{run}");
    }
    assert!(runs[0].contains(r#"<a:hlinkClick r:id="rId2"/>"#));
    assert!(runs[2].contains(r#"<a:hlinkClick r:id="rId3"/>"#));
    assert!(body.contains(r#"<a:endParaRPr lang="de-DE" sz="1700"/>"#));
    assert!(!body.contains(r#"<a:endParaRPr lang="en-US""#));
    assert!(xml.contains("marL=\"1234\""));
    assert_eq!(
        replacement_story(&DeckSession::open(&saved, 14).unwrap()).plain_text(),
        "First link first tailSecond link second tail"
    );
    assert!(deck.undo());
    assert_eq!(replacement_story(&deck), story);
    assert!(deck.redo());
    assert_eq!(deck.save().unwrap(), saved);
    let split = DeckSession::open(&saved, 20).unwrap();
    split
        .insert_paragraph_break(&context(), &story.id, 0)
        .unwrap();
    split
        .delete_paragraph_break(&context(), &story.id, 0)
        .unwrap();
    let rejoined = part_text(&parts(&split.save().unwrap()), "ppt/slides/slide2.xml");
    assert!(rejoined.contains(r#"<a:hlinkClick r:id="rId2"/>"#));
    assert!(rejoined.contains(r#"<a:hlinkClick r:id="rId3"/>"#));
    assert!(rejoined.contains(r#"<a:endParaRPr lang="de-DE" sz="1700"/>"#));
}

#[test]
fn joining_a_split_paragraph_preserves_the_styled_paragraph_mark() {
    let end = r#"<a:endParaRPr lang="de-DE" sz="1700"/>"#;
    let mut source = fixture_parts(256);
    let slide = source
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide2.xml")
        .unwrap();
    slide.1 = slide.1.replace(
        "<a:p><a:r><a:t>Second</a:t></a:r></a:p>",
        &format!(
            "<a:p><a:r><a:t>First</a:t></a:r>{end}</a:p><a:p><a:r><a:t>Second</a:t></a:r></a:p>"
        ),
    );
    let deck = DeckSession::open(&zip(source), 13).unwrap();
    let story = replacement_story(&deck);
    deck.insert_paragraph_break(&context(), &story.id, 9)
        .unwrap();
    deck.delete_paragraph_break(&context(), &story.id, 5)
        .unwrap();
    assert_eq!(replacement_story(&deck).plain_text(), "FirstSec\nond");

    let saved = deck.save().unwrap();
    let xml = part_text(&parts(&saved), "ppt/slides/slide2.xml");
    let first = xml.split("<a:p>").nth(1).unwrap();
    let first = first.split("</a:p>").next().unwrap();
    assert_eq!(first.matches(end).count(), 1, "{first}");
    assert!(first.ends_with(end), "{first}");
    assert_eq!(
        replacement_story(&DeckSession::open(&saved, 14).unwrap()).plain_text(),
        "FirstSec\nond"
    );
}

#[test]
fn plain_delete_batch_delete_and_proposals_refuse_paragraph_joins() {
    for field in [false, true] {
        let mut source = fixture_parts(256);
        let slide = source
            .iter_mut()
            .find(|(path, _)| path == "ppt/slides/slide2.xml")
            .unwrap();
        let second = if field {
            r#"<a:fld id="{D038279B-FC19-497E-A7D1-5ADD9CAF016F}" type="slidenum"><a:t>1</a:t></a:fld>"#
        } else {
            "<a:r><a:t>Second</a:t></a:r>"
        };
        slide.1 = slide.1.replace(
            "<a:p><a:r><a:t>Second</a:t></a:r></a:p>",
            &format!("<a:p><a:r><a:t>First</a:t></a:r></a:p><a:p>{second}</a:p>"),
        );
        let bytes = zip(source);
        let deck = DeckSession::open(&bytes, 15).unwrap();
        let story = replacement_story(&deck);
        let before = deck.encode_state_as_update_v1();
        assert!(matches!(
            deck.delete_text(&context(), &story.id, 5, 6),
            Err(EditError::ParagraphBoundary { .. })
        ));
        let request: pptx_edit::EditRequest = serde_json::from_value(serde_json::json!({
            "expectVersion": deck.version(),
            "steps": [{"op": "deleteText", "target": {"kind": "range", "slideId": deck.snapshot().unwrap().slides[1].id, "shapeId": deck.snapshot().unwrap().slides[1].shapes[0].id, "storyId": story.id, "start": 5, "end": 6}}]
        })).unwrap();
        assert!(deck.apply_edits(&request).unwrap().is_err());
        let proposal = deck.propose(pptx_edit::ProposalRequest {
            agent_id: "editor".into(),
            note: None,
            edits: vec![pptx_edit::ProposalEdit::ReplaceText {
                story_id: story.id.clone(),
                start: 5,
                end: 6,
                text: String::new(),
                style: None,
            }],
        });
        assert!(proposal.is_err());
        assert_eq!(deck.encode_state_as_update_v1(), before);
        assert!(!deck.can_undo());
    }
}

#[test]
fn deleting_a_paragraph_break_keeps_first_properties() {
    let deck = DeckSession::open(&replacement_source(&["First", "Second"]), 16).unwrap();
    let story = replacement_story(&deck);
    deck.set_paragraph_alignment(&context(), &story.id, 0, 0, Some("ctr"))
        .unwrap();
    deck.set_paragraph_alignment(&context(), &story.id, 6, 6, Some("r"))
        .unwrap();
    let before = replacement_story(&deck);
    deck.add_undo_barrier();
    deck.delete_paragraph_break(&context(), &story.id, 5)
        .unwrap();
    let joined = replacement_story(&deck);
    assert_eq!(joined.plain_text(), "FirstSecond");
    assert_eq!(joined.paragraphs[0].alignment.as_deref(), Some("ctr"));
    assert_eq!(joined.paragraphs[0].id, before.paragraphs[0].id);
    assert!(deck.undo());
    assert_eq!(replacement_story(&deck), before);
}

#[test]
fn splitting_a_joined_paragraph_keeps_current_properties() {
    let mut source = fixture_parts(256);
    let slide = source
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide2.xml")
        .unwrap();
    slide.1 = slide.1.replace(
        "<a:p><a:r><a:t>Second</a:t></a:r></a:p>",
        concat!(
            r#"<a:p><a:pPr algn="ctr" lvl="1" marL="1234"><a:buChar char="*"/></a:pPr><a:r><a:t>First</a:t></a:r></a:p>"#,
            r#"<a:p><a:pPr algn="r" lvl="2"><a:buNone/></a:pPr><a:r><a:t>Second</a:t></a:r></a:p>"#,
        ),
    );
    let deck = DeckSession::open(&zip(source), 20).unwrap();
    let story = replacement_story(&deck);
    deck.delete_paragraph_break(&context(), &story.id, 5)
        .unwrap();
    deck.set_paragraph_alignment(&context(), &story.id, 0, 0, Some("just"))
        .unwrap();
    let joined = replacement_story(&deck);
    deck.add_undo_barrier();
    deck.insert_paragraph_break(&context(), &story.id, 8)
        .unwrap();
    let split = replacement_story(&deck);
    assert_eq!(split.plain_text(), "FirstSec\nond");
    let trailing = &split.paragraphs[1];
    assert_eq!(trailing.alignment, joined.paragraphs[0].alignment);
    assert_eq!(trailing.level, joined.paragraphs[0].level);
    assert_eq!(trailing.bullet_json, joined.paragraphs[0].bullet_json);
    assert_eq!(
        trailing.source_paragraphs,
        joined.paragraphs[0].source_paragraphs
    );
    assert_ne!(split.paragraphs[0].id, trailing.id);
    let saved = deck.save().unwrap();
    let xml = part_text(&parts(&saved), "ppt/slides/slide2.xml");
    assert_eq!(xml.matches(r#"marL="1234""#).count(), 2);
    let reopened = DeckSession::open(&saved, 21).unwrap();
    let reopened = replacement_story(&reopened);
    for paragraph in &reopened.paragraphs {
        assert_eq!(paragraph.alignment, joined.paragraphs[0].alignment);
        assert_eq!(paragraph.level, joined.paragraphs[0].level);
        assert_eq!(paragraph.bullet_json, joined.paragraphs[0].bullet_json);
    }
    assert!(deck.undo());
    assert_eq!(replacement_story(&deck), joined);
    assert!(deck.redo());
    assert_eq!(replacement_story(&deck), split);
}

#[test]
fn duplicating_a_joined_story_keeps_its_paragraph_properties_and_sources() {
    let deck = DeckSession::open(&replacement_source(&["First", "Second"]), 18).unwrap();
    let story = replacement_story(&deck);
    deck.set_paragraph_alignment(&context(), &story.id, 0, 0, Some("ctr"))
        .unwrap();
    deck.delete_paragraph_break(&context(), &story.id, 5)
        .unwrap();
    let joined = replacement_story(&deck);
    let slide = deck.snapshot().unwrap().slides[1].id.clone();
    deck.duplicate_slide(&context(), &slide, 2).unwrap();
    let copy = deck.snapshot().unwrap().slides[2].shapes[0].text_stories[0].clone();
    assert_eq!(copy.plain_text(), joined.plain_text());
    assert_eq!(copy.paragraphs[0].alignment, joined.paragraphs[0].alignment);
    assert_eq!(
        copy.paragraphs[0].source_paragraphs,
        joined.paragraphs[0].source_paragraphs
    );
    assert_eq!(copy.paragraphs[0].runs, joined.paragraphs[0].runs);
    let saved = deck.save().unwrap();
    let reopened = DeckSession::open(&saved, 19).unwrap();
    let copy = &reopened.snapshot().unwrap().slides[2].shapes[0].text_stories[0];
    assert_eq!(copy.plain_text(), joined.plain_text());
    assert_eq!(copy.paragraphs[0].alignment.as_deref(), Some("ctr"));
}

#[test]
fn explicit_joins_and_break_deletion_protect_fields_on_both_sides() {
    for field_first in [false, true] {
        for field_text in ["1", ""] {
            let mut source = fixture_parts(256);
            let slide = source
                .iter_mut()
                .find(|(path, _)| path == "ppt/slides/slide2.xml")
                .unwrap();
            let field = format!(
                r#"<a:p><a:fld id="{{D038279B-FC19-497E-A7D1-5ADD9CAF016F}}" type="slidenum"><a:t>{field_text}</a:t></a:fld></a:p>"#
            );
            let plain = "<a:p><a:r><a:t>Plain</a:t></a:r></a:p>";
            let paragraphs = if field_first {
                format!("{field}{plain}")
            } else {
                format!("{plain}{field}")
            };
            slide.1 = slide
                .1
                .replace("<a:p><a:r><a:t>Second</a:t></a:r></a:p>", &paragraphs);
            let deck = DeckSession::open(&zip(source), 17).unwrap();
            let story = replacement_story(&deck);
            let boundary = if field_first {
                field_text.len() as u32
            } else {
                5
            };
            let before = deck.encode_state_as_update_v1();
            assert!(matches!(
                deck.delete_text_joining_paragraphs(&context(), &story.id, boundary, boundary + 1),
                Err(EditError::ParagraphBoundary { .. })
            ));
            assert!(matches!(
                deck.delete_paragraph_break(&context(), &story.id, boundary),
                Err(EditError::ParagraphBoundary { .. })
            ));
            assert_eq!(deck.encode_state_as_update_v1(), before);
            assert!(!deck.can_undo());
        }
    }
}

fn replacement_drafts(texts: &[&str]) -> Vec<pptx_edit::TextParagraphDraft> {
    texts
        .iter()
        .map(|text| pptx_edit::TextParagraphDraft {
            alignment: None,
            bullet: None,
            runs: vec![pptx_edit::TextRunDraft {
                text: (*text).into(),
                style: TextStyle::default(),
            }],
        })
        .collect()
}

fn snapshot_drafts(story: &pptx_edit::StorySnapshot) -> Vec<pptx_edit::TextParagraphDraft> {
    story
        .paragraphs
        .iter()
        .map(|paragraph| pptx_edit::TextParagraphDraft {
            alignment: None,
            bullet: None,
            runs: paragraph
                .runs
                .iter()
                .map(|run| pptx_edit::TextRunDraft {
                    text: run.text.clone(),
                    style: run.style.clone(),
                })
                .collect(),
        })
        .collect()
}

#[test]
fn story_replacement_beside_fields_refuses_without_mutation() {
    for field_first in [false, true] {
        for field_text in ["1", ""] {
            let mut source = fixture_parts(256);
            let slide = source
                .iter_mut()
                .find(|(path, _)| path == "ppt/slides/slide2.xml")
                .unwrap();
            let field = format!(
                r#"<a:p><a:fld id="{{D038279B-FC19-497E-A7D1-5ADD9CAF016F}}" type="slidenum"><a:t>{field_text}</a:t></a:fld></a:p>"#
            );
            let plain = "<a:p><a:r><a:t>Plain</a:t></a:r></a:p>";
            let paragraphs = if field_first {
                format!("{field}{plain}")
            } else {
                format!("{plain}{field}")
            };
            slide.1 = slide
                .1
                .replace("<a:p><a:r><a:t>Second</a:t></a:r></a:p>", &paragraphs);
            let deck = DeckSession::open(&zip(source), 22).unwrap();
            let original = replacement_story(&deck);
            let drafts = snapshot_drafts(&original);
            let field_index = usize::from(!field_first);
            let before = deck.encode_state_as_update_v1();
            assert!(matches!(
                deck.set_story_paragraphs(
                    &context(),
                    &original.id,
                    &drafts[field_index..field_index + 1],
                ),
                Err(EditError::ParagraphBoundary { .. })
            ));
            assert_eq!(replacement_story(&deck), original);
            assert_eq!(deck.encode_state_as_update_v1(), before);
            assert!(!deck.can_undo());
        }
    }
}

fn sync_replacement_peers(a: &DeckSession, b: &DeckSession) {
    a.apply_update_v1(&b.encode_state_as_update_v1()).unwrap();
    b.apply_update_v1(&a.encode_state_as_update_v1()).unwrap();
}

fn undo_replacement_peers(a: &DeckSession, b: &DeckSession, b_first: bool) {
    let (first, second) = if b_first { (b, a) } else { (a, b) };
    assert!(first.undo());
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    assert!(second.undo());
    first
        .apply_update_v1(&second.encode_state_as_update_v1())
        .unwrap();
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
}

#[test]
fn story_replacement_concurrent_undos_keep_original_markers() {
    let source = replacement_source(&["A", "B"]);
    for (a_id, b_id) in [(11, 12), (12, 11)] {
        for b_first in [false, true] {
            let a = DeckSession::open(&source, a_id).unwrap();
            let b = DeckSession::open(&source, b_id).unwrap();
            let original = replacement_story(&a);
            for (deck, texts) in [
                (&a, ["First left", "First right"]),
                (&b, ["Other left", "Other right"]),
            ] {
                deck.set_story_paragraphs(&context(), &original.id, &replacement_drafts(&texts))
                    .unwrap();
            }
            sync_replacement_peers(&a, &b);
            undo_replacement_peers(&a, &b, b_first);
            let current = a.story(&original.id).unwrap();
            assert_eq!(current, b.story(&original.id).unwrap());
            assert_eq!(
                current.paragraphs.len(),
                2,
                "clients {a_id}/{b_id}, B first {b_first}"
            );
            let ids = current.paragraphs.iter().map(|p| &p.id).collect::<Vec<_>>();
            assert_eq!(
                ids.iter().collect::<std::collections::HashSet<_>>().len(),
                2
            );
            assert_eq!(
                ids,
                original
                    .paragraphs
                    .iter()
                    .map(|p| &p.id)
                    .collect::<Vec<_>>()
            );
            for deck in [&a, &b] {
                let reopened = DeckSession::open(&deck.save().unwrap(), 13).unwrap();
                assert_eq!(reopened.story(&original.id).unwrap().paragraphs.len(), 2);
            }
        }
    }
}

#[test]
fn story_replacement_concurrent_undos_match_text_primitives() {
    let source = replacement_source(&["A", "B"]);
    for (a_id, b_id) in [(11, 12), (12, 11)] {
        for b_first in [false, true] {
            let mut final_texts = Vec::new();
            for primitives in [false, true] {
                let a = DeckSession::open(&source, a_id).unwrap();
                let b = DeckSession::open(&source, b_id).unwrap();
                let original = replacement_story(&a);
                for (deck, texts) in [
                    (&a, ["First left", "First right"]),
                    (&b, ["Other left", "Other right"]),
                ] {
                    if primitives {
                        deck.set_undo_capture_mode(pptx_edit::UndoCaptureMode::Manual);
                        deck.add_undo_barrier();
                        for (index, text) in texts.iter().enumerate().rev() {
                            let start = index as u32 * 2;
                            deck.delete_text(&context(), &original.id, start, start + 1)
                                .unwrap();
                            deck.insert_text(
                                &context(),
                                &original.id,
                                start,
                                text,
                                &TextStyle::default(),
                            )
                            .unwrap();
                        }
                        deck.add_undo_barrier();
                    } else {
                        deck.set_story_paragraphs(
                            &context(),
                            &original.id,
                            &replacement_drafts(&texts),
                        )
                        .unwrap();
                    }
                }
                sync_replacement_peers(&a, &b);
                undo_replacement_peers(&a, &b, b_first);
                let current = a.story(&original.id).unwrap();
                assert_eq!(current, b.story(&original.id).unwrap());
                final_texts.push(current.plain_text());
            }
            assert_eq!(
                final_texts[0], final_texts[1],
                "clients {a_id}/{b_id}, B first {b_first}"
            );
        }
    }
}

#[test]
fn story_replacement_preserves_surviving_paragraph_ids() {
    for (old, new, survivors) in [
        (
            vec!["A", "B", "C"],
            vec!["A", "X", "C"],
            vec![(0, 0), (1, 1), (2, 2)],
        ),
        (
            vec!["A", "B", "C", "D"],
            vec!["A", "C", "D"],
            vec![(0, 0), (2, 1), (3, 2)],
        ),
        (
            vec!["A", "B", "C"],
            vec!["A", "X", "B", "C"],
            vec![(0, 0), (1, 2), (2, 3)],
        ),
    ] {
        let deck = DeckSession::open(&replacement_source(&old), 11).unwrap();
        let original = replacement_story(&deck);
        let current = deck
            .set_story_paragraphs(&context(), &original.id, &replacement_drafts(&new))
            .unwrap();
        assert_eq!(current.plain_text(), new.join("\n"));
        assert_eq!(current.paragraphs.len(), new.len());
        for (before, after) in survivors {
            assert_eq!(current.paragraphs[after].id, original.paragraphs[before].id);
            if old[before] == new[after] {
                assert_eq!(current.paragraphs[after], original.paragraphs[before]);
            }
        }
    }
}

#[test]
fn story_replacement_noop_preserves_version_history_and_remote() {
    let source = replacement_source(&["A😀B", "Second"]);
    for (a_id, b_id) in [(11, 12), (12, 11)] {
        for existing_history in [false, true] {
            let a = DeckSession::open(&source, a_id).unwrap();
            let b = DeckSession::open(&source, b_id).unwrap();
            let story = replacement_story(&a);
            if existing_history {
                a.insert_text(&context(), &story.id, 0, "!", &TextStyle::default())
                    .unwrap();
                sync_replacement_peers(&a, &b);
            }
            let original = a.story(&story.id).unwrap();
            let version = a.version();
            let can_undo = a.can_undo();
            let update = a.encode_state_as_update_v1();
            let remote_version = b.version();
            assert_eq!(
                a.set_story_paragraphs(&context(), &story.id, &snapshot_drafts(&original))
                    .unwrap(),
                original
            );
            assert_eq!(a.version(), version);
            assert_eq!(a.can_undo(), can_undo);
            assert_eq!(a.encode_state_as_update_v1(), update);
            b.apply_update_v1(&a.encode_state_as_update_v1()).unwrap();
            assert_eq!(b.version(), remote_version);
            assert_eq!(b.story(&story.id).unwrap(), original);
        }
    }
}

#[test]
fn story_replacement_ignores_empty_runs_and_equal_run_boundaries() {
    let deck = DeckSession::open(&replacement_source(&["A😀B", "Second"]), 11).unwrap();
    let original = replacement_story(&deck);
    let mut drafts = snapshot_drafts(&original);
    drafts[0].runs = vec![
        pptx_edit::TextRunDraft {
            text: "A".into(),
            style: TextStyle::default(),
        },
        pptx_edit::TextRunDraft {
            text: String::new(),
            style: TextStyle {
                bold: Some(true),
                ..TextStyle::default()
            },
        },
        pptx_edit::TextRunDraft {
            text: "😀B".into(),
            style: TextStyle::default(),
        },
    ];
    let version = deck.version();
    assert_eq!(
        deck.set_story_paragraphs(&context(), &original.id, &drafts)
            .unwrap(),
        original
    );
    assert_eq!(deck.version(), version);
    assert!(!deck.can_undo());
}

#[test]
fn story_replacement_keeps_concurrent_insertions_in_untouched_text() {
    for (old, new, kept) in [
        (vec!["Alpha😀end", "Beta"], vec!["Alpha😀end", "X"], vec![0]),
        (
            vec!["Alpha😀end", "Beta"],
            vec!["Alpha😀end", "Beta", "C"],
            vec![0, 1],
        ),
        (
            vec!["Alpha😀end", "Beta", "C"],
            vec!["Alpha😀end", "Beta"],
            vec![0, 1],
        ),
    ] {
        for (a_id, b_id) in [(11, 12), (12, 11)] {
            for &index in &kept {
                let source = replacement_source(&old);
                let a = DeckSession::open(&source, a_id).unwrap();
                let b = DeckSession::open(&source, b_id).unwrap();
                let original = replacement_story(&a);
                let start = if index == 0 { 0 } else { 11 };
                a.insert_text(
                    &context(),
                    &original.id,
                    start + 2,
                    "!",
                    &TextStyle::default(),
                )
                .unwrap();
                let expected = a.story(&original.id).unwrap().paragraphs[index]
                    .runs
                    .clone();
                let mut drafts = replacement_drafts(&new);
                drafts[0].runs = snapshot_drafts(&original)[0].runs.clone();
                b.set_story_paragraphs(&context(), &original.id, &drafts)
                    .unwrap();
                sync_replacement_peers(&a, &b);
                let current = a.story(&original.id).unwrap();
                assert_eq!(current, b.story(&original.id).unwrap());
                assert_eq!(current.paragraphs[index].runs, expected);
                let mut expected_texts = new
                    .iter()
                    .map(|text| (*text).to_owned())
                    .collect::<Vec<_>>();
                expected_texts[index].insert(2, '!');
                assert_eq!(current.plain_text(), expected_texts.join("\n"));
            }
        }
    }
}

#[test]
fn story_replacement_keeps_concurrent_formatting_in_untouched_text() {
    let source = replacement_source(&["Alpha😀end", "Beta"]);
    for (a_id, b_id) in [(11, 12), (12, 11)] {
        let a = DeckSession::open(&source, a_id).unwrap();
        let b = DeckSession::open(&source, b_id).unwrap();
        let original = replacement_story(&a);
        a.format_text(
            &context(),
            &original.id,
            1,
            3,
            &pptx_edit::TextStylePatch {
                bold: Some(true),
                ..Default::default()
            },
        )
        .unwrap();
        let expected = a.story(&original.id).unwrap().paragraphs[0].clone();
        let mut drafts = snapshot_drafts(&original);
        drafts[1] = replacement_drafts(&["X"]).remove(0);
        b.set_story_paragraphs(&context(), &original.id, &drafts)
            .unwrap();
        sync_replacement_peers(&a, &b);
        let current = a.story(&original.id).unwrap();
        assert_eq!(current, b.story(&original.id).unwrap());
        assert_eq!(current.paragraphs[0], expected);
        assert_eq!(current.paragraphs[1].runs[0].text, "X");
    }
}

#[test]
fn story_replacement_single_undo_and_redo_restore_exact_snapshots() {
    for (old, new) in [
        (vec!["A", "B", "C"], vec!["A", "X", "C"]),
        (vec!["A", "B", "C", "D"], vec!["A", "C", "D"]),
        (vec!["A", "B", "C"], vec!["A", "X", "B", "C"]),
        (vec!["A", "B"], vec!["A", "B", "C"]),
        (vec!["A", "B", "C"], vec!["A", "B"]),
    ] {
        let deck = DeckSession::open(&replacement_source(&old), 11).unwrap();
        let original = replacement_story(&deck);
        let replaced = deck
            .set_story_paragraphs(&context(), &original.id, &replacement_drafts(&new))
            .unwrap();
        assert!(deck.undo());
        assert_eq!(deck.story(&original.id).unwrap(), original);
        assert!(!deck.can_undo());
        assert!(deck.redo());
        assert_eq!(deck.story(&original.id).unwrap(), replaced);
    }
}

#[test]
fn story_replacement_keeps_styled_prefix_and_suffix_characters() {
    let source = replacement_source(&["ab😀OLD🚀yz", "Last"]);
    for (a_id, b_id) in [(11, 12), (12, 11)] {
        let a = DeckSession::open(&source, a_id).unwrap();
        let b = DeckSession::open(&source, b_id).unwrap();
        let original = replacement_story(&a);
        a.insert_text(&context(), &original.id, 1, "!", &TextStyle::default())
            .unwrap();
        a.format_text(
            &context(),
            &original.id,
            10,
            12,
            &pptx_edit::TextStylePatch {
                bold: Some(true),
                ..Default::default()
            },
        )
        .unwrap();
        let mut drafts = snapshot_drafts(&original);
        drafts[0].runs = vec![
            pptx_edit::TextRunDraft {
                text: "ab😀".into(),
                style: TextStyle::default(),
            },
            pptx_edit::TextRunDraft {
                text: "NEW".into(),
                style: TextStyle {
                    bold: Some(true),
                    ..Default::default()
                },
            },
            pptx_edit::TextRunDraft {
                text: "🚀yz".into(),
                style: TextStyle::default(),
            },
        ];
        b.set_story_paragraphs(&context(), &original.id, &drafts)
            .unwrap();
        sync_replacement_peers(&a, &b);
        let current = a.story(&original.id).unwrap();
        assert_eq!(current, b.story(&original.id).unwrap());
        assert_eq!(current.plain_text(), "a!b😀NEW🚀yz\nLast");
        assert_eq!(
            current.paragraphs[0].runs,
            vec![
                pptx_edit::TextRunSnapshot {
                    text: "a!b😀".into(),
                    style: TextStyle::default(),
                },
                pptx_edit::TextRunSnapshot {
                    text: "NEW".into(),
                    style: TextStyle {
                        bold: Some(true),
                        ..Default::default()
                    },
                },
                pptx_edit::TextRunSnapshot {
                    text: "🚀".into(),
                    style: TextStyle::default(),
                },
                pptx_edit::TextRunSnapshot {
                    text: "yz".into(),
                    style: TextStyle {
                        bold: Some(true),
                        ..Default::default()
                    },
                },
            ]
        );
    }
}

#[test]
fn story_replacement_templates_preserve_distinct_paragraph_properties() {
    let mut source = fixture_parts(256);
    let slide = source
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide2.xml")
        .unwrap();
    let xml = [
        r#"<a:p><a:pPr algn="l" lvl="1" marL="1000"><a:buNone/></a:pPr><a:r><a:t>A</a:t></a:r><a:endParaRPr sz="1100"/></a:p>"#,
        r#"<a:p><a:pPr algn="r" lvl="2" marL="2000"><a:buChar char="*"/></a:pPr><a:r><a:t>B</a:t></a:r><a:endParaRPr sz="2200"/></a:p>"#,
        r#"<a:p><a:pPr algn="ctr" lvl="3" marL="3000"><a:buChar char="+"/></a:pPr><a:r><a:t>C</a:t></a:r><a:endParaRPr sz="3300"/></a:p>"#,
    ];
    slide.1 = slide
        .1
        .replace("<a:p><a:r><a:t>Second</a:t></a:r></a:p>", &xml.join(""));
    let source = zip(source);
    for (texts, templates) in [
        (vec!["A", "X", "B", "C"], vec![0, 0, 1, 2]),
        (vec!["X", "A", "B", "C"], vec![0, 0, 1, 2]),
        (vec!["A", "B", "C", "D"], vec![0, 1, 2, 2]),
        (vec!["A", "B"], vec![0, 1]),
        (vec!["A"], vec![0]),
        (vec!["C"], vec![2]),
    ] {
        let deck = DeckSession::open(&source, 11).unwrap();
        let original = replacement_story(&deck);
        let current = deck
            .set_story_paragraphs(&context(), &original.id, &replacement_drafts(&texts))
            .unwrap();
        assert_eq!(current.plain_text(), texts.join("\n"));
        for (paragraph, &template) in current.paragraphs.iter().zip(&templates) {
            let expected = &original.paragraphs[template];
            assert_eq!(paragraph.level, expected.level);
            assert_eq!(paragraph.alignment, expected.alignment);
            assert_eq!(paragraph.bullet_json, expected.bullet_json);
        }
        let saved = deck.save().unwrap();
        let slide = part_text(&parts(&saved), "ppt/slides/slide2.xml");
        for index in 0..3 {
            let count = templates
                .iter()
                .filter(|&&template| template == index)
                .count();
            assert_eq!(
                slide.matches(&format!("marL=\"{}000\"", index + 1)).count(),
                count
            );
            assert_eq!(
                slide
                    .matches(&format!("<a:endParaRPr sz=\"{}00\"/>", (index + 1) * 11))
                    .count(),
                count
            );
        }
        let reopened = DeckSession::open(&saved, 12).unwrap();
        let reopened = reopened.story(&original.id).unwrap();
        assert_eq!(reopened.plain_text(), current.plain_text());
        for (paragraph, expected) in reopened.paragraphs.iter().zip(&current.paragraphs) {
            assert_eq!(paragraph.level, expected.level);
            assert_eq!(paragraph.alignment, expected.alignment);
            assert_eq!(paragraph.bullet_json, expected.bullet_json);
        }
        assert!(deck.undo());
        assert_eq!(deck.story(&original.id).unwrap(), original);
        assert!(deck.redo());
        assert_eq!(deck.story(&original.id).unwrap(), current);
    }
}

#[test]
fn concurrent_paragraph_replacements_keep_unique_ids_and_formatting_through_history() {
    let properties = r#"<a:pPr algn="r" lvl="2" marL="3000" indent="-1000"><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buNone/></a:pPr>"#;
    let end = r#"<a:endParaRPr lang="de-DE" sz="1700"/>"#;
    let mut source = fixture_parts(256);
    let slide = source
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide2.xml")
        .unwrap();
    slide.1 = slide.1.replace("<a:p><a:r><a:t>Second</a:t></a:r></a:p>", &format!("<a:p>{properties}<a:r><a:rPr b=\"1\" lang=\"en-US\" strike=\"sngStrike\"/><a:t>Second</a:t></a:r>{end}</a:p>"));
    let source = zip(source);
    let first = DeckSession::open(&source, 11).unwrap();
    let second = DeckSession::open(&source, 12).unwrap();
    let story = first.snapshot().unwrap().slides[1].shapes[0].text_stories[0].clone();
    for (deck, text) in [
        (&first, "First replacement"),
        (&second, "Second replacement"),
    ] {
        deck.set_story_paragraphs(
            &context(),
            &story.id,
            &[pptx_edit::TextParagraphDraft {
                bullet: None,
                alignment: None,
                runs: vec![pptx_edit::TextRunDraft {
                    text: text.into(),
                    style: story.paragraphs[0].runs[0].style.clone(),
                }],
            }],
        )
        .unwrap();
    }
    first
        .apply_update_v1(&second.encode_state_as_update_v1())
        .unwrap();
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        first.story(&story.id).unwrap(),
        second.story(&story.id).unwrap()
    );
    let check = |deck: &DeckSession| {
        let current = deck.story(&story.id).unwrap();
        let ids: std::collections::HashSet<_> = current.paragraphs.iter().map(|p| &p.id).collect();
        assert_eq!(ids.len(), current.paragraphs.len());
        let xml = part_text(&parts(&deck.save().unwrap()), "ppt/slides/slide2.xml");
        for markup in [
            "algn=\"r\"",
            "marL=\"3000\"",
            "indent=\"-1000\"",
            end,
            "lang=\"en-US\"",
            "strike=\"sngStrike\"",
        ] {
            if markup == "lang=\"en-US\"" || markup == "strike=\"sngStrike\"" {
                assert!(
                    xml.matches(markup).count() >= current.paragraphs.len(),
                    "{markup}"
                );
            } else {
                assert_eq!(
                    xml.matches(markup).count(),
                    current.paragraphs.len(),
                    "{markup}"
                );
            }
        }
    };
    check(&first);
    check(&second);
    assert!(first.undo());
    check(&first);
    assert!(first.redo());
    check(&first);
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    check(&second);
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
fn added_text_boxes_save_their_identity_without_changing_existing_shapes() {
    for text in ["", "Saved note"] {
        let session = open();
        let slide_id = session.snapshot().unwrap().slides[1].id.clone();
        let rect = ShapeRect {
            x: 100_000,
            y: 200_000,
            width: 3_000_000,
            height: 500_000,
        };
        session
            .add_text_box(
                &context(),
                &slide_id,
                &pptx_edit::ShapeDraft {
                    name: "Note".to_owned(),
                    rect,
                    text: text.to_owned(),
                    style: TextStyle::default(),
                },
            )
            .unwrap();
        session
            .add_shape(
                &context(),
                &slide_id,
                &PresetShapeDraft {
                    name: "Text Box".to_owned(),
                    geometry: "rect".to_owned(),
                    rect,
                    fill: None,
                },
            )
            .unwrap();

        let saved = session.save().unwrap();
        let saved_parts = parts(&saved);
        let slide = part_text(&saved_parts, "ppt/slides/slide2.xml");
        let mut added = slide.rsplit("<p:sp>");
        let rectangle = added.next().unwrap();
        let text_box = added.next().unwrap();
        assert!(text_box.contains(r#"<p:cNvSpPr txBox="1"/>"#));
        assert!(text_box.contains(r#"<a:off x="100000" y="200000"/>"#));
        assert!(text_box.contains(r#"<a:ext cx="3000000" cy="500000"/>"#));
        assert!(text_box.contains("<a:bodyPr/>"));
        if !text.is_empty() {
            assert!(text_box.contains(&format!("<a:t>{text}</a:t>")));
        }
        assert!(rectangle.contains("<p:cNvSpPr/>"));
        assert!(!rectangle.contains("txBox="));
        for source in SLIDE2.split("<p:sp>").skip(1) {
            let (shape, _) = source.split_once("</p:sp>").unwrap();
            assert!(slide.contains(&format!("<p:sp>{shape}</p:sp>")));
        }
        for (path, bytes) in parts(&fixture(256)) {
            if path != "ppt/slides/slide2.xml" {
                assert_eq!(saved_parts.get(&path), Some(&bytes), "{path}");
            }
        }
        let reopened = DeckSession::open(&saved, 12).unwrap();
        assert_eq!(parts(&reopened.save().unwrap()), saved_parts);
    }
}

#[test]
fn a_remote_labelled_shape_without_a_text_box_marker_saves_as_a_shape() {
    use yrs::updates::decoder::Decode;
    use yrs::{Doc, Map, ReadTxn, Transact, Update};

    let source = open();
    let slide_id = source.snapshot().unwrap().slides[1].id.clone();
    let receipt = source
        .add_text_box(
            &context(),
            &slide_id,
            &pptx_edit::ShapeDraft {
                name: "Labelled ellipse".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 200_000,
                    width: 3_000_000,
                    height: 500_000,
                },
                text: "Ellipse label".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    let doc = Doc::new();
    let mut txn = doc.transact_mut();
    txn.apply_update(Update::decode_v1(&source.encode_state_as_update_v1()).unwrap())
        .unwrap();
    let shapes = txn.get_map("pptx:shapes").unwrap();
    let shape = shapes
        .get(&txn, &receipt.shape_id)
        .unwrap()
        .cast::<yrs::MapRef>()
        .unwrap();
    shape.remove(&mut txn, "textBox");
    shape.insert(&mut txn, "geometry", "ellipse");
    let update = txn.encode_state_as_update_v1(&Default::default());
    drop(txn);

    let peer = DeckSession::open(&fixture(256), 12).unwrap();
    peer.apply_update_v1(&update).unwrap();
    let snapshot = peer.snapshot().unwrap();
    let shape = snapshot.slides[1].shapes.last().unwrap();
    assert_eq!(shape.text_stories[0].plain_text(), "Ellipse label");
    let saved = parts(&peer.save().unwrap());
    let slide = part_text(&saved, "ppt/slides/slide2.xml");
    let added = slide.rsplit("<p:sp>").next().unwrap();
    assert!(added.contains(r#"<a:prstGeom prst="ellipse">"#));
    assert!(added.contains("<a:t>Ellipse label</a:t>"));
    assert!(!added.contains("txBox="));
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
fn an_added_picture_mints_its_media_part_content_type_and_relationship() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let png_bytes = vec![0x89, b'P', b'N', b'G', 0x0D, 0x0A, 0x1A, 0x0A, 1, 2, 3, 4];
    session
        .add_picture(
            &context(),
            &slide_id,
            &pptx_edit::PictureDraft {
                name: "Logo".to_owned(),
                rect: pptx_edit::ShapeRect {
                    x: 10,
                    y: 20,
                    width: 3_000,
                    height: 4_000,
                },
                content_type: "image/png".to_owned(),
                media_bytes: png_bytes.clone(),
            },
        )
        .unwrap();

    let saved = parts(&session.save().unwrap());
    assert_relationships_resolve(&saved);

    let content_types = part_text(&saved, "[Content_Types].xml");
    assert!(
        content_types.contains(r#"Extension="png""#),
        "{content_types}"
    );
    assert!(
        content_types.contains(r#"ContentType="image/png""#),
        "{content_types}"
    );

    let media = saved
        .get("ppt/media/image1.png")
        .expect("the picture's bytes land in a fresh media part");
    assert_eq!(media, &png_bytes);

    let slide_rels = part_text(&saved, "ppt/slides/_rels/slide1.xml.rels");
    assert!(
        slide_rels
            .contains("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"),
        "{slide_rels}"
    );
    assert!(
        slide_rels.contains(r#"Target="../media/image1.png""#),
        "{slide_rels}"
    );
    // The slide already had rId1 (layout) and rId2 (hyperlink); the image gets its own id.
    assert!(slide_rels.contains(r#"Id="rId3""#), "{slide_rels}");
    assert!(slide_rels.contains(r#"Id="rId1""#), "{slide_rels}");
    assert!(slide_rels.contains(r#"Id="rId2""#), "{slide_rels}");

    let slide = part_text(&saved, "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<p:pic>"#), "{slide}");
    assert!(slide.contains(r#"name="Logo""#), "{slide}");
    assert!(slide.contains(r#"r:embed="rId3""#), "{slide}");
    assert!(slide.contains(r#"<a:off x="10" y="20"/>"#), "{slide}");
    assert!(slide.contains(r#"<a:ext cx="3000" cy="4000"/>"#), "{slide}");
}

#[test]
fn an_unsupported_or_oversized_picture_is_rejected_before_it_touches_the_deck() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let rect = pptx_edit::ShapeRect {
        x: 0,
        y: 0,
        width: 100,
        height: 100,
    };

    let error = session
        .add_picture(
            &context(),
            &slide_id,
            &pptx_edit::PictureDraft {
                name: "Bad type".to_owned(),
                rect,
                content_type: "image/avif".to_owned(),
                media_bytes: vec![1, 2, 3, 4],
            },
        )
        .unwrap_err();
    assert!(matches!(error, EditError::InvalidState(_)), "{error:?}");

    let error = session
        .add_picture(
            &context(),
            &slide_id,
            &pptx_edit::PictureDraft {
                name: "Too big".to_owned(),
                rect,
                content_type: "image/png".to_owned(),
                media_bytes: vec![0; 8 * 1024 * 1024 + 1],
            },
        )
        .unwrap_err();
    assert!(matches!(error, EditError::InvalidState(_)), "{error:?}");

    // Neither rejected draft left a shape behind.
    assert_eq!(parts(&session.save().unwrap()), parts(&fixture(256)));
}

#[test]
fn shape_z_order_operations_reorder_within_the_slide() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = snapshot.slides[0].clone();
    let names: Vec<&str> = slide
        .shapes
        .iter()
        .map(|shape| shape.name.as_str())
        .collect();
    assert_eq!(
        names,
        [
            "Title",
            "Connector",
            "Box",
            "Halfway",
            "Script",
            "Tracked",
            "Linked"
        ]
    );
    let box_id = slide.shapes[2].id.clone();

    let names_after = |session: &DeckSession| -> Vec<String> {
        session.snapshot().unwrap().slides[0]
            .shapes
            .iter()
            .map(|shape| shape.name.clone())
            .collect()
    };

    let receipt = session
        .bring_to_front(&context(), &slide.id, &box_id)
        .unwrap();
    assert_eq!((receipt.from_index, receipt.to_index), (2, 6));
    assert_eq!(
        names_after(&session),
        [
            "Title",
            "Connector",
            "Halfway",
            "Script",
            "Tracked",
            "Linked",
            "Box"
        ]
    );

    let receipt = session
        .send_to_back(&context(), &slide.id, &box_id)
        .unwrap();
    assert_eq!((receipt.from_index, receipt.to_index), (6, 0));
    assert_eq!(
        names_after(&session),
        [
            "Box",
            "Title",
            "Connector",
            "Halfway",
            "Script",
            "Tracked",
            "Linked"
        ]
    );

    let receipt = session
        .bring_forward(&context(), &slide.id, &box_id)
        .unwrap();
    assert_eq!((receipt.from_index, receipt.to_index), (0, 1));
    assert_eq!(
        names_after(&session),
        [
            "Title",
            "Box",
            "Connector",
            "Halfway",
            "Script",
            "Tracked",
            "Linked"
        ]
    );

    let receipt = session
        .send_backward(&context(), &slide.id, &box_id)
        .unwrap();
    assert_eq!((receipt.from_index, receipt.to_index), (1, 0));
    assert_eq!(
        names_after(&session),
        [
            "Box",
            "Title",
            "Connector",
            "Halfway",
            "Script",
            "Tracked",
            "Linked"
        ]
    );

    // Already at the edge: stepping further is a clamped no-op.
    let receipt = session
        .send_backward(&context(), &slide.id, &box_id)
        .unwrap();
    assert_eq!((receipt.from_index, receipt.to_index), (0, 0));

    let saved = parts(&session.save().unwrap());
    let slide_xml = part_text(&saved, "ppt/slides/slide1.xml");
    let box_pos = slide_xml.find(r#"name="Box""#).unwrap();
    let title_pos = slide_xml.find(r#"name="Title""#).unwrap();
    assert!(box_pos < title_pos, "{slide_xml}");

    let error = session
        .bring_to_front(&context(), &slide.id, "missing-shape")
        .unwrap_err();
    assert!(matches!(error, EditError::ShapeNotFound(_)));
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

fn tracked_story(session: &DeckSession) -> String {
    session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Tracked")
        .unwrap()
        .text_stories[0]
        .id
        .clone()
}

#[test]
fn an_edit_inside_a_tracked_run_keeps_its_letter_spacing() {
    let session = open();
    let story_id = tracked_story(&session);
    session
        .insert_text(&context(), &story_id, 2, "X", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let slide = part_text(&parts(&saved), "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<a:rPr spc="300"/><a:t>Wi</a:t>"#));
    assert!(slide.contains(r#"<a:rPr spc="300"/><a:t>de</a:t>"#));
    assert!(slide.contains(r#"<a:rPr b="1" spc="300"/><a:t>Caps</a:t>"#));
    assert!(slide.contains(r#"<a:rPr/><a:t>X</a:t>"#));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    let story = reopened.story(&tracked_story(&reopened)).unwrap();
    let spacing = story.paragraphs[0]
        .runs
        .iter()
        .map(|run| (run.text.clone(), run.style.spacing_pt))
        .collect::<Vec<_>>();
    assert_eq!(
        spacing,
        vec![
            ("Wi".to_owned(), Some(3.0)),
            ("X".to_owned(), None),
            ("de".to_owned(), Some(3.0)),
            ("Caps".to_owned(), Some(3.0)),
        ]
    );
}

#[test]
fn an_edit_across_two_tracked_runs_rebuilds_their_letter_spacing() {
    let session = open();
    let story_id = tracked_story(&session);
    session.delete_text(&context(), &story_id, 3, 5).unwrap();

    let saved = session.save().unwrap();
    let slide = part_text(&parts(&saved), "ppt/slides/slide1.xml");
    assert!(slide.contains("<a:t>Wid</a:t>"));
    assert!(slide.contains("<a:t>aps</a:t>"));
    assert_eq!(slide.matches(r#"spc="300""#).count(), 2);

    let reopened = DeckSession::open(&saved, 12).unwrap();
    let story = reopened.story(&tracked_story(&reopened)).unwrap();
    assert_eq!(story.plain_text(), "Widaps");
    assert!(
        story.paragraphs[0]
            .runs
            .iter()
            .all(|run| run.style.spacing_pt == Some(3.0))
    );
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

fn script_story(session: &DeckSession) -> String {
    session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Script")
        .unwrap()
        .text_stories[0]
        .id
        .clone()
}

fn run_baselines(session: &DeckSession, story_id: &str) -> Vec<(String, Option<f64>)> {
    session.story(story_id).unwrap().paragraphs[0]
        .runs
        .iter()
        .map(|run| (run.text.clone(), run.style.baseline_pct))
        .collect()
}

#[test]
fn an_edit_inside_a_raised_run_keeps_its_baseline() {
    let session = open();
    let story_id = script_story(&session);
    session
        .insert_text(&context(), &story_id, 2, "X", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let slide = part_text(&parts(&saved), "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<a:rPr baseline="30000" sz="1200"/><a:t>s</a:t>"#));
    assert!(slide.contains(r#"<a:rPr/><a:t>X</a:t>"#));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert_eq!(
        run_baselines(&reopened, &script_story(&reopened)),
        vec![
            ("1".to_owned(), None),
            ("s".to_owned(), Some(30.0)),
            ("X".to_owned(), None),
            ("t".to_owned(), Some(30.0)),
            (" place, H".to_owned(), None),
            ("2".to_owned(), Some(-25.0)),
            ("O".to_owned(), None),
        ]
    );
}

#[test]
fn an_edit_spanning_two_runs_keeps_both_baselines() {
    let session = open();
    let story_id = script_story(&session);
    session.delete_text(&context(), &story_id, 2, 4).unwrap();

    let saved = session.save().unwrap();
    let slide = part_text(&parts(&saved), "ppt/slides/slide1.xml");
    assert!(slide.contains(r#"<a:rPr baseline="30000" sz="1200"/><a:t>s</a:t>"#));
    assert!(slide.contains(r#"baseline="-25000""#));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert_eq!(
        run_baselines(&reopened, &script_story(&reopened)),
        vec![
            ("1".to_owned(), None),
            ("s".to_owned(), Some(30.0)),
            ("place, H".to_owned(), None),
            ("2".to_owned(), Some(-25.0)),
            ("O".to_owned(), None),
        ]
    );
}

const LINK_RUN: &str = r#"<a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rId2"/></a:rPr>"#;
const SLIDE_NUMBER_FIELD: &str = r#"<a:fld id="{5C2A3F1E-8B7D-4C6A-9E0F-1A2B3C4D5E6F}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld>"#;

fn linked_story(session: &DeckSession) -> String {
    session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Linked")
        .unwrap()
        .text_stories[0]
        .id
        .clone()
}

fn linked_shape_xml(saved: &[u8]) -> String {
    part_text(&parts(saved), "ppt/slides/slide1.xml")
        .split(r#"name="Linked""#)
        .nth(1)
        .unwrap()
        .split("</p:sp>")
        .next()
        .unwrap()
        .to_owned()
}

fn linked_runs(saved: &[u8]) -> Vec<(String, Option<String>, Option<bool>)> {
    let package = pptx_parse::parse_pptx(saved).unwrap();
    let shape = package.slides[0]
        .shapes
        .iter()
        .find_map(|node| match node {
            pptx_parse::ShapeNode::Shape(shape) if shape.base.name == "Linked" => Some(shape),
            _ => None,
        })
        .unwrap();
    shape.text.as_ref().unwrap().paragraphs[0]
        .runs
        .iter()
        .map(|run| {
            (
                run.text.clone(),
                run.properties.hyperlink_relationship_id.clone(),
                run.properties.bold,
            )
        })
        .collect()
}

#[test]
fn an_edit_spanning_plain_and_linked_runs_keeps_the_link_and_the_field() {
    let session = open();
    let story_id = linked_story(&session);
    session.delete_text(&context(), &story_id, 3, 4).unwrap();
    session.delete_text(&context(), &story_id, 10, 11).unwrap();
    session.delete_text(&context(), &story_id, 11, 16).unwrap();
    session
        .insert_text(&context(), &story_id, 11, "now", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let shape = linked_shape_xml(&saved);
    assert!(
        shape.contains(r#"<a:r><a:rPr lang="en-US"/><a:t>See</a:t></a:r>"#),
        "{shape}"
    );
    assert!(
        shape.contains(&format!("{LINK_RUN}<a:t>the doc</a:t></a:r>")),
        "{shape}"
    );
    assert!(
        shape.contains(r#"<a:r><a:rPr lang="en-US"/><a:t> now </a:t></a:r>"#),
        "{shape}"
    );
    assert!(shape.contains(SLIDE_NUMBER_FIELD), "{shape}");
    assert_eq!(
        linked_runs(&saved),
        vec![
            ("See".to_owned(), None, None),
            ("the doc".to_owned(), Some("rId2".to_owned()), None),
            (" now ".to_owned(), None, None),
            ("1".to_owned(), None, None),
        ]
    );

    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert_eq!(
        reopened.story(&story_id).unwrap().plain_text(),
        "Seethe doc now 1"
    );
}

#[test]
fn a_field_in_an_edited_span_keeps_its_binding_until_its_text_changes() {
    let session = open();
    let story_id = linked_story(&session);
    session.delete_text(&context(), &story_id, 13, 18).unwrap();

    let saved = session.save().unwrap();
    let shape = linked_shape_xml(&saved);
    assert!(shape.contains(SLIDE_NUMBER_FIELD), "{shape}");
    assert!(
        shape.contains(&format!("{LINK_RUN}<a:t>the docs</a:t></a:r>")),
        "{shape}"
    );
    assert!(
        shape.contains(r#"<a:r><a:rPr lang="en-US"/><a:t>  </a:t></a:r>"#),
        "{shape}"
    );
    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert_eq!(
        reopened.story(&story_id).unwrap().plain_text(),
        "See the docs  1"
    );

    let session = open();
    let story_id = linked_story(&session);
    session.delete_text(&context(), &story_id, 19, 20).unwrap();
    session
        .insert_text(&context(), &story_id, 19, "2", &TextStyle::default())
        .unwrap();

    let saved = session.save().unwrap();
    let shape = linked_shape_xml(&saved);
    assert!(!shape.contains("<a:fld"), "{shape}");
    assert!(
        shape.contains(r#"<a:r><a:rPr lang="en-US"/><a:t>2</a:t></a:r>"#),
        "{shape}"
    );
    assert!(
        shape.contains(&format!("{LINK_RUN}<a:t>the docs</a:t></a:r>")),
        "{shape}"
    );
    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert_eq!(
        reopened.story(&story_id).unwrap().plain_text(),
        "See the docs today 2"
    );
}

#[test]
fn text_typed_at_the_end_of_a_link_stays_linked() {
    let session = open();
    let story_id = linked_story(&session);
    session
        .insert_text(
            &context(),
            &story_id,
            12,
            "X",
            &TextStyle {
                bold: Some(true),
                ..TextStyle::default()
            },
        )
        .unwrap();

    let saved = session.save().unwrap();
    let shape = linked_shape_xml(&saved);
    assert!(
        shape.contains(&format!(
            r#"{LINK_RUN}<a:t>the docs</a:t></a:r><a:r><a:rPr b="1" lang="en-US"><a:hlinkClick r:id="rId2"/></a:rPr><a:t>X</a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> today </a:t></a:r>"#
        )),
        "{shape}"
    );
    assert!(shape.contains(SLIDE_NUMBER_FIELD), "{shape}");
    assert_eq!(
        linked_runs(&saved),
        vec![
            ("See ".to_owned(), None, None),
            ("the docs".to_owned(), Some("rId2".to_owned()), None),
            ("X".to_owned(), Some("rId2".to_owned()), Some(true)),
            (" today ".to_owned(), None, None),
            ("1".to_owned(), None, None),
        ]
    );

    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert_eq!(
        reopened.story(&story_id).unwrap().plain_text(),
        "See the docsX today 1"
    );
}

#[test]
fn deleting_a_linked_run_drops_its_link() {
    let session = open();
    let story_id = linked_story(&session);
    session.delete_text(&context(), &story_id, 4, 12).unwrap();

    let saved = session.save().unwrap();
    let shape = linked_shape_xml(&saved);
    assert!(!shape.contains("hlinkClick"), "{shape}");
    assert!(
        shape.contains(
            r#"<a:r><a:rPr lang="en-US"/><a:t>See </a:t></a:r><a:r><a:rPr lang="en-US"/><a:t> today </a:t></a:r>"#
        ),
        "{shape}"
    );
    assert!(shape.contains(SLIDE_NUMBER_FIELD), "{shape}");

    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert_eq!(
        reopened.story(&story_id).unwrap().plain_text(),
        "See  today 1"
    );
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

const CUSTOM_SHAPE: &str = r#"<p:sp><p:nvSpPr><p:cNvPr id="8" name="Custom"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="300" cy="400"/></a:xfrm><a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="0" t="0" r="21600" b="21600"/><a:pathLst><a:path w="21600" h="21600"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="21600" y="0"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Custom</a:t></a:r></a:p></p:txBody></p:sp>"#;

fn cust_geom_fixture() -> Vec<u8> {
    let mut part_list = fixture_parts(256);
    for (path, body) in part_list.iter_mut() {
        if path == "ppt/slides/slide1.xml" {
            *body = body.replace("</p:spTree>", &format!("{CUSTOM_SHAPE}</p:spTree>"));
        }
    }
    zip(part_list)
}

#[test]
fn adjust_edits_on_custom_geometry_are_refused_and_leave_state_untouched() {
    let session = DeckSession::open(&cust_geom_fixture(), 11).unwrap();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let shape = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Custom")
        .unwrap();
    assert_eq!(shape.geometry, "custom");
    let before = session.snapshot().unwrap();
    let mut adjustments = BTreeMap::new();
    adjustments.insert("adj".to_owned(), 0.25);
    let error = session
        .set_shape_adjust(&context(), &slide.id, &shape.id, &adjustments)
        .unwrap_err();
    assert!(matches!(error, EditError::InvalidGeometry(_)));
    assert_eq!(session.snapshot().unwrap(), before);
}

#[test]
fn adjust_edits_on_shapes_without_preset_geometry_are_refused() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    let slide = &snapshot.slides[0];
    let missing = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Tracked")
        .unwrap();
    assert_eq!(missing.geometry, "rect");
    let before = session.snapshot().unwrap();
    let mut adjustments = BTreeMap::new();
    adjustments.insert("adj".to_owned(), 0.25);
    let error = session
        .set_shape_adjust(&context(), &slide.id, &missing.id, &adjustments)
        .unwrap_err();
    assert!(matches!(error, EditError::InvalidGeometry(_)));
    assert_eq!(session.snapshot().unwrap(), before);

    let preset = slide
        .shapes
        .iter()
        .find(|shape| shape.name == "Box")
        .unwrap();
    assert_eq!(preset.geometry, "roundRect");
    session
        .set_shape_adjust(&context(), &slide.id, &preset.id, &adjustments)
        .unwrap();
}

#[test]
fn an_exhausted_shape_id_space_errors_instead_of_panicking() {
    let mut part_list = fixture_parts(256);
    for (path, body) in part_list.iter_mut() {
        if path == "ppt/slides/slide1.xml" {
            *body = body.replace(
                r#"<p:cNvPr id="7" name="Tracked""#,
                r#"<p:cNvPr id="4294967295" name="Tracked""#,
            );
        }
    }
    let session = DeckSession::open(&zip(part_list), 11).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    session
        .add_shape(
            &context(),
            &slide_id,
            &PresetShapeDraft {
                name: "Overflow".to_owned(),
                geometry: "rect".to_owned(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                fill: None,
            },
        )
        .unwrap();
    let error = session.save().unwrap_err();
    assert!(matches!(error, EditError::Write(message) if message.contains("shape id")));
}

#[test]
fn an_exhausted_shape_id_space_still_saves_edits_that_allocate_nothing() {
    let mut part_list = fixture_parts(256);
    for (path, body) in part_list.iter_mut() {
        if path == "ppt/slides/slide1.xml" {
            *body = body.replace(
                r#"<p:cNvPr id="7" name="Tracked""#,
                r#"<p:cNvPr id="4294967295" name="Tracked""#,
            );
        }
    }
    let session = DeckSession::open(&zip(part_list.clone()), 11).unwrap();
    let story_id = tracked_story(&session);
    session
        .insert_text(&context(), &story_id, 0, "X", &TextStyle::default())
        .unwrap();
    let saved = session.save().unwrap();
    let slide = part_text(&parts(&saved), "ppt/slides/slide1.xml");
    assert!(slide.contains("<a:t>X</a:t>"), "{slide}");
    let reopened = DeckSession::open(&saved, 13).unwrap();
    assert_eq!(
        reopened
            .story(&tracked_story(&reopened))
            .unwrap()
            .plain_text(),
        "XWideCaps"
    );

    let session = DeckSession::open(&zip(part_list), 12).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    session
        .add_shape(
            &context(),
            &slide_id,
            &PresetShapeDraft {
                name: "Overflow".to_owned(),
                geometry: "rect".to_owned(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 1_000_000,
                    height: 1_000_000,
                },
                fill: None,
            },
        )
        .unwrap();
    let error = session.save().unwrap_err();
    assert!(
        matches!(error, EditError::Write(ref message) if message.contains("shape id space is exhausted")),
        "{error:?}"
    );
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

    let edited = session.snapshot().unwrap();
    let resized = edited.slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Halfway")
        .unwrap();
    assert_eq!(
        (resized.x, resized.y, resized.rotation_deg),
        (123_456, 654_321, 20.0)
    );

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

#[test]
fn notes_read_from_an_existing_part_default_to_empty() {
    let session = open();
    let snapshot = session.snapshot().unwrap();
    assert_eq!(snapshot.slides[1].notes, "");
}

#[test]
fn setting_notes_patches_an_existing_notes_part_in_place() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[1].id.clone();
    session
        .set_slide_notes(&context(), &slide_id, "Remember the demo flow.")
        .unwrap();

    let saved = session.save().unwrap();
    let notes_xml = part_text(&parts(&saved), "ppt/notesSlides/notesSlide1.xml");
    assert!(notes_xml.contains("Remember the demo flow."));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.slides[1].notes, "Remember the demo flow.");
    assert_relationships_resolve(&parts(&saved));
}

#[test]
fn setting_notes_mints_a_new_part_for_a_slide_that_has_none() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    session
        .set_slide_notes(&context(), &slide_id, "Line one\nLine two")
        .unwrap();

    let saved = session.save().unwrap();
    let saved_parts = parts(&saved);
    let notes_xml = part_text(&saved_parts, "ppt/notesSlides/notesSlide2.xml");
    assert!(notes_xml.contains("Line one"));
    assert!(notes_xml.contains("Line two"));
    let slide1_rels = part_text(&saved_parts, "ppt/slides/_rels/slide1.xml.rels");
    assert!(slide1_rels.contains("relationships/notesSlide"));
    let content_types = part_text(&saved_parts, "[Content_Types].xml");
    assert!(content_types.contains("notesSlide2.xml"));
    let notes_rels = part_text(&saved_parts, "ppt/notesSlides/_rels/notesSlide2.xml.rels");
    assert!(notes_rels.contains("relationships/slide\""));
    assert!(notes_rels.contains("Target=\"../slides/slide1.xml\""));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.slides[0].notes, "Line one\nLine two");
    assert_relationships_resolve(&parts(&saved));
}

#[test]
fn reverting_notes_does_not_mint_a_part() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    session
        .set_slide_notes(&context(), &slide_id, "Temporary note.")
        .unwrap();
    session.set_slide_notes(&context(), &slide_id, "").unwrap();

    let saved = session.save().unwrap();
    let saved_parts = parts(&saved);
    assert!(!saved_parts.contains_key("ppt/notesSlides/notesSlide2.xml"));

    let reopened = DeckSession::open(&saved, 12).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.slides[0].notes, "");
    assert_relationships_resolve(&parts(&saved));
}

#[test]
fn invalid_notes_leave_state_and_history_unchanged() {
    let session = open();
    let before = session.snapshot().unwrap();
    for text in ["bad\0note", "bad\u{b}note", "bad\u{ffff}note"] {
        assert!(matches!(
            session.set_slide_notes(&context(), &before.slides[0].id, text),
            Err(EditError::InvalidText(_))
        ));
        assert_eq!(session.snapshot().unwrap(), before);
        assert!(!session.can_undo());
    }
    assert!(matches!(
        session.set_slide_notes(&context(), "missing-slide", "notes"),
        Err(EditError::SlideNotFound(_))
    ));
}

#[test]
fn notes_roundtrip_xml_sensitive_text_and_support_history() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    let text = "  <&> \"notes\"\t\r\nSecond line 😀\n";
    session
        .set_slide_notes(&context(), &slide_id, text)
        .unwrap();
    let saved = session.save().unwrap();
    let reopened = DeckSession::open(&saved, 12).unwrap();
    assert_eq!(reopened.snapshot().unwrap().slides[0].notes, text);
    assert_eq!(
        parts(&saved)["ppt/slides/slide1.xml"],
        parts(&fixture(256))["ppt/slides/slide1.xml"]
    );
    assert!(session.undo());
    assert!(session.snapshot().unwrap().slides[0].notes.is_empty());
    assert!(session.redo());
    assert_eq!(session.snapshot().unwrap().slides[0].notes, text);
}

#[test]
fn clearing_saved_notes_preserves_the_notes_page_and_relationships() {
    let session = open();
    let slide_id = session.snapshot().unwrap().slides[1].id.clone();
    session
        .set_slide_notes(&context(), &slide_id, "Temporary notes")
        .unwrap();
    let saved = session.save().unwrap();
    let reopened = DeckSession::open(&saved, 12).unwrap();
    let id = reopened.snapshot().unwrap().slides[1].id.clone();
    reopened.set_slide_notes(&context(), &id, "").unwrap();
    let cleared = reopened.save().unwrap();
    let saved_parts = parts(&saved);
    let cleared_parts = parts(&cleared);
    assert_eq!(
        saved_parts.keys().collect::<Vec<_>>(),
        cleared_parts.keys().collect::<Vec<_>>()
    );
    for (path, bytes) in &saved_parts {
        if path != "ppt/notesSlides/notesSlide1.xml" {
            assert_eq!(&cleared_parts[path], bytes, "{path}");
        }
    }
    assert!(
        DeckSession::open(&cleared, 13)
            .unwrap()
            .snapshot()
            .unwrap()
            .slides[1]
            .notes
            .is_empty()
    );
    assert_relationships_resolve(&cleared_parts);
}

#[test]
fn notes_follow_inserted_and_reordered_slides_and_are_deleted_with_them() {
    let session = open();
    session.insert_slide(&context(), 1, None).unwrap();
    let id = session.snapshot().unwrap().slides[1].id.clone();
    session
        .set_slide_notes(&context(), &id, "New slide notes")
        .unwrap();
    session.move_slide(&context(), &id, 0).unwrap();
    let saved = session.save().unwrap();
    assert_relationships_resolve(&parts(&saved));
    let reopened = DeckSession::open(&saved, 12).unwrap();
    let snapshot = reopened.snapshot().unwrap();
    assert_eq!(snapshot.slides[0].notes, "New slide notes");
    assert!(snapshot.slides[1].notes.is_empty());
    reopened
        .delete_slide(&context(), &snapshot.slides[0].id)
        .unwrap();
    let deleted = reopened.save().unwrap();
    assert_relationships_resolve(&parts(&deleted));
    assert!(
        !parts(&deleted)
            .values()
            .any(|bytes| String::from_utf8_lossy(bytes).contains("New slide notes"))
    );
}

#[test]
fn slide_ids_stay_with_their_slides_across_reorder_insert_and_save() {
    let slide_list = |bytes: &[u8]| {
        let presentation = part_text(&parts(bytes), "ppt/presentation.xml");
        let start = presentation.find("<p:sldIdLst>").unwrap();
        let end = presentation.find("</p:sldIdLst>").unwrap();
        presentation[start..end].to_owned()
    };
    let source = fixture(300);
    let custom_show = r#"<p:sldLst><p:sld r:id="rId2"/><p:sld r:id="rId3"/></p:sldLst>"#;
    let mut first = None;
    for _ in 0..20 {
        let session = DeckSession::open(&source, 11).unwrap();
        let last = session.snapshot().unwrap().slides[1].id.clone();
        session.move_slide(&context(), &last, 0).unwrap();
        session.insert_slide(&context(), 1, None).unwrap();
        let saved = session.save().unwrap();
        let list = slide_list(&saved);
        assert!(
            list.starts_with(r#"<p:sldIdLst><p:sldId id="257" r:id="rId3"/><p:sldId id="301" "#)
        );
        assert!(list.ends_with(r#"<p:sldId id="300" r:id="rId2"/>"#));
        assert!(part_text(&parts(&saved), "ppt/presentation.xml").contains(custom_show));
        assert_eq!(first.get_or_insert_with(|| list.clone()), &list);

        let reopened = DeckSession::open(&saved, 12).unwrap();
        let moved = reopened.snapshot().unwrap().slides[2].id.clone();
        reopened.move_slide(&context(), &moved, 0).unwrap();
        let resaved = slide_list(&reopened.save().unwrap());
        let mut entries: Vec<&str> = list.split("<p:sldId ").skip(1).collect();
        let tail = entries.pop().unwrap();
        entries.insert(0, tail);
        assert_eq!(
            resaved,
            format!("<p:sldIdLst><p:sldId {}", entries.join("<p:sldId "))
        );
    }
}

#[test]
fn an_older_collaboration_document_recovers_notes_without_resurrecting_cleared_text() {
    use yrs::updates::decoder::Decode;
    use yrs::{Any, Doc, Map, Out, ReadTxn, Transact, Update};

    let initial = open();
    let id = initial.snapshot().unwrap().slides[0].id.clone();
    initial
        .set_slide_notes(&context(), &id, "Source notes")
        .unwrap();
    let source = initial.save().unwrap();
    let session = DeckSession::open(&source, 24).unwrap();
    let id = session.snapshot().unwrap().slides[0].id.clone();
    let doc = Doc::new();
    let mut txn = doc.transact_mut();
    txn.apply_update(Update::decode_v1(&session.encode_state_as_update_v1()).unwrap())
        .unwrap();
    let slides = txn.get_map("pptx:slides").unwrap();
    let slide = slides
        .get(&txn, &id)
        .unwrap()
        .cast::<yrs::MapRef>()
        .unwrap();
    slide.remove(&mut txn, "notes");
    let meta = txn.get_map("pptx:meta").unwrap();
    let Some(Out::Any(Any::Buffer(bytes))) = meta.get(&txn, "packageJson") else {
        panic!("package")
    };
    let mut package: serde_json::Value = serde_json::from_slice(&bytes).unwrap();
    for slide in package["slides"].as_array_mut().unwrap() {
        slide.as_object_mut().unwrap().remove("notes");
    }
    meta.insert(
        &mut txn,
        "packageJson",
        Any::Buffer(std::sync::Arc::from(serde_json::to_vec(&package).unwrap())),
    );
    let update = txn.encode_state_as_update_v1(&Default::default());
    drop(txn);
    let restored = DeckSession::open_from_update_with_source(&update, &source, 25).unwrap();
    assert_eq!(restored.snapshot().unwrap().slides[0].notes, "Source notes");
    let detached =
        DeckSession::open_from_update(&restored.encode_state_as_update_v1(), 26).unwrap();
    assert_eq!(detached.snapshot().unwrap().slides[0].notes, "Source notes");
    restored.set_slide_notes(&context(), &id, "").unwrap();
    let cleared = DeckSession::open_from_update_with_source(
        &restored.encode_state_as_update_v1(),
        &source,
        27,
    )
    .unwrap();
    assert!(cleared.snapshot().unwrap().slides[0].notes.is_empty());
    assert!(
        DeckSession::open(&cleared.save().unwrap(), 28)
            .unwrap()
            .snapshot()
            .unwrap()
            .slides[0]
            .notes
            .is_empty()
    );
}

fn picture_draft() -> pptx_edit::PictureDraft {
    pptx_edit::PictureDraft {
        name: "Inserted picture".to_owned(),
        rect: ShapeRect {
            x: 10,
            y: 20,
            width: 3000,
            height: 4000,
        },
        content_type: "image/png".to_owned(),
        media_bytes: vec![0x89, b'P', b'N', b'G', 13, 10, 26, 10],
    }
}

#[test]
fn concurrent_shape_arrangement_can_be_rearranged_and_deleted() {
    let first = open();
    let second = DeckSession::open(&fixture(256), 12).unwrap();
    let slide = first.snapshot().unwrap().slides.remove(0);
    let shape_id = &slide.shapes[2].id;
    first
        .bring_to_front(&context(), &slide.id, shape_id)
        .unwrap();
    second
        .send_to_back(&context(), &slide.id, shape_id)
        .unwrap();
    first
        .apply_update_v1(&second.encode_state_as_update_v1())
        .unwrap();
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
    first
        .bring_to_front(&context(), &slide.id, shape_id)
        .unwrap();
    assert_eq!(
        first.snapshot().unwrap().slides[0]
            .shapes
            .last()
            .unwrap()
            .id,
        *shape_id
    );
    first.remove_shape(&context(), &slide.id, shape_id).unwrap();
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
    first.save().unwrap();
}

#[test]
fn concurrent_shape_arrangement_and_deletion_converge() {
    let first = open();
    let second = DeckSession::open(&fixture(256), 12).unwrap();
    let slide = first.snapshot().unwrap().slides.remove(0);
    let shape_id = &slide.shapes[2].id;
    first
        .bring_to_front(&context(), &slide.id, shape_id)
        .unwrap();
    second
        .remove_shape(&context(), &slide.id, shape_id)
        .unwrap();
    first
        .apply_update_v1(&second.encode_state_as_update_v1())
        .unwrap();
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(first.snapshot().unwrap(), second.snapshot().unwrap());
    assert!(
        !first.snapshot().unwrap().slides[0]
            .shapes
            .iter()
            .any(|s| s.id == *shape_id)
    );
    first.save().unwrap();
    first.search_text("Title", false, None).unwrap();
    first.delete_slide(&context(), &slide.id).unwrap();
    first.save().unwrap();
}

#[test]
fn inserted_picture_preserves_conflicting_content_type_defaults() {
    let mut source = fixture_parts(256);
    source[0].1 = source[0].1.replace(
        "</Types>",
        r#"<Default Extension="png" ContentType="application/octet-stream"/></Types>"#,
    );
    let session = DeckSession::open(&zip(source), 11).unwrap();
    let slide = session.snapshot().unwrap().slides.remove(0);
    session
        .add_picture(&context(), &slide.id, &picture_draft())
        .unwrap();
    let saved = session.save().unwrap();
    let package = pptx_parse::parse_pptx(&saved).unwrap();
    assert_eq!(
        package
            .media
            .iter()
            .find(|m| m.part_path.ends_with(".png"))
            .unwrap()
            .content_type,
        "image/png"
    );
    assert!(
        part_text(&parts(&saved), "[Content_Types].xml")
            .contains(r#"ContentType="application/octet-stream""#)
    );
}

#[test]
fn pictures_on_existing_and_new_slides_survive_sync_undo_and_repeated_save() {
    let first = open();
    let second = DeckSession::open(&fixture(256), 12).unwrap();
    let slide = first.snapshot().unwrap().slides[0].id.clone();
    let added = first.insert_slide(&context(), 1, None).unwrap();
    first
        .add_picture(&context(), &slide, &picture_draft())
        .unwrap();
    first.add_undo_barrier();
    first
        .add_picture(&context(), &added.slide_id, &picture_draft())
        .unwrap();
    let expected = first.snapshot().unwrap();
    assert!(first.undo());
    assert!(first.snapshot().unwrap().slides[1].shapes.is_empty());
    assert!(first.redo());
    assert_eq!(first.snapshot().unwrap(), expected);
    second
        .apply_update_v1(&first.encode_state_as_update_v1())
        .unwrap();
    assert_eq!(second.snapshot().unwrap(), expected);
    let saved = first.save().unwrap();
    assert_eq!(parts(&saved), parts(&first.save().unwrap()));
    assert_eq!(parts(&saved), parts(&second.save().unwrap()));
    assert_relationships_resolve(&parts(&saved));
    let reopened = DeckSession::open(&saved, 13).unwrap();
    for slide in &reopened.snapshot().unwrap().slides[..2] {
        let picture = slide
            .shapes
            .iter()
            .find(|s| s.name == picture_draft().name)
            .unwrap();
        assert_eq!(picture.kind, pptx_edit::ShapeKind::Picture);
        assert!(picture.media_part_path.is_some());
    }
}

#[test]
fn picture_insertion_preserves_prefixed_metadata_and_exhausted_numeric_names() {
    let mut source = fixture_parts(256);
    for (path, xml) in &mut source {
        if path == "[Content_Types].xml" {
            *xml = xml
                .replace("xmlns=", "xmlns:ct=")
                .replace("<Types", "<ct:Types")
                .replace("</Types>", "</ct:Types>")
                .replace("<Default", "<ct:Default")
                .replace("<Override", "<ct:Override");
        }
        if path == "ppt/slides/_rels/slide1.xml.rels" {
            *xml = xml
                .replace("xmlns=", "xmlns:rel=")
                .replace("<Relationships", "<rel:Relationships")
                .replace("</Relationships>", "</rel:Relationships>")
                .replace("<Relationship ", "<rel:Relationship ")
                .replace("rId2", "rId18446744073709551615");
        }
        if path == "ppt/slides/slide1.xml" {
            *xml = xml
                .replace("rId2", "rId18446744073709551615")
                .replace("xmlns:r=", "xmlns:links=")
                .replace("r:id=", "links:id=")
                .replace("<p:sld ", r#"<p:sld xmlns:r="urn:preserved" "#);
        }
    }
    source.push((
        "ppt/media/image18446744073709551615.png".to_owned(),
        "existing".to_owned(),
    ));
    let session = DeckSession::open(&zip(source), 11).unwrap();
    let slide_id = session.snapshot().unwrap().slides[0].id.clone();
    session
        .add_picture(&context(), &slide_id, &picture_draft())
        .unwrap();
    let saved = session.save().unwrap();
    let package = pptx_parse::parse_pptx(&saved).unwrap();
    let image = package
        .media
        .iter()
        .find(|media| media.bytes == picture_draft().media_bytes)
        .unwrap();
    assert_eq!(image.content_type, "image/png");
    let saved_parts = parts(&saved);
    assert!(part_text(&saved_parts, "[Content_Types].xml").contains("<ct:Default"));
    assert!(
        part_text(&saved_parts, "ppt/slides/_rels/slide1.xml.rels")
            .contains("<rel:Relationship Id=\"rId2\"")
    );
    assert!(
        part_text(&saved_parts, "ppt/slides/slide1.xml").contains(r#"xmlns:r="urn:preserved""#)
    );
    assert_eq!(
        saved_parts["ppt/media/image18446744073709551615.png"],
        b"existing"
    );
}

#[test]
fn concurrent_replacements_keep_terminal_marker() {
    let properties = r#"<a:pPr algn="r" lvl="2" marL="3000" indent="-1000"><a:spcBef><a:spcPts val="500"/></a:spcBef><a:buNone/></a:pPr>"#;
    let end = r#"<a:endParaRPr lang="de-DE" sz="1700"/>"#;
    let mut source = fixture_parts(256);
    let slide = source
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide2.xml")
        .unwrap();
    let paragraph = format!(
        "<a:p>{properties}<a:r><a:rPr b=\"1\" lang=\"en-US\" strike=\"sngStrike\"/><a:t>Source</a:t></a:r>{end}</a:p>"
    );
    slide.1 = slide.1.replace(
        "<a:p><a:r><a:t>Second</a:t></a:r></a:p>",
        &paragraph.repeat(2),
    );
    let source = zip(source);
    for (a_id, b_id) in [(11, 12), (12, 11)] {
        for (a_count, b_count) in [(1, 2), (2, 1), (1, 3), (3, 2), (3, 3)] {
            let a = DeckSession::open(&source, a_id).unwrap();
            let b = DeckSession::open(&source, b_id).unwrap();
            let story = a.snapshot().unwrap().slides[1].shapes[0].text_stories[0].clone();
            for (deck, count, prefix) in [(&a, a_count, "A"), (&b, b_count, "B")] {
                let paragraphs = (0..count)
                    .map(|index| pptx_edit::TextParagraphDraft {
                        bullet: None,
                        alignment: None,
                        runs: vec![pptx_edit::TextRunDraft {
                            text: format!("{prefix} replacement {index}"),
                            style: story.paragraphs[0].runs[0].style.clone(),
                        }],
                    })
                    .collect::<Vec<_>>();
                deck.set_story_paragraphs(&context(), &story.id, &paragraphs)
                    .unwrap();
            }
            let a_update = a.encode_state_as_update_v1();
            let b_update = b.encode_state_as_update_v1();
            let a_result = a.apply_update_v1(&b_update);
            let b_result = b.apply_update_v1(&a_update);
            assert!(a_result.is_ok(), "A: {a_result:?}");
            assert!(b_result.is_ok(), "B: {b_result:?}");
            assert_eq!(a.story(&story.id).unwrap(), b.story(&story.id).unwrap());
            let merged_text = a.story(&story.id).unwrap().plain_text();
            for (count, prefix) in [(a_count, "A"), (b_count, "B")] {
                for index in 0..count {
                    assert!(merged_text.contains(&format!("{prefix} replacement {index}")));
                }
            }
            let check = |deck: &DeckSession| {
                let current = deck.story(&story.id).unwrap();
                let ids = current
                    .paragraphs
                    .iter()
                    .map(|p| &p.id)
                    .collect::<std::collections::HashSet<_>>();
                assert_eq!(ids.len(), current.paragraphs.len());
                assert_eq!(
                    current.paragraphs.last().unwrap().id,
                    story.paragraphs.last().unwrap().id
                );
                let xml = part_text(&parts(&deck.save().unwrap()), "ppt/slides/slide2.xml");
                for markup in ["algn=\"r\"", "marL=\"3000\"", "indent=\"-1000\"", end] {
                    assert_eq!(
                        xml.matches(markup).count(),
                        current.paragraphs.len(),
                        "{markup}, counts {a_count}/{b_count}"
                    );
                }
                for markup in ["lang=\"en-US\"", "strike=\"sngStrike\""] {
                    assert!(
                        xml.matches(markup).count() >= current.paragraphs.len(),
                        "{markup}"
                    );
                }
            };
            check(&a);
            check(&b);
            assert!(a.undo());
            check(&a);
            assert!(a.redo());
            check(&a);
            b.apply_update_v1(&a.encode_state_as_update_v1()).unwrap();
            check(&b);
            assert_eq!(a.story(&story.id).unwrap(), b.story(&story.id).unwrap());
        }
    }
}

fn layout_change_source() -> Vec<u8> {
    let mut source = fixture_parts(256);
    for (path, xml) in &mut source {
        if path == "ppt/slideMasters/_rels/slideMaster1.xml.rels" {
            *xml = xml.replace("</Relationships>", r#"<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout2.xml"/></Relationships>"#);
        }
        if path == "ppt/slideMasters/slideMaster1.xml" {
            *xml = xml.replace(
                "</p:sldLayoutIdLst>",
                r#"<p:sldLayoutId id="2147483650" r:id="rId3"/></p:sldLayoutIdLst>"#,
            );
        }
    }
    source.push((
        "ppt/slideLayouts/slideLayout2.xml".into(),
        LAYOUT
            .replace(
                "type=\"body\"",
                "type=\"obj\" idx=\"7\" orient=\"vert\" sz=\"half\"",
            )
            .replace("type=\"title\"", "type=\"title\" idx=\"9\""),
    ));
    source.push((
        "ppt/slideLayouts/_rels/slideLayout2.xml.rels".into(),
        LAYOUT_RELS.into(),
    ));
    let body_start = LAYOUT.find("<p:sp><p:nvSpPr><p:cNvPr id=\"3\"").unwrap();
    let body_end = body_start + LAYOUT[body_start..].find("</p:sp>").unwrap() + "</p:sp>".len();
    let title_only = format!("{}{}", &LAYOUT[..body_start], &LAYOUT[body_end..]);
    source.push(("ppt/slideLayouts/slideLayout3.xml".into(), title_only));
    source.push((
        "ppt/slideLayouts/_rels/slideLayout3.xml.rels".into(),
        LAYOUT_RELS.into(),
    ));
    source.push(("ppt/slideLayouts/slideLayout4.xml".into(), LAYOUT.into()));
    source.push((
        "ppt/slideLayouts/_rels/slideLayout4.xml.rels".into(),
        LAYOUT_RELS.replace("slideMaster1.xml", "slideMaster2.xml"),
    ));
    source.push(("ppt/slideMasters/slideMaster2.xml".into(), MASTER.into()));
    zip(source)
}

#[test]
fn set_slide_layout_maps_placeholders_and_preserves_untouched_bytes() {
    let source = layout_change_source();
    let deck = DeckSession::open(&source, 11).unwrap();
    let before = deck.snapshot().unwrap();
    let slide = &before.slides[0];
    let mapping = [
        (slide.shapes[0].id.clone(), 2),
        (slide.shapes[3].id.clone(), 3),
    ];
    deck.set_slide_layout(
        &context(),
        &slide.id,
        "ppt/slideLayouts/slideLayout2.xml",
        &mapping,
    )
    .unwrap();
    let changed = deck.snapshot().unwrap();
    assert_eq!(
        changed.slides[0].shapes[3]
            .placeholder
            .as_ref()
            .unwrap()
            .index,
        Some(7)
    );
    let saved = deck.save().unwrap();
    let original = parts(&source);
    let output = parts(&saved);
    for (path, bytes) in &original {
        if path != "ppt/slides/slide1.xml" && path != "ppt/slides/_rels/slide1.xml.rels" {
            assert_eq!(output.get(path), Some(bytes), "{path}");
        }
    }
    assert_eq!(
        part_text(&output, "ppt/slides/_rels/slide1.xml.rels"),
        SLIDE1_RELS.replace("slideLayout1.xml", "slideLayout2.xml")
    );
    let xml = part_text(&output, "ppt/slides/slide1.xml");
    assert_eq!(
        xml,
        SLIDE1
            .replace("<p:ph type=\"title\"/>", "<p:ph type=\"title\" idx=\"9\"/>")
            .replace(
                "<p:ph type=\"body\"/>",
                "<p:ph type=\"obj\" idx=\"7\" orient=\"vert\" sz=\"half\"/>"
            )
    );
    let reopened = DeckSession::open(&saved, 12).unwrap().snapshot().unwrap();
    assert_eq!(
        reopened.slides[0].layout_part_path,
        changed.slides[0].layout_part_path
    );
    assert_eq!(
        reopened.slides[0].shapes[0].placeholder,
        changed.slides[0].shapes[0].placeholder
    );
    assert!(deck.undo());
    assert_eq!(deck.snapshot().unwrap(), before);
    assert!(deck.redo());
    assert_eq!(deck.snapshot().unwrap(), changed);
}

#[test]
fn set_slide_layout_maps_nested_group_placeholders() {
    let mut source: Vec<_> = parts(&layout_change_source())
        .into_iter()
        .map(|(path, bytes)| (path, String::from_utf8(bytes).unwrap()))
        .collect();
    let (_, xml) = source
        .iter_mut()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap();
    let start = xml.find("<p:sp><p:nvSpPr>").unwrap();
    let end = start + xml[start..].find("</p:sp>").unwrap() + "</p:sp>".len();
    let mut grouped = xml[start..end].to_owned();
    for id in [20, 21] {
        grouped = format!(
            r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{id}" name="Group {id}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{grouped}</p:grpSp>"#
        );
    }
    xml.replace_range(start..end, &grouped);
    let source = zip(source);
    let deck = DeckSession::open(&source, 11).unwrap();
    let before = deck.snapshot().unwrap();
    let slide = &before.slides[0];
    let placeholder = &slide.shapes[0].children[0].children[0];
    let update = deck.encode_state_as_update_v1();
    let layout = "ppt/slideLayouts/slideLayout2.xml";
    for (slide_id, mapping) in [
        (&before.slides[1].id, vec![(placeholder.id.clone(), 2)]),
        (
            &slide.id,
            vec![(placeholder.id.clone(), 2), ("missing-shape".into(), 3)],
        ),
    ] {
        assert!(matches!(
            deck.set_slide_layout(&context(), slide_id, layout, &mapping),
            Err(EditError::ShapeNotFound(_))
        ));
        assert_eq!(deck.snapshot().unwrap(), before);
        assert_eq!(deck.encode_state_as_update_v1(), update);
    }
    deck.set_slide_layout(
        &context(),
        &slide.id,
        layout,
        &[(placeholder.id.clone(), 2)],
    )
    .unwrap();
    let changed = deck.snapshot().unwrap();
    let mapped = &changed.slides[0].shapes[0].children[0].children[0];
    assert_eq!(mapped.placeholder.as_ref().unwrap().index, Some(9));
    assert_eq!(mapped.text_stories, placeholder.text_stories);
    let saved = deck.save().unwrap();
    let original = parts(&source);
    let output = parts(&saved);
    assert_eq!(
        part_text(&output, "ppt/slides/slide1.xml"),
        part_text(&original, "ppt/slides/slide1.xml")
            .replace("<p:ph type=\"title\"/>", "<p:ph type=\"title\" idx=\"9\"/>")
    );
    let reopened = DeckSession::open(&saved, 12).unwrap().snapshot().unwrap();
    assert_eq!(
        reopened.slides[0].shapes[0].children[0].children[0].placeholder,
        mapped.placeholder
    );
    assert_eq!(reopened.slides[0].layout_part_path.as_deref(), Some(layout));
    assert!(deck.undo());
    assert_eq!(deck.snapshot().unwrap(), before);
    assert!(deck.redo());
    assert_eq!(deck.snapshot().unwrap(), changed);
}

#[test]
fn set_slide_layout_retargets_strict_relationships() {
    let strict_rels = SLIDE1_RELS.replace(
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
        "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout",
    );
    let source = zip(parts(&layout_change_source())
        .into_iter()
        .map(|(path, bytes)| {
            let xml = if path == "ppt/slides/_rels/slide1.xml.rels" {
                strict_rels.clone()
            } else {
                String::from_utf8(bytes).unwrap()
            };
            (path, xml)
        })
        .collect());
    let deck = DeckSession::open(&source, 11).unwrap();
    let slide_id = deck.snapshot().unwrap().slides[0].id.clone();
    deck.set_slide_layout(
        &context(),
        &slide_id,
        "ppt/slideLayouts/slideLayout2.xml",
        &[],
    )
    .unwrap();
    deck.duplicate_slide(&context(), &slide_id, 1).unwrap();
    let saved = deck.save().unwrap();
    let output = parts(&saved);
    let expected_rels = strict_rels.replace("slideLayout1.xml", "slideLayout2.xml");
    assert_eq!(
        part_text(&output, "ppt/slides/_rels/slide1.xml.rels"),
        expected_rels
    );
    let package = pptx_parse::parse_pptx(&saved).unwrap();
    assert_eq!(
        package.relationships["ppt/slides/slide3.xml"],
        package.relationships["ppt/slides/slide1.xml"]
    );
    assert_eq!(part_text(&output, "ppt/slides/slide1.xml"), SLIDE1);
    let reopened = DeckSession::open(&saved, 12).unwrap().snapshot().unwrap();
    for slide in &reopened.slides[..2] {
        assert_eq!(
            slide.layout_part_path.as_deref(),
            Some("ppt/slideLayouts/slideLayout2.xml")
        );
    }
}

#[test]
fn set_slide_layout_refusals_and_noop_leave_state_unchanged() {
    let deck = DeckSession::open(&layout_change_source(), 11).unwrap();
    let before = deck.snapshot().unwrap();
    let slide = &before.slides[0];
    let update = deck.encode_state_as_update_v1();
    deck.set_slide_layout(
        &context(),
        &slide.id,
        slide.layout_part_path.as_deref().unwrap(),
        &[],
    )
    .unwrap();
    assert_eq!(deck.encode_state_as_update_v1(), update);
    for (layout, mapping) in [
        ("missing", vec![]),
        ("ppt/slideLayouts/slideLayout4.xml", vec![]),
        (
            "ppt/slideLayouts/slideLayout2.xml",
            vec![(slide.shapes[2].id.clone(), 2)],
        ),
        (
            "ppt/slideLayouts/slideLayout2.xml",
            vec![(slide.shapes[0].id.clone(), 99)],
        ),
        (
            "ppt/slideLayouts/slideLayout2.xml",
            vec![
                (slide.shapes[0].id.clone(), 2),
                (slide.shapes[3].id.clone(), 2),
            ],
        ),
    ] {
        assert!(
            deck.set_slide_layout(&context(), &slide.id, layout, &mapping)
                .is_err()
        );
        assert_eq!(deck.snapshot().unwrap(), before);
        assert_eq!(deck.encode_state_as_update_v1(), update);
    }
}

#[test]
fn set_slide_layout_transfers_and_clones_with_notes() {
    let source = layout_change_source();
    let deck = DeckSession::open(&source, 11).unwrap();
    let fork =
        DeckSession::open_from_update_with_source(&deck.encode_state_as_update_v1(), &source, 12)
            .unwrap();
    let slide = fork.snapshot().unwrap().slides[0].clone();
    fork.set_slide_layout(
        &context(),
        &slide.id,
        "ppt/slideLayouts/slideLayout2.xml",
        &[(slide.shapes[0].id.clone(), 2)],
    )
    .unwrap();
    deck.apply_update_v1(&fork.encode_state_as_update_v1())
        .unwrap();
    deck.set_slide_notes(&context(), &slide.id, "Updated notes")
        .unwrap();
    deck.duplicate_slide(&context(), &slide.id, 1).unwrap();
    let saved = deck.save().unwrap();
    let rels = part_text(&parts(&saved), "ppt/slides/_rels/slide1.xml.rels");
    let prefix = SLIDE1_RELS
        .replace("slideLayout1.xml", "slideLayout2.xml")
        .replace("</Relationships>", "");
    assert!(rels.starts_with(&prefix), "{rels}");
    let reopened = DeckSession::open(&deck.save().unwrap(), 13)
        .unwrap()
        .snapshot()
        .unwrap();
    for slide in &reopened.slides[..2] {
        assert_eq!(
            slide.layout_part_path.as_deref(),
            Some("ppt/slideLayouts/slideLayout2.xml")
        );
        assert_eq!(slide.shapes[0].placeholder.as_ref().unwrap().index, Some(9));
        assert_eq!(slide.notes, "Updated notes");
    }
    assert_eq!(
        parts(&source),
        parts(&DeckSession::open(&source, 14).unwrap().save().unwrap())
    );
}

#[test]
fn set_slide_layout_title_only_keeps_unmapped_body() {
    let source = layout_change_source();
    let deck = DeckSession::open(&source, 11).unwrap();
    let slide = deck.snapshot().unwrap().slides[0].clone();
    deck.set_slide_layout(
        &context(),
        &slide.id,
        "ppt/slideLayouts/slideLayout3.xml",
        &[(slide.shapes[0].id.clone(), 2)],
    )
    .unwrap();
    assert_eq!(
        deck.snapshot().unwrap().slides[0].shapes[3].placeholder,
        slide.shapes[3].placeholder
    );
    assert_eq!(
        part_text(&parts(&deck.save().unwrap()), "ppt/slides/slide1.xml"),
        SLIDE1
    );
}

#[test]
fn set_slide_layout_rejects_unknown_remote_layout_atomically() {
    use yrs::updates::decoder::Decode;
    use yrs::{Doc, Map, ReadTxn, Transact, Update};
    let source = layout_change_source();
    let deck = DeckSession::open(&source, 11).unwrap();
    let before = deck.snapshot().unwrap();
    let doc = Doc::new();
    let mut txn = doc.transact_mut();
    txn.apply_update(Update::decode_v1(&deck.encode_state_as_update_v1()).unwrap())
        .unwrap();
    let slide = txn
        .get_map("pptx:slides")
        .unwrap()
        .get(&txn, &before.slides[0].id)
        .unwrap()
        .cast::<yrs::MapRef>()
        .unwrap();
    slide.insert(&mut txn, "layoutPartPath", "unknown-layout");
    let update = txn.encode_state_as_update_v1(&Default::default());
    drop(txn);
    assert!(deck.apply_update_v1(&update).is_err());
    assert_eq!(deck.snapshot().unwrap(), before);
    assert!(!deck.can_undo());
}

#[test]
fn set_slide_layout_writer_requires_exactly_one_layout_relationship() {
    for count in [0, 2] {
        let source: Vec<_> = parts(&layout_change_source())
            .into_iter()
            .map(|(path, bytes)| {
                let mut xml = String::from_utf8(bytes).unwrap();
                if path == "ppt/slides/_rels/slide1.xml.rels" {
                    let start = xml.find("<Relationship Id=\"rId1\"").unwrap();
                    let end = start + xml[start..].find("/>").unwrap() + 2;
                    let relationship = xml[start..end].to_owned();
                    xml.replace_range(start..end, "");
                    if count == 2 {
                        xml = xml.replace(
                            "</Relationships>",
                            &format!(
                                "{}{}{}",
                                relationship,
                                relationship.replace("rId1", "rId3").replace(
                                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout",
                                    "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout",
                                ),
                                "</Relationships>"
                            ),
                        );
                    }
                }
                (path, xml)
            })
            .collect();
        let package = pptx_parse::parse_pptx(&zip(source)).unwrap();
        let write = pptx_parse::DeckWrite {
            slides: vec![
                pptx_parse::SlideWrite::Patch {
                    part_path: "ppt/slides/slide1.xml".into(),
                    shapes: (0..package.slides[0].shapes.len())
                        .map(|source_index| pptx_parse::ShapeWrite::Keep { source_index })
                        .collect(),
                    background: None,
                    layout_part_path: Some("ppt/slideLayouts/slideLayout2.xml".into()),
                },
                pptx_parse::SlideWrite::Keep {
                    part_path: "ppt/slides/slide2.xml".into(),
                },
            ],
            comments: None,
            notes: None,
        };
        assert!(
            matches!(pptx_parse::write_pptx_with_edits(&package, &write), Err(pptx_parse::PptxError::InvalidSlideLayoutRelationships { count: actual, .. }) if actual == count)
        );
    }
}

#[test]
fn set_slide_layout_without_mapping_keeps_slide_xml_verbatim() {
    let source = parts(&layout_change_source()).into_iter().map(|(path, bytes)| {
        let mut xml = String::from_utf8(bytes).unwrap();
        if path == "ppt/slides/slide2.xml" {
            xml = xml.replace(" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"", "");
        }
        (path, xml)
    }).collect();
    let source = zip(source);
    let deck = DeckSession::open(&source, 11).unwrap();
    let slide = deck.snapshot().unwrap().slides[1].id.clone();
    deck.set_slide_layout(&context(), &slide, "ppt/slideLayouts/slideLayout2.xml", &[])
        .unwrap();
    let original = parts(&source);
    let saved = parts(&deck.save().unwrap());
    assert_eq!(
        saved["ppt/slides/slide2.xml"],
        original["ppt/slides/slide2.xml"]
    );
}
