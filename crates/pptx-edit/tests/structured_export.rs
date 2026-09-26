//! Structured export against a synthetic deck: inherited and numbered lists, a link, a field, a
//! soft line break and an equation, nested and hidden groups, a merged table, a chart, SmartArt,
//! an embedded object, a picture, a video, an ink part, a hidden slide, a layout with its own
//! drawing and prompt text, speaker notes with an extra notes-page box, and a comment.

use std::cell::Cell;
use std::rc::Rc;

use pptx_edit::structured::{
    AnchorScope, ExportDiagnosticCode, ExportFailureCode, ExportList, ExportObjectKind,
    ExportRunKind, ExportShape, ExportShapeKind, ExportSlide, PptxAnchor, PptxExportOptions,
    PptxMarkdownOptions, PptxStructuredContent, export_pptx_markdown, export_pptx_structured,
    render_pptx_markdown,
};
use pptx_edit::{DeckSession, EditCtx, ReadRequest, ShapeDraft, ShapeRect, TextStyle};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

const NS: &str = r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main""#;
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Default Extension="png" ContentType="image/png"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
<Override PartName="/ppt/notesSlides/notesSlide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>
<Override PartName="/ppt/comments/comment1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.comments+xml"/>
<Override PartName="/ppt/commentAuthors.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
</Types>"#;

const ROOT_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>"#;

fn presentation() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation {NS}>
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/></p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>"#
    )
}

fn presentation_rels() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="{REL}/slide" Target="slides/slide1.xml"/>
<Relationship Id="rId3" Type="{REL}/slide" Target="slides/slide2.xml"/>
<Relationship Id="rId4" Type="{REL}/slide" Target="slides/slide3.xml"/>
<Relationship Id="rId5" Type="{REL}/commentAuthors" Target="commentAuthors.xml"/>
</Relationships>"#
    )
}

fn text_shape(id: u32, name: &str, extra: &str, paragraphs: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"{extra}/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="100000" y="100000"/><a:ext cx="3000000" cy="400000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</p:txBody></p:sp>"#
    )
}

fn placeholder_shape(id: u32, name: &str, placeholder: &str, paragraphs: &str) -> String {
    format!(
        r#"<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="{name}"/><p:cNvSpPr/><p:nvPr>{placeholder}</p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>{paragraphs}</p:txBody></p:sp>"#
    )
}

fn frame(id: u32, name: &str, extra: &str, graphic: &str) -> String {
    format!(
        r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{id}" name="{name}"{extra}/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="3000000" y="3000000"/><a:ext cx="4000000" cy="600000"/></p:xfrm><a:graphic>{graphic}</a:graphic></p:graphicFrame>"#
    )
}

fn run(text: &str) -> String {
    format!(r#"<a:r><a:rPr lang="en-US"/><a:t>{text}</a:t></a:r>"#)
}

fn numbered(scheme: &str, start: Option<u32>, text: &str) -> String {
    let start = start.map_or(String::new(), |start| format!(r#" startAt="{start}""#));
    let runs = if text.is_empty() {
        String::new()
    } else {
        run(text)
    };
    format!(r#"<a:p><a:pPr><a:buAutoNum type="{scheme}"{start}/></a:pPr>{runs}</a:p>"#)
}

fn ink(id: u32, name: &str, extra: &str, relationship: &str) -> String {
    format!(
        r#"<p:contentPart r:id="{relationship}"><p14:nvContentPartPr xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"><p14:cNvPr id="{id}" name="{name}"{extra}/></p14:nvContentPartPr></p:contentPart>"#
    )
}

fn slide1() -> String {
    let title = placeholder_shape(
        2,
        "Title",
        r#"<p:ph type="title"/>"#,
        &format!("<a:p>{}</a:p>", run("Quarterly review")),
    );
    let body = placeholder_shape(
        3,
        "Body",
        r#"<p:ph type="body" idx="1"/>"#,
        &format!(
            r#"<a:p>{}</a:p><a:p><a:pPr lvl="1"/>{}</a:p><a:p>{}</a:p>"#,
            run("First point"),
            run("Sub point 😀"),
            run("Second point")
        ),
    );
    let numbered_shape = text_shape(
        4,
        "Numbered",
        "",
        &[
            numbered("arabicPeriod", Some(3), "Three"),
            numbered("arabicPeriod", Some(3), "Four"),
            numbered("arabicPeriod", Some(3), ""),
            numbered("arabicPeriod", Some(3), "Five"),
        ]
        .concat(),
    );
    let roman = text_shape(
        5,
        "Roman",
        "",
        &[
            numbered("romanUcPeriod", None, "One"),
            numbered("romanUcPeriod", None, "Two"),
        ]
        .concat(),
    );
    let circled = text_shape(6, "Circled", "", &numbered("circleNumDbPlain", None, "Odd"));
    let linked = text_shape(
        7,
        "Linked",
        "",
        &format!(
            r#"<a:p><a:r><a:rPr lang="en-US"><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Docs</a:t></a:r>{}<a:fld id="{{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld><a:br><a:rPr lang="en-US"/></a:br>{}<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14"><a14:m><m:oMathPara xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"/></a14:m></mc:Choice><mc:Fallback><a:r><a:t>[math]</a:t></a:r></mc:Fallback></mc:AlternateContent>{}</a:p>"#,
            run(" page "),
            run("next"),
            run(" end &lt;script&gt;"),
        ),
    );
    let group = format!(
        r#"<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/><a:chOff x="0" y="0"/><a:chExt cx="100" cy="100"/></a:xfrm></p:grpSpPr>{}<p:grpSp><p:nvGrpSpPr><p:cNvPr id="12" name="Hidden inner" hidden="1"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}{}</p:grpSp></p:grpSp>"#,
        text_shape(
            11,
            "Nested",
            "",
            &format!("<a:p>{}</a:p>", run("Nested text"))
        ),
        text_shape(
            13,
            "Secret",
            "",
            &format!("<a:p>{}</a:p>", run("Secret text"))
        ),
        ink(14, "Inner ink", "", "rIdInk"),
    );
    let hidden = text_shape(
        20,
        "Hidden shape",
        r#" hidden="1""#,
        &format!("<a:p>{}</a:p>", run("Hidden text")),
    );
    let ink = ink(30, "Ink", r#" descr="Signature""#, "rIdInk");
    let cell = |text: &str, attributes: &str| {
        format!(
            r#"<a:tc{attributes}><a:txBody><a:bodyPr/><a:lstStyle/><a:p>{}</a:p></a:txBody><a:tcPr/></a:tc>"#,
            run(text)
        )
    };
    let table = frame(
        40,
        "Table",
        "",
        &format!(
            r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid><a:tr h="300000">{}{}</a:tr><a:tr h="300000">{}{}</a:tr></a:tbl></a:graphicData>"#,
            cell("Merged header", r#" gridSpan="2""#),
            cell("ghost", r#" hMerge="1""#),
            cell("A &lt;b&gt;", ""),
            cell("B", ""),
        ),
    );
    let chart = frame(
        50,
        "Revenue chart",
        r#" descr="Revenue by quarter""#,
        r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rIdChart"/></a:graphicData>"#,
    );
    let smart_art = frame(
        60,
        "Process",
        "",
        r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdDm" r:lo="rIdLo" r:qs="rIdQs" r:cs="rIdCs"/></a:graphicData>"#,
    );
    let ole = frame(
        70,
        "Embedded",
        "",
        r#"<a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj name="Worksheet" r:id="rIdOle"/></a:graphicData>"#,
    );
    let picture = |id: u32, name: &str, extra: &str, media: &str| {
        format!(
            r#"<p:pic><p:nvPicPr><p:cNvPr id="{id}" name="{name}"{extra}/><p:cNvPicPr/><p:nvPr>{media}</p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100000" cy="100000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"#
        )
    };
    let logo = picture(
        80,
        "Logo",
        r#" descr="Company logo" title="Logo title""#,
        "",
    );
    let clip = picture(
        90,
        "Clip",
        r#" descr="Product demo""#,
        r#"<a:videoFile r:link="rIdVideo"/>"#,
    );
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld {NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{title}{body}{numbered_shape}{roman}{circled}{linked}{group}{hidden}{ink}{table}{chart}{smart_art}{ole}{logo}{clip}</p:spTree></p:cSld></p:sld>"#
    )
}

fn slide1_rels() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="{REL}/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="{REL}/notesSlide" Target="../notesSlides/notesSlide1.xml"/>
<Relationship Id="rId3" Type="{REL}/comments" Target="../comments/comment1.xml"/>
<Relationship Id="rIdLink" Type="{REL}/hyperlink" Target="https://example.com/docs?a=1&amp;b=2" TargetMode="External"/>
<Relationship Id="rIdChart" Type="{REL}/chart" Target="../charts/chart1.xml"/>
<Relationship Id="rIdDm" Type="{REL}/diagramData" Target="../diagrams/data1.xml"/>
<Relationship Id="rIdLo" Type="{REL}/diagramLayout" Target="../diagrams/layout1.xml"/>
<Relationship Id="rIdQs" Type="{REL}/diagramQuickStyle" Target="../diagrams/quickStyle1.xml"/>
<Relationship Id="rIdCs" Type="{REL}/diagramColors" Target="../diagrams/colors1.xml"/>
<Relationship Id="rIdOle" Type="{REL}/oleObject" Target="../embeddings/oleObject1.bin"/>
<Relationship Id="rIdImg" Type="{REL}/image" Target="../media/image1.png"/>
<Relationship Id="rIdVideo" Type="{REL}/video" Target="https://example.com/clip.mp4" TargetMode="External"/>
<Relationship Id="rIdInk" Type="http://schemas.microsoft.com/office/2011/relationships/inkAction" Target="../ink/ink1.xml"/>
</Relationships>"#
    )
}

fn slide2() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld {NS} show="0"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}</p:spTree></p:cSld></p:sld>"#,
        text_shape(
            2,
            "Backup",
            "",
            &format!("<a:p>{}</a:p>", run("Backup slide"))
        )
    )
}

fn slide3() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld {NS}><p:cSld name="Closing"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}{}{}{}</p:spTree></p:cSld></p:sld>"#,
        placeholder_shape(
            2,
            "Empty body",
            r#"<p:ph type="body" idx="1"/>"#,
            r#"<a:p><a:endParaRPr lang="en-US"/></a:p>"#
        ),
        text_shape(3, "Plain", "", &format!("<a:p>{}</a:p>", run("Last slide"))),
        tall_table(),
        ink(5, "Hidden ink", r#" hidden="1""#, "rIdInk"),
    )
}

/// A column whose first cell spans all three rows.
fn tall_table() -> String {
    let cell = |text: &str, attributes: &str| {
        format!(
            r#"<a:tc{attributes}><a:txBody><a:bodyPr/><a:lstStyle/><a:p>{}</a:p></a:txBody><a:tcPr/></a:tc>"#,
            run(text)
        )
    };
    let rows: String = [
        (cell("Tall", r#" rowSpan="3""#), "a"),
        (cell("x", r#" vMerge="1""#), "b"),
        (cell("", r#" vMerge="1""#), "c"),
    ]
    .into_iter()
    .map(|(first, second)| format!(r#"<a:tr h="300000">{first}{}</a:tr>"#, cell(second, "")))
    .collect();
    frame(
        4,
        "Tall table",
        "",
        &format!(
            r#"<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>{rows}</a:tbl></a:graphicData>"#
        ),
    )
}

/// A relationships part of `(id, type, target)` internal relationships.
fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let entries: String = entries
        .iter()
        .map(|(id, kind, target)| {
            format!(r#"<Relationship Id="{id}" Type="{REL}/{kind}" Target="{target}"/>"#)
        })
        .collect();
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{entries}</Relationships>"#
    )
}

/// A notes page with `body` as its notes and a separate text box reading `extra`.
fn notes(body: &str, extra: &str) -> String {
    let body = if body.is_empty() {
        "<a:p><a:endParaRPr/></a:p>".to_owned()
    } else {
        format!("<a:p>{}</a:p>", run(body))
    };
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes {NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}{}</p:spTree></p:cSld></p:notes>"#,
        placeholder_shape(2, "Notes", r#"<p:ph type="body" idx="1"/>"#, &body),
        text_shape(3, "Extra", "", &format!("<a:p>{}</a:p>", run(extra))),
    )
}

fn comments() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmLst {NS}><p:cm authorId="0" dt="2026-01-02T03:04:05.000" idx="1"><p:pos x="10" y="10"/><p:text>Check numbers</p:text></p:cm></p:cmLst>"#
    )
}

fn comment_authors() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:cmAuthorLst {NS}><p:cmAuthor id="0" name="Ada" initials="A" lastIdx="1" clrIdx="0"/></p:cmAuthorLst>"#
    )
}

fn layout() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout {NS} type="obj"><p:cSld name="Content"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}{}{}</p:spTree></p:cSld></p:sldLayout>"#,
        placeholder_shape(
            2,
            "Title",
            r#"<p:ph type="title"/>"#,
            &format!("<a:p>{}</a:p>", run("Layout title prompt"))
        ),
        placeholder_shape(
            3,
            "Body",
            r#"<p:ph type="body" idx="1"/>"#,
            &format!("<a:p>{}</a:p>", run("Layout prompt text"))
        ),
        text_shape(4, "Layout logo", "", &format!("<a:p>{}</a:p>", run("ACME"))),
    )
}

fn master() -> String {
    format!(
        r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster {NS}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{}{}</p:spTree></p:cSld>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
<p:txStyles><p:titleStyle><a:lvl1pPr><a:defRPr b="1"/></a:lvl1pPr></p:titleStyle><p:bodyStyle><a:lvl1pPr><a:buFont typeface="Arial"/><a:buChar char="•"/></a:lvl1pPr><a:lvl2pPr><a:buChar char="–"/></a:lvl2pPr></p:bodyStyle><p:otherStyle><a:lvl1pPr/></p:otherStyle></p:txStyles></p:sldMaster>"#,
        placeholder_shape(2, "Title", r#"<p:ph type="title"/>"#, "<a:p/>"),
        placeholder_shape(3, "Body", r#"<p:ph type="body" idx="1"/>"#, "<a:p/>"),
    )
}

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

fn parts() -> Vec<(String, Vec<u8>)> {
    let layout = [("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml")];
    [
        ("[Content_Types].xml", CONTENT_TYPES.to_owned()),
        ("_rels/.rels", ROOT_RELS.to_owned()),
        ("ppt/presentation.xml", presentation()),
        ("ppt/_rels/presentation.xml.rels", presentation_rels()),
        ("ppt/slides/slide1.xml", slide1()),
        ("ppt/slides/_rels/slide1.xml.rels", slide1_rels()),
        ("ppt/slides/slide2.xml", slide2()),
        ("ppt/slides/_rels/slide2.xml.rels", relationships(&layout)),
        ("ppt/slides/slide3.xml", slide3()),
        (
            "ppt/slides/_rels/slide3.xml.rels",
            relationships(&[
                layout[0],
                ("rId2", "notesSlide", "../notesSlides/notesSlide2.xml"),
                ("rIdInk", "customXml", "../ink/ink2.xml"),
            ]),
        ),
        (
            "ppt/notesSlides/notesSlide1.xml",
            notes("Talk track", "Extra box"),
        ),
        ("ppt/notesSlides/notesSlide2.xml", notes("", "Aside box")),
        (
            "ppt/notesSlides/_rels/notesSlide2.xml.rels",
            relationships(&[("rId1", "slide", "../slides/slide3.xml")]),
        ),
        (
            "ppt/notesSlides/_rels/notesSlide1.xml.rels",
            relationships(&[("rId1", "slide", "../slides/slide1.xml")]),
        ),
        ("ppt/comments/comment1.xml", comments()),
        ("ppt/commentAuthors.xml", comment_authors()),
        ("ppt/slideLayouts/slideLayout1.xml", self::layout()),
        (
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
            relationships(&[("rId1", "slideMaster", "../slideMasters/slideMaster1.xml")]),
        ),
        ("ppt/slideMasters/slideMaster1.xml", master()),
        (
            "ppt/slideMasters/_rels/slideMaster1.xml.rels",
            relationships(&[
                ("rId1", "slideLayout", "../slideLayouts/slideLayout1.xml"),
                ("rId2", "theme", "../theme/theme1.xml"),
            ]),
        ),
        ("ppt/theme/theme1.xml", THEME.to_owned()),
    ]
    .into_iter()
    .map(|(path, body)| (path.to_owned(), body.into_bytes()))
    .chain([(
        "ppt/media/image1.png".to_owned(),
        vec![0x89, b'P', b'N', b'G', 0x0d, 0x0a, 0x1a, 0x0a],
    )])
    .collect()
}

fn fixture() -> Vec<u8> {
    ooxml_opc::rezip_parts(&parts()).unwrap()
}

fn options(value: Value) -> PptxExportOptions {
    serde_json::from_value(value).unwrap()
}

fn export(bytes: &[u8], value: Value) -> PptxStructuredContent {
    export_pptx_structured(bytes, &options(value)).unwrap()
}

fn session_export(session: &DeckSession, value: Value) -> PptxStructuredContent {
    session
        .export_structured(&options(value))
        .unwrap()
        .unwrap()
        .content
}

fn find<'a>(shapes: &'a [ExportShape], name: &str) -> Option<&'a ExportShape> {
    shapes.iter().find_map(|shape| {
        if shape.name == name {
            Some(shape)
        } else {
            find(&shape.children, name)
        }
    })
}

fn shape<'a>(content: &'a PptxStructuredContent, name: &str) -> &'a ExportShape {
    content
        .slides
        .iter()
        .find_map(|slide| find(&slide.shapes, name))
        .unwrap_or_else(|| panic!("no shape named {name}"))
}

fn paragraph_texts(shape: &ExportShape) -> Vec<String> {
    shape
        .stories
        .iter()
        .flat_map(|story| &story.paragraphs)
        .map(|paragraph| {
            paragraph
                .runs
                .iter()
                .map(|run| match &run.content {
                    ExportRunKind::Text { text } | ExportRunKind::Field { text, .. } => {
                        text.clone()
                    }
                    ExportRunKind::LineBreak => "\n".to_owned(),
                    ExportRunKind::Unsupported { .. } => String::new(),
                })
                .collect()
        })
        .collect()
}

fn codes(content: &PptxStructuredContent) -> Vec<ExportDiagnosticCode> {
    content
        .diagnostics
        .iter()
        .map(|diagnostic| diagnostic.code)
        .collect()
}

fn slide_names(content: &PptxStructuredContent) -> Vec<(u32, Vec<String>)> {
    content
        .slides
        .iter()
        .map(|slide: &ExportSlide| {
            (
                slide.index,
                slide
                    .shapes
                    .iter()
                    .map(|shape| shape.name.clone())
                    .collect(),
            )
        })
        .collect()
}

#[test]
fn slides_and_shapes_follow_the_current_shape_tree() {
    let deck = fixture();
    let content = export(&deck, json!({}));
    assert_eq!(content.anchor_scope, AnchorScope::Snapshot);
    assert!(!content.truncated);
    assert_eq!(
        slide_names(&content),
        [
            (
                0,
                [
                    "Title",
                    "Body",
                    "Numbered",
                    "Roman",
                    "Circled",
                    "Linked",
                    "Group",
                    "Ink",
                    "Table",
                    "Revenue chart",
                    "Process",
                    "Embedded",
                    "Logo",
                    "Clip",
                ]
                .map(str::to_owned)
                .to_vec()
            ),
            (
                2,
                ["Empty body", "Plain", "Tall table"]
                    .map(str::to_owned)
                    .to_vec()
            ),
        ]
    );
    let group = shape(&content, "Group");
    assert_eq!(group.kind, ExportShapeKind::Group);
    assert_eq!(
        group
            .children
            .iter()
            .map(|child| child.name.as_str())
            .collect::<Vec<_>>(),
        ["Nested"]
    );
    assert_eq!(group.children[0].id, "s0.h6.h0");

    let title = shape(&content, "Title");
    let provenance = title.provenance.as_ref().unwrap();
    let slide_part = parts()
        .into_iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .unwrap()
        .1;
    assert_eq!(provenance.part, "ppt/slides/slide1.xml");
    assert_eq!(
        provenance.part_sha256,
        format!("{:x}", Sha256::digest(&slide_part))
    );
    assert_eq!(provenance.path, [0, 0, 2]);
    assert_eq!(
        (provenance.sld_id, provenance.source_id),
        (Some(256), Some(2))
    );
    assert_eq!(
        content.slides[0].provenance.as_ref().unwrap().path,
        Vec::<u32>::new()
    );

    let ink = shape(&content, "Ink");
    assert_eq!(ink.kind, ExportShapeKind::Unknown);
    assert_eq!(ink.id, "s0.x0");
    assert_eq!(ink.description.as_deref(), Some("Signature"));
    assert_eq!(ink.object.as_ref().unwrap().element, "p:contentPart");
    assert!(matches!(
        &ink.anchor,
        PptxAnchor::SourcePart { part, path, .. } if part == "ppt/slides/slide1.xml" && path == &[0, 0, 10]
    ));

    let diagnostics = codes(&content);
    for code in [
        ExportDiagnosticCode::HiddenContentExcluded,
        ExportDiagnosticCode::StoriesOmitted,
        ExportDiagnosticCode::InheritedContentOmitted,
        ExportDiagnosticCode::UnsupportedContent,
        ExportDiagnosticCode::ImageDataOmitted,
        ExportDiagnosticCode::UnsupportedNumbering,
        ExportDiagnosticCode::FieldCachedResult,
    ] {
        assert!(diagnostics.contains(&code), "missing {code:?}");
    }
    let inherited = content
        .diagnostics
        .iter()
        .find(|diagnostic| diagnostic.code == ExportDiagnosticCode::InheritedContentOmitted)
        .unwrap();
    assert!(matches!(
        &inherited.anchor,
        Some(PptxAnchor::SourcePart { part, .. }) if part == "ppt/slideLayouts/slideLayout1.xml"
    ));
    assert!(
        !serde_json::to_string(&content)
            .unwrap()
            .contains("Layout prompt text")
    );
}

#[test]
fn lists_inherit_master_bullets_and_count_numbers_per_level() {
    let content = export(&fixture(), json!({}));
    let lists = |name: &str| -> Vec<Option<ExportList>> {
        shape(&content, name).stories[0]
            .paragraphs
            .iter()
            .map(|paragraph| paragraph.list.clone())
            .collect()
    };
    let bullet = |character: &str, font: Option<&str>| {
        Some(ExportList::Bullet {
            character: character.to_owned(),
            font: font.map(str::to_owned),
        })
    };
    assert_eq!(
        lists("Body"),
        [
            bullet("•", Some("Arial")),
            bullet("–", None),
            bullet("•", Some("Arial"))
        ]
    );
    let number = |scheme: &str, start_at: u32, value: u32, marker: Option<&str>| {
        Some(ExportList::Number {
            scheme: scheme.to_owned(),
            start_at,
            value,
            marker: marker.map(str::to_owned),
        })
    };
    assert_eq!(
        lists("Numbered"),
        [
            number("arabicPeriod", 3, 3, Some("3.")),
            number("arabicPeriod", 3, 4, Some("4.")),
            None,
            number("arabicPeriod", 3, 5, Some("5.")),
        ]
    );
    assert_eq!(
        lists("Roman"),
        [
            number("romanUcPeriod", 1, 1, Some("I.")),
            number("romanUcPeriod", 1, 2, Some("II."))
        ]
    );
    assert_eq!(lists("Circled"), [number("circleNumDbPlain", 1, 1, None)]);
    let body = &shape(&content, "Body").stories[0].paragraphs;
    assert_eq!(body[1].level, 1);
    assert_eq!(body[0].bullet_json, None);
    let title = &shape(&content, "Title").stories[0].paragraphs[0];
    assert_eq!(
        title.runs[0].marks.as_deref(),
        Some([pptx_edit::structured::ExportMark::Bold].as_slice())
    );
}

#[test]
fn runs_carry_links_fields_breaks_and_unsupported_inlines() {
    let content = export(&fixture(), json!({}));
    let linked = &shape(&content, "Linked").stories[0].paragraphs[0];
    let kinds: Vec<(&ExportRunKind, Option<&str>, (u32, u32))> = linked
        .runs
        .iter()
        .map(|run| {
            let PptxAnchor::Text { range, .. } = &run.anchor else {
                panic!("run anchors are text ranges")
            };
            (
                &run.content,
                run.link.as_ref().map(|link| link.href.as_str()),
                (range.start, range.end),
            )
        })
        .collect();
    let text = |text: &str| ExportRunKind::Text {
        text: text.to_owned(),
    };
    assert_eq!(
        kinds,
        [
            (
                &text("Docs"),
                Some("https://example.com/docs?a=1&b=2"),
                (0, 4)
            ),
            (&text(" page "), None, (4, 10)),
            (
                &ExportRunKind::Field {
                    field_type: Some("slidenum".to_owned()),
                    text: "1".to_owned()
                },
                None,
                (10, 11)
            ),
            (&ExportRunKind::LineBreak, None, (11, 12)),
            (&text("next"), None, (12, 16)),
            (
                &ExportRunKind::Unsupported {
                    element: "a14:m".to_owned()
                },
                None,
                (16, 16)
            ),
            (&text(" end <script>"), None, (16, 29)),
        ]
    );
    assert!(linked.runs[0].link.as_ref().unwrap().external);
}

#[test]
fn text_anchors_agree_with_batch_reads() {
    let session = DeckSession::open(&fixture(), 81).unwrap();
    let content = session_export(&session, json!({}));
    let read = session
        .read_content(&ReadRequest::default())
        .unwrap()
        .unwrap();
    let mut checked = 0;
    for story in &read.stories {
        let exported = content
            .slides
            .iter()
            .flat_map(|slide| stories_of(&slide.shapes))
            .find(|(_, anchor)| {
                matches!(anchor, PptxAnchor::Text { story_id, .. } if story_id == &story.story_id)
            });
        let Some((paragraphs, _)) = exported else {
            continue;
        };
        let units: Vec<u16> = story.text.encode_utf16().collect();
        for (paragraph, expected) in paragraphs.iter().zip(&story.paragraphs) {
            let PptxAnchor::Text { range, .. } = &paragraph.anchor else {
                panic!("paragraph anchors are text ranges")
            };
            assert_eq!((range.start, range.end), (expected.start, expected.end));
            assert_eq!(paragraph.paragraph_id, expected.paragraph_id);
            for run in &paragraph.runs {
                let PptxAnchor::Text { range, .. } = &run.anchor else {
                    panic!("run anchors are text ranges")
                };
                let slice =
                    String::from_utf16(&units[range.start as usize..range.end as usize]).unwrap();
                match &run.content {
                    ExportRunKind::Text { text } | ExportRunKind::Field { text, .. } => {
                        assert_eq!(&slice, text)
                    }
                    ExportRunKind::LineBreak => assert_eq!(slice, "\n"),
                    ExportRunKind::Unsupported { .. } => assert!(slice.is_empty()),
                }
            }
            checked += 1;
        }
    }
    assert_eq!(checked, 25);
}

fn stories_of(
    shapes: &[ExportShape],
) -> Vec<(&[pptx_edit::structured::ExportParagraph], &PptxAnchor)> {
    let mut output = Vec::new();
    for shape in shapes {
        for story in &shape.stories {
            output.push((story.paragraphs.as_slice(), &story.anchor));
        }
        if let Some(table) = &shape.table {
            for cell in table.rows.iter().flat_map(|row| &row.cells) {
                if let Some(story) = &cell.story {
                    output.push((story.paragraphs.as_slice(), &story.anchor));
                }
            }
        }
        output.extend(stories_of(&shape.children));
    }
    output
}

#[test]
fn tables_keep_merges_and_read_current_cell_text() {
    let session = DeckSession::open(&fixture(), 82).unwrap();
    let table_shape = session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Table")
        .unwrap()
        .clone();
    let cell_story = &table_shape.text_stories[2];
    session
        .insert_text(
            &EditCtx::local("test"),
            &cell_story.id,
            0,
            "Edited ",
            &TextStyle::default(),
        )
        .unwrap();
    let content = session_export(&session, json!({}));
    let table = shape(&content, "Table").table.as_ref().unwrap();
    assert_eq!(table.columns, 2);
    let header = &table.rows[0].cells;
    assert_eq!((header[0].grid_span, header[0].merged), (2, false));
    assert!(header[1].merged);
    assert_eq!(
        header[1]
            .merge_origin
            .map(|origin| (origin.row, origin.column)),
        Some((0, 0))
    );
    let edited = table.rows[1].cells[0].story.as_ref().unwrap();
    assert_eq!(edited.paragraphs[0].runs.len(), 1);
    assert!(matches!(
        &edited.paragraphs[0].runs[0].content,
        ExportRunKind::Text { text } if text == "Edited A <b>"
    ));
    assert_eq!(table.rows[1].cells[0].id, "s0.h8.r1c0");

    let markdown = session
        .export_markdown(&options(json!({})))
        .unwrap()
        .unwrap()
        .content;
    assert!(markdown.markdown.contains(r#"<td colspan="2">"#));
    assert!(markdown.markdown.contains("Edited A &lt;b&gt;"));
    assert!(!markdown.markdown.contains("ghost"));
    assert!(
        markdown
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.code
                == ExportDiagnosticCode::MergeContinuationContentOmitted)
    );
}

#[test]
fn unrepresented_objects_are_placeholders_with_alternative_text() {
    let content = export(&fixture(), json!({}));
    let object = |name: &str| shape(&content, name).object.clone().unwrap();
    let logo = shape(&content, "Logo");
    assert_eq!(logo.title.as_deref(), Some("Logo title"));
    assert_eq!(logo.description.as_deref(), Some("Company logo"));
    let picture = object("Logo");
    assert_eq!(picture.kind, ExportObjectKind::Picture);
    assert_eq!(picture.relationship_ids, ["rIdImg"]);
    assert_eq!(picture.parts, ["ppt/media/image1.png"]);
    assert!(picture.external_targets.is_empty());
    let clip = object("Clip");
    assert_eq!(clip.kind, ExportObjectKind::Video);
    assert_eq!(clip.relationship_ids, ["rIdVideo", "rIdImg"]);
    assert_eq!(clip.parts, ["ppt/media/image1.png"]);
    assert_eq!(clip.external_targets, ["https://example.com/clip.mp4"]);
    let embedded = object("Embedded");
    assert_eq!(embedded.relationship_ids, ["rIdOle"]);
    assert_eq!(embedded.parts, ["ppt/embeddings/oleObject1.bin"]);
    let ink = object("Ink");
    assert_eq!(ink.relationship_ids, ["rIdInk"]);
    assert_eq!(ink.parts, ["ppt/ink/ink1.xml"]);
    let chart = object("Revenue chart");
    assert_eq!(chart.kind, ExportObjectKind::Chart);
    assert_eq!(chart.parts, ["ppt/charts/chart1.xml"]);
    assert_eq!(
        shape(&content, "Revenue chart").description.as_deref(),
        Some("Revenue by quarter")
    );
    let smart_art = object("Process");
    assert_eq!(smart_art.kind, ExportObjectKind::SmartArt);
    assert_eq!(smart_art.relationship_ids.len(), 4);
    assert!(
        smart_art
            .parts
            .contains(&"ppt/diagrams/data1.xml".to_owned())
    );
    let embedded = object("Embedded");
    assert_eq!(embedded.kind, ExportObjectKind::EmbeddedObject);
    assert_eq!(
        embedded.uri.as_deref(),
        Some("http://schemas.openxmlformats.org/presentationml/2006/ole")
    );
    let anchored = |name: &str| {
        let anchor = &shape(&content, name).anchor;
        content
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.anchor.as_ref() == Some(anchor))
    };
    for name in [
        "Logo",
        "Clip",
        "Revenue chart",
        "Process",
        "Embedded",
        "Ink",
    ] {
        assert!(anchored(name), "{name} has no diagnostic");
    }
}

#[test]
fn hidden_slides_and_shapes_are_excluded_unless_requested() {
    let deck = fixture();
    let content = export(
        &deck,
        json!({"includeHiddenSlides": true, "includeHiddenShapes": true}),
    );
    assert_eq!(
        content
            .slides
            .iter()
            .map(|slide| (slide.index, slide.hidden))
            .collect::<Vec<_>>(),
        [(0, Some(false)), (1, Some(true)), (2, Some(false))]
    );
    assert!(shape(&content, "Hidden shape").hidden);
    let secret = shape(&content, "Secret");
    assert!(secret.hidden);
    assert!(!shape(&content, "Hidden inner").children.is_empty());
    assert!(!shape(&content, "Nested").hidden);
    let inner_ink = shape(&content, "Inner ink");
    assert_eq!(
        (inner_ink.kind, inner_ink.hidden),
        (ExportShapeKind::Unknown, true)
    );
    assert!(shape(&content, "Hidden ink").hidden);

    let default = export(&deck, json!({}));
    let hidden_shapes: Vec<&str> = default
        .diagnostics
        .iter()
        .filter(|diagnostic| {
            diagnostic.code == ExportDiagnosticCode::HiddenContentExcluded
                && diagnostic.anchor.is_some()
        })
        .map(|diagnostic| diagnostic.message.as_str())
        .collect();
    assert_eq!(hidden_shapes.len(), 2);
    assert!(hidden_shapes[0].starts_with("2 hidden shape(s)"));
    assert!(hidden_shapes[1].starts_with("1 hidden shape(s)"));
    let serialized = serde_json::to_string(&default).unwrap();
    assert!(!serialized.contains("Inner ink") && !serialized.contains("Hidden ink"));
}

#[test]
fn notes_and_comments_are_opt_in_stories() {
    let deck = fixture();
    let default = export(&deck, json!({}));
    assert!(default.slides[0].notes.is_none());
    assert!(default.slides[0].comments.is_empty());
    let omitted = default
        .diagnostics
        .iter()
        .filter(|diagnostic| diagnostic.code == ExportDiagnosticCode::StoriesOmitted)
        .count();
    assert_eq!(omitted, 2);

    let content = export(
        &deck,
        json!({"includeNotes": true, "includeComments": true}),
    );
    let slide = &content.slides[0];
    let notes = slide.notes.as_ref().unwrap();
    assert_eq!(notes.text, "Talk track");
    assert_eq!(
        notes.anchor,
        PptxAnchor::Notes {
            slide_id: slide_id(slide),
            range: pptx_edit::structured::TextSpan { start: 0, end: 10 }
        }
    );
    assert_eq!(
        notes.provenance.as_ref().unwrap().part,
        "ppt/notesSlides/notesSlide1.xml"
    );
    let comment = &slide.comments[0];
    assert_eq!(comment.author.as_deref(), Some("Ada"));
    assert_eq!(comment.text, "Check numbers");
    assert_eq!(comment.date.as_deref(), Some("2026-01-02T03:04:05.000"));
    let notes_diagnostics: Vec<_> = content
        .diagnostics
        .iter()
        .filter(|diagnostic| diagnostic.anchor.as_ref() == Some(&notes.anchor))
        .collect();
    assert_eq!(notes_diagnostics.len(), 1);
    assert!(notes_diagnostics[0].message.contains("1 more text shape"));
    assert!(codes(&content).contains(&ExportDiagnosticCode::NotesStructureOmitted));

    let closing = &content.slides[1];
    assert!(closing.notes.is_none());
    let aside = PptxAnchor::Notes {
        slide_id: slide_id(closing),
        range: pptx_edit::structured::TextSpan { start: 0, end: 0 },
    };
    assert!(content.diagnostics.iter().any(|diagnostic| {
        diagnostic.anchor.as_ref() == Some(&aside)
            && diagnostic.message.contains("1 more text shape")
    }));
}

fn slide_id(slide: &ExportSlide) -> String {
    match &slide.anchor {
        PptxAnchor::Slide { slide_id } => slide_id.clone(),
        other => panic!("slide anchored at {other:?}"),
    }
}

#[test]
fn bytes_exports_are_deterministic_and_match_a_fresh_session() {
    let deck = fixture();
    let first = export(
        &deck,
        json!({"includeNotes": true, "includeComments": true}),
    );
    let second = export(
        &deck,
        json!({"includeNotes": true, "includeComments": true}),
    );
    assert_eq!(
        serde_json::to_string(&first).unwrap(),
        serde_json::to_string(&second).unwrap()
    );
    let session = DeckSession::open(&deck, 4242).unwrap();
    let mut live = session_export(
        &session,
        json!({"includeNotes": true, "includeComments": true}),
    );
    assert_eq!(live.anchor_scope, AnchorScope::Session);
    live.anchor_scope = AnchorScope::Snapshot;
    assert_eq!(live, first);
}

#[test]
fn session_exports_read_the_current_deck_and_change_nothing() {
    let deck = fixture();
    let session = DeckSession::open(&deck, 83).unwrap();
    let context = EditCtx::local("test");
    let slides = session.slide_ids().unwrap();
    session.move_slide(&context, &slides[2], 0).unwrap();
    session
        .add_text_box(
            &context,
            &slides[0],
            &ShapeDraft {
                name: "Added".to_owned(),
                rect: ShapeRect {
                    x: 0,
                    y: 0,
                    width: 100,
                    height: 100,
                },
                text: "Fresh text".to_owned(),
                style: TextStyle::default(),
            },
        )
        .unwrap();
    session
        .set_slide_notes(&context, &slides[0], "Rewritten notes")
        .unwrap();
    let updates = Rc::new(Cell::new(0));
    let counted = Rc::clone(&updates);
    let _subscription = session
        .observe_update_v1(move |_| counted.set(counted.get() + 1))
        .unwrap();
    let version = session.version();
    let state = session.encode_state_as_update_v1();
    let (can_undo, can_redo) = (session.can_undo(), session.can_redo());
    let saved = session.save().unwrap();

    let read = session
        .export_structured(&options(json!({"includeNotes": true})))
        .unwrap()
        .unwrap();
    let markdown = session
        .export_markdown(&options(json!({"includeNotes": true})))
        .unwrap()
        .unwrap();
    assert_eq!(read.version, version);
    assert_eq!(markdown.version, version);
    assert_eq!(session.version(), version);
    assert_eq!(session.encode_state_as_update_v1(), state);
    assert_eq!(
        (session.can_undo(), session.can_redo()),
        (can_undo, can_redo)
    );
    assert_eq!(updates.get(), 0);
    assert_eq!(session.save().unwrap(), saved);

    let content = read.content;
    assert_eq!(
        slide_names(&content)
            .into_iter()
            .map(|(index, _)| index)
            .collect::<Vec<_>>(),
        [0, 1]
    );
    assert_eq!(slide_id(&content.slides[0]), slides[2]);
    let added = shape(&content, "Added");
    assert!(added.provenance.is_none());
    assert_eq!(paragraph_texts(added), ["Fresh text"]);
    assert_eq!(
        content.slides[1].notes.as_ref().unwrap().text,
        "Rewritten notes"
    );
    assert!(markdown.content.markdown.contains("Fresh text"));

    let witness = DeckSession::open(&deck, 83).unwrap();
    let draft = ShapeDraft {
        name: "Probe".to_owned(),
        rect: ShapeRect {
            x: 0,
            y: 0,
            width: 10,
            height: 10,
        },
        text: String::new(),
        style: TextStyle::default(),
    };
    let fresh = witness.slide_ids().unwrap();
    witness.move_slide(&context, &fresh[2], 0).unwrap();
    witness.add_text_box(&context, &fresh[0], &draft).unwrap();
    let exported_then_added = session.add_text_box(&context, &slides[0], &draft).unwrap();
    let added_directly = witness.add_text_box(&context, &fresh[0], &draft).unwrap();
    assert_eq!(exported_then_added.shape_id, added_directly.shape_id);
}

#[test]
fn limits_stop_at_whole_records_and_bad_options_refuse() {
    let deck = fixture();
    let content = export(&deck, json!({"maxBlocks": 4}));
    assert!(content.truncated);
    assert_eq!(
        content.diagnostics.last().unwrap().code,
        ExportDiagnosticCode::Truncated
    );
    assert_eq!(content.slides.len(), 1);
    let shapes = &content.slides[0].shapes;
    assert_eq!(shapes.len(), 2);
    assert_eq!(shapes[0].stories[0].paragraphs.len(), 1);
    assert!(shapes[1].stories[0].paragraphs.is_empty());

    for max_bytes in [1_024, 4_096, 20_000] {
        let content = export(&deck, json!({"maxBytes": max_bytes}));
        assert!(content.truncated);
        assert!(serde_json::to_vec(&content).unwrap().len() <= max_bytes);
        assert_eq!(
            content.diagnostics.last().unwrap().code,
            ExportDiagnosticCode::Truncated
        );
    }

    let session = DeckSession::open(&deck, 84).unwrap();
    let version = session.version();
    for (value, code) in [
        (json!({"maxBytes": 10}), ExportFailureCode::InvalidOptions),
        (json!({"maxBlocks": 0}), ExportFailureCode::InvalidOptions),
        (
            json!({"maxBytes": 67_108_865}),
            ExportFailureCode::LimitExceeded,
        ),
        (
            json!({"maxBlocks": 1_000_001}),
            ExportFailureCode::LimitExceeded,
        ),
    ] {
        let refusal = session
            .export_structured(&options(value))
            .unwrap()
            .unwrap_err();
        assert_eq!(refusal.version, version);
        assert_eq!(refusal.failure.code, code);
        assert_eq!(refusal.failure.target, None);
        assert_eq!(
            serde_json::to_value(&refusal.failure).unwrap()["target"],
            Value::Null
        );
    }
    assert!(serde_json::from_value::<PptxExportOptions>(json!({"stories": []})).is_err());
    assert!(export_pptx_structured(b"not a deck", &PptxExportOptions::default()).is_err());
}

#[test]
fn markdown_renders_slides_lists_tables_and_markers_escaped() {
    let deck = fixture();
    let markdown = export_pptx_markdown(
        &deck,
        &options(json!({"includeNotes": true, "includeComments": true})),
    )
    .unwrap();
    let text = &markdown.markdown;
    assert!(text.contains("## Slide 1\n"));
    assert!(text.contains("## Slide 3: Closing\n"));
    assert!(text.contains("### **Quarterly review**"));
    assert!(text.contains("- First point"));
    assert!(text.contains("    - Sub point 😀"));
    assert!(text.contains("3. Three"));
    assert!(text.contains("- I. One"));
    assert!(text.contains("[Docs](https://example.com/docs?a=1&amp;b=2)"));
    assert!(text.contains(r"end \<script\>"));
    assert!(text.contains("<!-- pptx-unsupported: a14:m -->"));
    assert!(text.contains("![Company logo]()"));
    assert!(text.contains(r"\[Chart: Revenue by quarter\]"));
    assert!(text.contains("> Talk track"));
    assert!(text.contains("- **Ada** (2026-01-02T03:04:05.000): Check numbers"));
    assert!(!text.contains("Layout prompt text"));
    assert!(!text.contains("Backup slide"));
    let markers = text.matches("<!-- pptx-export:").count();
    assert_eq!(markers, markdown.anchors.len());
    for (index, anchor) in markdown.anchors.iter().enumerate() {
        assert_eq!(anchor.marker, format!("pptx-export:{index}"));
        assert!(text.contains(&format!("<!-- pptx-export:{index} -->")));
    }
    assert!(
        markdown
            .diagnostics
            .iter()
            .any(|diagnostic| diagnostic.code == ExportDiagnosticCode::MarkdownLossy)
    );

    let content = export(&deck, json!({}));
    let rendered = render_pptx_markdown(&content, &PptxMarkdownOptions::default()).unwrap();
    assert_eq!(
        rendered,
        export_pptx_markdown(&deck, &PptxExportOptions::default()).unwrap()
    );
    let short = render_pptx_markdown(
        &content,
        &PptxMarkdownOptions {
            max_bytes: Some(1_024),
        },
    )
    .unwrap();
    assert!(short.truncated);
    assert!(short.markdown.len() <= 1_024);
    assert_eq!(
        short.markdown.matches("<!-- pptx-export:").count(),
        short.anchors.len()
    );
}

#[test]
fn the_renderer_validates_what_it_is_given() {
    let content = export(&fixture(), json!({}));
    let mut value = serde_json::to_value(&content).unwrap();
    value["schemaVersion"] = json!(2);
    assert!(serde_json::from_value::<PptxStructuredContent>(value).is_err());
    let mut broken = content.clone();
    let paragraph = &mut broken.slides[0].shapes[0].stories[0].paragraphs[0];
    if let PptxAnchor::Text { range, .. } = &mut paragraph.anchor {
        range.start = range.end + 1;
    }
    let failure = render_pptx_markdown(&broken, &PptxMarkdownOptions::default()).unwrap_err();
    assert_eq!(failure.code, ExportFailureCode::InvalidContent);
    let mut grid = content.clone();
    let table = grid.slides[0].shapes[8].table.as_mut().unwrap();
    table.rows[0].cells[0].grid_span = 5;
    assert_eq!(
        render_pptx_markdown(&grid, &PptxMarkdownOptions::default())
            .unwrap_err()
            .code,
        ExportFailureCode::InvalidContent
    );
    assert_eq!(
        render_pptx_markdown(
            &content,
            &PptxMarkdownOptions {
                max_bytes: Some(10)
            }
        )
        .unwrap_err()
        .code,
        ExportFailureCode::InvalidOptions
    );
}

#[test]
fn collaboration_sessions_without_source_bytes_diagnose_provenance() {
    let deck = fixture();
    let seeded = DeckSession::open(&deck, 85).unwrap();
    let joined = DeckSession::open_from_update(&seeded.encode_state_as_update_v1(), 86).unwrap();
    let content = session_export(&joined, json!({}));
    assert_eq!(
        slide_names(&content)
            .into_iter()
            .map(|(index, _)| index)
            .collect::<Vec<_>>(),
        [0, 2]
    );
    assert!(shape(&content, "Title").provenance.is_none());
    assert!(content.slides[0].provenance.is_none());
    assert_eq!(
        shape(&content, "Revenue chart").description.as_deref(),
        Some("Revenue by quarter")
    );
    assert!(codes(&content).contains(&ExportDiagnosticCode::ProvenanceUnavailable));
    assert!(
        !content.slides[0]
            .shapes
            .iter()
            .any(|shape| shape.name == "Ink")
    );

    let update = include_bytes!("fixtures/deck-schema-v2-connectors.update.bin");
    let legacy = DeckSession::open_from_update(update, 87).unwrap();
    let unknown = session_export(&legacy, json!({}));
    assert!(!unknown.slides.is_empty());
    assert!(unknown.slides.iter().all(|slide| slide.hidden.is_none()));
    assert!(codes(&unknown).contains(&ExportDiagnosticCode::VisibilityUnknown));

    let source = include_bytes!("fixtures/deck-schema-v2-connectors.pptx");
    let reattached = DeckSession::open_from_update_with_source(update, source, 88).unwrap();
    let recovered = session_export(&reattached, json!({}));
    assert_eq!(recovered.slides.len(), unknown.slides.len());
    assert!(
        recovered
            .slides
            .iter()
            .all(|slide| slide.hidden == Some(false) && slide.provenance.is_some())
    );
    assert!(!codes(&recovered).contains(&ExportDiagnosticCode::VisibilityUnknown));
}

/// Pins the full JSON and Markdown of the fixture; `PPTX_EXPORT_GOLDEN_UPDATE=1` rewrites them.
#[test]
fn golden_json_and_markdown_pin_the_contract() {
    let deck = fixture();
    let options = options(
        json!({"includeHiddenSlides": true, "includeNotes": true, "includeComments": true}),
    );
    let content = export_pptx_structured(&deck, &options).unwrap();
    let json = serde_json::to_string_pretty(&content).unwrap() + "\n";
    let markdown = export_pptx_markdown(&deck, &options).unwrap().markdown;
    let fixtures = std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures");
    let json_path = fixtures.join("structured-export.golden.json");
    let markdown_path = fixtures.join("structured-export.golden.md");
    if std::env::var("PPTX_EXPORT_GOLDEN_UPDATE").as_deref() == Ok("1") {
        std::fs::write(&json_path, &json).unwrap();
        std::fs::write(&markdown_path, &markdown).unwrap();
    }
    assert_eq!(json, std::fs::read_to_string(json_path).unwrap());
    assert_eq!(markdown, std::fs::read_to_string(markdown_path).unwrap());
}

#[test]
fn edited_paragraphs_keep_what_can_still_be_located() {
    let session = DeckSession::open(&fixture(), 88).unwrap();
    let linked = session.snapshot().unwrap().slides[0]
        .shapes
        .iter()
        .find(|shape| shape.name == "Linked")
        .unwrap()
        .text_stories[0]
        .id
        .clone();
    session
        .delete_text(&EditCtx::local("test"), &linked, 5, 11)
        .unwrap();
    let content = session_export(&session, json!({"includeFormatting": false}));
    let paragraph = &shape(&content, "Linked").stories[0].paragraphs[0];
    assert!(paragraph.runs.iter().all(|run| run.marks.is_none()));
    assert_eq!(
        paragraph.runs[0]
            .link
            .as_ref()
            .map(|link| link.href.as_str()),
        Some("https://example.com/docs?a=1&b=2")
    );
    assert!(
        !paragraph
            .runs
            .iter()
            .any(|run| matches!(run.content, ExportRunKind::Field { .. }))
    );
    let unsupported = paragraph
        .runs
        .iter()
        .find(|run| matches!(run.content, ExportRunKind::Unsupported { .. }))
        .unwrap();
    assert!(matches!(
        &unsupported.anchor,
        PptxAnchor::Text { range, .. } if (range.start, range.end) == (10, 10)
    ));
    assert!(content.diagnostics.iter().any(|diagnostic| {
        diagnostic.code == ExportDiagnosticCode::ProvenanceUnavailable
            && diagnostic.anchor.as_ref() == Some(&paragraph.anchor)
    }));
    assert!(codes(&content).contains(&ExportDiagnosticCode::FormattingOmitted));
}

#[test]
fn limits_are_checked_before_slides_and_stories_are_read() {
    let session = DeckSession::open(&fixture(), 89).unwrap();
    let slide = session.slide_ids().unwrap()[0].clone();
    let context = EditCtx::local("test");
    let mut draft = ShapeDraft {
        name: "Huge".to_owned(),
        rect: ShapeRect {
            x: 0,
            y: 0,
            width: 100,
            height: 100,
        },
        text: "x".repeat(2_000_000),
        style: TextStyle::default(),
    };
    session.add_text_box(&context, &slide, &draft).unwrap();
    draft.text = "y".repeat(10);
    for index in 0..200 {
        draft.name = format!("Box {index}");
        session.add_text_box(&context, &slide, &draft).unwrap();
    }
    let session_ids = session.slide_ids().unwrap();
    session.move_slide(&context, &session_ids[0], 2).unwrap();

    let content = session_export(&session, json!({"maxBytes": 100_000}));
    assert!(content.truncated);
    let huge = shape(&content, "Huge");
    assert!(huge.stories.is_empty());
    assert!(serde_json::to_vec(&content).unwrap().len() <= 100_000);
    assert!(!serde_json::to_string(&content).unwrap().contains("Box 0"));

    let content = session_export(&session, json!({"maxBlocks": 3}));
    assert!(content.truncated);
    assert_eq!(content.slides.len(), 1);
}

#[test]
fn truncated_tables_stay_renderable() {
    let deck = fixture();
    let full = serde_json::to_vec(&export(&deck, json!({}))).unwrap().len();
    let mut checked = 0;
    let limits = (1..=80).map(|blocks| json!({"maxBlocks": blocks})).chain(
        (1_024..full)
            .step_by(211)
            .map(|bytes| json!({"maxBytes": bytes})),
    );
    for limit in limits {
        let content = export(&deck, limit.clone());
        let rendered = render_pptx_markdown(&content, &PptxMarkdownOptions::default());
        assert!(rendered.is_ok(), "{limit}: {:?}", rendered.err());
        if let Some(table) = content
            .slides
            .iter()
            .find_map(|slide| find(&slide.shapes, "Tall table"))
            .and_then(|shape| shape.table.as_ref())
        {
            let rows = table.rows.len() as u32;
            assert!(
                table
                    .rows
                    .iter()
                    .flat_map(|row| &row.cells)
                    .all(|cell| cell.row + cell.row_span <= rows)
            );
            checked += usize::from(rows < 3);
        }
    }
    assert!(checked > 0);
    let whole = export(&deck, json!({}));
    let tall = shape(&whole, "Tall table").table.as_ref().unwrap();
    assert_eq!(tall.rows[0].cells[0].row_span, 3);
    assert_eq!(
        tall.rows[2].cells[0]
            .merge_origin
            .map(|origin| (origin.row, origin.column)),
        Some((0, 0))
    );
}

fn title(content: &mut PptxStructuredContent) -> &mut pptx_edit::structured::ExportParagraph {
    &mut content.slides[0].shapes[0].stories[0].paragraphs[0]
}

/// The first story with three or more paragraphs.
fn list_story(content: &mut PptxStructuredContent) -> &mut pptx_edit::structured::ExportStory {
    content.slides[0]
        .shapes
        .iter_mut()
        .flat_map(|shape| &mut shape.stories)
        .find(|story| story.paragraphs.len() >= 3)
        .unwrap()
}

#[test]
fn the_renderer_refuses_content_that_breaks_its_contract() {
    let content = export(&fixture(), json!({}));
    let refused = |mutate: &dyn Fn(&mut PptxStructuredContent)| {
        let mut broken = content.clone();
        mutate(&mut broken);
        render_pptx_markdown(&broken, &PptxMarkdownOptions::default())
            .unwrap_err()
            .code
    };
    type Mutation = Box<dyn Fn(&mut PptxStructuredContent)>;
    let mutations: Vec<Mutation> = vec![
        Box::new(|content| content.schema_version = 2),
        Box::new(|content| {
            title(content).anchor = PptxAnchor::Slide {
                slide_id: "slide:0:256".to_owned(),
            }
        }),
        Box::new(|content| {
            if let PptxAnchor::Text { shape_id, .. } = &mut title(content).runs[0].anchor {
                *shape_id = "elsewhere".to_owned();
            }
        }),
        Box::new(|content| {
            if let ExportRunKind::Text { text } = &mut title(content).runs[0].content {
                text.push('!');
            }
        }),
        Box::new(|content| {
            let id = content.slides[0].shapes[0].id.clone();
            content.slides[0].shapes[1].id = id;
        }),
        Box::new(|content| {
            let table = content.slides[0].shapes[8].table.as_mut().unwrap();
            table.rows[0].cells[1].merge_origin =
                Some(pptx_edit::structured::CellPosition { row: 1, column: 0 });
        }),
        Box::new(|content| {
            title(content).runs[0].marks =
                Some(vec![pptx_edit::structured::ExportMark::Bold; 1_000]);
        }),
        Box::new(|content| {
            title(content).runs[0].marks = Some(vec![
                pptx_edit::structured::ExportMark::Italic,
                pptx_edit::structured::ExportMark::Italic,
            ]);
        }),
        Box::new(|content| {
            list_story(content).paragraphs.remove(1);
        }),
        Box::new(|content| {
            list_story(content).paragraphs.pop();
        }),
        Box::new(|content| {
            let table = content.slides[0].shapes[8].table.as_mut().unwrap();
            table.rows[1].cells.swap(0, 1);
        }),
        Box::new(|content| {
            let mut shape = content.slides[0].shapes[0].clone();
            for depth in 0..130 {
                let mut parent = shape.clone();
                parent.id = format!("deep{depth}");
                parent.stories.clear();
                parent.children = vec![shape];
                shape = parent;
            }
            content.slides[0].shapes.push(shape);
        }),
    ];
    for mutate in &mutations {
        assert_eq!(refused(mutate.as_ref()), ExportFailureCode::InvalidContent);
    }
}

#[test]
fn caller_content_renders_without_overflow_or_block_syntax() {
    let mut content = export(
        &fixture(),
        json!({"includeNotes": true, "includeComments": true}),
    );
    content.slides[0].index = u32::MAX;
    let span = |text: &str| pptx_edit::structured::TextSpan {
        start: 0,
        end: text.encode_utf16().count() as u32,
    };
    let notes = content.slides[0].notes.as_mut().unwrap();
    notes.text = "# Heading\n- item\n1. one\n    code\n\t- nested\n  # indented\n===".to_owned();
    if let PptxAnchor::Notes {
        range: anchored, ..
    } = &mut notes.anchor
    {
        *anchored = span(&notes.text);
    }
    let comment = &mut content.slides[0].comments[0];
    (comment.author, comment.date) = (None, None);
    comment.text = "  # Heading\n- item".to_owned();
    if let PptxAnchor::Comment {
        range: anchored, ..
    } = &mut comment.anchor
    {
        *anchored = span(&comment.text);
    }
    let markdown = render_pptx_markdown(&content, &PptxMarkdownOptions::default())
        .unwrap()
        .markdown;
    assert!(markdown.contains("## Slide 4294967296\n"), "{markdown}");
    for line in [
        r"> \# Heading",
        r"> \- item",
        r"> 1\. one",
        "> \u{a0}\u{a0}\u{a0}\u{a0}code",
        "> \u{a0}\u{a0}\u{a0}\u{a0}\\- nested",
        "> \u{a0}\u{a0}\\# indented",
        r"> \===",
    ] {
        assert!(
            markdown.lines().any(|candidate| candidate == line),
            "{line}: {markdown}"
        );
    }
    assert!(
        markdown.contains("- \u{a0}\u{a0}\\# Heading<br>- item"),
        "{markdown}"
    );
}

#[test]
fn links_with_encoded_script_schemes_are_not_linked() {
    let mut content = export(&fixture(), json!({}));
    let linked = content.slides[0]
        .shapes
        .iter()
        .position(|shape| shape.name == "Linked")
        .unwrap();
    let render = |content: &mut PptxStructuredContent, href: &str| {
        let link = Some(pptx_edit::structured::ExportLink {
            href: href.to_owned(),
            external: true,
        });
        content.slides[0].shapes[linked].stories[0].paragraphs[0].runs[0].link = link.clone();
        let table = content.slides[0].shapes[8].table.as_mut().unwrap();
        table.rows[0].cells[0].story.as_mut().unwrap().paragraphs[0].runs[0].link = link;
        render_pptx_markdown(content, &PptxMarkdownOptions::default())
            .unwrap()
            .markdown
    };
    let spaced = format!("{}javascript:alert(1)", " ".repeat(256));
    let tabbed = format!("{}javascript:alert(1)", "\t".repeat(300));
    let nested = format!("&{}#106;avascript:alert(1)", "amp;".repeat(20));
    for href in [
        "jav&#x61;script:alert(1)",
        "javascript&colon;alert(1)",
        "javascript&#58;alert(1)",
        "java&#x09;script:alert(1)",
        "&#106;avascript:alert(1)",
        "&#106avascript:alert(1)",
        "&#x6A&#x61vascript&#x3A;alert(1)",
        "%6Aavascript:alert(1)",
        "java%0Ascript:alert(1)",
        "JaVaScRiPt:alert(1)",
        " JavaScript:alert(1)",
        spaced.as_str(),
        tabbed.as_str(),
        nested.as_str(),
        "\u{1}javascript:alert(1)",
        "java\u{0}script:alert(1)",
        "java\rscript:alert(1)",
        "java\u{7f}script:alert(1)",
        "vbscript:msgbox(1)",
        "file:///etc/passwd",
        "data:text/html,<script>",
    ] {
        let markdown = render(&mut content, href);
        assert!(!markdown.contains("[Docs]"), "{href:?}");
        assert!(markdown.contains("Docs page 1"), "{href:?}");
        assert!(!markdown.contains("<a "), "{href:?}");
        assert!(markdown.contains("<p>Merged header</p>"), "{href:?}");
    }
    let markdown = render(&mut content, "https://a.test/?a=1&b=2");
    assert!(markdown.contains("[Docs](https://a.test/?a=1&amp;b=2)"));
    assert!(markdown.contains(r#"<a href="https://a.test/?a=1&amp;b=2">Merged header</a>"#));
    let markdown = render(&mut content, "https://a.test/a b(1)?t=<x>");
    assert!(markdown.contains("[Docs](https://a.test/a%20b%281%29?t=%3Cx%3E)"));
    assert!(markdown.contains(r#"<a href="https://a.test/a%20b(1)?t=&lt;x&gt;">"#));
    for href in [
        "slides/2#top",
        "//example.test/x",
        "tel:+15551234",
        "ftp://example.test/f",
    ] {
        let markdown = render(&mut content, href);
        assert!(!markdown.contains("[Docs]"), "{href:?}");
        assert!(markdown.contains("Docs page 1"), "{href:?}");
    }
}

/// The fixture with `(part, from, to)` substitutions in its XML.
fn fixture_with(changes: &[(&str, &str, &str)]) -> Vec<u8> {
    let mut parts = parts();
    for (part, from, to) in changes {
        let (_, bytes) = parts.iter_mut().find(|(path, _)| path == part).unwrap();
        let xml = String::from_utf8(std::mem::take(bytes)).unwrap();
        assert!(xml.contains(from), "{from}");
        *bytes = xml.replace(from, to).into_bytes();
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn text_and_metadata_stay_inert_in_markdown_and_html_tables() {
    const IMG: &str = "&#10;&#13;&#10;&#x2028;&lt;img src=x onerror=alert(1)&gt;";
    const SCRIPT: &str = "&lt;/a&gt;&lt;script&gt;alert(1)&lt;/script&gt;";
    let slide1 = "ppt/slides/slide1.xml";
    let slide3 = "ppt/slides/slide3.xml";
    let deck = fixture_with(&[
        (
            slide1,
            ">Quarterly review<",
            &format!(">Quarterly review{IMG}<"),
        ),
        (
            slide1,
            r#"descr="Company logo""#,
            &format!(r#"descr="Company logo{IMG}""#),
        ),
        (
            slide3,
            r#"name="Closing""#,
            &format!(r#"name="Closing{IMG}""#),
        ),
        (
            slide3,
            ">Last slide<",
            &format!(">Last slide{SCRIPT}{IMG}<"),
        ),
        (slide3, ">Tall<", &format!(">Tall{SCRIPT}{IMG}<")),
    ]);
    let markdown = export_pptx_markdown(&deck, &PptxExportOptions::default())
        .unwrap()
        .markdown;
    let live = markdown.replace(r"\<", "");
    for tag in ["<img", "<script", "</a>"] {
        assert!(!live.contains(tag), "{tag} in {markdown}");
    }
    let img = r"\<img src=x onerror=alert(1)\>";
    let line = |prefix: &str| {
        markdown
            .lines()
            .find(|line| line.starts_with(prefix))
            .unwrap_or_else(|| panic!("no {prefix:?} line in {markdown}"))
    };
    assert!(line("### ").contains(img));
    assert!(line("## Slide 3: Closing").ends_with(img));
    assert!(line("![Company logo").ends_with(&format!("{img}]()")));
    let paragraph = &markdown[markdown.find("Last slide").unwrap()..];
    let paragraph = &paragraph[..paragraph.find("\n\n").unwrap()];
    assert!(paragraph.ends_with(img), "{paragraph}");
    assert!(markdown.contains(
        "<p>Tall&lt;/a&gt;&lt;script&gt;alert(1)&lt;/script&gt;<br> <br> &lt;img src=x onerror=alert(1)&gt;</p>"
    ));
}
