#![cfg(feature = "raster")]

use std::io::Cursor;

use betteroffice_pptx::{EditCtx, PngTarget, Presentation, Primitive, RenderOptions};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NS: &str = r#"xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships""#;
const REL: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

fn fixture() -> Vec<u8> {
    let slide = |text: &str, strike: &str| {
        format!(
            r#"<p:sld {NS}><p:cSld><p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="4267200" cy="1371600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0" wrap="none" anchor="t"/><a:lstStyle/><a:p><a:r><a:rPr sz="7200"{strike}><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#
        )
    };
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="{REL}/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#)),
        ("ppt/presentation.xml", format!(r#"<p:presentation {NS}><p:sldIdLst><p:sldId id="256" r:id="slide1"/><p:sldId id="257" r:id="slide2"/></p:sldIdLst><p:sldSz cx="6096000" cy="3429000"/></p:presentation>"#)),
        ("ppt/_rels/presentation.xml.rels", format!(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="slide1" Type="{REL}/slide" Target="slides/slide1.xml"/><Relationship Id="slide2" Type="{REL}/slide" Target="slides/slide2.xml"/></Relationships>"#)),
        ("ppt/slides/slide1.xml", slide("Struck", r#" strike="sngStrike""#)),
        ("ppt/slides/slide2.xml", slide("Unrelated", "")),
    ]
    .into_iter()
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
    .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn open(bytes: &[u8]) -> Presentation {
    let mut presentation = Presentation::open(bytes).unwrap();
    presentation
        .register_font("Arial", false, false, FONT)
        .unwrap();
    presentation
}

fn slide_xml(bytes: &[u8], number: usize) -> String {
    let path = format!("ppt/slides/slide{number}.xml");
    String::from_utf8(
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(name, _)| name == &path)
            .unwrap()
            .1,
    )
    .unwrap()
}

fn assert_struck_run(xml: &str, text: &str) {
    let text_xml = format!("<a:t>{text}</a:t>");
    let run = xml
        .split("<a:r>")
        .skip(1)
        .map(|tail| tail.split("</a:r>").next().unwrap())
        .find(|run| run.contains(&text_xml))
        .expect("expected text run in saved slide XML");
    let properties = run
        .split("<a:rPr")
        .nth(1)
        .expect("expected run properties")
        .split('>')
        .next()
        .unwrap();
    assert!(
        properties.contains(r#" strike="sngStrike""#),
        "strike missing from {text:?}: {run}"
    );
}

#[test]
fn single_strike_paints_a_continuous_rule_through_the_run() {
    let presentation = open(&fixture());
    let layout = presentation.render_slide(0).unwrap();
    let (baseline, font_size) = layout
        .display_list
        .primitives
        .iter()
        .find_map(|primitive| match primitive {
            Primitive::TextBox { lines, .. } => lines.iter().find_map(|line| {
                line.runs
                    .iter()
                    .find(|run| run.text == "Struck")
                    .map(|run| (line.baseline, run.font_size_px))
            }),
            _ => None,
        })
        .expect("expected the Struck run in the display list");
    let rendered = presentation
        .render_png(PngTarget { slide: 0 }, &RenderOptions::default())
        .unwrap();
    let mut reader = png::Decoder::new(Cursor::new(&rendered.bytes))
        .read_info()
        .unwrap();
    let mut pixels = vec![0; reader.output_buffer_size().unwrap()];
    let info = reader.next_frame(&mut pixels).unwrap();
    assert_eq!(info.color_type, png::ColorType::Rgba);
    assert_eq!((info.width, info.height), (640, 360));
    let width = info.width as usize;
    let height = info.height as usize;
    let dark = |x: usize, y: usize| {
        let pixel = &pixels[(y * width + x) * 4..(y * width + x) * 4 + 4];
        pixel[0] < 64 && pixel[1] < 64 && pixel[2] < 64 && pixel[3] > 200
    };
    let (mut left, mut top, mut right, mut bottom) = (width, height, 0, 0);
    for y in 0..height {
        for x in 0..width {
            if dark(x, y) {
                left = left.min(x);
                top = top.min(y);
                right = right.max(x);
                bottom = bottom.max(y);
            }
        }
    }
    assert!(left < right && top < bottom, "expected visible text ink");
    let strike_y = baseline - font_size * 0.3;
    assert!(
        (top as f32) < strike_y && strike_y < bottom as f32,
        "strike row must lie between the glyph top and bottom"
    );
    let start = ((strike_y - font_size * 0.06).floor() as usize).max(top + 1);
    let end = ((strike_y + font_size * 0.06).ceil() as usize).min(bottom - 1);
    let mut longest = 0;
    for y in start..=end {
        let mut span = 0;
        for x in left..=right {
            if dark(x, y) {
                span += 1;
                longest = longest.max(span);
            } else {
                span = 0;
            }
        }
    }
    let ink_width = right - left + 1;
    assert!(
        longest * 10 >= ink_width * 9,
        "expected a strike across at least 90% of {ink_width}px ink width near y={strike_y}, found {longest}px"
    );
}

#[test]
fn editing_another_slide_preserves_strike_xml() {
    let bytes = fixture();
    let presentation = open(&bytes);
    let before = presentation.snapshot().unwrap();
    assert_eq!(before.slides.len(), 2);
    let story_id = &before.slides[1].shapes[0].text_stories[0].id;
    presentation
        .delete_text(&EditCtx::local("strike-red"), story_id, 0, 2)
        .unwrap();
    let saved = presentation.save().unwrap();
    let reopened = open(&saved);
    let after = reopened.snapshot().unwrap();
    assert_eq!(after.slides[0], before.slides[0]);
    assert_eq!(
        after.slides[1].shapes[0].text_stories[0].plain_text(),
        "related"
    );
    let xml = slide_xml(&saved, 1);
    assert_struck_run(&xml, "Struck");
    assert_eq!(xml, slide_xml(&bytes, 1));
}

#[test]
fn editing_the_struck_runs_text_preserves_strike_xml() {
    let bytes = fixture();
    let presentation = open(&bytes);
    let before = presentation.snapshot().unwrap();
    let story_id = &before.slides[0].shapes[0].text_stories[0].id;
    presentation
        .delete_text(&EditCtx::local("strike-red"), story_id, 2, 3)
        .unwrap();
    let saved = presentation.save().unwrap();
    let reopened = open(&saved);
    let after = reopened.snapshot().unwrap();
    assert_eq!(
        after.slides[0].shapes[0].text_stories[0].plain_text(),
        "Stuck"
    );
    assert_eq!(after.slides[1], before.slides[1]);
    assert_struck_run(&slide_xml(&saved, 1), "Stuck");
    assert_eq!(slide_xml(&saved, 2), slide_xml(&bytes, 2));
}
