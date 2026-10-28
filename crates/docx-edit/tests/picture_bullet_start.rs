use docx_edit::bridge::{RenderEnv, yrs_doc_to_mapped_layout_blocks};
use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position, seed_from_docx};
use docx_parse::serializer::{S13SaveRequest, write_docx_s13};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const NS: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:v="urn:schemas-microsoft-com:vml""#
);

fn paragraph(id: u32, text: &str) -> String {
    format!(r#"<w:p w14:paraId="{id:08X}"><w:r><w:t>{text}</w:t></w:r></w:p>"#)
}

fn numbering(bullet: &str, referenced: bool) -> String {
    let reference = if referenced {
        r#"<w:lvlPicBulletId w:val="7"/>"#
    } else {
        ""
    };
    format!(
        r#"<w:numbering {NS}>{bullet}<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="*"/>{reference}</w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num></w:numbering>"#
    )
}

fn bullet(style: &str) -> String {
    format!(
        r#"<w:numPicBullet w:numPicBulletId="7"><w:pict><v:shape style="{style}"/></w:pict></w:numPicBullet>"#
    )
}

fn document(body: &str, numbering: Option<&str>) -> Vec<u8> {
    let mut parts = vec![
        (
            "[Content_Types].xml".to_owned(),
            br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/numbering.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml"/></Types>"#.to_vec(),
        ),
        (
            "_rels/.rels".to_owned(),
            br#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_vec(),
        ),
        (
            "word/document.xml".to_owned(),
            format!(
                r#"<w:document {NS}><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#
            )
            .into_bytes(),
        ),
    ];
    if let Some(numbering) = numbering {
        parts.push(("word/numbering.xml".to_owned(), numbering.as_bytes().to_vec()));
    }
    let numbering_rel = if numbering.is_some() {
        r#"<Relationship Id="numbering" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" Target="numbering.xml"/>"#
    } else {
        ""
    };
    parts.push((
        "word/_rels/document.xml.rels".to_owned(),
        format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{numbering_rel}<Relationship Id="image" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image.svg"/></Relationships>"#
        )
        .into_bytes(),
    ));
    parts.push((
        "word/media/image.svg".to_owned(),
        br#"<svg xmlns="http://www.w3.org/2000/svg" width="10" height="10"><rect width="10" height="10" fill="red"/></svg>"#.to_vec(),
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn render(bytes: &[u8]) -> (EngineSession, Value, docx_layout::display_list::DisplayList) {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(79101);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body",
                "renderEnv": {},
                "regions": {"sections": [{"properties": package.document.final_section_properties}]},
                "measurement": {
                    "fontChains": {"arial|0|0": [font]},
                    "defaults": {"fontFamily": "Arial", "fontSize": 12},
                    "authoritativeShaping": true
                }
            })
            .to_string(),
        )
        .unwrap();
    let display = serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    (engine, serde_json::from_str(&output).unwrap(), display)
}

fn origin(display: &docx_layout::display_list::DisplayList, text: &str) -> (f64, f64) {
    let display = serde_json::to_value(display).unwrap();
    let primitive = display["pages"][0]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .find(|primitive| {
            primitive["text"]
                .as_str()
                .is_some_and(|value| !value.is_empty() && text.starts_with(value))
        })
        .unwrap();
    if let Some(glyph) = primitive["glyphs"].as_array().and_then(|glyphs| glyphs.first()) {
        (glyph["x"].as_f64().unwrap(), glyph["y"].as_f64().unwrap())
    } else {
        (primitive["x"].as_f64().unwrap(), primitive["baselineY"].as_f64().unwrap())
    }
}

fn near(actual: f64, expected: f64) {
    assert!((actual - expected).abs() < 0.1, "{actual} vs {expected}");
}

#[test]
fn sized_picture_bullet_advances_only_the_first_paragraph_without_numbered_text() {
    let body = format!("{}{}", paragraph(1, "First"), paragraph(2, "Second"));
    let (_, base_layout, base_display) = render(&document(&body, None));
    let (base_x, _) = origin(&base_display, "First");
    for (style, width, height) in [
        ("width:3in;height:3in", 216.0, 216.0),
        ("width:1in;height:1in", 72.0, 72.0),
        ("width:2in;height:1in", 144.0, 72.0),
    ] {
        for referenced in [false, true] {
            let xml = numbering(&bullet(style), referenced);
            let (_, layout, display) = render(&document(&body, Some(&xml)));
            let (x, baseline) = origin(&display, "First");
            near((x - base_x) * 0.75, width);
            let top = layout["layout"]["pages"][0]["fragments"][0]["y"]
                .as_f64()
                .unwrap();
            near((baseline - top) * 0.75, height);
            near(origin(&display, "Second").0, origin(&base_display, "Second").0);
            assert_eq!(
                layout["measured"][0]["block"]["runs"],
                base_layout["measured"][0]["block"]["runs"]
            );
        }
    }
}

#[test]
fn picture_bullet_moves_an_existing_inline_picture_to_its_baseline() {
    let body = r#"<w:p w14:paraId="00000001"><w:r><w:pict><v:shape style="width:209pt;height:48.6pt"><v:imagedata r:id="image"/></v:shape></w:pict></w:r></w:p>"#;
    let (_, _, base) = render(&document(body, None));
    let xml = numbering(&bullet("width:3in;height:3in"), false);
    let (_, _, display) = render(&document(body, Some(&xml)));
    let image = |display: &docx_layout::display_list::DisplayList| {
        let value = serde_json::to_value(display).unwrap();
        value["pages"][0]["primitives"]
            .as_array()
            .unwrap()
            .iter()
            .find(|primitive| primitive["kind"] == "image")
            .unwrap()
            .clone()
    };
    let base = image(&base);
    let image = image(&display);
    near((image["x"].as_f64().unwrap() - base["x"].as_f64().unwrap()) * 0.75, 216.0);
    near((image["y"].as_f64().unwrap() - base["y"].as_f64().unwrap()) * 0.75, 167.4);
}

#[test]
fn missing_bullet_or_invalid_size_leaves_layout_unchanged() {
    let body = paragraph(1, "First");
    let (_, base, _) = render(&document(&body, None));
    for picture in [
        String::new(),
        bullet(""),
        bullet("width:3in"),
        bullet("height:3in"),
        bullet("width:0pt;height:3in"),
        bullet("width:3in;height:-1pt"),
    ] {
        let xml = numbering(&picture, false);
        let (_, actual, _) = render(&document(&body, Some(&xml)));
        assert_eq!(actual, base);
    }
}

#[test]
fn table_first_picture_bullet_wraps_cell_text_below_its_line() {
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="4000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>{}</w:tc></w:tr></w:tbl>{}"#,
        paragraph(1, "Cell"),
        paragraph(2, "After")
    );
    let (_, base, base_display) = render(&document(&body, None));
    let xml = numbering(&bullet("width:3in;height:3in"), false);
    let (_, layout, display) = render(&document(&body, Some(&xml)));
    let base_cell = &base["measured"][0]["measure"]["rows"][0]["cells"][0]["blocks"][0];
    let cell = &layout["measured"][0]["measure"]["rows"][0]["cells"][0]["blocks"][0];
    near(cell["lines"][0]["ascent"].as_f64().unwrap() * 0.75, 216.0);
    assert!(cell["lines"].as_array().unwrap().len() > base_cell["lines"].as_array().unwrap().len());
    assert!(origin(&display, "Cell").1 - origin(&base_display, "Cell").1 >= 288.0);
    near(origin(&display, "After").0, origin(&base_display, "After").0);
    let start = layout["measured"][0]["block"]["rows"][0]["cells"][0]["blocks"][0]["runs"][0]
        ["pmStart"]
        .as_f64()
        .unwrap() as i64;
    let caret = docx_layout::hit::caret_rect(&display, start).unwrap();
    near(caret.x, origin(&display, "Cell").0);
}

#[test]
fn picture_bullet_keeps_source_offsets_and_caret_positions() {
    let body = paragraph(1, "First");
    let xml = numbering(&bullet("width:3in;height:3in"), false);
    let (base, _, base_display) = render(&document(&body, None));
    let (engine, _, display) = render(&document(&body, Some(&xml)));
    let (base_blocks, base_map) =
        yrs_doc_to_mapped_layout_blocks(base.doc(), "body", &RenderEnv::default()).unwrap();
    let (blocks, map) =
        yrs_doc_to_mapped_layout_blocks(engine.doc(), "body", &RenderEnv::default()).unwrap();
    assert_eq!(map, base_map);
    assert_eq!(
        engine.doc().paragraphs("body").unwrap(),
        base.doc().paragraphs("body").unwrap()
    );
    assert_eq!(
        engine.doc().story_len("body").unwrap(),
        base.doc().story_len("body").unwrap()
    );
    let blocks = serde_json::to_value(blocks).unwrap();
    let base_blocks = serde_json::to_value(base_blocks).unwrap();
    assert_eq!(blocks[0]["runs"], base_blocks[0]["runs"]);
    for position in 1..=6 {
        let caret = docx_layout::hit::caret_rect(&display, position).unwrap();
        let base_caret = docx_layout::hit::caret_rect(&base_display, position).unwrap();
        near((caret.x - base_caret.x) * 0.75, 216.0);
        assert_eq!(
            docx_layout::hit::hit_test(
                &display,
                caret.page_index,
                caret.x,
                caret.y + caret.height / 2.0,
            ),
            Some(position)
        );
    }
    engine
        .doc()
        .insert_text(
            &EditCtx::local("Owner", "2026-10-28T00:00:00Z"),
            Position::new("body", 0),
            "X",
            FormatPolicy::Inherit,
        )
        .unwrap();
    assert_eq!(engine.doc().paragraphs("body").unwrap()[0].text, "XFirst");
}

#[test]
fn empty_first_paragraph_reserves_the_box_and_other_stories_do_not() {
    let body = format!("{}{}", paragraph(1, ""), paragraph(2, "Second"));
    let xml = numbering(&bullet("width:3in;height:3in"), false);
    let (engine, layout, _) = render(&document(&body, Some(&xml)));
    near(layout["measured"][0]["measure"]["lines"][0]["ascent"].as_f64().unwrap(), 288.0);
    assert!(layout["measured"][1]["block"]["attrs"]["leadingInlineBox"].is_null());
    for story in ["header:1", "footer:1", "fn:1", "en:1"] {
        engine.doc().create_story(story, "Other", "Normal", "left").unwrap();
        let (blocks, _) =
            yrs_doc_to_mapped_layout_blocks(engine.doc(), story, &RenderEnv::default()).unwrap();
        let blocks = serde_json::to_value(blocks).unwrap();
        assert!(blocks[0]["attrs"]["leadingInlineBox"].is_null());
    }
}

#[test]
fn untouched_save_preserves_picture_bullet_numbering_bytes() {
    let xml = numbering(&bullet("width:3in;height:3in"), false);
    let bytes = document(&paragraph(1, "First"), Some(&xml));
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let request: S13SaveRequest = serde_json::from_value(json!({
        "determinism": {"seed": "0".repeat(64), "now": "2026-10-28T00:00:00.000Z"},
        "document": package.document,
        "numbering": package.numbering,
        "options": {"updateModifiedDate": false}
    }))
    .unwrap();
    let saved = write_docx_s13(request, &bytes).unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let numbering = &parts
        .iter()
        .find(|(path, _)| path == "word/numbering.xml")
        .unwrap()
        .1;
    assert_eq!(numbering, xml.as_bytes());
}
