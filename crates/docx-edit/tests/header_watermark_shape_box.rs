use docx_edit::layout_request::{measurement, region_layout_request};
use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const HEADER_START: &str = concat!(
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office">"#,
    r#"<w:sdt><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Watermarks"/></w:docPartObj></w:sdtPr><w:sdtContent><w:p><w:r><w:pict>"#,
    r#"<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" adj="10800"><v:path textpathok="t"/><v:textpath on="t" fitshape="t"/></v:shapetype>"#,
);

const HEADER_END: &str = r#"</w:pict></w:r></w:p></w:sdtContent></w:sdt></w:hdr>"#;

fn watermarked_document(shape: [&str; 5]) -> Vec<u8> {
    watermarked_document_with_header_tail(shape, "")
}

fn watermarked_document_with_header_tail(
    [width, height, family, size, text]: [&str; 5],
    tail: &str,
) -> Vec<u8> {
    let shape = format!(
        r##"<v:shape id="PowerPlusWaterMarkObject1" o:spid="_x0000_s2049" type="#_x0000_t136" style="position:absolute;margin-left:0;margin-top:0;width:{width};height:{height};rotation:315;z-index:-251658752;mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin" o:allowincell="f" fillcolor="silver" stroked="f"><v:fill opacity=".5"/><v:textpath style="font-family:&quot;{family}&quot;;font-size:{size}" string="{text}"/></v:shape>"##
    );
    let header = format!("{HEADER_START}{shape}{HEADER_END}")
        .replace("</w:hdr>", &format!("{tail}</w:hdr>"));
    let parts = vec![
        ("[Content_Types].xml".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#.to_owned()),
        ("_rels/.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="rIdHeader"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>"#.to_owned()),
        ("word/_rels/document.xml.rels".to_owned(), r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdHeader" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/></Relationships>"#.to_owned()),
        ("word/header1.xml".to_owned(), header),
    ];
    ooxml_opc::rezip_parts(
        &parts
            .into_iter()
            .map(|(name, value)| (name, value.into_bytes()))
            .collect::<Vec<_>>(),
    )
    .unwrap()
}

fn painted_watermark(bytes: &[u8], font: u32) -> Value {
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let package = serde_json::to_value(package).unwrap();
    let mut request = region_layout_request(&package, 24.0);
    request["measurement"] = measurement(&package);
    request["measurement"]["fontChains"] = json!({"arial|0|0": [font], "calibri|0|0": [font]});
    let engine = EngineSession::new(73103);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(&request.to_string())
        .unwrap();
    let mut display_input: Value = serde_json::from_str(&output).unwrap();
    display_input["fontChains"] = request["measurement"]["fontChains"].clone();
    let display: Value = serde_json::from_str(
        &engine
            .build_display_list_json(&display_input.to_string())
            .unwrap(),
    )
    .unwrap();
    let page = &display["pages"][0];
    assert_eq!(page["watermarkPrimitiveCount"], json!(1));
    page["primitives"][0].clone()
}

fn face(run: &Value) -> &str {
    if run["kind"] == "glyphRun" {
        run["fallbackFont"].as_str().unwrap()
    } else {
        run["font"].as_str().unwrap()
    }
}

fn em_px(run: &Value) -> f64 {
    if run["kind"] == "glyphRun" {
        return run["size"].as_f64().unwrap();
    }
    face(run)
        .split_whitespace()
        .find_map(|part| part.strip_suffix("px")?.parse().ok())
        .unwrap()
}

struct PaintedRun {
    x: f64,
    width: f64,
    pivot_x: f64,
    baseline: f64,
}

fn painted_run(run: &Value) -> PaintedRun {
    if run["kind"] != "glyphRun" {
        let x = run["x"].as_f64().unwrap();
        let width = run["width"].as_f64().unwrap();
        return PaintedRun {
            x,
            width,
            pivot_x: x + width / 2.0,
            baseline: run["baselineY"].as_f64().unwrap(),
        };
    }
    let glyphs = run["glyphs"].as_array().unwrap();
    let left = glyphs
        .iter()
        .map(|glyph| glyph["x"].as_f64().unwrap())
        .fold(f64::MAX, f64::min);
    let right = glyphs
        .iter()
        .map(|glyph| glyph["x"].as_f64().unwrap() + glyph["advance"].as_f64().unwrap())
        .fold(f64::MIN, f64::max);
    let scale = run["horizontalScale"].as_f64().unwrap_or(100.0) / 100.0;
    PaintedRun {
        x: left,
        width: (right - left) * scale,
        pivot_x: (left + right) / 2.0,
        baseline: glyphs[0]["y"].as_f64().unwrap(),
    }
}

fn rotated_em_box_points(run: &Value) -> [f64; 4] {
    let em = em_px(run);
    let PaintedRun {
        x,
        width,
        pivot_x,
        baseline,
    } = painted_run(run);
    let pivot_y = baseline - 0.2 * em;
    let angle = run["rotationDeg"].as_f64().unwrap().to_radians();
    let (sin, cos) = angle.sin_cos();
    let top = baseline - 0.8 * em;
    let mut bounds = [f64::MAX, f64::MAX, f64::MIN, f64::MIN];
    for (corner_x, corner_y) in [
        (x, top),
        (x + width, top),
        (x, top + em),
        (x + width, top + em),
    ] {
        let (dx, dy) = (corner_x - pivot_x, corner_y - pivot_y);
        let turned_x = pivot_x + dx * cos - dy * sin;
        let turned_y = pivot_y + dx * sin + dy * cos;
        bounds = [
            bounds[0].min(turned_x),
            bounds[1].min(turned_y),
            bounds[2].max(turned_x),
            bounds[3].max(turned_y),
        ];
    }
    bounds.map(|value| value * 0.75)
}

#[test]
fn text_watermark_fills_its_shape_box_centred_on_the_margin_box() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let cases = [
        (
            ["592.1pt", "47.35pt", "Arial", "1pt", "FOR REVIEW"],
            [592.1, 47.35],
            [79.92, 169.92, 532.08, 622.08],
        ),
        (
            ["672pt", "58.5pt", "Calibri", "48pt", "Sample copy"],
            [672.0, 58.5],
            [47.73, 137.73, 564.27, 654.27],
        ),
        (
            ["412.4pt", "247.45pt", "Calibri", "1pt", "DRAFT"],
            [412.4, 247.45],
            [72.71, 162.71, 539.29, 629.29],
        ),
    ];
    for (shape, [width_pt, height_pt], expected) in cases {
        let [_, _, family, _, text] = shape;
        let run = painted_watermark(&watermarked_document(shape), font);
        assert_eq!(run["text"], text);
        assert_eq!(run["rotationDeg"].as_f64(), Some(-45.0), "{text}");
        let bounds = rotated_em_box_points(&run);
        for (got, want) in bounds.iter().zip(expected) {
            assert!(
                (got - want).abs() < 0.05,
                "{text}: box {bounds:?}, expected {expected:?}"
            );
        }
        let painted_width = painted_run(&run).width * 0.75;
        assert!((painted_width - width_pt).abs() < 0.01, "{text}: {run}");
        let em_pt = em_px(&run) * 0.75;
        assert!((em_pt - height_pt).abs() < 0.01, "{text}: {run}");
        assert!(face(&run).contains(family), "{text}: {run}");
        let stretch = run["horizontalScale"].as_f64();
        assert!(stretch.is_some_and(|scale| scale > 0.0), "{text}: {run}");
    }
}

#[test]
fn text_watermark_paints_the_glyphs_it_was_stretched_with() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let run = painted_watermark(
        &watermarked_document(["672pt", "58.5pt", "Calibri", "48pt", "Sample copy"]),
        font,
    );
    assert_eq!(run["kind"], "glyphRun", "{run}");
    assert_eq!(run["fontId"].as_u64(), Some(u64::from(font)), "{run}");
    assert_eq!(run["text"], "Sample copy");
    assert_eq!(run["rotationDeg"].as_f64(), Some(-45.0), "{run}");
    assert_eq!(run["opacity"].as_f64(), Some(0.5), "{run}");
    let scale = run["horizontalScale"].as_f64().unwrap() / 100.0;
    let advances: f64 = run["glyphs"]
        .as_array()
        .unwrap()
        .iter()
        .map(|glyph| glyph["advance"].as_f64().unwrap())
        .sum();
    assert!(
        (advances * scale * 0.75 - 672.0).abs() < 0.05,
        "{advances} x {scale}: {run}"
    );
}

fn stretched_text_run(text: &str) -> Value {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let run = painted_watermark(
        &watermarked_document(["672pt", "58.5pt", "Calibri", "48pt", text]),
        font,
    );
    assert_eq!(run["text"], text);
    assert_eq!(run["kind"], "text", "{run}");
    assert!(run["horizontalScale"].as_f64().is_some(), "{run}");
    assert!(
        (run["width"].as_f64().unwrap() * 0.75 - 672.0).abs() < 0.01,
        "{run}"
    );
    run
}

#[test]
fn text_watermark_with_a_character_its_first_face_lacks_keeps_the_text_run() {
    stretched_text_run("DRAFT \u{4E2D}");
}

#[test]
fn right_to_left_text_watermark_keeps_the_text_run() {
    stretched_text_run("\u{05D8}\u{05D9}\u{05D5}\u{05D8}\u{05D4}");
}

fn mark_centre(run: &Value) -> (f64, f64) {
    let painted = painted_run(run);
    (painted.pivot_x, painted.baseline - 0.2 * em_px(run))
}

#[test]
fn text_watermark_centres_on_the_authored_margins_under_a_header_taller_than_the_top_margin() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let shape = ["672pt", "58.5pt", "Calibri", "48pt", "Sample copy"];
    let tall = r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial"/><w:sz w:val="288"/></w:rPr><w:t>Tall</w:t></w:r></w:p>"#;
    let plain = mark_centre(&painted_watermark(&watermarked_document(shape), font));
    let under_tall_header = mark_centre(&painted_watermark(
        &watermarked_document_with_header_tail(shape, tall),
        font,
    ));
    assert!(
        (plain.0 - under_tall_header.0).abs() < 0.001
            && (plain.1 - under_tall_header.1).abs() < 0.001,
        "{plain:?} vs {under_tall_header:?}"
    );
}

#[test]
fn text_watermark_with_a_degenerate_box_uses_the_unsized_watermark() {
    docx_layout::clear_measure_fonts();
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let unsized_run = painted_watermark(
        &watermarked_document(["auto", "auto", "Calibri", "48pt", "Sample copy"]),
        font,
    );
    assert!(unsized_run["horizontalScale"].is_null(), "{unsized_run}");
    for [width, height] in [
        ["672pt", ".0001pt"],
        ["0.0001pt", "58.5pt"],
        ["1000000000000px", "58.5pt"],
    ] {
        let run = painted_watermark(
            &watermarked_document([width, height, "Calibri", "48pt", "Sample copy"]),
            font,
        );
        assert_eq!(run, unsized_run, "{width} x {height}");
    }
}

#[test]
fn text_watermark_with_a_positioned_mark_keeps_the_text_run() {
    stretched_text_run("DRAFT\u{0301}");
}
