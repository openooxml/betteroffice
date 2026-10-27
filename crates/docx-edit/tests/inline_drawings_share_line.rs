//! Inline charts and inline text boxes that share a paragraph line with
//! another inline drawing or with visible text flow on that line, as Word
//! draws inline pictures. A paragraph holding only one of them keeps the
//! block lowering. The documents are built here; no corpus files.

use docx_edit::{EngineSession, seed_from_docx};
use serde_json::{Value, json};

const FONT: &[u8] = include_bytes!("../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");

const NAMESPACES: &str = concat!(
    r#"xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" "#,
    r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" "#,
    r#"xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" "#,
    r#"xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" "#,
    r#"xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
    r#"xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" "#,
    r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006""#
);

/// 5648400 x 3049200 EMU: 593 x 320 px at 96 dpi, the corpus pair's extent.
const EXTENT: &str = r#"cx="5648400" cy="3049200""#;
const BOX_W: f64 = 593.0;
const BOX_H: f64 = 320.0;
const BOX_TITLE: &str = "Inline box title";
const BOX_BODY: &str = "Inline box body text.";

const CHART: &str = concat!(
    r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" "#,
    r#"xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart>"#,
    r#"<c:title><c:tx><c:rich><a:p><a:r><a:t>Sales</a:t></a:r></a:p></c:rich></c:tx></c:title>"#,
    r#"<c:plotArea><c:barChart><c:barDir val="col"/><c:grouping val="clustered"/>"#,
    r#"<c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>North</c:v></c:tx>"#,
    r#"<c:cat><c:strLit><c:ptCount val="2"/><c:pt idx="0"><c:v>Q1</c:v></c:pt>"#,
    r#"<c:pt idx="1"><c:v>Q2</c:v></c:pt></c:strLit></c:cat>"#,
    r#"<c:val><c:numLit><c:ptCount val="2"/><c:pt idx="0"><c:v>2</c:v></c:pt>"#,
    r#"<c:pt idx="1"><c:v>4</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart>"#,
    r#"</c:plotArea></c:chart></c:chartSpace>"#
);

fn chart_drawing() -> String {
    format!(
        r#"<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent {EXTENT}/><wp:effectExtent l="0" t="0" r="9525" b="18415"/><wp:docPr id="22" name="Chart 22"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
}

fn text_box_drawing() -> String {
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent {EXTENT}/><wp:effectExtent l="0" t="0" r="28575" b="18415"/><wp:docPr id="4" name="Text Box 4"/><wp:cNvGraphicFramePr><a:graphicFrameLocks/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr><wps:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="5648400" cy="3049200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent><w:p><w:r><w:t>{BOX_TITLE}</w:t></w:r></w:p><w:p><w:r><w:t>{BOX_BODY}</w:t></w:r></w:p></w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:t>Fallback</w:t></mc:Fallback></mc:AlternateContent>"#
    )
}

fn paragraph(id: &str, content: &str) -> String {
    format!(r#"<w:p w14:paraId="{id}">{content}</w:p>"#)
}

fn run(content: &str) -> String {
    format!("<w:r>{content}</w:r>")
}

fn text_run(text: &str) -> String {
    format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#)
}

/// A one-section package holding `body` on a page of `page_w` x `page_h`
/// twips with `margin` twips on every side.
fn document(body: &str, page_w: u32, page_h: u32, margin: u32) -> Vec<u8> {
    let parts = [
        (
            "[Content_Types].xml",
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#
                .to_owned(),
        ),
        (
            "_rels/.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#
                .to_owned(),
        ),
        (
            "word/_rels/document.xml.rels",
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>"#
                .to_owned(),
        ),
        ("word/charts/chart1.xml", CHART.to_owned()),
        (
            "word/document.xml",
            format!(
                r#"<w:document {NAMESPACES}><w:body>{body}<w:sectPr><w:pgSz w:w="{page_w}" w:h="{page_h}"/><w:pgMar w:top="{margin}" w:right="{margin}" w:bottom="{margin}" w:left="{margin}" w:header="360" w:footer="360" w:gutter="0"/></w:sectPr></w:body></w:document>"#
            ),
        ),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

/// The chart and text box pair of test (a): a 1344 x 816 px landscape page
/// (20160 x 12240 twips, 1080-twip margins, 1200 px of content width).
fn pair_document(page_w: u32) -> Vec<u8> {
    let body = paragraph(
        "5F401DA4",
        &format!("{}{}", run(&chart_drawing()), run(&text_box_drawing())),
    );
    document(&body, page_w, 12240, 1080)
}

fn render(bytes: &[u8], client_id: u64) -> (Value, Value) {
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let package = docx_parse::parse_docx_s9_wire(bytes, Default::default())
        .unwrap()
        .document
        .package;
    let engine = EngineSession::new(client_id);
    seed_from_docx(engine.doc(), bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body",
                "renderEnv": {},
                "regions": {"sections": [{"properties": package.document.final_section_properties}]},
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    let layout: Value = serde_json::from_str(&output).unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    (layout, display)
}

fn page_count(layout: &Value) -> usize {
    layout["layout"]["pages"].as_array().unwrap().len()
}

fn primitives(display: &Value) -> &Vec<Value> {
    display["pages"][0]["primitives"].as_array().unwrap()
}

fn num(value: &Value, key: &str) -> f64 {
    value[key].as_f64().unwrap()
}

/// The chart's painted canvas: the one rect carrying chart data at box size.
fn chart_canvas(display: &Value) -> Value {
    primitives(display)
        .iter()
        .find(|primitive| {
            primitive["kind"] == "rect"
                && primitive["chart"].is_object()
                && (num(primitive, "w") - BOX_W).abs() < 2.0
                && (num(primitive, "h") - BOX_H).abs() < 2.0
        })
        .cloned()
        .unwrap_or_else(|| panic!("no chart canvas in {display}"))
}

/// The text box's painted frame.
fn box_shape(display: &Value) -> Value {
    primitives(display)
        .iter()
        .find(|primitive| {
            primitive["kind"] == "shape"
                && (num(primitive, "w") - BOX_W).abs() < 2.0
                && (num(primitive, "h") - BOX_H).abs() < 2.0
        })
        .cloned()
        .unwrap_or_else(|| panic!("no text box shape in {display}"))
}

fn paragraph_fragment<'a>(layout: &'a Value, para_id: &str) -> Option<&'a Value> {
    layout["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap().iter())
        .find(|fragment| fragment["blockId"] == para_id)
}

/// (x, baseline y) of the first painted run whose text contains `needle`.
/// Painted text arrives as one primitive per character (a `text`
/// primitive, or a glyph run when fonts shape it), so the needle is
/// located in the concatenation of all text primitives in paint order.
fn text_origin(display: &Value, needle: &str) -> (f64, f64) {
    let mut joined = String::new();
    let mut entries: Vec<(usize, f64, f64)> = Vec::new();
    for primitive in primitives(display) {
        let Some(text) = primitive["text"].as_str() else {
            continue;
        };
        if text.is_empty() {
            continue;
        }
        let (x, baseline) =
            if let Some(glyph) = primitive["glyphs"].as_array().and_then(|g| g.first()) {
                (num(glyph, "x"), num(glyph, "y"))
            } else if primitive["x"].is_number() {
                (num(primitive, "x"), num(primitive, "baselineY"))
            } else {
                continue;
            };
        entries.push((joined.len(), x, baseline));
        joined.push_str(text);
    }
    let start = joined
        .find(needle)
        .unwrap_or_else(|| panic!("no painted text {needle:?} in {display}"));
    let (_, x, baseline) = entries
        .iter()
        .rev()
        .find(|(offset, _, _)| *offset <= start)
        .unwrap();
    (*x, *baseline)
}

fn near(actual: f64, expected: f64, tolerance: f64, what: &str) {
    assert!(
        (actual - expected).abs() <= tolerance,
        "{what}: {actual} vs {expected} (tolerance {tolerance})"
    );
}

#[test]
fn chart_and_text_box_share_one_line_on_a_wide_page() {
    let (layout, display) = render(&pair_document(20160), 78201);
    assert_eq!(page_count(&layout), 1, "the pair fits one page: {layout}");
    let chart = chart_canvas(&display);
    let boxed = box_shape(&display);
    near(num(&chart, "x"), 72.0, 1.5, "chart at the left margin");
    near(
        num(&chart, "y"),
        num(&boxed, "y"),
        1.5,
        "same top on one line",
    );
    // Word adds the first object's effectExtent (1 px here) between the
    // boxes; abutting placement is within 2 px of Word's box x.
    near(
        num(&boxed, "x"),
        num(&chart, "x") + BOX_W,
        2.0,
        "text box right of the chart",
    );
    let fragment = paragraph_fragment(&layout, "5F401DA4")
        .expect("the pair stays one paragraph, not stacked blocks");
    near(num(fragment, "height"), BOX_H, 2.0, "one line tall");
}

#[test]
fn chart_and_text_box_wrap_on_a_narrow_page() {
    // 816 px wide (12240 twips): 672 px of content fits one 593 px box.
    let (layout, display) = render(&pair_document(12240), 78202);
    assert_eq!(page_count(&layout), 1, "both lines fit one page: {layout}");
    let chart = chart_canvas(&display);
    let boxed = box_shape(&display);
    near(
        num(&boxed, "x"),
        num(&chart, "x"),
        1.5,
        "wrapped box back at the line start",
    );
    near(
        num(&boxed, "y"),
        num(&chart, "y") + BOX_H,
        2.0,
        "wrapped box one line below the chart",
    );
    let fragment = paragraph_fragment(&layout, "5F401DA4")
        .expect("the pair stays one paragraph, not stacked blocks");
    near(num(fragment, "height"), BOX_H * 2.0, 3.0, "two lines tall");
}

#[test]
fn a_single_inline_chart_keeps_the_block_layout() {
    let body = paragraph("10000001", &run(&chart_drawing()));
    let (layout, display) = render(&document(&body, 20160, 12240, 1080), 78203);
    assert_eq!(page_count(&layout), 1);
    let chart = chart_canvas(&display);
    near(num(&chart, "w"), BOX_W, 2.0, "chart width");
    near(num(&chart, "h"), BOX_H, 2.0, "chart height");
    let blocks: Vec<_> = layout["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap().iter())
        .filter(|fragment| fragment["kind"] == "chart")
        .collect();
    assert_eq!(blocks.len(), 1, "one chart block as today: {layout}");
    assert!(
        paragraph_fragment(&layout, "10000001").is_none(),
        "no paragraph fragment replaces the chart block"
    );
}

#[test]
fn a_single_inline_text_box_keeps_the_block_layout() {
    let body = paragraph("10000002", &run(&text_box_drawing()));
    let (layout, display) = render(&document(&body, 20160, 12240, 1080), 78204);
    assert_eq!(page_count(&layout), 1);
    let _ = box_shape(&display);
    let blocks: Vec<_> = layout["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap().iter())
        .filter(|fragment| fragment["kind"] == "shape")
        .collect();
    assert_eq!(blocks.len(), 1, "one shape block as today: {layout}");
}

#[test]
fn text_after_an_inline_chart_stays_on_the_chart_line() {
    let body = paragraph(
        "10000003",
        &format!("{}{}", run(&chart_drawing()), text_run("After the chart")),
    );
    let (layout, display) = render(&document(&body, 20160, 12240, 1080), 78205);
    assert_eq!(page_count(&layout), 1);
    let chart = chart_canvas(&display);
    let (text_x, baseline) = text_origin(&display, "After the chart");
    assert!(
        text_x > num(&chart, "x") + BOX_W - 5.0,
        "text starts at the chart's right edge, got x {text_x}: {display}"
    );
    assert!(
        baseline <= num(&chart, "y") + BOX_H + 5.0,
        "text baseline on the chart's line (bottom), got {baseline}: {display}"
    );
    assert!(
        paragraph_fragment(&layout, "10000003").is_some(),
        "chart and text stay one paragraph"
    );
}

#[test]
fn text_box_text_paints_inside_the_side_by_side_box() {
    let (_, display) = render(&pair_document(20160), 78206);
    let boxed = box_shape(&display);
    for needle in [BOX_TITLE, BOX_BODY] {
        let (text_x, baseline) = text_origin(&display, needle);
        assert!(
            text_x >= num(&boxed, "x") && text_x < num(&boxed, "x") + BOX_W,
            "{needle:?} inside the box horizontally, got x {text_x}: {display}"
        );
        assert!(
            baseline > num(&boxed, "y") && baseline <= num(&boxed, "y") + BOX_H,
            "{needle:?} inside the box vertically, got baseline {baseline}: {display}"
        );
    }
}
