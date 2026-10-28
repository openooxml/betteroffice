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
    r#"xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" "#,
    r#"xmlns:v="urn:schemas-microsoft-com:vml" "#,
    r#"xmlns:o="urn:schemas-microsoft-com:office:office""#
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

/// The same chart in a floating (`wp:anchor`, wrapNone) drawing.
fn anchored_chart_drawing() -> String {
    anchored_chart_drawing_with("<wp:wrapNone/>", EXTENT, 0)
}

fn anchored_chart_drawing_with(wrap: &str, extent: &str, vertical_offset: i64) -> String {
    format!(
        r#"<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>{vertical_offset}</wp:posOffset></wp:positionV><wp:extent {extent}/><wp:effectExtent l="0" t="0" r="9525" b="18415"/>{wrap}<wp:docPr id="22" name="Chart 22"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:anchor></w:drawing>"#
    )
}

fn chart_drawing() -> String {
    format!(
        r#"<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent {EXTENT}/><wp:effectExtent l="0" t="0" r="9525" b="18415"/><wp:docPr id="22" name="Chart 22"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rIdChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
}

fn text_box_drawing() -> String {
    text_box_drawing_with(
        EXTENT,
        &format!(
            "<w:p><w:r><w:t>{BOX_TITLE}</w:t></w:r></w:p><w:p><w:r><w:t>{BOX_BODY}</w:t></w:r></w:p>"
        ),
    )
}

/// An inline text box with extent attribute string `extent` (for example
/// `cx="5943600" cy="3810000"`) holding `content` as its txbxContent.
fn text_box_drawing_with(extent: &str, content: &str) -> String {
    text_box_drawing_fb(extent, content, "<w:t>Fallback</w:t>")
}

/// The same text box with `fallback` as its AlternateContent fallback.
fn text_box_drawing_fb(extent: &str, content: &str, fallback: &str) -> String {
    format!(
        r#"<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent {extent}/><wp:effectExtent l="0" t="0" r="28575" b="18415"/><wp:docPr id="4" name="Text Box 4"/><wp:cNvGraphicFramePr><a:graphicFrameLocks/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:cNvSpPr txBox="1"><a:spLocks noChangeArrowheads="1"/></wps:cNvSpPr><wps:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="5648400" cy="3049200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln></wps:spPr><wps:txbx><w:txbxContent>{content}</w:txbxContent></wps:txbx><wps:bodyPr rot="0" vert="horz" wrap="square" lIns="91440" tIns="45720" rIns="91440" bIns="45720" anchor="t" anchorCtr="0"><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback>{fallback}</mc:Fallback></mc:AlternateContent>"#
    )
}

/// A VML fallback text box of `width` by `height` points holding `text`.
fn vml_fallback(width: &str, height: &str, text: &str) -> String {
    format!(
        r##"<w:pict><v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype><v:shape id="Text Box 10" type="#_x0000_t202" style="width:{width}pt;height:{height}pt;visibility:visible;mso-wrap-style:square;mso-position-horizontal:absolute;mso-position-horizontal-relative:char;mso-position-vertical:absolute;mso-position-vertical-relative:line;v-text-anchor:top"><v:textbox><w:txbxContent><w:p><w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="36"/></w:rPr><w:t>{text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict>"##
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
                "measurement": {"fontChains": {"calibri|0|0": [font], "calibri|1|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
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

/// A full-width text box after paragraph text wraps to a line of its
/// own in Word, where the block layout is what Word draws; the box keeps
/// the block lowering at single and at multiple line spacing, so the
/// spacing never multiplies the box height.
fn text_then_full_width_box(spacing: &str, client_id: u64) -> (Value, Value) {
    let content = format!(
        "{}{}",
        text_run("Please provide examples and any information which you think would be useful"),
        run(&text_box_drawing_with(
            r#"cx="5943600" cy="3810000""#,
            "<w:p><w:r><w:t>Answer space</w:t></w:r></w:p>",
        )),
    );
    let body = format!(r#"<w:p w14:paraId="20000001"><w:pPr>{spacing}</w:pPr>{content}</w:p>"#);
    // An 816 px wide A4 page: 624 px of content, 931 px tall.
    render(&document(&body, 12240, 16845, 1440), client_id)
}

fn assert_full_width_box_stays_a_block(spacing: &str, max_text_height: f64, client_id: u64) {
    let (layout, display) = text_then_full_width_box(spacing, client_id);
    assert_eq!(page_count(&layout), 1, "one page: {layout}");
    let shapes: Vec<_> = layout["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap().iter())
        .filter(|fragment| fragment["kind"] == "shape")
        .collect();
    assert_eq!(shapes.len(), 1, "one shape block, not a carrier: {layout}");
    near(num(shapes[0], "width"), 624.0, 2.0, "box width");
    near(num(shapes[0], "height"), 400.0, 2.0, "box height unscaled");
    let text =
        paragraph_fragment(&layout, "20000001").expect("the text stays its own paragraph fragment");
    assert!(
        num(text, "height") < max_text_height,
        "text fragment holds only its lines, got height {}: {layout}",
        num(text, "height")
    );
    assert!(
        num(shapes[0], "y") >= num(text, "y") + num(text, "height") - 1.0,
        "box below the text: {layout}"
    );
    let _ = display;
}

#[test]
fn text_then_full_width_text_box_keeps_the_block_layout() {
    assert_full_width_box_stays_a_block("", 60.0, 78207);
}

#[test]
fn text_then_full_width_text_box_ignores_multiple_line_spacing() {
    assert_full_width_box_stays_a_block(
        r#"<w:spacing w:line="480" w:lineRule="auto"/>"#,
        120.0,
        78210,
    );
}

#[test]
fn table_text_paints_inside_a_shared_line_text_box() {
    let table = concat!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="dxa"/></w:tblPr>"#,
        r#"<w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid><w:tr><w:tc>"#,
        r#"<w:tcPr><w:tcW w:w="5000" w:type="dxa"/></w:tcPr>"#,
        r#"<w:p><w:r><w:t>Cell text in box</w:t></w:r></w:p>"#,
        r#"</w:tc></w:tr></w:tbl>"#
    );
    let body = paragraph(
        "20000002",
        &format!(
            "{}{}",
            run(&chart_drawing()),
            run(&text_box_drawing_with(EXTENT, table)),
        ),
    );
    let (_, display) = render(&document(&body, 20160, 12240, 1080), 78208);
    let boxed = box_shape(&display);
    let (text_x, baseline) = text_origin(&display, "Cell text in box");
    assert!(
        text_x >= num(&boxed, "x") && text_x < num(&boxed, "x") + BOX_W,
        "cell text inside the box horizontally, got x {text_x}: {display}"
    );
    assert!(
        baseline > num(&boxed, "y") && baseline <= num(&boxed, "y") + BOX_H,
        "cell text inside the box vertically, got baseline {baseline}: {display}"
    );
}

#[test]
fn inline_chart_atoms_have_source_spans() {
    use docx_edit::bridge::{RenderEnv, yrs_doc_to_mapped_layout_blocks};

    let body = paragraph(
        "20000003",
        &format!(
            "{}{}{}",
            run(&chart_drawing()),
            run(&chart_drawing()),
            text_run("After")
        ),
    );
    let bytes = document(&body, 20160, 12240, 1080);
    let engine = EngineSession::new(78209);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let (blocks, map) =
        yrs_doc_to_mapped_layout_blocks(engine.doc(), "body", &RenderEnv::default()).unwrap();
    let paragraphs: Vec<_> = blocks
        .iter()
        .filter(|block| matches!(block, docx_layout::types::LayoutBlock::Paragraph(_)))
        .collect();
    assert_eq!(paragraphs.len(), 1, "one paragraph holds both charts");
    let atoms: Vec<_> = map
        .spans
        .iter()
        .filter(|span| span.atom && span.paragraph == 0)
        .map(|span| (span.pm_start, span.pm_end, span.raw_start, span.raw_end))
        .collect();
    assert_eq!(atoms, [(1, 2, 0, 1), (2, 3, 1, 2)], "chart atom spans");
    // A caret between the charts and one right after the second chart
    // map to the story positions after the first and second atoms.
    assert_eq!(map.span_at(2).map(|span| span.raw_start), Some(1));
    assert_eq!(map.span_at(3).map(|span| span.raw_start), Some(2));
}

fn chart_block_fragments(layout: &Value) -> Vec<&Value> {
    layout["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap().iter())
        .filter(|fragment| fragment["kind"] == "chart")
        .collect()
}

/// A narrow label box beside a wide box (144 and 466 px in 624 px of
/// content): the label text wraps inside its own box and never reaches
/// into the wide box. The label's AlternateContent fallback is a VML
/// text box of the same size, the corpus form.
#[test]
fn narrow_label_box_text_wraps_inside_its_box() {
    let label = text_box_drawing_fb(
        r#"cx="1371600" cy="1760855""#,
        r#"<w:p><w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="36"/></w:rPr><w:t>Guiding Principles</w:t></w:r></w:p>"#,
        &vml_fallback("108", "138.65", "Guiding Principles"),
    );
    let wide = text_box_drawing_with(
        r#"cx="4438650" cy="1760855""#,
        "<w:p><w:r><w:t>Information and data are managed as assets</w:t></w:r></w:p>",
    );
    let body = paragraph("20000006", &format!("{}{}", run(&label), run(&wide)));
    let (_, display) = render(&document(&body, 12240, 16845, 1440), 78213);
    let (guiding_x, guiding_baseline) = text_origin(&display, "Guiding");
    let (principles_x, principles_baseline) = text_origin(&display, "Principles");
    near(guiding_x, 105.6, 3.0, "label first line at the box inset");
    near(
        principles_x,
        guiding_x,
        3.0,
        "label second line back at the box inset",
    );
    assert!(
        principles_baseline > guiding_baseline + 10.0,
        "label wraps to a second line: baselines {guiding_baseline} and {principles_baseline}"
    );
}

#[test]
fn a_converted_label_box_paints_one_representation() {
    let label = text_box_drawing_fb(
        r#"cx="1371600" cy="1760855""#,
        r#"<w:p><w:r><w:rPr><w:b/><w:sz w:val="32"/><w:szCs w:val="36"/></w:rPr><w:t>Guiding Principles</w:t></w:r></w:p>"#,
        &vml_fallback("108", "138.65", "Guiding Principles"),
    );
    let wide = text_box_drawing_with(
        r#"cx="4438650" cy="1760855""#,
        "<w:p><w:r><w:t>Information and data are managed as assets</w:t></w:r></w:p>",
    );
    let body = paragraph("20000007", &format!("{}{}", run(&label), run(&wide)));
    let (_, display) = render(&document(&body, 12240, 16845, 1440), 78218);
    let painted: String = primitives(&display)
        .iter()
        .filter_map(|primitive| primitive["text"].as_str())
        .collect();
    assert_eq!(
        painted.matches("Guiding").count(),
        1,
        "one label representation paints: {display}"
    );
    let (_, guiding_baseline) = text_origin(&display, "Guiding");
    let (_, principles_baseline) = text_origin(&display, "Principles");
    assert!(
        principles_baseline > guiding_baseline + 10.0,
        "label wraps to a second line: baselines {guiding_baseline} and {principles_baseline}"
    );
}

fn assert_chart_uses_main_block_layout(drawing: &str, client_id: u64) {
    let body = paragraph(
        "90000001",
        &format!(
            "{}{}{}",
            text_run("Before the chart"),
            run(drawing),
            text_run("After the chart"),
        ),
    );
    let (layout, display) = render(&document(&body, 20160, 12240, 1080), client_id);
    let charts = chart_block_fragments(&layout);
    assert_eq!(charts.len(), 1, "one chart block: {layout}");
    let chart = charts[0];
    assert!(chart["isAnchored"].is_null(), "block placement: {layout}");
    assert!(chart["zIndex"].is_null(), "block paint order: {layout}");
    let blocks = layout["measured"].as_array().unwrap();
    let chart_index = blocks
        .iter()
        .position(|entry| entry["block"]["kind"] == "chart")
        .unwrap();
    assert_eq!(blocks[chart_index - 1]["block"]["kind"], "paragraph");
    assert_eq!(blocks[chart_index + 1]["block"]["kind"], "paragraph");
    let block = &blocks[chart_index]["block"];
    for field in [
        "position",
        "wrapType",
        "wrapText",
        "relativeHeight",
        "behindDoc",
    ] {
        assert!(block[field].is_null(), "main chart field {field}: {layout}");
    }
    near(num(chart, "x"), 72.0, 0.01, "block left margin");
    let (_, before) = text_origin(&display, "Before the chart");
    let (after_x, after) = text_origin(&display, "After the chart");
    assert!(before < num(chart, "y"), "source block order: {display}");
    assert!(after >= num(chart, "y") + num(chart, "height"));
    near(after_x, 72.0, 0.01, "text after the block");
}

#[test]
fn front_wrap_none_anchored_chart_keeps_main_block_layout() {
    assert_chart_uses_main_block_layout(&anchored_chart_drawing(), 78211);
}

#[test]
fn anchored_chart_pairs_stack_as_main_blocks() {
    for (wrap, behind) in [
        ("<wp:wrapNone/>", false),
        ("<wp:wrapNone/>", true),
        ("<wp:wrapTopAndBottom/>", false),
        (r#"<wp:wrapSquare wrapText="bothSides"/>"#, false),
    ] {
        let mut drawing = anchored_chart_drawing_with(wrap, EXTENT, 0);
        if behind {
            drawing = drawing.replace("behindDoc=\"0\"", "behindDoc=\"1\"");
        }
        let body = paragraph("20000004", &format!("{}{}", run(&drawing), run(&drawing)));
        let (layout, _) = render(&document(&body, 20160, 12240, 1080), 78212);
        assert_eq!(
            page_count(&layout),
            1,
            "{wrap}, behindDoc={behind}: {layout}"
        );
        let charts = chart_block_fragments(&layout);
        assert_eq!(charts.len(), 2, "{wrap}, behindDoc={behind}: {layout}");
        for chart in &charts {
            assert!(chart["isAnchored"].is_null(), "{layout}");
            assert!(chart["zIndex"].is_null(), "{layout}");
        }
        near(
            num(charts[1], "y"),
            num(charts[0], "y") + num(charts[0], "height"),
            0.01,
            "main block stacking",
        );
        let blocks = layout["measured"].as_array().unwrap();
        assert_eq!(
            blocks
                .iter()
                .filter(|entry| entry["block"]["kind"] == "chart")
                .count(),
            2
        );
        assert!(blocks.iter().all(|entry| {
            entry["block"]["runs"]
                .as_array()
                .is_none_or(|runs| runs.iter().all(|run| run["inlineChart"].is_null()))
        }));
    }
}

#[test]
fn top_and_bottom_anchored_chart_keeps_main_block_layout() {
    for relative in ["paragraph", "page", "margin"] {
        let drawing = anchored_chart_drawing_with("<wp:wrapTopAndBottom/>", EXTENT, 914400)
            .replace(
                "relativeFrom=\"paragraph\"",
                &format!("relativeFrom=\"{relative}\""),
            );
        assert_chart_uses_main_block_layout(&drawing, 78216);
    }
}

#[test]
fn square_anchored_chart_keeps_main_block_layout() {
    let drawing =
        anchored_chart_drawing_with(r#"<wp:wrapSquare wrapText="bothSides"/>"#, EXTENT, 914400);
    assert_chart_uses_main_block_layout(&drawing, 78217);
}

#[test]
fn tight_anchored_chart_keeps_main_block_layout() {
    let drawing =
        anchored_chart_drawing_with(r#"<wp:wrapTight wrapText="bothSides"/>"#, EXTENT, 914400);
    assert_chart_uses_main_block_layout(&drawing, 78232);
}

#[test]
fn through_anchored_chart_keeps_main_block_layout() {
    let drawing =
        anchored_chart_drawing_with(r#"<wp:wrapThrough wrapText="bothSides"/>"#, EXTENT, 914400);
    assert_chart_uses_main_block_layout(&drawing, 78233);
}

#[test]
fn behind_doc_anchored_chart_keeps_main_block_layout() {
    let drawing = anchored_chart_drawing_with("<wp:wrapNone/>", EXTENT, -914400)
        .replace("behindDoc=\"0\"", "behindDoc=\"1\"");
    assert_chart_uses_main_block_layout(&drawing, 78234);
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

fn document_with_picture(body: &str) -> Vec<u8> {
    let mut parts = ooxml_opc::unzip_parts(&document(body, 20160, 12240, 1080)).unwrap();
    for (name, bytes) in &mut parts {
        let (end, entry) = match name.as_str() {
            "word/_rels/document.xml.rels" => (
                "</Relationships>",
                r#"<Relationship Id="rIdPicture" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/picture.png"/>"#,
            ),
            "[Content_Types].xml" => (
                "</Types>",
                r#"<Default Extension="png" ContentType="image/png"/>"#,
            ),
            _ => continue,
        };
        let xml = String::from_utf8(std::mem::take(bytes)).unwrap();
        *bytes = xml.replace(end, &format!("{entry}{end}")).into_bytes();
    }
    parts.push((
        "word/media/picture.png".to_owned(),
        vec![
            137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1,
            8, 6, 0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 11, 73, 68, 65, 84, 120, 156, 99, 96, 0, 2,
            0, 0, 5, 0, 1, 169, 245, 232, 39, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
        ],
    ));
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn ordinary_picture() -> &'static str {
    concat!(
        r#"<w:pict><v:shape id="ordinary-picture" style="width:444.75pt;height:240pt">"#,
        r#"<v:imagedata r:id="rIdPicture"/></v:shape></w:pict>"#
    )
}

fn text_box_picture_twin(extent: &str) -> String {
    format!(
        r#"<w:drawing><wp:inline><wp:extent {extent}/><wp:docPr id="4" name="Text Box 4"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:blipFill><a:blip r:embed="rIdPicture"/></pic:blipFill><pic:spPr><a:prstGeom prst="rect"/></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    )
}

fn text_box_with_opaque_picture_fallback(fallback: &str) -> String {
    let unsupported = concat!(
        r#"<w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>"#,
        r#"<wp:docPr id="99" name="Unsupported chart"/><a:graphic><a:graphicData "#,
        r#"uri="http://schemas.microsoft.com/office/drawing/2014/chartex">"#,
        r#"<cx:chart xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" "#,
        r#"r:id="rIdUnknown"/></a:graphicData></a:graphic></wp:inline></w:drawing>"#
    );
    text_box_drawing_fb(
        EXTENT,
        "<w:p><w:r><w:t>Selected Choice text</w:t></w:r></w:p>",
        fallback,
    )
    .replace("</mc:Choice>", &format!("{unsupported}</mc:Choice>"))
}

#[test]
fn ordinary_picture_before_matching_inline_chart_survives() {
    let body = paragraph(
        "40000001",
        &format!("{}{}", run(ordinary_picture()), run(&chart_drawing())),
    );
    let (layout, display) = render(&document_with_picture(&body), 78219);
    let pictures: Vec<_> = primitives(&display)
        .iter()
        .filter(|primitive| primitive["kind"] == "image")
        .collect();
    assert_eq!(pictures.len(), 1, "ordinary picture survives: {display}");
    let chart = chart_canvas(&display);
    near(
        num(&chart, "x"),
        num(pictures[0], "x") + BOX_W,
        2.0,
        "chart follows the ordinary picture",
    );
    assert_eq!(page_count(&layout), 1);
}

#[test]
fn converted_choice_keeps_a_differently_sized_composite_preview() {
    let fallback = text_box_picture_twin(r#"cx="1270000" cy="635000""#);
    let text_box = text_box_with_opaque_picture_fallback(&fallback);
    let body = paragraph(
        "40000006",
        &format!("{}{}", run(ordinary_picture()), run(&text_box)),
    );
    let (_, display) = render(&document_with_picture(&body), 78226);
    assert_eq!(
        primitives(&display)
            .iter()
            .filter(|primitive| primitive["kind"] == "image")
            .count(),
        2,
        "the ordinary picture and composite preview paint: {display}"
    );
    let _ = box_shape(&display);
    let _ = text_origin(&display, "Selected Choice text");
}

#[test]
fn matching_doc_pr_keeps_a_possibly_composite_preview() {
    let text_box = text_box_with_opaque_picture_fallback(&text_box_picture_twin(EXTENT));
    let body = paragraph(
        "40000007",
        &format!("{}{}", run(&text_box), text_run("After the Choice")),
    );
    let bytes = document_with_picture(&body);
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let package = serde_json::to_value(package).unwrap();
    let mut pending = vec![&package];
    let mut sources = Vec::new();
    while let Some(value) = pending.pop() {
        if value["type"] == "opaqueDrawing" && value["kind"] == "alternateContent" {
            sources.push(value["xml"].as_str().unwrap());
        }
        match value {
            Value::Array(values) => pending.extend(values),
            Value::Object(values) => pending.extend(values.values()),
            _ => {}
        }
    }
    assert_eq!(sources.len(), 1);
    assert!(sources[0].contains(r#"<wp:docPr id="4" name="Text Box 4"/>"#));
    let (_, display) = render(&bytes, 78227);
    assert_eq!(
        primitives(&display)
            .iter()
            .filter(|primitive| primitive["kind"] == "image")
            .count(),
        1,
        "matching drawing properties do not prove the preview content: {display}"
    );
    let _ = box_shape(&display);
    let _ = text_origin(&display, "Selected Choice text");
    let _ = text_origin(&display, "After the Choice");
}

#[test]
fn converting_a_text_box_keeps_an_unsupported_charts_grouped_fallback() {
    let fallback = concat!(
        r#"<w:pict><v:group coordsize="1000,1000" style="width:100pt;height:50pt">"#,
        r#"<v:shape id="box-preview" style="position:absolute;left:0;top:0;width:500;height:1000">"#,
        r#"<v:imagedata r:id="rIdPicture"/></v:shape>"#,
        r#"<v:shape id="chart-preview" style="position:absolute;left:500;top:0;width:500;height:1000">"#,
        r#"<v:imagedata r:id="rIdPicture"/></v:shape></v:group></w:pict>"#
    );
    let drawing = text_box_with_opaque_picture_fallback(fallback);
    let body = paragraph(
        "60000001",
        &format!(
            "{}{}",
            run(&drawing),
            text_run("After the grouped fallback")
        ),
    );
    let (layout, display) = render(&document_with_picture(&body), 78228);
    assert!(
        layout["measured"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|entry| entry["block"]["runs"].as_array().into_iter().flatten())
            .any(|run| run["inlineShape"].is_object()),
        "the supported text box converts: {layout}"
    );
    assert_eq!(
        primitives(&display)
            .iter()
            .filter(|primitive| primitive["kind"] == "image")
            .count(),
        2,
        "the grouped fallback keeps the unsupported chart visible: {display}"
    );
    let _ = text_origin(&display, "Selected Choice text");
}

#[test]
fn converting_a_text_box_keeps_an_unsupported_charts_composite_preview() {
    for fallback in [
        ordinary_picture().to_owned(),
        text_box_picture_twin(EXTENT),
        text_box_picture_twin(EXTENT).replace(r#"id="4""#, r#"id="99""#),
        text_box_picture_twin(EXTENT).replace("Text Box 4", "Composite preview"),
    ] {
        let drawing = text_box_with_opaque_picture_fallback(&fallback);
        let body = paragraph(
            "80000001",
            &format!(
                "{}{}",
                run(&drawing),
                text_run("After the composite preview")
            ),
        );
        let (layout, display) = render(&document_with_picture(&body), 78230);
        assert!(
            layout["measured"]
                .as_array()
                .unwrap()
                .iter()
                .flat_map(|entry| entry["block"]["runs"].as_array().into_iter().flatten())
                .any(|run| run["inlineShape"].is_object()),
            "the supported text box converts: {layout}"
        );
        assert_eq!(
            primitives(&display)
                .iter()
                .filter(|primitive| primitive["kind"] == "image")
                .count(),
            1,
            "the composite preview keeps the unsupported chart visible: {display}"
        );
        let _ = text_origin(&display, "Selected Choice text");
    }
}

#[test]
fn unrelated_vml_box_before_matching_inline_box_survives() {
    let unrelated = vml_fallback("444.75", "240", "Independent VML box");
    let body = paragraph(
        "40000002",
        &format!(
            "{}{}{}",
            run(&unrelated),
            run(&text_box_drawing()),
            run(&chart_drawing()),
        ),
    );
    let (_, display) = render(&document(&body, 20160, 12240, 1080), 78220);
    let _ = text_origin(&display, "Independent VML box");
    let _ = text_origin(&display, BOX_TITLE);
}

#[test]
fn shared_text_box_keeps_text_after_section_width_remeasurement() {
    let bytes = pair_document(20160);
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(78221);
    seed_from_docx(engine.doc(), &bytes).unwrap();
    let mut request = json!({
        "bodyStory": "body",
        "renderEnv": {},
        "regions": {"sections": [{
            "pageSize": {"w": 1344, "h": 816},
            "margins": {"top": 72, "right": 72, "bottom": 72, "left": 72}
        }]},
        "measurement": {
            "fontChains": {"calibri|0|0": [font], "calibri|1|0": [font]},
            "defaults": {"fontFamily": "Calibri", "fontSize": 11}
        }
    });
    for width in [1344, 1280, 1344] {
        request["regions"]["sections"][0]["pageSize"]["w"] = json!(width);
        let output = engine
            .layout_document_with_regions_json(&request.to_string())
            .unwrap();
        let layout: Value = serde_json::from_str(&output).unwrap();
        let shape = layout["measured"]
            .as_array()
            .unwrap()
            .iter()
            .flat_map(|entry| entry["block"]["runs"].as_array().into_iter().flatten())
            .find_map(|run| run.get("inlineShape"))
            .expect("inline text box");
        assert!(shape["innerMeasures"].is_array(), "width {width}: {layout}");
        let display: Value =
            serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
        let _ = text_origin(&display, BOX_TITLE);
        let _ = text_origin(&display, BOX_BODY);
    }
}

#[test]
fn shared_text_box_requests_its_unique_font_and_script() {
    let text_box = text_box_drawing_with(
        EXTENT,
        concat!(
            r#"<w:p><w:r><w:rPr><w:rFonts w:ascii="Box Font" w:hAnsi="Box Font" "#,
            r#"w:eastAsia="Box Font"/><w:b/><w:i/></w:rPr><w:t>日本語かな</w:t></w:r></w:p>"#
        ),
    );
    let body = paragraph(
        "40000003",
        &format!("{}{}", run(&chart_drawing()), run(&text_box)),
    );
    let engine = EngineSession::new(78222);
    seed_from_docx(engine.doc(), &document(&body, 20160, 12240, 1080)).unwrap();
    let requirements: Vec<Value> = serde_json::from_str(
        &engine
            .layout_font_requirements_json(&json!({"bodyStory": "body"}).to_string())
            .unwrap(),
    )
    .unwrap();
    assert!(
        requirements.iter().any(|requirement| {
            requirement["key"] == "box font|1|1" && requirement["scripts"] == json!(["cjk-jp"])
        }),
        "text box font and script: {requirements:?}"
    );
}

#[test]
fn page_break_before_keeps_anchored_charts_in_main_block_order() {
    for (wrap, behind) in [
        ("<wp:wrapNone/>", false),
        ("<wp:wrapNone/>", true),
        ("<wp:wrapTopAndBottom/>", false),
        (r#"<wp:wrapSquare wrapText="bothSides"/>"#, false),
        (r#"<wp:wrapTight wrapText="bothSides"/>"#, false),
        (r#"<wp:wrapThrough wrapText="bothSides"/>"#, false),
    ] {
        let mut drawing = anchored_chart_drawing_with(wrap, r#"cx="2857500" cy="1905000""#, 0);
        if behind {
            drawing = drawing.replace("behindDoc=\"0\"", "behindDoc=\"1\"");
        }
        let body = format!(
            r#"{}<w:p w14:paraId="80000003"><w:pPr><w:pageBreakBefore/></w:pPr>{}{}</w:p>"#,
            paragraph("80000002", &text_run("Previous page")),
            text_run("Wrapped owner text"),
            run(&drawing),
        );
        let (layout, display) = render(&document(&body, 20160, 12240, 1080), 78231);
        let pages = layout["layout"]["pages"].as_array().unwrap();
        assert_eq!(pages.len(), 2, "{wrap}: {layout}");
        let chart = pages[1]["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .find(|fragment| fragment["kind"] == "chart")
            .expect("chart on the owner's page");
        let page = json!({"pages": [display["pages"][1].clone()]});
        let (x, baseline) = text_origin(&page, "Wrapped owner text");
        assert!(chart["isAnchored"].is_null(), "{wrap}: {layout}");
        assert!(
            baseline < num(chart, "y"),
            "text precedes the chart block: {display}"
        );
        near(x, 72.0, 0.01, "full-width text before the block");
    }
}
