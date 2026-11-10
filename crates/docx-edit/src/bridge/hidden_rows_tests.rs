use serde_json::{Value, json};

use crate::EngineSession;

const FONT: &[u8] = include_bytes!("../../../ooxml-text/tests/fonts/LiberationSans-Regular.ttf");
const BORDER: &str = r#"w:val="single" w:sz="24" w:space="0" w:color="000000""#;

const HIDDEN: &str = "<w:trPr><w:hidden/></w:trPr>";
const VANISH: &str = "<w:rPr><w:vanish/></w:rPr>";

fn row(text: &str, tr_pr: &str) -> String {
    row_with_run(&format!("<w:r><w:t>{text}</w:t></w:r>"), tr_pr)
}

fn hidden_row(text: &str) -> String {
    row_with_run(&format!("<w:r>{VANISH}<w:t>{text}</w:t></w:r>"), HIDDEN)
}

fn row_with_run(run: &str, tr_pr: &str) -> String {
    format!(
        r#"<w:tr>{tr_pr}<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p>{run}</w:p></w:tc></w:tr>"#
    )
}

fn package(rows: &[String]) -> Vec<u8> {
    let rows = rows.concat();
    let body = format!(
        r#"<w:p><w:r><w:t>Before</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="4680" w:type="dxa"/><w:tblBorders><w:top {BORDER}/><w:left {BORDER}/><w:bottom {BORDER}/><w:right {BORDER}/><w:insideH {BORDER}/><w:insideV {BORDER}/></w:tblBorders></w:tblPr><w:tblGrid><w:gridCol w:w="4680"/></w:tblGrid>{rows}</w:tbl><w:p><w:r><w:t>After</w:t></w:r></w:p>"#
    );
    let parts = [
        ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_owned()),
        ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="doc" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#.to_owned()),
        ("word/document.xml", format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr></w:body></w:document>"#)),
    ];
    let parts: Vec<_> = parts
        .into_iter()
        .map(|(name, value)| (name.to_owned(), value.into_bytes()))
        .collect();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

struct Rendered {
    table: Option<Value>,
    primitives: Vec<Value>,
}

fn render(rows: &[String], show_hidden_text: bool) -> Rendered {
    let bytes = package(rows);
    let package = docx_parse::parse_docx_s9_wire(&bytes, Default::default())
        .unwrap()
        .document
        .package;
    let font = docx_layout::register_measure_font(FONT).unwrap();
    let engine = EngineSession::new(75311);
    crate::seed::seed_with_layout_tokens(engine.doc(), &bytes).unwrap();
    let output = engine
        .layout_document_with_regions_json(
            &json!({
                "bodyStory": "body",
                "renderEnv": {"showHiddenText": show_hidden_text},
                "regions": {"sections": [{"properties": package.document.final_section_properties}]},
                "measurement": {"fontChains": {"calibri|0|0": [font]}, "defaults": {"fontFamily": "Calibri", "fontSize": 11}}
            })
            .to_string(),
        )
        .unwrap();
    let display: Value =
        serde_json::from_str(&engine.build_display_list_json(&output).unwrap()).unwrap();
    let output: Value = serde_json::from_str(&output).unwrap();
    let table = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|item| item["block"]["kind"] == "table")
        .cloned();
    assert_eq!(display["pages"].as_array().unwrap().len(), 1);
    Rendered {
        table,
        primitives: display["pages"][0]["primitives"]
            .as_array()
            .unwrap()
            .clone(),
    }
}

fn texts(primitives: &[Value]) -> impl Iterator<Item = &Value> {
    primitives
        .iter()
        .filter(|primitive| matches!(primitive["kind"].as_str(), Some("text" | "glyphRun")))
}

fn painted(primitives: &[Value], word: &str) -> bool {
    texts(primitives)
        .filter_map(|primitive| primitive["text"].as_str())
        .collect::<String>()
        .contains(word)
}

fn cell_primitives(primitives: &[Value]) -> usize {
    primitives
        .iter()
        .filter(|primitive| primitive.get("cell").is_some())
        .count()
}

fn last_text_y(primitives: &[Value]) -> f64 {
    let primitive = texts(primitives).last().expect("painted text");
    primitive["baselineY"]
        .as_f64()
        .or_else(|| primitive["glyphs"][0]["y"].as_f64())
        .unwrap()
}

#[test]
fn a_hidden_row_is_left_out_of_the_layout_table() {
    let rows = [row("Alpha", ""), hidden_row("Bravo"), row("Charlie", "")];
    let hidden = render(&rows, false);
    let table = hidden.table.unwrap();
    let block_rows = table["block"]["rows"].as_array().unwrap();
    assert_eq!(block_rows.len(), 2);
    assert!(block_rows.iter().all(|row| row.get("hidden").is_none()));
    assert!(!painted(&hidden.primitives, "Bravo"));

    let without = render(&[row("Alpha", ""), row("Charlie", "")], false);
    let without_table = without.table.unwrap();
    assert_eq!(
        table["measure"]["totalHeight"],
        without_table["measure"]["totalHeight"]
    );
    assert_eq!(
        cell_primitives(&hidden.primitives),
        cell_primitives(&without.primitives)
    );
    assert_eq!(
        last_text_y(&hidden.primitives),
        last_text_y(&without.primitives)
    );
}

#[test]
fn shown_hidden_text_keeps_a_hidden_row_and_marks_it() {
    let rows = [row("Alpha", ""), hidden_row("Bravo"), row("Charlie", "")];
    let shown = render(&rows, true);
    let table = shown.table.unwrap();
    let block_rows = table["block"]["rows"].as_array().unwrap();
    assert_eq!(block_rows.len(), 3);
    assert_eq!(block_rows[1]["hidden"], json!(true));
    assert!(block_rows[0].get("hidden").is_none());
    assert!(painted(&shown.primitives, "Bravo"));

    let plain = render(
        &[row("Alpha", ""), row("Bravo", ""), row("Charlie", "")],
        true,
    );
    assert_eq!(
        table["measure"]["totalHeight"],
        plain.table.unwrap()["measure"]["totalHeight"]
    );
}

#[test]
fn a_false_hidden_value_keeps_the_row() {
    let rows = [
        row("Alpha", ""),
        row("Bravo", r#"<w:trPr><w:hidden w:val="0"/></w:trPr>"#),
        row("Charlie", r#"<w:trPr><w:hidden w:val="false"/></w:trPr>"#),
    ];
    let rendered = render(&rows, false);
    let table = rendered.table.unwrap();
    let block_rows = table["block"]["rows"].as_array().unwrap();
    assert_eq!(block_rows.len(), 3);
    assert!(block_rows.iter().all(|row| row.get("hidden").is_none()));
    assert!(painted(&rendered.primitives, "Bravo"));
}

#[test]
fn a_hidden_marked_row_with_visible_text_still_lays_out_with_its_borders() {
    let rows = [row("Alpha", ""), row("Bravo", HIDDEN), row("Charlie", "")];
    let marked = render(&rows, false);
    let table = marked.table.unwrap();
    let block_rows = table["block"]["rows"].as_array().unwrap();
    assert_eq!(block_rows.len(), 3);
    assert_eq!(block_rows[1]["hidden"], json!(true));
    assert!(painted(&marked.primitives, "Bravo"));

    let plain = render(
        &[row("Alpha", ""), row("Bravo", ""), row("Charlie", "")],
        false,
    );
    assert_eq!(
        table["measure"]["totalHeight"],
        plain.table.unwrap()["measure"]["totalHeight"]
    );
    assert_eq!(
        cell_primitives(&marked.primitives),
        cell_primitives(&plain.primitives)
    );
    let borders = |primitives: &[Value]| {
        primitives
            .iter()
            .filter(|primitive| primitive["role"] == "table-border")
            .count()
    };
    assert_eq!(borders(&marked.primitives), borders(&plain.primitives));
    assert!(borders(&marked.primitives) > 0);
}

#[test]
fn an_empty_row_without_the_mark_is_kept() {
    let rows = [
        row("Alpha", ""),
        row_with_run(&format!("<w:r>{VANISH}<w:t>Bravo</w:t></w:r>"), ""),
        row_with_run("", ""),
        row("Charlie", ""),
    ];
    let rendered = render(&rows, false);
    let table = rendered.table.unwrap();
    assert_eq!(table["block"]["rows"].as_array().unwrap().len(), 4);
}

#[test]
fn a_fully_hidden_table_takes_no_space_and_paints_nothing() {
    let rows = [hidden_row("Alpha"), hidden_row("Bravo")];
    let hidden = render(&rows, false);
    assert!(hidden.table.is_none());
    assert_eq!(cell_primitives(&hidden.primitives), 0);
    let primitives = &hidden.primitives;
    assert!(painted(primitives, "Before") && painted(primitives, "After"));
    assert!(!painted(primitives, "Alpha") && !painted(primitives, "Bravo"));
    assert!(
        hidden
            .primitives
            .iter()
            .all(|primitive| primitive["role"] != "table-border")
    );

    let shown = render(&rows, true);
    assert!(last_text_y(&shown.primitives) > last_text_y(&hidden.primitives));
}

fn merged(text: &str, merge: &str, tr_pr: &str) -> String {
    let r_pr = if tr_pr == HIDDEN { VANISH } else { "" };
    format!(
        r#"<w:tr>{tr_pr}<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/><w:vMerge{merge}/></w:tcPr><w:p><w:r>{r_pr}<w:t>{text}</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr><w:p><w:r>{r_pr}<w:t>Side{text}</w:t></w:r></w:p></w:tc></w:tr>"#
    )
}

#[test]
fn a_merge_across_a_hidden_row_spans_only_the_visible_rows() {
    let through = render(
        &[
            merged("Alpha", r#" w:val="restart""#, ""),
            merged("", "", HIDDEN),
            merged("", "", ""),
        ],
        false,
    );
    let rows = through.table.unwrap()["block"]["rows"].clone();
    assert_eq!(rows.as_array().unwrap().len(), 2);
    assert_eq!(rows[0]["cells"][0]["rowSpan"], json!(2.0));
    assert_eq!(rows[1]["cells"].as_array().unwrap().len(), 1);
}

fn cell_rows(primitives: &[Value]) -> Vec<u64> {
    let mut rows: Vec<u64> = primitives
        .iter()
        .filter_map(|primitive| primitive["cell"]["row"].as_u64())
        .collect();
    rows.sort_unstable();
    rows.dedup();
    rows
}

#[test]
fn cells_after_a_removed_row_keep_their_source_row_index() {
    let rows = [row("Alpha", ""), hidden_row("Bravo"), row("Charlie", "")];
    let hidden = render(&rows, false);
    let table = hidden.table.unwrap();
    assert_eq!(table["block"]["rows"][1]["sourceRow"], json!(2));
    assert_eq!(cell_rows(&hidden.primitives), vec![0, 2]);
    let charlie = hidden
        .primitives
        .iter()
        .filter(|primitive| primitive["cell"]["row"] == json!(2))
        .filter_map(|primitive| primitive["text"].as_str())
        .collect::<String>();
    assert!(charlie.contains("Charlie"), "{charlie}");
    assert!(
        hidden
            .primitives
            .iter()
            .filter_map(|primitive| primitive["table"]["rowCount"].as_u64())
            .all(|count| count == 3)
    );

    let shown = render(&rows, true);
    assert_eq!(cell_rows(&shown.primitives), vec![0, 1, 2]);
    assert!(
        shown.table.unwrap()["block"]["rows"][1]
            .get("sourceRow")
            .is_none()
    );
}

fn kept_rows(hidden: String) -> usize {
    render(&[row("Alpha", ""), hidden, row("Charlie", "")], false)
        .table
        .unwrap()["block"]["rows"]
        .as_array()
        .unwrap()
        .len()
}

#[test]
fn a_hidden_row_holding_only_a_vanished_empty_field_is_removed() {
    let field = |r_pr: &str, result: &str| {
        row_with_run(
            &format!(r#"<w:fldSimple w:instr=" DATE "><w:r>{r_pr}{result}</w:r></w:fldSimple>"#),
            HIDDEN,
        )
    };
    let rows = [
        row("Alpha", ""),
        field(VANISH, "<w:t/>"),
        row("Charlie", ""),
    ];
    let hidden = render(&rows, false);
    assert_eq!(
        hidden.table.unwrap()["block"]["rows"]
            .as_array()
            .unwrap()
            .len(),
        2
    );
    assert_eq!(cell_rows(&hidden.primitives), vec![0, 2]);
    assert_eq!(kept_rows(field("", "<w:t/>")), 3);
    assert_eq!(kept_rows(field("", "<w:t>9 November 2026</w:t>")), 3);
}

#[test]
fn a_hidden_row_keeps_a_field_mixing_vanished_and_visible_runs() {
    let run = format!(
        r#"<w:fldSimple w:instr=" DATE "><w:r>{VANISH}<w:t>Hidden</w:t></w:r><w:r><w:t>Shown</w:t></w:r></w:fldSimple>"#
    );
    assert_eq!(kept_rows(row_with_run(&run, HIDDEN)), 3);
}

#[test]
fn a_hidden_row_keeps_a_visible_form_checkbox_without_a_cached_result() {
    let run = r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="Check1"/><w:enabled/><w:calcOnExit w:val="0"/><w:checkBox><w:sizeAuto/><w:default w:val="0"/></w:checkBox></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMCHECKBOX </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#;
    assert_eq!(kept_rows(row_with_run(run, HIDDEN)), 3);
}

#[test]
fn a_hidden_row_keeps_an_inline_control_holding_a_drawing() {
    let run = r#"<w:sdt><w:sdtContent><w:r><w:drawing xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wp:inline><wp:extent cx="914400" cy="457200"/><wp:docPr id="7" name="Shape"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp><wps:spPr><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:sdtContent></w:sdt>"#;
    assert_eq!(kept_rows(row_with_run(run, HIDDEN)), 3);
}

#[test]
fn a_hidden_row_of_vanished_complex_field_runs_is_removed() {
    let run = format!(
        r#"<w:r>{VANISH}<w:fldChar w:fldCharType="begin"/></w:r><w:r>{VANISH}<w:instrText xml:space="preserve"> REF Target </w:instrText></w:r><w:r>{VANISH}<w:fldChar w:fldCharType="separate"/></w:r><w:r>{VANISH}<w:t>Result</w:t></w:r><w:r>{VANISH}<w:fldChar w:fldCharType="end"/></w:r>"#
    );
    assert_eq!(kept_rows(row_with_run(&run, HIDDEN)), 2);
}

#[test]
fn a_hidden_row_starting_a_merge_into_a_visible_row_stays() {
    let rendered = render(
        &[
            merged("Alpha", "", ""),
            merged("Bravo", r#" w:val="restart""#, HIDDEN),
            merged("", "", ""),
        ],
        false,
    );
    let table = rendered.table.unwrap();
    let rows = table["block"]["rows"].as_array().unwrap();
    assert_eq!(rows.len(), 3);
    assert_eq!(rows[1]["hidden"], json!(true));
    assert_eq!(rows[1]["cells"][0]["rowSpan"], json!(2.0));
    assert!(rows[2].get("sourceRow").is_none());
}

fn row_with_nested_table(tbl_pr: &str, tc_pr: &str) -> String {
    format!(
        r#"<w:tr>{HIDDEN}<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p><w:r>{VANISH}<w:t>Hidden</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/>{tbl_pr}</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/>{tc_pr}</w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:p/></w:tc></w:tr>"#
    )
}

#[test]
fn a_hidden_row_keeps_a_nested_table() {
    let rows_of = |nested: String| render(&[row("Alpha", ""), nested, row("Charlie", "")], false);
    let count = |rendered: Rendered| {
        rendered.table.unwrap()["block"]["rows"]
            .as_array()
            .unwrap()
            .len()
    };
    let shaded = row_with_nested_table(
        "",
        r#"<w:shd w:val="clear" w:color="auto" w:fill="FF0000"/>"#,
    );
    assert_eq!(count(rows_of(shaded)), 3);
    let bordered = row_with_nested_table(
        r#"<w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/><w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/></w:tblBorders>"#,
        "",
    );
    assert_eq!(count(rows_of(bordered)), 3);
    assert_eq!(count(rows_of(row_with_nested_table("", ""))), 3);
}
