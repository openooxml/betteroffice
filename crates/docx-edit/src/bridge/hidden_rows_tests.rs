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

#[test]
fn a_merge_across_a_hidden_row_spans_only_the_visible_rows() {
    let merged = |text: &str, merge: &str, tr_pr: &str| {
        let r_pr = if tr_pr == HIDDEN { VANISH } else { "" };
        format!(
            r#"<w:tr>{tr_pr}<w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/><w:vMerge{merge}/></w:tcPr><w:p><w:r>{r_pr}<w:t>{text}</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2340" w:type="dxa"/></w:tcPr><w:p><w:r>{r_pr}<w:t>Side{text}</w:t></w:r></w:p></w:tc></w:tr>"#
        )
    };
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

    let starting = render(
        &[
            merged("Alpha", "", ""),
            merged("Bravo", r#" w:val="restart""#, HIDDEN),
            merged("", "", ""),
        ],
        false,
    );
    let rows = starting.table.unwrap()["block"]["rows"].clone();
    assert_eq!(rows.as_array().unwrap().len(), 2);
    assert_eq!(rows[1]["cells"].as_array().unwrap().len(), 2);
    assert_eq!(rows[1]["cells"][0]["rowSpan"], json!(1.0));
    assert!(!painted(&starting.primitives, "Bravo"));
    assert!(painted(&starting.primitives, "SideAlpha"));
}
