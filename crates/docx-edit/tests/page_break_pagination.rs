#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use docx_edit::{EditCtx, EngineSession, FormatPolicy, Position};
use serde_json::Value;

const UNEQUAL_COLUMNS: &str = r#"<w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="3256" w:space="2870"/><w:col w:w="3954"/></w:cols>"#;

fn layout_page_body(width: u32, height: u32, columns: &str, body: &str) -> Value {
    let body = format!(
        r#"{body}<w:sectPr><w:pgSz w:w="{width}" w:h="{height}"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/>{columns}</w:sectPr>"#,
    );
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), 9403);
    serde_json::from_str(&engine.layout_document_with_regions_json(&request).unwrap()).unwrap()
}

fn column_layout(columns: &str, height: u32, first: &str, second: &str) -> Value {
    let body = format!(
        "{}{}",
        fixture::p(
            "71000001",
            &format!(
                r#"<w:pPr><w:spacing w:before="0" w:after="0"/><w:widowControl w:val="0"/></w:pPr>{}"#,
                fixture::r(first)
            ),
        ),
        fixture::p(
            "71000002",
            &format!(
                r#"<w:pPr><w:jc w:val="right"/><w:spacing w:before="0" w:after="0"/></w:pPr><w:r><w:br w:type="column"/></w:r>{}"#,
                fixture::r(second)
            ),
        ),
    );
    layout_page_body(12240, height, columns, &body)
}

fn assert_column_paragraphs(
    output: &Value,
    widths: [f64; 2],
    second_x: f64,
    fragment_widths: [f64; 2],
) {
    let paragraphs: Vec<_> = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|block| {
            block["block"]["kind"] == "paragraph"
                && block["block"]["runs"].as_array().is_some_and(|runs| {
                    runs.iter()
                        .any(|run| run["text"].as_str().is_some_and(|text| !text.is_empty()))
                })
        })
        .collect();
    let first_id = &paragraphs[0]["block"]["id"];
    let second_id = &paragraphs[1]["block"]["id"];
    let fragments: Vec<_> = output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["kind"] == "paragraph")
        .collect();
    for (id, x, width) in [
        (first_id, 72.0, fragment_widths[0]),
        (second_id, second_x, fragment_widths[1]),
    ] {
        let fragments: Vec<_> = fragments
            .iter()
            .filter(|fragment| &fragment["blockId"] == id)
            .collect();
        assert!(!fragments.is_empty());
        for fragment in fragments {
            assert!((fragment["x"].as_f64().unwrap() - x).abs() < 0.01);
            assert!((fragment["width"].as_f64().unwrap() - width).abs() < 0.01);
        }
    }
    let first_lines = paragraphs[0]["measure"]["lines"].as_array().unwrap();
    assert!(first_lines.len() > 1);
    assert!(
        first_lines
            .iter()
            .all(|line| visible_line_width(paragraphs[0], line) <= widths[0] + 0.01)
    );

    let right_edges = painted_line_right_edges(output, "71000002");
    assert!(!right_edges.is_empty());
    let second_lines = paragraphs[1]["measure"]["lines"].as_array().unwrap();
    assert_eq!(right_edges.len(), second_lines.len());
    for (index, right) in &right_edges {
        let line = &second_lines[*index as usize];
        assert!(visible_line_width(paragraphs[1], line) <= widths[1] + 0.01);
        let expected_right = second_x + fragment_widths[1].max(line["width"].as_f64().unwrap());
        assert!((right - expected_right).abs() < 0.1, "line ends at {right}");
    }
}

fn visible_line_width(measured: &Value, line: &Value) -> f64 {
    let trailing_width: f64 = line["clusterAdvances"]
        .as_array()
        .unwrap()
        .iter()
        .rev()
        .take_while(|cluster| {
            let run = cluster["runIndex"].as_u64().unwrap() as usize;
            let text = measured["block"]["runs"][run]["text"].as_str().unwrap();
            let start = cluster["startChar"].as_u64().unwrap() as usize;
            let end = cluster["endChar"].as_u64().unwrap() as usize;
            start < end
                && text
                    .encode_utf16()
                    .skip(start)
                    .take(end - start)
                    .all(|unit| unit == u16::from(b' '))
        })
        .map(|cluster| cluster["advance"].as_f64().unwrap())
        .sum();
    line["width"].as_f64().unwrap() - trailing_width
}

fn painted_line_right_edges(output: &Value, para_id: &str) -> std::collections::BTreeMap<u64, f64> {
    let display = docx_layout::build_display_list_value(&output.to_string()).unwrap();
    let display = serde_json::to_value(display).unwrap();
    let mut right_edges = std::collections::BTreeMap::<u64, f64>::new();
    for primitive in display["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["primitives"].as_array().unwrap())
    {
        if primitive["paraId"] != para_id {
            continue;
        }
        let right = match primitive["kind"].as_str() {
            Some("glyphRun") => primitive["glyphs"]
                .as_array()
                .unwrap()
                .iter()
                .map(|glyph| glyph["x"].as_f64().unwrap() + glyph["advance"].as_f64().unwrap())
                .fold(f64::NEG_INFINITY, f64::max),
            Some("text") => primitive["x"].as_f64().unwrap() + primitive["width"].as_f64().unwrap(),
            _ => continue,
        };
        let line = primitive["lineIndex"].as_u64().unwrap();
        right_edges
            .entry(line)
            .and_modify(|edge| *edge = edge.max(right))
            .or_insert(right);
    }
    right_edges
}

#[test]
fn unequal_columns_measure_and_paint_inside_the_active_column() {
    let output = column_layout(
        UNEQUAL_COLUMNS,
        15840,
        &"Column text wraps within its own width. ".repeat(6),
        &"Upper right title with several words. ".repeat(5),
    );
    assert_column_paragraphs(
        &output,
        [3256.0 / 15.0, 3954.0 / 15.0],
        72.0 + 6126.0 / 15.0,
        [3256.0 / 15.0, 3954.0 / 15.0],
    );

    let control = layout_page_body(
        6114,
        15840,
        "",
        &fixture::p(
            "71000002",
            &fixture::r(&"Upper right title with several words. ".repeat(5)),
        ),
    );
    let title = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|measured| {
            measured["block"]["paraId"] == "71000002"
                && measured["block"]["runs"].as_array().is_some_and(|runs| {
                    runs.iter()
                        .any(|run| run["text"].as_str().is_some_and(|text| !text.is_empty()))
                })
        })
        .unwrap();
    assert_eq!(line_ranges(title), line_ranges(&control["measured"][0]));
}

fn line_ranges(measured: &Value) -> Vec<(u64, u64, u64, u64)> {
    measured["measure"]["lines"]
        .as_array()
        .unwrap()
        .iter()
        .map(|line| {
            (
                line["headRun"].as_u64().unwrap(),
                line["headChar"].as_u64().unwrap(),
                line["tailRun"].as_u64().unwrap(),
                line["tailChar"].as_u64().unwrap(),
            )
        })
        .collect()
}

#[test]
fn equal_columns_keep_their_widths_and_origins() {
    let output = column_layout(
        r#"<w:cols w:num="2" w:space="720"/>"#,
        15840,
        &"Column text wraps within its own width. ".repeat(6),
        "Upper right title",
    );
    assert_column_paragraphs(&output, [312.0, 312.0], 432.0, [672.0, 672.0]);
}

#[test]
fn a_paragraph_rewraps_when_it_flows_between_unequal_columns() {
    let body = fixture::p(
        "71000003",
        &format!(
            r#"<w:pPr><w:jc w:val="right"/><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>{}"#,
            fixture::r(&"Words continue through columns with different widths. ".repeat(20))
        ),
    );
    let output = layout_page_body(12240, 3360, UNEQUAL_COLUMNS, &body);
    assert_spanning_columns(&output);
}

fn assert_spanning_columns(output: &Value) {
    let lines = output["measured"][0]["measure"]["lines"]
        .as_array()
        .unwrap();
    let right_edges = painted_line_right_edges(output, "71000003");
    assert_eq!(right_edges.len(), lines.len());
    let mut next_line = 0;
    let mut columns_seen = [false; 2];
    for fragment in output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["kind"] == "paragraph")
    {
        let x = fragment["x"].as_f64().unwrap();
        let column = usize::from(x > 73.0);
        columns_seen[column] = true;
        let width = [3256.0 / 15.0, 3954.0 / 15.0][column];
        assert!((fragment["width"].as_f64().unwrap() - width).abs() < 0.01);
        let from = fragment["fromLine"].as_u64().unwrap() as usize;
        let to = fragment["toLine"].as_u64().unwrap() as usize;
        assert_eq!(from, next_line);
        next_line = to;
        for (index, line) in lines.iter().enumerate().take(to).skip(from) {
            assert!(visible_line_width(&output["measured"][0], line) <= width + 0.01);
            assert!(line["rightOffset"].as_f64().unwrap_or(0.0).abs() < 0.01);
            let expected_right = x + width.max(line["width"].as_f64().unwrap());
            assert!((right_edges[&(index as u64)] - expected_right).abs() < 0.1);
        }
    }
    assert_eq!(next_line, lines.len());
    assert_eq!(columns_seen, [true, true]);

    let second_column_lines: Vec<_> = output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["x"].as_f64().unwrap() > 73.0)
        .flat_map(|fragment| {
            let from = fragment["fromLine"].as_u64().unwrap() as usize;
            let to = fragment["toLine"].as_u64().unwrap() as usize;
            &lines[from..to]
        })
        .collect();
    assert!(
        second_column_lines
            .iter()
            .any(|line| line["width"].as_f64().unwrap() > 3256.0 / 15.0 + 1.0)
    );
}

#[test]
fn balancing_rewraps_a_paragraph_in_each_unequal_column() {
    let body = fixture::p(
        "71000003",
        &format!(
            r#"<w:pPr><w:jc w:val="right"/><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>{}"#,
            fixture::r(&"Words continue through columns with different widths. ".repeat(6))
        ),
    );
    let output = layout_page_body(12240, 15840, UNEQUAL_COLUMNS, &body);
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 1);
    assert_spanning_columns(&output);
}

#[test]
fn keep_lines_remeasures_a_paragraph_after_moving_it_to_column_two() {
    let title = fixture::p(
        "71000008",
        &format!(
            r#"<w:pPr><w:keepLines/><w:spacing w:before="0" w:after="0"/></w:pPr>{}"#,
            fixture::r(&"A paragraph kept together inside its column. ".repeat(3))
        ),
    );
    let body = format!(
        "{}{}",
        fixture::p(
            "71000009",
            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="720" w:lineRule="exact"/></w:pPr><w:r><w:t>One</w:t><w:br/><w:t>Two</w:t><w:br/><w:t>Three</w:t></w:r>"#,
        ),
        title,
    );
    let output = layout_page_body(12240, 4560, UNEQUAL_COLUMNS, &body);
    let control = layout_page_body(6114, 15840, "", &title);
    let measured = &output["measured"][1];
    assert_eq!(line_ranges(measured), line_ranges(&control["measured"][0]));
    let fragments: Vec<_> = output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["blockId"] == measured["block"]["id"])
        .collect();
    assert_eq!(fragments.len(), 1);
    assert!((fragments[0]["x"].as_f64().unwrap() - (72.0 + 6126.0 / 15.0)).abs() < 0.01);
    assert!((fragments[0]["width"].as_f64().unwrap() - 3954.0 / 15.0).abs() < 0.01);
}

#[test]
fn keep_lines_column_width_oscillation_uses_one_stable_measure() {
    let title = fixture::p(
        "7100000A",
        &format!(
            r#"<w:pPr><w:keepLines/><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="exact"/></w:pPr>{}"#,
            fixture::r(&"A paragraph kept together inside its column. ".repeat(4)),
        ),
    );
    let narrow = layout_page_body(3960, 15840, "", &title);
    let wide = layout_page_body(10080, 15840, "", &title);
    let narrow_height = narrow["measured"][0]["measure"]["totalHeight"]
        .as_f64()
        .unwrap();
    let wide_height = wide["measured"][0]["measure"]["totalHeight"]
        .as_f64()
        .unwrap();
    let remainder = (narrow_height + wide_height) / 2.0;
    assert!(wide_height < remainder && remainder < narrow_height);
    let filler_height = narrow_height + 32.0;
    let body = format!(
        "{}{}{}",
        fixture::p(
            "7100000B",
            &format!(
                r#"<w:pPr><w:spacing w:before="0" w:after="0" w:line="{}" w:lineRule="exact"/></w:pPr>{}"#,
                (filler_height * 15.0).round() as u32,
                fixture::r("Filler"),
            ),
        ),
        title,
        fixture::p("7100000C", r#"<w:r><w:br w:type="page"/></w:r>"#),
    );
    let columns = r#"<w:cols w:num="2" w:space="360" w:equalWidth="0"><w:col w:w="1800" w:space="360"/><w:col w:w="7920"/></w:cols>"#;
    let height = 2160 + ((filler_height + remainder) * 15.0).round() as u32;
    let output = layout_page_body(12240, height, columns, &body);
    let measured = &output["measured"][1];
    assert_eq!(line_ranges(measured), line_ranges(&narrow["measured"][0]));
    let fragments: Vec<_> = output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["blockId"] == measured["block"]["id"])
        .collect();
    assert_eq!(fragments.len(), 1);
    assert_eq!(fragments[0]["x"], 216.0);
    assert_eq!(fragments[0]["width"], 528.0);
    assert_eq!(fragments[0]["fromLine"], 0);
    assert_eq!(
        fragments[0]["toLine"].as_u64().unwrap() as usize,
        measured["measure"]["lines"].as_array().unwrap().len(),
    );
}

#[test]
fn single_column_ignores_an_unequal_column_definition() {
    let body = fixture::p(
        "71000004",
        &fixture::r(&"Ordinary paragraph text. ".repeat(8)),
    );
    let plain = layout_page_body(12240, 15840, "", &body);
    let authored = layout_page_body(
        12240,
        15840,
        r#"<w:cols w:num="1" w:equalWidth="0"><w:col w:w="3256" w:space="2870"/></w:cols>"#,
        &body,
    );
    assert_eq!(plain["measured"], authored["measured"]);
    assert_eq!(
        plain["layout"]["pages"][0]["fragments"],
        authored["layout"]["pages"][0]["fragments"]
    );
    assert_eq!(
        authored["layout"]["pages"][0]["fragments"][0]["width"],
        672.0
    );
}

#[test]
fn a_table_after_a_column_break_uses_the_second_column_width() {
    let body = format!(
        r#"{}{}<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="3954"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/></w:tcPr>{}</w:tc></w:tr></w:tbl>"#,
        fixture::p("71000005", &fixture::r("First column")),
        fixture::p("71000006", r#"<w:r><w:br w:type="column"/></w:r>"#),
        fixture::p("71000007", &fixture::r("Table text")),
    );
    let output = layout_page_body(12240, 15840, UNEQUAL_COLUMNS, &body);
    let table = output["measured"]
        .as_array()
        .unwrap()
        .iter()
        .find(|block| block["block"]["kind"] == "table")
        .unwrap();
    assert!((table["measure"]["totalWidth"].as_f64().unwrap() - 3954.0 / 15.0).abs() < 0.01);
    let fragment = output["layout"]["pages"][0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .find(|fragment| fragment["kind"] == "table")
        .unwrap();
    assert!((fragment["width"].as_f64().unwrap() - 3954.0 / 15.0).abs() < 0.01);
}

#[test]
fn table_fragments_stay_inside_unequal_columns() {
    let rows: String = (0..8)
        .map(|index| {
            format!(
                r#"<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="360" w:hRule="exact"/></w:trPr><w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/><w:shd w:fill="FFFF00"/></w:tcPr>{}</w:tc></w:tr>"#,
                fixture::p(&format!("720000{index:02X}"), &fixture::r("Table row")),
            )
        })
        .collect();
    let body = format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid><w:gridCol w:w="7920"/></w:tblGrid>{rows}</w:tbl>{}"#,
        fixture::p("72000010", r#"<w:r><w:br w:type="page"/></w:r>"#),
    );
    let columns = r#"<w:cols w:num="2" w:space="360" w:equalWidth="0"><w:col w:w="7920" w:space="360"/><w:col w:w="1800"/></w:cols>"#;
    let output = layout_page_body(12240, 3360, columns, &body);
    let mut columns_seen = [false; 2];
    for fragment in output["layout"]["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["fragments"].as_array().unwrap())
        .filter(|fragment| fragment["kind"] == "table")
    {
        let x = fragment["x"].as_f64().unwrap();
        let column = usize::from(x > 73.0);
        columns_seen[column] = true;
        let right = [600.0, 744.0][column];
        assert!(x + fragment["width"].as_f64().unwrap() <= right + 0.01);
    }
    assert_eq!(columns_seen, [true, true]);

    let display = docx_layout::build_display_list_value(&output.to_string()).unwrap();
    let display = serde_json::to_value(display).unwrap();
    let mut clipped_in_narrow_column = false;
    for primitive in display["pages"]
        .as_array()
        .unwrap()
        .iter()
        .flat_map(|page| page["primitives"].as_array().unwrap())
        .filter(|primitive| primitive["table"].is_object())
    {
        let clip = &primitive["clipGroup"]["clip"];
        let x = clip["x"].as_f64().unwrap();
        let right = if x > 601.0 { 744.0 } else { 600.0 };
        assert!(x + clip["w"].as_f64().unwrap() <= right + 0.01);
        clipped_in_narrow_column |= x > 601.0;
    }
    assert!(clipped_in_narrow_column);
}

#[test]
fn a_continuous_section_applies_changed_authored_column_widths() {
    let body = format!(
        "{}{}",
        fixture::p(
            "71000010",
            &format!(
                r#"<w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/>{UNEQUAL_COLUMNS}</w:sectPr></w:pPr>{}"#,
                fixture::r("First region")
            ),
        ),
        fixture::p("71000011", &fixture::r("Second region")),
    );
    let output = layout_page_body(
        12240,
        15840,
        r#"<w:type w:val="continuous"/><w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="3954" w:space="2870"/><w:col w:w="3256"/></w:cols>"#,
        &body,
    );
    assert_eq!(output["layout"]["pages"].as_array().unwrap().len(), 1);
    let fragments: Vec<_> = output["layout"]["pages"][0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|fragment| fragment["kind"] == "paragraph")
        .collect();
    assert_eq!(fragments.len(), 2);
    assert!((fragments[0]["width"].as_f64().unwrap() - 3256.0 / 15.0).abs() < 0.01);
    assert!((fragments[1]["width"].as_f64().unwrap() - 3954.0 / 15.0).abs() < 0.01);
    assert_eq!(fragments[1]["x"], 72.0);
    assert!(fragments[1]["y"].as_f64().unwrap() > fragments[0]["y"].as_f64().unwrap());
}

fn staggered_row(exact: bool) -> Vec<u8> {
    let cells: String = [("left", 90, 240), ("right", 85, 260)]
        .into_iter()
        .enumerate()
        .map(|(cell, (prefix, count, spacing))| {
            let paragraphs: String = (0..count)
                .map(|index| {
                    fixture::p(
                        &format!("{:08X}", 0x7100_0000 + cell * 0x100 + index),
                        &format!(
                            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="{spacing}" w:lineRule="exact"/></w:pPr>{}"#,
                            fixture::r(&format!("{prefix}{index:03}"))
                        ),
                    )
                })
                .collect();
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="4320" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>{paragraphs}</w:tc>"#
            )
        })
        .collect();
    let height = if exact {
        r#"<w:trPr><w:trHeight w:val="7200" w:hRule="exact"/></w:trPr>"#
    } else {
        ""
    };
    fixture::with_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid><w:gridCol w:w="4320"/><w:gridCol w:w="4320"/></w:tblGrid><w:tr>{height}{cells}</w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#
    ))
}

fn staggered_row_layout(exact: bool) -> (Value, Value) {
    let (engine, request) = fixture::laid_out(&staggered_row(exact), 9403);
    let output = engine.layout_document_with_regions_json(&request).unwrap();
    let display = engine.build_display_list_json(&output).unwrap();
    (
        serde_json::from_str(&output).unwrap(),
        serde_json::from_str(&display).unwrap(),
    )
}

#[test]
fn centred_cells_with_staggered_lines_split_within_the_page_body() {
    let (output, display) = staggered_row_layout(false);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), display["pages"].as_array().unwrap().len());
    let mut fragment_count = 0;
    let mut texts = [String::new(), String::new()];
    for (page, painted) in pages.iter().zip(display["pages"].as_array().unwrap()) {
        let margins = page.get("bodyMargins").unwrap_or(&page["margins"]);
        let body_top = margins["top"].as_f64().unwrap();
        let body_bottom = page["size"]["h"].as_f64().unwrap() - margins["bottom"].as_f64().unwrap();
        let fragments: Vec<_> = page["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|fragment| fragment["kind"] == "table")
            .collect();
        if fragments.is_empty() {
            continue;
        }
        assert_eq!(fragments.len(), 1);
        fragment_count += 1;
        let fragment = fragments[0];
        let top = fragment["y"].as_f64().unwrap();
        let bottom = top + fragment["height"].as_f64().unwrap();
        assert!(top >= body_top - 0.02);
        assert!(bottom <= body_bottom + 0.02, "{fragment}");
        let clips = fragment["cellClips"].as_array().unwrap();
        assert_eq!(clips.len(), 2);
        assert_ne!(clips[0]["bottom"], clips[1]["bottom"]);
        for row in fragment["rowStart"].as_u64().unwrap()..fragment["rowEnd"].as_u64().unwrap() {
            assert!(
                painted["primitives"]
                    .as_array()
                    .unwrap()
                    .iter()
                    .any(|primitive| {
                        primitive["cell"]["row"].as_u64() == Some(row)
                            && primitive["text"]
                                .as_str()
                                .is_some_and(|text| !text.is_empty())
                    }),
                "row {row} has no line text in {fragment}"
            );
        }
        for primitive in painted["primitives"].as_array().unwrap() {
            let Some(text) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
                continue;
            };
            let cell = primitive["cell"]["col"].as_u64().unwrap() as usize;
            let measured_cell = &output["measured"][0]["measure"]["rows"][0]["cells"][cell];
            let line = &measured_cell["blocks"][0]["lines"][0];
            let baseline = primitive["baselineY"]
                .as_f64()
                .or_else(|| primitive["glyphs"][0]["y"].as_f64())
                .unwrap();
            let line_top = baseline - line["ascent"].as_f64().unwrap();
            let line_bottom = line_top + line["lineHeight"].as_f64().unwrap();
            let clip = &primitive["clipGroup"]["clip"];
            let clip_top = clip["y"].as_f64().unwrap();
            let clip_bottom = clip_top + clip["h"].as_f64().unwrap();
            assert!(line_top >= top.max(clip_top) - 0.02, "{text}: {line_top}");
            assert!(
                line_bottom <= bottom.min(clip_bottom) + 0.02,
                "{text}: {line_bottom}"
            );
            texts[cell].push_str(text);
        }
    }
    assert!(fragment_count > 1);
    for (cell, (prefix, count)) in [("left", 90), ("right", 85)].into_iter().enumerate() {
        assert_eq!(
            texts[cell],
            (0..count)
                .map(|index| format!("{prefix}{index:03}"))
                .collect::<String>()
        );
    }
}

#[test]
fn a_short_bottom_aligned_cell_stays_on_the_last_row_fragment() {
    assert_short_aligned_cells(&["top", "top", "bottom"]);
}

#[test]
fn short_centred_and_bottom_cells_keep_their_whole_row_alignment() {
    assert_short_aligned_cells(&["top", "top", "center", "bottom"]);
}

fn assert_short_aligned_cells(alignments: &[&str]) {
    let width = 8640 / alignments.len();
    let counts = [225, 216, 1, 1];
    let cells: String = alignments
        .iter()
        .enumerate()
        .map(|(cell, alignment)| {
            let spacing = if cell == 1 { 250 } else { 240 };
            let paragraphs: String = (0..counts[cell])
                .map(|index| {
                    fixture::p(
                        &format!("{:08X}", 0x7200_0000 + cell * 0x100 + index),
                        &format!(
                            r#"<w:pPr><w:widowControl w:val="0"/><w:spacing w:before="0" w:after="0" w:line="{spacing}" w:lineRule="exact"/></w:pPr>{}"#,
                            fixture::r(&format!("cell{cell}line{index:03}"))
                        ),
                    )
                })
                .collect();
            format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/><w:vAlign w:val="{alignment}"/></w:tcPr>{paragraphs}</w:tc>"#
            )
        })
        .collect();
    let grid = format!(r#"<w:gridCol w:w="{width}"/>"#).repeat(alignments.len());
    let bytes = fixture::with_body(&format!(
        r#"<w:tbl><w:tblPr><w:tblW w:w="8640" w:type="dxa"/><w:tblLayout w:type="fixed"/><w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/></w:tblCellMar></w:tblPr><w:tblGrid>{grid}</w:tblGrid><w:tr>{cells}</w:tr></w:tbl><w:sectPr><w:pgSz w:w="12240" w:h="15840"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>"#
    ));
    let (engine, request) = fixture::laid_out(&bytes, 9404);
    let output = engine.layout_document_with_regions_json(&request).unwrap();
    let display = engine.build_display_list_json(&output).unwrap();
    let output: Value = serde_json::from_str(&output).unwrap();
    let display: Value = serde_json::from_str(&display).unwrap();
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 5);
    assert_eq!(pages.len(), display["pages"].as_array().unwrap().len());
    let row_height = output["measured"][0]["measure"]["rows"][0]["height"]
        .as_f64()
        .unwrap();
    let mut consumed = 0.0;
    let mut texts = vec![String::new(); alignments.len()];
    for (page_index, (page, painted)) in pages
        .iter()
        .zip(display["pages"].as_array().unwrap())
        .enumerate()
    {
        let margins = page.get("bodyMargins").unwrap_or(&page["margins"]);
        let body_top = margins["top"].as_f64().unwrap();
        let body_bottom = page["size"]["h"].as_f64().unwrap() - margins["bottom"].as_f64().unwrap();
        let fragments: Vec<_> = page["fragments"]
            .as_array()
            .unwrap()
            .iter()
            .filter(|fragment| fragment["kind"] == "table")
            .collect();
        assert_eq!(fragments.len(), 1);
        let fragment = fragments[0];
        let top = fragment["y"].as_f64().unwrap();
        let bottom = top + fragment["height"].as_f64().unwrap();
        assert!(top >= body_top - 0.02);
        assert!(bottom <= body_bottom + 0.02, "{fragment}");
        for primitive in painted["primitives"].as_array().unwrap() {
            let Some(text) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
                continue;
            };
            let cell = primitive["cell"]["col"].as_u64().unwrap() as usize;
            if cell > 1 {
                let expected_page = if alignments[cell] == "center" { 2 } else { 4 };
                assert_eq!(page_index, expected_page, "{text}");
            }
            let line = &output["measured"][0]["measure"]["rows"][0]["cells"][cell]["blocks"][0]["lines"]
                [0];
            let baseline = primitive["baselineY"]
                .as_f64()
                .or_else(|| primitive["glyphs"][0]["y"].as_f64())
                .unwrap();
            let line_top = baseline - line["ascent"].as_f64().unwrap();
            let line_height = line["lineHeight"].as_f64().unwrap();
            let line_bottom = line_top + line_height;
            if cell > 1 {
                let slack = row_height - line_height;
                let offset = if alignments[cell] == "center" {
                    slack / 2.0
                } else {
                    slack
                };
                let expected_top = top + offset - consumed;
                assert!((line_top - expected_top).abs() < 0.02, "{text}: {line_top}");
            }
            let clip = &primitive["clipGroup"]["clip"];
            let clip_top = clip["y"].as_f64().unwrap();
            let clip_bottom = clip_top + clip["h"].as_f64().unwrap();
            assert!(line_top >= top.max(clip_top) - 0.02, "{text}: {line_top}");
            assert!(
                line_bottom <= bottom.min(clip_bottom) + 0.02,
                "{text}: {line_bottom}"
            );
            texts[cell].push_str(text);
        }
        consumed += fragment["height"].as_f64().unwrap();
    }
    for (cell, text) in texts.iter().enumerate() {
        assert_eq!(
            *text,
            (0..counts[cell])
                .map(|index| format!("cell{cell}line{index:03}"))
                .collect::<String>()
        );
    }
}

#[test]
fn an_exact_height_row_with_staggered_lines_stays_whole_and_clips() {
    let (output, display) = staggered_row_layout(true);
    let pages = output["layout"]["pages"].as_array().unwrap();
    assert_eq!(pages.len(), 1);
    let fragments: Vec<_> = pages[0]["fragments"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|fragment| fragment["kind"] == "table")
        .collect();
    assert_eq!(fragments.len(), 1);
    let fragment = fragments[0];
    assert_eq!(fragment["rowStart"], 0);
    assert_eq!(fragment["rowEnd"], 1);
    assert_eq!(fragment["height"], 480.0);
    assert_ne!(fragment["carriedToNext"], true);
    assert!(fragment["cellClips"].is_null());
    let top = fragment["y"].as_f64().unwrap();
    let bottom = top + fragment["height"].as_f64().unwrap();
    let mut text = String::new();
    for primitive in display["pages"][0]["primitives"].as_array().unwrap() {
        let Some(run) = primitive["text"].as_str().filter(|text| !text.is_empty()) else {
            continue;
        };
        let clip = &primitive["clipGroup"]["clip"];
        assert!((clip["y"].as_f64().unwrap() - top).abs() < 0.02);
        assert!(clip["y"].as_f64().unwrap() + clip["h"].as_f64().unwrap() <= bottom + 0.02);
        text.push_str(run);
    }
    assert!(text.contains("left000"));
    assert!(text.contains("right000"));
    assert!(!text.contains("left089"));
    assert!(!text.contains("right084"));
}

#[test]
fn an_edit_before_page_breaks_repaginates_only_its_own_page() {
    let paragraph = |id: u32| fixture::p(&format!("{id:08X}"), &fixture::r("text on its own page"));
    let page_break =
        |id: u32| fixture::p(&format!("{id:08X}"), r#"<w:r><w:br w:type="page"/></w:r>"#);
    let body: String = (0..6_u32)
        .map(|index| {
            format!(
                "{}{}",
                paragraph(0x7000_0000 + index * 2),
                page_break(0x7000_0001 + index * 2)
            )
        })
        .collect();
    let (engine, request) = fixture::laid_out(&fixture::with_body(&body), 9401);
    let before = engine.stats();
    assert!(before.retained_pages >= 6);

    engine
        .doc()
        .insert_text(
            &EditCtx::local("", ""),
            Position::new("body", 0),
            "x",
            FormatPolicy::Inherit,
        )
        .unwrap();
    let layout = engine.layout_document_with_regions_json(&request).unwrap();
    let after = engine.stats();
    assert_eq!(
        after.incremental_pagination_calls - before.incremental_pagination_calls,
        1
    );
    assert_eq!(after.rebuilt_pages, 1, "later pages converge");

    let fresh = EngineSession::new(9402);
    fresh
        .doc()
        .apply_update_v1(&engine.doc().encode_state_as_update_v1())
        .unwrap();
    assert_eq!(
        fresh.layout_document_with_regions_json(&request).unwrap(),
        layout
    );
}
