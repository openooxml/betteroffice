#[allow(dead_code)]
#[path = "support/page_fixture.rs"]
mod fixture;

use serde_json::Value;

type Spec = (&'static str, u32, &'static str);

const S: Spec = ("single", 6, "000000");
const NIL: Spec = ("nil", 0, "");
const INNER_XS: [f64; 4] = [216.0, 336.0, 456.0, 576.0];
const SECT_PR: &str = concat!(
    r#"<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>"#,
    r#"<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" "#,
    r#"w:header="720" w:footer="720" w:gutter="0"/></w:sectPr>"#,
);

fn border_xml(side: &str, (val, sz, color): Spec) -> String {
    if matches!(val, "nil" | "none") {
        format!(r#"<w:{side} w:val="{val}"/>"#)
    } else {
        format!(r#"<w:{side} w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>"#)
    }
}

fn build_table(
    rows: usize,
    cols: usize,
    tbl_borders: &str,
    mut cell_sides: impl FnMut(usize, usize) -> Vec<(&'static str, Spec)>,
) -> String {
    let col_w = 1800;
    let tbl_w = col_w * cols;
    let mut xml = format!(r#"<w:tbl><w:tblPr><w:tblW w:w="{tbl_w}" w:type="dxa"/>"#);
    xml.push_str(tbl_borders);
    xml.push_str(r#"<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>"#);
    for _ in 0..cols {
        xml.push_str(&format!(r#"<w:gridCol w:w="{col_w}"/>"#));
    }
    xml.push_str("</w:tblGrid>");
    for r in 0..rows {
        xml.push_str("<w:tr>");
        for c in 0..cols {
            let borders: String = cell_sides(r, c)
                .into_iter()
                .map(|(side, spec)| border_xml(side, spec))
                .collect();
            let tc_borders = if borders.is_empty() {
                String::new()
            } else {
                format!("<w:tcBorders>{borders}</w:tcBorders>")
            };
            xml.push_str(&format!(
                r#"<w:tc><w:tcPr><w:tcW w:w="{col_w}" w:type="dxa"/>{tc_borders}</w:tcPr>"#
            ));
            xml.push_str(&format!(
                r#"<w:p><w:r><w:t>r{r}c{c}</w:t></w:r></w:p></w:tc>"#
            ));
        }
        xml.push_str("</w:tr>");
    }
    xml.push_str("</w:tbl>");
    xml
}

fn table_border_lines(table_xml: &str) -> Vec<Value> {
    page_table_border_lines(table_xml, 0)
}

fn page_table_border_lines(table_xml: &str, page: usize) -> Vec<Value> {
    let body = format!("{table_xml}<w:p/>{SECT_PR}");
    let bytes = fixture::with_body(&body);
    let (engine, _) = fixture::laid_out(&bytes, 76800);
    engine.build_display_list_frame("{}", 0).unwrap();
    let display = engine
        .with_display_list(|display| serde_json::to_value(display).unwrap())
        .unwrap();
    display["pages"][page]["primitives"]
        .as_array()
        .unwrap()
        .iter()
        .filter(|p| p["kind"] == "line" && p["role"] == "table-border")
        .cloned()
        .collect()
}

fn find_vertical(lines: &[Value], row: usize, x: f64) -> &Value {
    lines
        .iter()
        .find(|line| {
            line["cell"]["row"].as_u64() == Some(row as u64)
                && (line["x1"].as_f64().unwrap() - x).abs() < 0.01
                && (line["x2"].as_f64().unwrap() - x).abs() < 0.01
                && line["y2"].as_f64().unwrap() > line["y1"].as_f64().unwrap()
        })
        .unwrap_or_else(|| panic!("missing vertical border at row {row}, x={x}"))
}

fn assert_line(line: &Value, width: f64, color: &str, style: &str) {
    let actual_w = line["strokeWidth"].as_f64().unwrap();
    assert!((actual_w - width).abs() < 0.01, "{actual_w} vs {width}");
    assert_eq!(line["color"], color);
    assert_eq!(line["borderStyle"], style);
}

#[test]
fn inner_verticals_resolve_from_either_neighbour() {
    let issue = build_table(3, 5, "", |_, c| {
        let mut sides = vec![("top", S), ("left", S), ("bottom", S)];
        if c == 4 {
            sides.push(("right", S));
        }
        sides
    });
    let mirror = build_table(3, 5, "", |_, c| {
        let mut sides = vec![("top", S), ("right", S), ("bottom", S)];
        if c == 0 {
            sides.push(("left", S));
        }
        sides
    });
    let nil_right = build_table(3, 5, "", |_, c| {
        vec![
            ("top", S),
            ("left", S),
            ("bottom", S),
            ("right", if c == 4 { S } else { NIL }),
        ]
    });
    for xml in [issue, mirror, nil_right] {
        let lines = table_border_lines(&xml);
        for row in 0..3 {
            for x in INNER_XS {
                assert_line(find_vertical(&lines, row, x), 1.0, "#000000", "solid");
            }
        }
    }
}

#[test]
fn inner_horizontals_resolve_from_lower_cell_top_border() {
    let horizontal = build_table(3, 5, "", |r, _| {
        let mut sides = vec![("top", S), ("left", S), ("right", S)];
        if r == 2 {
            sides.push(("bottom", S));
        }
        sides
    });
    let lines = table_border_lines(&horizontal);
    for row in 0..2 {
        for col in 0..5 {
            let x1 = 96.0 + col as f64 * 120.0;
            let x2 = x1 + 120.0;
            let line = lines
                .iter()
                .find(|l| {
                    l["cell"]["row"].as_u64() == Some(row as u64)
                        && l["cell"]["col"].as_u64() == Some(col as u64)
                        && (l["x1"].as_f64().unwrap() - x1).abs() < 0.01
                        && (l["x2"].as_f64().unwrap() - x2).abs() < 0.01
                        && (l["y1"].as_f64().unwrap() - l["y2"].as_f64().unwrap()).abs() < 0.01
                })
                .unwrap_or_else(|| panic!("missing inner horizontal at row {row}, col {col}"));
            assert_line(line, 1.0, "#000000", "solid");
        }
    }
}

#[test]
fn conflicting_cell_borders_follow_word_weight_and_darkness_rules() {
    let rows_spec: [(Spec, Spec, (f64, &str, &str)); 4] = [
        (
            ("single", 4, "FF0000"),
            ("single", 24, "0000FF"),
            (4.0, "#0000FF", "solid"),
        ),
        (
            ("double", 4, "FF0000"),
            ("single", 8, "0000FF"),
            (0.667, "#FF0000", "double"),
        ),
        (
            ("single", 12, "808080"),
            ("single", 12, "000000"),
            (2.0, "#000000", "solid"),
        ),
        (
            ("single", 24, "FF0000"),
            ("single", 4, "0000FF"),
            (4.0, "#FF0000", "solid"),
        ),
    ];
    let xml = build_table(4, 5, "", |r, c| {
        let (right_spec, left_spec, _) = rows_spec[r];
        vec![
            ("top", S),
            ("bottom", S),
            ("left", if c == 0 { S } else { left_spec }),
            ("right", if c == 4 { S } else { right_spec }),
        ]
    });
    let lines = table_border_lines(&xml);
    for (row, &(_, _, (width, color, style))) in rows_spec.iter().enumerate() {
        for x in INNER_XS {
            assert_line(find_vertical(&lines, row, x), width, color, style);
        }
    }
}

#[test]
fn cell_borders_resolve_against_inherited_tbl_borders_inside_v() {
    let thick: Spec = ("single", 12, "000000");
    let tbl_borders = format!(
        "<w:tblBorders>{}</w:tblBorders>",
        ["top", "left", "bottom", "right", "insideH", "insideV"]
            .into_iter()
            .map(|side| border_xml(side, thick))
            .collect::<String>()
    );
    let xml = build_table(4, 5, &tbl_borders, |r, c| match (r, c) {
        (0, 2) => vec![("left", ("single", 2, "FF0000"))],
        (1, 2) => vec![("left", NIL)],
        (2, 1) => vec![("right", NIL)],
        (3, 2) => vec![("left", ("double", 6, "FF0000"))],
        _ => Vec::new(),
    });
    let lines = table_border_lines(&xml);
    let expected = [
        (2.0, "#000000", "solid"),
        (2.0, "#000000", "solid"),
        (2.0, "#000000", "solid"),
        (1.0, "#FF0000", "double"),
    ];
    for (row, (width, color, style)) in expected.into_iter().enumerate() {
        assert_line(find_vertical(&lines, row, 336.0), width, color, style);
    }
}

#[test]
fn repeated_header_bottom_resolves_against_the_first_body_row() {
    let rows = 70;
    let xml = build_table(rows, 5, "", |r, _| {
        let mut sides = vec![("top", S), ("left", S), ("right", S)];
        if r == rows - 1 {
            sides.push(("bottom", S));
        }
        sides
    })
    .replacen("<w:tr>", "<w:tr><w:trPr><w:tblHeader/></w:trPr>", 1);
    let header_ys: Vec<f64> = page_table_border_lines(&xml, 1)
        .iter()
        .filter(|l| l["cell"]["row"] == 0 && l["cell"]["col"] == 0 && l["y1"] == l["y2"])
        .map(|l| l["y1"].as_f64().unwrap())
        .collect();
    assert_eq!(
        header_ys.len(),
        2,
        "repeated header top and bottom: {header_ys:?}"
    );
}
