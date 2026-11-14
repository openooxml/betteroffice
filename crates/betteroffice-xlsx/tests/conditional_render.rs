use betteroffice_xlsx::{DrawCmd, GridGeometry, SheetId, Viewport, Workbook};

const RED: &str = "#ffc7ce";
const GREEN: &str = "#c6efce";
const BLUE: &str = "#bdd7ee";

fn workbook(values: &[(&str, &str)], formats: &str, dxfs: &[&str]) -> Workbook {
    let cells: String = values
        .iter()
        .map(|(at, value)| {
            let row: String = at.chars().filter(char::is_ascii_digit).collect();
            let cell = match value.parse::<f64>() {
                Ok(_) => format!(r#"<c r="{at}"><v>{value}</v></c>"#),
                Err(_) => format!(r#"<c r="{at}" t="inlineStr"><is><t>{value}</t></is></c>"#),
            };
            format!(r#"<row r="{row}">{cell}</row>"#)
        })
        .collect();
    let dxfs: String = dxfs.iter().map(|dxf| format!("<dxf>{dxf}</dxf>")).collect();
    let parts: Vec<(String, Vec<u8>)> = vec![
        ("xl/workbook.xml".into(), br#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_vec()),
        ("xl/_rels/workbook.xml.rels".into(), br#"<Relationships><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#.to_vec()),
        ("xl/worksheets/sheet1.xml".into(), format!("<worksheet><sheetData>{cells}</sheetData>{formats}</worksheet>").into_bytes()),
        ("xl/styles.xml".into(), format!(r#"<styleSheet><cellXfs count="1"><xf/></cellXfs><dxfs count="{}">{dxfs}</dxfs></styleSheet>"#, dxfs.len()).into_bytes()),
    ];
    Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap()
}

fn fill(color: &str) -> String {
    format!(
        r#"<fill><patternFill><bgColor rgb="FF{}"/></patternFill></fill>"#,
        &color[1..]
    )
}

/// The rule-colored fills a display list paints, by `(row, col)`.
fn fills(workbook: &Workbook, viewport: &Viewport) -> Vec<((u32, u32), String)> {
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    let list = workbook.display_list(viewport).unwrap();
    list.commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::FillRect { x, y, color, .. } if [RED, GREEN, BLUE].contains(&&**color) => {
                let at = |offset: f32, start: fn(&GridGeometry, u32) -> f32| {
                    (0..200).rfind(|&index| start(&geometry, index) <= offset + 0.5)
                };
                Some((
                    (
                        at(y + viewport.y, GridGeometry::row_y)?,
                        at(x + viewport.x, GridGeometry::col_x)?,
                    ),
                    color.to_string(),
                ))
            }
            _ => None,
        })
        .collect()
}

/// `(text, color, bold, italic)` of each text run a display list paints.
fn texts(workbook: &Workbook) -> Vec<(String, String, bool, bool)> {
    let list = workbook.display_list(&viewport(0.0)).unwrap();
    list.commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::Text {
                text,
                color,
                bold,
                italic,
                ..
            } => Some((text.to_string(), color.to_string(), *bold, *italic)),
            _ => None,
        })
        .collect()
}

fn viewport(y: f32) -> Viewport {
    Viewport {
        x: 0.0,
        y,
        width: 400.0,
        height: 200.0,
    }
}

fn red(rows: &[u32]) -> Vec<((u32, u32), String)> {
    rows.iter().map(|&row| ((row, 0), RED.to_owned())).collect()
}

#[test]
fn a_cell_is_rule_paints_only_the_cells_it_matches() {
    let workbook = workbook(
        &[("A1", "1"), ("A2", "10"), ("A3", "3"), ("A4", "20")],
        r#"<conditionalFormatting sqref="A1:A4"><cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>"#,
        &[&format!(
            r#"<font><b/><i/><color rgb="FF9C0006"/></font>{}"#,
            fill(RED)
        )],
    );
    assert_eq!(fills(&workbook, &viewport(0.0)), red(&[1, 3]));
    let styled = |text: &str| (text.to_owned(), "#9c0006".to_owned(), true, true);
    let texts = texts(&workbook);
    assert!(texts.contains(&styled("10")), "{texts:?}");
    assert!(texts.contains(&styled("20")), "{texts:?}");
    assert!(!texts.iter().any(|run| run.0 == "1" && run.2), "{texts:?}");
}

#[test]
fn an_expression_rule_shifts_its_relative_references_per_cell() {
    let workbook = workbook(
        &[("B1", "1"), ("B2", "0"), ("B3", "1"), ("B4", "0")],
        r#"<conditionalFormatting sqref="A1:A4"><cfRule type="expression" dxfId="0" priority="1"><formula>AND(B1&gt;0,$B$2=0)</formula></cfRule></conditionalFormatting>"#,
        &[&fill(RED)],
    );
    assert_eq!(fills(&workbook, &viewport(0.0)), red(&[0, 2]));
}

#[test]
fn priority_and_stop_if_true_pick_each_property() {
    let workbook = workbook(
        &[("A1", "1"), ("A2", "10"), ("A3", "20")],
        &format!(
            "{}{}",
            r#"<conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" dxfId="2" priority="3" operator="greaterThan"><formula>0</formula></cfRule></conditionalFormatting>"#,
            r#"<conditionalFormatting sqref="A1:A3"><cfRule type="cellIs" dxfId="0" priority="1" stopIfTrue="1" operator="greaterThan"><formula>15</formula></cfRule><cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>5</formula></cfRule></conditionalFormatting>"#,
        ),
        &[
            &fill(RED),
            &fill(GREEN),
            &format!("<font><b/></font>{}", fill(BLUE)),
        ],
    );
    assert_eq!(
        fills(&workbook, &viewport(0.0)),
        [
            ((0, 0), BLUE.to_owned()),
            ((1, 0), GREEN.to_owned()),
            ((2, 0), RED.to_owned())
        ]
    );
    let bold: Vec<_> = texts(&workbook)
        .into_iter()
        .map(|(text, _, bold, _)| (text, bold))
        .collect();
    assert_eq!(
        bold,
        [
            ("1".into(), true),
            ("10".into(), true),
            ("20".into(), false)
        ]
    );
}

#[test]
fn a_top_rule_ranks_its_whole_range_when_part_is_in_view() {
    let values: Vec<(String, String)> = (1..=60)
        .map(|row| (format!("A{row}"), row.to_string()))
        .collect();
    let values: Vec<(&str, &str)> = values
        .iter()
        .map(|(at, value)| (at.as_str(), value.as_str()))
        .collect();
    let workbook = workbook(
        &values,
        r#"<conditionalFormatting sqref="A1:A60"><cfRule type="top10" dxfId="0" priority="1" rank="3"/><cfRule type="top10" dxfId="0" priority="2" rank="5" percent="1" bottom="1"/></conditionalFormatting>"#,
        &[&fill(RED)],
    );
    assert_eq!(fills(&workbook, &viewport(0.0)), red(&[0, 1, 2]));
    let geometry = GridGeometry::new(
        workbook.sheet(SheetId(0)).unwrap(),
        &workbook.model().styles,
    );
    assert_eq!(
        fills(&workbook, &viewport(geometry.row_y(50))),
        red(&[57, 58, 59])
    );
}

#[test]
fn average_duplicate_and_text_rules_paint_their_matches() {
    let workbook = workbook(
        &[
            ("A1", "1"),
            ("A2", "2"),
            ("A3", "9"),
            ("B1", "x"),
            ("B2", "Y"),
            ("B3", "y"),
            ("C1", "overdue"),
            ("C2", "paid"),
        ],
        r#"<conditionalFormatting sqref="A1:A3"><cfRule type="aboveAverage" dxfId="0" priority="1"/></conditionalFormatting><conditionalFormatting sqref="B1:B3"><cfRule type="duplicateValues" dxfId="0" priority="2"/></conditionalFormatting><conditionalFormatting sqref="C1:C2"><cfRule type="containsText" dxfId="0" priority="3" operator="containsText" text="due"><formula>NOT(ISERROR(SEARCH("due",C1)))</formula></cfRule></conditionalFormatting>"#,
        &[&fill(RED)],
    );
    let mut painted: Vec<_> = fills(&workbook, &viewport(0.0))
        .into_iter()
        .map(|(at, _)| at)
        .collect();
    painted.sort();
    assert_eq!(painted, [(0, 2), (1, 1), (2, 0), (2, 1)]);
}
