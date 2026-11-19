use betteroffice_xlsx::{
    CalculationOptions, CellRef, DrawCmd, GridGeometry, SheetId, Viewport, Workbook,
};

const RED: &str = "#ffc7ce";
const GREEN: &str = "#c6efce";
const BLUE: &str = "#bdd7ee";

fn workbook(values: &[(&str, &str)], formats: &str, dxfs: &[&str]) -> Workbook {
    workbook_named(values, formats, dxfs, "")
}

/// [`workbook`] whose workbook part also holds `names`, a `definedNames`
/// element or nothing.
fn workbook_named(values: &[(&str, &str)], formats: &str, dxfs: &[&str], names: &str) -> Workbook {
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
        ("xl/workbook.xml".into(), format!(r#"<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>{names}</workbook>"#).into_bytes()),
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

fn rule(sqref: &str, rules: &str) -> String {
    format!(r#"<conditionalFormatting sqref="{sqref}">{rules}</conditionalFormatting>"#)
}

fn expression(priority: u32, formula: &str) -> String {
    format!(
        r#"<cfRule type="expression" dxfId="0" priority="{priority}"><formula>{formula}</formula></cfRule>"#
    )
}

/// A rule above zero painting `dxfId` 0 at `priority`.
fn positive(priority: u32) -> String {
    format!(
        r#"<cfRule type="cellIs" dxfId="0" priority="{priority}" operator="greaterThan"><formula>0</formula></cfRule>"#
    )
}

/// A time period rule painting `dxfId` 1, a kind not evaluated.
fn time_period(priority: u32, stop: bool) -> String {
    format!(
        r#"<cfRule type="timePeriod" dxfId="1" priority="{priority}" stopIfTrue="{}" timePeriod="today"><formula>FLOOR(A1,1)=TODAY()</formula></cfRule>"#,
        u8::from(stop)
    )
}

fn assert_unpainted(workbook: &Workbook, case: &str) {
    let painted = fills(workbook, &viewport(0.0));
    assert!(painted.is_empty(), "{case}: {painted:?}");
}

#[test]
fn a_not_between_rule_with_an_error_bound_paints_nothing() {
    let workbook = workbook(
        &[("A1", "5"), ("A2", "50")],
        &rule(
            "A1:A2",
            r#"<cfRule type="cellIs" dxfId="0" priority="1" operator="notBetween"><formula>1/0</formula><formula>10</formula></cfRule>"#,
        ),
        &[&fill(RED)],
    );
    assert_unpainted(&workbook, "notBetween");
    assert_eq!(workbook.conditional_stats().withheld_error_bound, 2);
}

#[test]
fn a_rule_formula_unreadable_at_the_sheet_edge_withholds_the_rules_below_it() {
    let workbook = workbook(
        &[("A1", "1"), ("B1", "1"), ("C1", "1")],
        &rule(
            "A1:C1",
            &format!(
                r#"<cfRule type="expression" dxfId="1" priority="1" stopIfTrue="1"><formula>ISERROR(SUM(XFB:XFC))</formula></cfRule>{}"#,
                positive(2)
            ),
        ),
        &[&fill(RED), &fill(GREEN)],
    );
    assert_eq!(
        fills(&workbook, &viewport(0.0)),
        [((0, 0), RED.to_owned()), ((0, 1), RED.to_owned())]
    );
    assert_eq!(workbook.conditional_stats().withheld_unparsable, 1);
}

#[test]
fn a_rule_the_calc_engine_refuses_for_size_paints_nothing() {
    let workbook = workbook(
        &[("A1", "1"), ("A2", "1")],
        &[
            rule("A1:A2", &expression(1, "ISERROR(SUM($B$1:$Z$50000))")),
            rule("C1:C2", &expression(2, "ISERROR(SUM(D1:Z50000))")),
        ]
        .concat(),
        &[&fill(RED)],
    );
    assert_unpainted(&workbook, "refused");
    assert_eq!(workbook.conditional_stats().withheld_refused, 4);
}

#[test]
fn an_unsupported_rule_withholds_every_format_on_the_cells_it_covers() {
    let book = |rules: String| {
        workbook(
            &[("A1", "1"), ("A2", "2")],
            &rule("A1:A2", &rules),
            &[&fill(RED), &fill(GREEN)],
        )
    };
    for unsupported in [
        time_period(1, true),
        time_period(1, false),
        r#"<cfRule type="cellIs" dxfId="1" priority="1" operator="sometimes"><formula>0</formula></cfRule>"#.to_owned(),
        r#"<cfRule type="top10" dxfId="1" priority="1"/>"#.to_owned(),
    ] {
        let workbook = book(unsupported.clone() + &positive(2));
        assert_unpainted(&workbook, &unsupported);
        assert_eq!(
            workbook.conditional_stats().withheld_unsupported,
            2,
            "{unsupported}"
        );
    }
    let scale = r#"<cfRule type="colorScale" priority="1"><colorScale><cfvo type="min"/><cfvo type="max"/><color rgb="FFFF0000"/><color rgb="FF00FF00"/></colorScale></cfRule>"#;
    assert_eq!(
        fills(&book(scale.to_owned() + &positive(2)), &viewport(0.0)),
        red(&[0, 1])
    );
}

#[test]
fn an_unknown_rule_a_cell_reaches_withholds_the_formats_of_certain_rules() {
    let book = |rules: String| {
        workbook(
            &[("A1", "1")],
            &rule("A1", &rules),
            &[&format!("<font><b/></font>{}", fill(RED)), &fill(GREEN)],
        )
    };
    let bold = |workbook: &Workbook| texts(workbook).iter().any(|run| run.0 == "1" && run.2);
    for rules in [
        positive(1) + &time_period(2, false),
        time_period(1, false) + &positive(2),
    ] {
        let workbook = book(rules.clone());
        assert_unpainted(&workbook, &rules);
        assert!(!bold(&workbook), "{rules}");
    }
    let stopping = r#"<cfRule type="cellIs" dxfId="0" priority="1" stopIfTrue="1" operator="greaterThan"><formula>0</formula></cfRule>"#;
    let stopped = book(stopping.to_owned() + &time_period(2, false));
    assert_eq!(fills(&stopped, &viewport(0.0)), red(&[0]));
    assert!(bold(&stopped));
}

/// A solid differential fill naming `foreground` and `background`, each
/// written only when given.
fn solid(foreground: Option<&str>, background: Option<&str>) -> String {
    let side = |name: &str, color: Option<&str>| {
        color.map_or(String::new(), |color| {
            format!(r#"<{name} rgb="FF{}"/>"#, &color[1..])
        })
    };
    format!(
        r#"<fill><patternFill patternType="solid">{}{}</patternFill></fill>"#,
        side("fgColor", foreground),
        side("bgColor", background)
    )
}

#[test]
fn a_conditional_fill_paints_one_color_and_withholds_two_different_ones() {
    for (dxf, shown) in [
        (fill(RED), true),
        (solid(None, Some(RED)), true),
        (solid(Some(RED), None), true),
        (solid(Some(RED), Some(RED)), true),
        (solid(Some(RED), Some(GREEN)), false),
    ] {
        let workbook = workbook(
            &[("A1", "1"), ("A2", "2")],
            &rule(
                "A1:A2",
                r#"<cfRule type="cellIs" dxfId="0" priority="1" operator="greaterThan"><formula>1</formula></cfRule><cfRule type="cellIs" dxfId="1" priority="2" operator="greaterThan"><formula>0</formula></cfRule>"#,
            ),
            &[&format!("<font><b/></font>{dxf}"), &fill(BLUE)],
        );
        let mut expected = vec![((0, 0), BLUE.to_owned())];
        if shown {
            expected.push(((1, 0), RED.to_owned()));
        }
        assert_eq!(fills(&workbook, &viewport(0.0)), expected, "{dxf}");
        let bold: Vec<_> = texts(&workbook)
            .into_iter()
            .map(|(text, _, bold, _)| (text, bold))
            .collect();
        assert_eq!(bold, [("1".into(), false), ("2".into(), true)], "{dxf}");
        assert_eq!(
            workbook.conditional_stats().withheld_fill,
            u64::from(!shown),
            "{dxf}"
        );
    }
}

#[test]
fn a_table_style_with_two_fill_colors_paints_and_saves_as_read_before() {
    let dxf = r#"<dxf><fill><patternFill patternType="solid"><fgColor rgb="FFFF0000"/><bgColor rgb="FF0000FF"/></patternFill></fill></dxf>"#;
    let rels = |body: &str| {
        format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{body}</Relationships>"#
        )
    };
    let link = |id: &str, kind: &str, target: &str| {
        format!(
            r#"<Relationship Id="{id}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{kind}" Target="{target}"/>"#
        )
    };
    let parts: Vec<(String, Vec<u8>)> = [
        (
            "xl/workbook.xml",
            r#"<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#.to_owned(),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            rels(&(link("rId1", "worksheet", "worksheets/sheet1.xml")
                + &link("rId2", "styles", "styles.xml"))),
        ),
        (
            "xl/styles.xml",
            format!(
                r#"<styleSheet><cellXfs count="1"><xf/></cellXfs><dxfs count="1">{dxf}</dxfs><tableStyles count="1"><tableStyle name="Mixed" pivot="0" count="1"><tableStyleElement type="headerRow" dxfId="0"/></tableStyle></tableStyles></styleSheet>"#
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            r#"<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1" t="inlineStr"><is><t>b</t></is></c></row><row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c></row></sheetData></worksheet>"#.to_owned(),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            rels(&link("rId1", "table", "../tables/table1.xml")),
        ),
        (
            "xl/tables/table1.xml",
            r#"<table id="1" name="T" displayName="T" ref="A1:B2"><tableColumns count="2"><tableColumn id="1" name="a"/><tableColumn id="2" name="b"/></tableColumns><tableStyleInfo name="Mixed" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/></table>"#.to_owned(),
        ),
    ]
    .into_iter()
    .map(|(path, xml)| (path.to_owned(), xml.into_bytes()))
    .collect();
    let workbook = Workbook::open(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
    let colors: Vec<String> = workbook
        .display_list(&viewport(0.0))
        .unwrap()
        .commands
        .iter()
        .filter_map(|command| match command {
            DrawCmd::FillRect { color, .. } => Some(color.to_string()),
            _ => None,
        })
        .collect();
    assert!(colors.iter().any(|color| color == "#ff0000"), "{colors:?}");
    assert!(!colors.iter().any(|color| color == "#0000ff"), "{colors:?}");
    let saved = ooxml_opc::unzip_parts(&workbook.save().unwrap()).unwrap();
    let styles = saved
        .iter()
        .find(|(path, _)| path == "xl/styles.xml")
        .map(|(_, bytes)| String::from_utf8_lossy(bytes).into_owned())
        .unwrap();
    assert!(styles.contains(dxf), "{styles}");
}

#[test]
fn a_rule_reading_random_numbers_paints_nothing() {
    let names = r#"<definedNames><definedName name="Draw">RAND()</definedName><definedName name="Later">Draw+1</definedName></definedNames>"#;
    for formula in [
        "SUM(RANDARRAY(1))&gt;=0",
        "RAND()&gt;=0",
        "RANDBETWEEN(1,1)=1",
        "ISERROR(Later)",
    ] {
        let workbook = workbook_named(
            &[("A1", "1")],
            &rule("A1", &expression(1, formula)),
            &[&fill(RED)],
            names,
        );
        assert_unpainted(&workbook, formula);
        assert_eq!(
            workbook.conditional_stats().withheld_volatile,
            1,
            "{formula}"
        );
    }
}

#[test]
fn a_range_rule_over_overlapping_areas_paints_nothing() {
    for kind in [
        r#"type="duplicateValues""#,
        r#"type="top10" rank="1""#,
        r#"type="aboveAverage""#,
    ] {
        let workbook = workbook(
            &[("A1", "1"), ("A2", "2"), ("A3", "3")],
            &rule(
                "A1:A2 A2:A3",
                &format!(r#"<cfRule {kind} dxfId="0" priority="1"/>"#),
            ),
            &[&fill(RED)],
        );
        assert_unpainted(&workbook, kind);
        assert_eq!(
            workbook.conditional_stats().withheld_overlapping,
            3,
            "{kind}"
        );
    }
}

/// The cells withheld for `reason`, read from the counters by name: none
/// when no counter has that name.
fn withheld(workbook: &Workbook, reason: &str) -> u64 {
    let stats = format!("{:?}", workbook.conditional_stats());
    let Some((_, rest)) = stats.split_once(&format!(" withheld_{reason}: ")) else {
        return 0;
    };
    let digits: String = rest.chars().take_while(char::is_ascii_digit).collect();
    digits.parse().unwrap()
}

/// A call to a function the calc engine does not implement, whose error a
/// handler turns into `TRUE` where Excel computes `FALSE`.
const GAP: &str = r#"ISERROR(REGEXTEST("abc","^a"))"#;

fn gap_book(sqref: &str, rules: &str) -> Workbook {
    workbook(
        &[
            ("A1", "1"),
            ("A2", "1"),
            ("A3", "1"),
            ("B1", "abc"),
            ("B2", "abc"),
            ("B3", "abc"),
        ],
        &rule(sqref, rules),
        &[&fill(RED), &fill(GREEN)],
    )
}

#[test]
fn a_rule_the_calc_engine_cannot_evaluate_in_full_paints_nothing() {
    for (sqref, formula, cells) in [
        ("A1", GAP.to_owned(), 1),
        ("A1:A3", GAP.to_owned(), 3),
        ("A1:A3", r#"ISERROR(REGEXTEST(B1,"^a"))"#.to_owned(), 3),
        ("A1:A3", format!("AND({GAP},SUM(1:1)&gt;=0)"), 3),
        ("A1:A3", "ISERROR(SUM('[1]Data'!$B$1:$B$2))".to_owned(), 3),
    ] {
        let workbook = gap_book(sqref, &expression(1, &formula));
        assert_unpainted(&workbook, &formula);
        assert_eq!(withheld(&workbook, "engine_gap"), cells, "{formula}");
    }
}

#[test]
fn a_rule_the_calc_engine_cannot_evaluate_withholds_the_certain_rules_beside_it() {
    for rules in [
        expression(1, GAP) + &positive(2),
        positive(1) + &expression(2, GAP),
    ] {
        let workbook = gap_book("A1", &rules);
        assert_unpainted(&workbook, &rules);
        assert_eq!(withheld(&workbook, "engine_gap"), 1, "{rules}");
    }
    let stopping = r#"<cfRule type="cellIs" dxfId="1" priority="1" stopIfTrue="1" operator="greaterThan"><formula>0</formula></cfRule>"#;
    let stopped = gap_book("A1", &(stopping.to_owned() + &expression(2, GAP)));
    assert_eq!(
        fills(&stopped, &viewport(0.0)),
        [((0, 0), GREEN.to_owned())]
    );
    assert_eq!(withheld(&stopped, "engine_gap"), 0);
}

#[test]
fn a_rule_whose_formula_meets_an_error_value_still_paints() {
    let workbook = gap_book("A1", &expression(1, "ISERROR(1/0)"));
    assert_eq!(fills(&workbook, &viewport(0.0)), red(&[0]));
    assert_eq!(withheld(&workbook, "engine_gap"), 0);
}

/// The cells the whole-range builds of one edit may read: twice the calc
/// engine's cell visits per evaluation.
const WHOLE_RANGE_BUDGET: u64 = 2 * xlsx_calc::eval::MAX_EVALUATION_CELL_VISITS;
/// The rows a budget sheet uses, its last cell being in this one.
const USED_ROWS: u64 = 50_000;

/// The letters of the 0-based column `index`.
fn column(index: u64) -> String {
    let mut letters = Vec::new();
    let mut rest = index + 1;
    while rest > 0 {
        rest -= 1;
        letters.push(char::from(b'A' + (rest % 26) as u8));
        rest /= 26;
    }
    letters.iter().rev().collect()
}

/// The last column of a band from column C whose used rows hold `cells`.
fn band_end(cells: u64) -> String {
    assert_eq!(cells % USED_ROWS, 0);
    column(2 + cells / USED_ROWS - 1)
}

/// The band from column C whose used rows hold `cells`.
fn band(cells: u64) -> String {
    format!("C1:{}{USED_ROWS}", band_end(cells))
}

fn duplicates(priority: u32) -> String {
    format!(r#"<cfRule type="duplicateValues" dxfId="0" priority="{priority}"/>"#)
}

/// A sheet using `USED_ROWS` rows whose duplicate pairs are 7 in A1:A2 and 8
/// in B1:B2, with `column_c` in C1:C2, `extra` and `rules`.
fn budget_book(column_c: [&str; 2], extra: &[(&str, &str)], rules: &[String]) -> Workbook {
    let end = format!("AU{USED_ROWS}");
    let mut values = vec![
        ("A1", "7"),
        ("A2", "7"),
        ("B1", "8"),
        ("B2", "8"),
        ("C1", column_c[0]),
        ("C2", column_c[1]),
        (end.as_str(), "1"),
    ];
    values.extend_from_slice(extra);
    workbook(&values, &rules.concat(), &[&fill(RED), &fill(GREEN)])
}

/// Duplicate rules over A and over B, each reading two cells once built.
fn first_rules() -> Vec<String> {
    vec![
        rule(&format!("A1:A{USED_ROWS}"), &duplicates(1)),
        rule(&format!("B1:B{USED_ROWS}"), &duplicates(2)),
    ]
}

fn sorted_fills(workbook: &Workbook) -> Vec<((u32, u32), String)> {
    let mut painted = fills(workbook, &viewport(0.0));
    painted.sort();
    painted
}

fn red_at(cells: &[(u32, u32)]) -> Vec<((u32, u32), String)> {
    cells.iter().map(|&at| (at, RED.to_owned())).collect()
}

const PAIRS_OF_A_AND_B: [(u32, u32); 4] = [(0, 0), (0, 1), (1, 0), (1, 1)];

/// Cells evaluated rendering `workbook` less those evaluated rendering
/// `without`, the same sheet without the rule past the budget.
fn evaluated_beyond(workbook: &Workbook, without: &Workbook) -> u64 {
    fills(without, &viewport(0.0));
    let beyond =
        workbook.conditional_stats().evaluated_cells - without.conditional_stats().evaluated_cells;
    assert!(beyond > 0);
    beyond
}

/// A direct edit of AV1, a cell no budget rule covers.
fn edit_outside(workbook: &mut Workbook) {
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::new(0, 47),
            "1",
            CalculationOptions::default(),
        )
        .unwrap();
}

#[test]
fn a_duplicate_rule_past_the_edit_budget_paints_nothing() {
    let mut rules = first_rules();
    rules.push(rule(&band(WHOLE_RANGE_BUDGET), &duplicates(3)));
    let mut workbook = budget_book(["9", "9"], &[], &rules);
    assert_eq!(sorted_fills(&workbook), red_at(&PAIRS_OF_A_AND_B));
    assert_eq!(workbook.conditional_stats().aggregate_builds, 2);
    let without = budget_book(["9", "9"], &[], &first_rules());
    let refused = workbook.conditional_stats().withheld_refused;
    assert_eq!(refused, evaluated_beyond(&workbook, &without));

    edit_outside(&mut workbook);
    assert_eq!(sorted_fills(&workbook), red_at(&PAIRS_OF_A_AND_B));
    let stats = workbook.conditional_stats();
    assert_eq!(stats.aggregate_builds, 4);
    assert_eq!(stats.withheld_refused, 2 * refused);
}

#[test]
fn a_count_table_past_the_edit_budget_is_not_built_and_paints_nothing() {
    let range = format!("$C$1:${}${USED_ROWS}", band_end(WHOLE_RANGE_BUDGET));
    let rules = [
        rule(&format!("A1:A{USED_ROWS}"), &duplicates(1)),
        rule(
            "C1:C2",
            &expression(2, &format!("COUNTIF({range},C1)&gt;1")),
        ),
    ];
    let workbook = budget_book(["9", "9"], &[], &rules);
    assert_eq!(sorted_fills(&workbook), red_at(&[(0, 0), (1, 0)]));
    assert_eq!(workbook.conditional_stats().withheld_refused, 2);
    assert_eq!(workbook.conditional_stats().aggregate_builds, 1);
}

#[test]
fn a_top_rule_past_the_edit_budget_paints_nothing() {
    let first = rule(&format!("A1:A{USED_ROWS}"), &duplicates(1));
    let top = rule(
        &band(WHOLE_RANGE_BUDGET),
        r#"<cfRule type="top10" dxfId="0" priority="2" rank="1"/>"#,
    );
    let workbook = budget_book(["9", "3"], &[], &[first.clone(), top]);
    assert_eq!(sorted_fills(&workbook), red_at(&[(0, 0), (1, 0)]));
    assert_eq!(workbook.conditional_stats().aggregate_builds, 1);
    let without = budget_book(["9", "3"], &[], &[first]);
    assert_eq!(
        workbook.conditional_stats().withheld_refused,
        evaluated_beyond(&workbook, &without)
    );
}

/// The duplicate pairs of A, B and C.
const PAIRS_OF_A_B_AND_C: [(u32, u32); 6] = [(0, 0), (0, 1), (0, 2), (1, 0), (1, 1), (1, 2)];

#[test]
fn whole_range_builds_within_the_edit_budget_all_paint() {
    let mut rules = first_rules();
    rules.push(rule(&band(WHOLE_RANGE_BUDGET - USED_ROWS), &duplicates(3)));
    let workbook = budget_book(["9", "9"], &[], &rules);
    assert_eq!(sorted_fills(&workbook), red_at(&PAIRS_OF_A_B_AND_C));
    assert_eq!(workbook.conditional_stats().aggregate_builds, 3);
    assert_eq!(workbook.conditional_stats().withheld_refused, 0);

    let wide = format!(
        "A1:{}{USED_ROWS}",
        column(WHOLE_RANGE_BUDGET / 2 / USED_ROWS - 1)
    );
    let sparse = budget_book(
        ["9", "9"],
        &[],
        &[
            rule(&wide, &duplicates(1)),
            rule(&wide, &duplicates(2)),
            rule(&wide, &duplicates(3)),
        ],
    );
    assert_eq!(sorted_fills(&sparse), red_at(&PAIRS_OF_A_B_AND_C));
    assert_eq!(sparse.conditional_stats().aggregate_builds, 3);
    assert_eq!(sparse.conditional_stats().withheld_refused, 0);
}

#[test]
fn a_rule_without_a_whole_range_build_paints_beside_a_rule_past_the_budget() {
    let mut rules = first_rules();
    rules.push(rule(&band(WHOLE_RANGE_BUDGET), &duplicates(3)));
    rules.push(rule(
        "A3:A4",
        r#"<cfRule type="cellIs" dxfId="1" priority="4" operator="greaterThan"><formula>0</formula></cfRule>"#,
    ));
    let workbook = budget_book(["9", "9"], &[("A3", "5"), ("A4", "6")], &rules);
    let green: Vec<_> = sorted_fills(&workbook)
        .into_iter()
        .filter(|(_, color)| color == GREEN)
        .collect();
    assert_eq!(
        green,
        [((2, 0), GREEN.to_owned()), ((3, 0), GREEN.to_owned())]
    );
}

#[test]
fn each_edit_starts_the_whole_range_budget_again() {
    let mut alone = budget_book(
        ["9", "9"],
        &[],
        &[rule(&band(WHOLE_RANGE_BUDGET), &duplicates(1))],
    );
    assert_eq!(sorted_fills(&alone), red_at(&[(0, 2), (1, 2)]));
    edit_outside(&mut alone);
    assert_eq!(sorted_fills(&alone), red_at(&[(0, 2), (1, 2)]));
    assert_eq!(alone.conditional_stats().aggregate_builds, 2);
}
