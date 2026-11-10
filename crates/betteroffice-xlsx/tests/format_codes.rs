//! Number format codes are `ST_Xstring` in styles (`numFmt@formatCode`) and in
//! chart caches (`c:formatCode`): read decoded and written the way Excel
//! writes them, so every format a cell can hold saves as valid XML 1.0 and
//! reads back unchanged.

use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, NumberFormatMutation, Op, SheetId, Workbook,
};
use xlsx_model::styles::FormatCode;

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STYLES: &str = "xl/styles.xml";
const CHART: &str = "xl/charts/chart1.xml";

fn relationships(entries: &[(&str, &str, &str)]) -> String {
    let mut xml =
        format!(r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}">"#);
    for (id, kind, target) in entries {
        xml.push_str(&format!(
            r#"<Relationship Id="{id}" Type="{DOC_RELS}/{kind}" Target="{target}"/>"#
        ));
    }
    xml + "</Relationships>"
}

/// Numbers in Data!B2:B4 styled with custom format 164, whose `formatCode`
/// attribute is `style_code` as the file spells it, and a bar chart plotting
/// them through a numeric cache whose `c:formatCode` is `chart_code` as the
/// file spells it.
fn book(style_code: &str, chart_code: &str) -> Vec<u8> {
    let chart = format!(
        concat!(
            r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/>"#,
            r#"<c:val><c:numRef><c:f>Data!$B$2:$B$4</c:f><c:numCache><c:formatCode>{chart_code}</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>"#,
            r#"</c:ser></c:barChart><c:catAx><c:axId val="1"/></c:catAx></c:plotArea></c:chart></c:chartSpace>"#
        ),
        chart_code = chart_code
    );
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            relationships(&[("rId1", "officeDocument", "xl/workbook.xml")]),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            relationships(&[
                ("rId1", "worksheet", "worksheets/sheet1.xml"),
                ("rId2", "styles", "styles.xml"),
            ]),
        ),
        (
            STYLES,
            format!(
                concat!(
                    r#"<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="{MAIN}"><numFmts count="1"><numFmt numFmtId="164" formatCode="{style_code}"/></numFmts>"#,
                    r#"<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>"#,
                    r#"<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>"#,
                    r#"<cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#
                ),
                MAIN = MAIN,
                style_code = style_code
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData><row r="2"><c r="B2" s="1"><v>1</v></c></row><row r="3"><c r="B3" s="1"><v>2</v></c></row><row r="4"><c r="B4" s="1"><v>3</v></c></row></sheetData><drawing r:id="d1"/></worksheet>"#
            ),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationships(&[("d1", "drawing", "../drawings/drawing1.xml")]),
        ),
        (
            "xl/drawings/drawing1.xml",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><c:chart r:id="c1"/></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>"#.to_owned(),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            relationships(&[("c1", "chart", "../charts/chart1.xml")]),
        ),
        (CHART, chart),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn part(bytes: &[u8], name: &str) -> String {
    String::from_utf8(
        ooxml_opc::unzip_parts(bytes)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == name)
            .map(|(_, bytes)| bytes)
            .unwrap_or_default(),
    )
    .unwrap()
}

/// The number format code cell `a1` displays with.
fn format_at(workbook: &Workbook, a1: &str) -> String {
    let model = workbook.model();
    let sheet = &model.sheets[0];
    match model
        .styles
        .resolved_format(sheet.effective_style(CellRef::parse_a1(a1).unwrap()))
        .number_format
    {
        FormatCode::Custom(code) => code.to_owned(),
        FormatCode::Builtin(id) => xlsx_model::numfmt::builtin_format_code(id)
            .unwrap_or("General")
            .to_owned(),
    }
}

/// `xml` holds only characters and character references XML 1.0 allows.
fn assert_xml10(xml: &str) {
    let legal = |code: u32| {
        matches!(code, 0x9 | 0xA | 0xD)
            || (0x20..=0xD7FF).contains(&code)
            || (0xE000..=0xFFFD).contains(&code)
            || (0x10000..=0x10FFFF).contains(&code)
    };
    assert!(
        xml.chars().all(|ch| legal(ch as u32)),
        "raw illegal character in {xml}"
    );
    let mut rest = xml;
    while let Some(at) = rest.find("&#") {
        let reference = &rest[at + 2..];
        let end = reference.find(';').unwrap();
        let digits = &reference[..end];
        let code = match digits.strip_prefix('x') {
            Some(hex) => u32::from_str_radix(hex, 16).unwrap(),
            None => digits.parse().unwrap(),
        };
        assert!(legal(code), "illegal reference &#{digits}; in {xml}");
        rest = &reference[end..];
    }
}

fn set_format(workbook: &mut Workbook, a1: &str, pattern: &str) {
    let at = CellRef::parse_a1(a1).unwrap();
    workbook
        .set_range_number_format(
            SheetId(0),
            CellRange::new(at, at),
            NumberFormatMutation::Custom {
                pattern: pattern.into(),
            },
            CalculationOptions::default(),
        )
        .unwrap();
}

fn insert_row(workbook: &mut Workbook) {
    workbook
        .apply_ops(
            vec![Op::InsertRows {
                sheet: SheetId(0),
                at: 2,
                count: 1,
            }],
            CalculationOptions::default(),
        )
        .unwrap();
}

/// Excel writes a line break in a custom number format as `_x000a_`. It reads
/// as the line break, and a save that leaves the format alone keeps its bytes.
#[test]
fn an_excel_line_break_in_a_number_format_reads_decoded_and_keeps_its_bytes() {
    let source = book(r"mmm\_x000a_yyyy", "General");
    let workbook = Workbook::open(&source).unwrap();
    assert_eq!(format_at(&workbook, "B2"), "mmm\\\nyyyy");
    assert_eq!(
        part(&workbook.save().unwrap(), STYLES),
        part(&source, STYLES)
    );

    let mut workbook = Workbook::open(&source).unwrap();
    set_format(&mut workbook, "B3", "0.0");
    let saved = workbook.save().unwrap();
    let styles = part(&saved, STYLES);
    assert!(
        styles.contains(r#"formatCode="mmm\_x000a_yyyy""#),
        "{styles}"
    );
    assert_xml10(&styles);
    let reopened = Workbook::open(&saved).unwrap();
    assert_eq!(format_at(&reopened, "B2"), "mmm\\\nyyyy");
    assert_eq!(format_at(&reopened, "B3"), "0.0");
}

/// Any number format set through the API saves as valid XML 1.0 in styles and
/// in the chart cache it refreshes, written as Excel writes `ST_Xstring`, and
/// reads back unchanged. Nothing is refused.
#[test]
fn any_number_format_saves_as_valid_xml_and_reads_back() {
    for (pattern, attribute, element) in [
        ("0\u{1}0", "0_x0001_0", "0_x0001_0"),
        ("0\t0\n0\r0", "0_x0009_0_x000A_0_x000D_0", "0\t0\n0_x000D_0"),
        ("0_x0041_0", "0_x005F_x0041_0", "0_x005F_x0041_0"),
        ("0_X0041_0", "0_x005F_X0041_0", "0_x005F_X0041_0"),
        ("0\u{0}0", "0_x0000_0", "0_x0000_0"),
        ("0\u{fffe}", "0_xFFFE_", "0_xFFFE_"),
        (
            r#"#,##0" & <a>""#,
            "#,##0&quot; &amp; &lt;a&gt;&quot;",
            "#,##0&quot; &amp; &lt;a&gt;&quot;",
        ),
        ("mmm\\\nyyyy", r"mmm\_x000A_yyyy", "mmm\\\nyyyy"),
    ] {
        let mut workbook = Workbook::open(&book("0.00", "0.00")).unwrap();
        set_format(&mut workbook, "B2", pattern);
        let saved = workbook
            .save()
            .unwrap_or_else(|error| panic!("{pattern:?}: {error}"));
        let styles = part(&saved, STYLES);
        assert!(
            styles.contains(&format!(r#"formatCode="{attribute}""#)),
            "{pattern:?}: {styles}"
        );
        assert_xml10(&styles);
        let chart = part(&saved, CHART);
        assert!(
            chart.contains(&format!("<c:formatCode>{element}</c:formatCode>")),
            "{pattern:?}: {chart}"
        );
        assert_xml10(&chart);
        let reopened = Workbook::open(&saved).unwrap();
        assert_eq!(format_at(&reopened, "B2"), pattern);
        assert_eq!(
            part(&reopened.save().unwrap(), STYLES),
            styles,
            "{pattern:?}"
        );
    }
}

/// A literal `_xHHHH_` in a format code is written with a leading `_x005F_`,
/// in styles and in a chart cache, and reads back as the literal; a bare
/// `_xHHHH_` in a file is the escaped character, as Excel reads it.
#[test]
fn a_literal_escape_in_a_format_code_survives_save_and_reopen() {
    let source = book("0_x005F_x0041_0", "0&quot;_x005F_x0041_&quot;");
    let mut workbook = Workbook::open(&source).unwrap();
    assert_eq!(format_at(&workbook, "B2"), "0_x0041_0");
    insert_row(&mut workbook);
    let saved = workbook.save().unwrap();
    assert!(
        part(&saved, STYLES).contains(r#"formatCode="0_x005F_x0041_0""#),
        "{}",
        part(&saved, STYLES)
    );
    let chart = part(&saved, CHART);
    assert!(chart.contains("<c:f>Data!$B$2:$B$5</c:f>"), "{chart}");
    assert!(
        chart.contains("<c:formatCode>0&quot;_x005F_x0041_&quot;</c:formatCode>"),
        "{chart}"
    );
    let mut reopened = Workbook::open(&saved).unwrap();
    assert_eq!(format_at(&reopened, "B2"), "0_x0041_0");
    insert_row(&mut reopened);
    let again = reopened.save().unwrap();
    assert!(
        part(&again, CHART).contains("<c:formatCode>0&quot;_x005F_x0041_&quot;</c:formatCode>"),
        "{}",
        part(&again, CHART)
    );

    let mut bare = Workbook::open(&book("0_x0041_0", "0&quot;_x0041_&quot;")).unwrap();
    assert_eq!(format_at(&bare, "B2"), "0A0");
    insert_row(&mut bare);
    let chart = part(&bare.save().unwrap(), CHART);
    assert!(
        chart.contains("<c:formatCode>0&quot;A&quot;</c:formatCode>"),
        "{chart}"
    );
}
