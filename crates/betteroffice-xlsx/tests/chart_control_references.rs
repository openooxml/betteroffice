//! A chart reference (`c:f`, plain `xsd:string`) holding a character xml 1.0
//! cannot carry can only come from a file that already wrote one, here as a
//! `&#1;` reference. Such a workbook opens and its chart is kept as authored;
//! an edit that would have to write the reference anew is refused before
//! anything changes, so a save never fails and never writes invalid xml.

use betteroffice_xlsx::{CalculationOptions, CellRef, Op, SheetId, Workbook};

const MAIN: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELS: &str = "http://schemas.openxmlformats.org/package/2006/relationships";
const DOC_RELS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CHART: &str = "xl/charts/chart1.xml";

/// Numbers in B2:B4 of a sheet named `A&#1;B` as the file spells it, and a bar
/// chart on that sheet plotting `'A&#1;B'!$B$2:$B$4`.
fn book() -> Vec<u8> {
    let chart = concat!(
        r#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:idx val="0"/><c:order val="0"/>"#,
        r#"<c:val><c:numRef><c:f>'A&#1;B'!$B$2:$B$4</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="3"/><c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt><c:pt idx="2"><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val>"#,
        r#"</c:ser></c:barChart><c:catAx><c:axId val="1"/></c:catAx></c:plotArea></c:chart></c:chartSpace>"#
    );
    let parts = [
        (
            "[Content_Types].xml",
            r#"<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/drawings/drawing1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/><Override PartName="/xl/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/></Types>"#.to_owned(),
        ),
        (
            "_rels/.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/officeDocument" Target="xl/workbook.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/workbook.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheets><sheet name="A&#1;B" sheetId="1" r:id="rId1"/></sheets></workbook>"#
            ),
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="rId1" Type="{DOC_RELS}/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/worksheets/sheet1.xml",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN}" xmlns:r="{DOC_RELS}"><sheetData><row r="2"><c r="B2"><v>1</v></c></row><row r="3"><c r="B3"><v>2</v></c></row><row r="4"><c r="B4"><v>3</v></c></row></sheetData><drawing r:id="d1"/></worksheet>"#
            ),
        ),
        (
            "xl/worksheets/_rels/sheet1.xml.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="d1" Type="{DOC_RELS}/drawing" Target="../drawings/drawing1.xml"/></Relationships>"#
            ),
        ),
        (
            "xl/drawings/drawing1.xml",
            r#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>12</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame><c:chart r:id="c1"/></xdr:graphicFrame></xdr:twoCellAnchor></xdr:wsDr>"#.to_owned(),
        ),
        (
            "xl/drawings/_rels/drawing1.xml.rels",
            format!(
                r#"<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="{RELS}"><Relationship Id="c1" Type="{DOC_RELS}/chart" Target="../charts/chart1.xml"/></Relationships>"#
            ),
        ),
        (CHART, chart.to_owned()),
    ];
    let parts = parts
        .into_iter()
        .map(|(name, xml)| (name.to_owned(), xml.into_bytes()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).unwrap()
}

fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    ooxml_opc::unzip_parts(bytes)
        .unwrap()
        .into_iter()
        .find(|(path, _)| path == name)
        .map(|(_, bytes)| bytes)
        .unwrap_or_default()
}

fn apply(workbook: &mut Workbook, op: Op) -> betteroffice_xlsx::Result<()> {
    workbook
        .apply_ops(vec![op], CalculationOptions::default())
        .map(|_| ())
}

/// The workbook opens, its chart reads the reference as authored, and saves
/// that leave the chart alone, plain or after a cell edit, keep its bytes.
#[test]
fn an_opened_chart_reference_holding_a_control_character_opens_and_is_kept() {
    let source = book();
    let workbook =
        Workbook::open(&source).unwrap_or_else(|error| panic!("refused at open: {error}"));
    let chart = &workbook.model().sheets[0].charts[0];
    assert_eq!(chart.refs[0].formula, "'A\u{1}B'!$B$2:$B$4");
    assert_eq!(part(&workbook.save().unwrap(), CHART), part(&source, CHART));

    let mut workbook = Workbook::open(&source).unwrap();
    workbook
        .edit_cell(
            SheetId(0),
            CellRef::parse_a1("D20").unwrap(),
            "7",
            CalculationOptions::default(),
        )
        .unwrap();
    let saved = workbook.save().unwrap();
    assert_eq!(part(&saved, CHART), part(&source, CHART));
    Workbook::open(&saved).unwrap();
}

/// A row inserted inside the plotted range would move the reference, which the
/// writer could only write back with the control character, so the edit is
/// refused up front and the workbook and its chart are left as they were.
#[test]
fn an_edit_that_would_rewrite_such_a_reference_is_refused_and_the_chart_kept() {
    let source = book();
    let mut workbook = Workbook::open(&source).unwrap();
    let before = workbook.model().clone();
    let error = apply(
        &mut workbook,
        Op::InsertRows {
            sheet: SheetId(0),
            at: 2,
            count: 1,
        },
    )
    .unwrap_err();
    assert!(error.to_string().contains("xml cannot carry"), "{error}");
    assert_eq!(workbook.model(), &before);
    assert_eq!(part(&workbook.save().unwrap(), CHART), part(&source, CHART));
}

/// Renaming the sheet writes a reference xml can carry; undo puts back the
/// authored one, which the save leaves as the source wrote it.
#[test]
fn undo_puts_back_the_authored_reference_without_refusal() {
    let source = book();
    let mut workbook = Workbook::open(&source).unwrap();
    apply(
        &mut workbook,
        Op::RenameSheet {
            sheet: SheetId(0),
            name: "Data".into(),
        },
    )
    .unwrap();
    let renamed = workbook.save().unwrap();
    let chart = String::from_utf8(part(&renamed, CHART)).unwrap();
    assert!(chart.contains("<c:f>Data!$B$2:$B$4</c:f>"), "{chart}");
    workbook
        .undo(CalculationOptions::default())
        .unwrap_or_else(|error| panic!("undo refused: {error}"));
    assert_eq!(
        workbook.model().sheets[0].charts[0].refs[0].formula,
        "'A\u{1}B'!$B$2:$B$4"
    );
    assert_eq!(part(&workbook.save().unwrap(), CHART), part(&source, CHART));
}

/// A snapshot of such a workbook hydrates, and the hydrated workbook saves its
/// chart as authored.
#[test]
fn a_snapshot_of_such_a_workbook_hydrates_and_saves_the_chart_as_authored() {
    let source = book();
    let workbook = Workbook::open(&source).unwrap();
    let budget = betteroffice_xlsx::SnapshotBudget::new(8, 65536).unwrap();
    let mut encoder =
        betteroffice_xlsx::WorkbookSnapshotEncoder::new(&workbook, None, budget).unwrap();
    let mut builder = betteroffice_xlsx::WorkbookSnapshotBuilder::new();
    while let Some(chunk) = encoder.next(&workbook).unwrap() {
        builder.push(&chunk).unwrap();
    }
    while !builder
        .advance(budget)
        .unwrap_or_else(|error| panic!("hydration refused: {error}"))
        .is_ready()
    {}
    let (restored, _) = builder.finish().unwrap().into_parts();
    assert_eq!(
        restored.model().sheets[0].charts[0].refs[0].formula,
        "'A\u{1}B'!$B$2:$B$4"
    );
    assert_eq!(part(&restored.save().unwrap(), CHART), part(&source, CHART));
}

/// A recovery checkpoint of such a workbook restores, and the restored
/// workbook saves its chart as authored.
#[test]
fn a_recovery_checkpoint_of_such_a_workbook_restores_and_saves_the_chart_as_authored() {
    let source = book();
    let workbook = Workbook::open(&source).unwrap();
    let checkpoint = workbook.recovery_checkpoint(0, None).unwrap();
    let (restored, ..) = Workbook::restore_recovery_checkpoint(&checkpoint)
        .unwrap_or_else(|error| panic!("restore refused: {error}"));
    assert_eq!(restored.model(), workbook.model());
    assert_eq!(part(&restored.save().unwrap(), CHART), part(&source, CHART));
}
