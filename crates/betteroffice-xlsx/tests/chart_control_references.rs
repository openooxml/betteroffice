//! A chart reference (`c:f`, plain `xsd:string`) holding a character xml 1.0
//! cannot carry can only come from a file that already wrote one, here as a
//! `&#1;` reference. Such a workbook opens and its chart is kept as authored;
//! an edit that would have to write the reference anew is refused before
//! anything changes, so a save never fails and never writes invalid xml.

use betteroffice_xlsx::{CalculationOptions, CellRef, Op, SheetId, Workbook};

include!("fixtures/chart_control_book.rs");

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
    assert!(error.to_string().contains("cannot be rewritten"), "{error}");
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
