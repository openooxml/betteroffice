use betteroffice_xlsx::{
    CalculationOptions, CellRange, CellRef, DisplayList, DrawCmd, NumberFormatMutation, Op,
    RenderError, SheetChart, SheetId, Viewport, Workbook,
};
use ooxml_drawingml::chart::ChartSpace;
use xlsx_parse::{chart_counters, preserved_chart_space, reset_chart_counters};

const FIXTURE: &[u8] = include_bytes!("../../../packages/xlsx/test-fixtures/charts.xlsx");
const VALUE: CellRef = CellRef {
    row: 1,
    col: 1,
    abs_row: false,
    abs_col: false,
};

fn viewport() -> Viewport {
    Viewport {
        x: 0.0,
        y: 0.0,
        width: 300.0,
        height: 400.0,
    }
}

struct Fixture {
    workbook: Workbook,
    parts: Vec<(String, Vec<u8>)>,
}

impl Fixture {
    fn new() -> Self {
        let mut workbook = Workbook::open(FIXTURE).unwrap();
        workbook
            .apply_ops(
                vec![Op::AddSheet {
                    index: 1,
                    name: "Other".to_owned(),
                }],
                CalculationOptions::default(),
            )
            .unwrap();
        Self {
            workbook,
            parts: ooxml_opc::unzip_parts(FIXTURE).unwrap(),
        }
    }

    fn reference_space(&self, chart: &SheetChart) -> ChartSpace {
        let part = self
            .parts
            .iter()
            .find(|(path, _)| path == &chart.part)
            .unwrap();
        let model = self.workbook.model();
        let owner = &model.sheet(SheetId(0)).unwrap().name;
        preserved_chart_space(&part.1, model, owner, &model.styles.theme).unwrap()
    }

    fn render(&self, counters: (u64, u64)) -> DisplayList {
        reset_chart_counters();
        let actual = self
            .workbook
            .display_list_for(SheetId(0), &viewport())
            .unwrap();
        assert_eq!(chart_counters(), counters);
        assert_eq!(actual.charts.len(), 1);
        assert!(!actual.charts[0].placeholder);
        let reference = xlsx_render::build_display_list_with_charts(
            self.workbook.model(),
            SheetId(0),
            &viewport(),
            |chart| Ok::<_, RenderError>(self.reference_space(chart)),
        )
        .unwrap();
        assert_eq!(
            serde_json::to_vec(&actual).unwrap(),
            serde_json::to_vec(&reference).unwrap(),
        );
        actual
    }

    fn edit(&mut self, sheet: SheetId, at: CellRef, input: &str) {
        self.workbook
            .edit_cell(sheet, at, input, CalculationOptions::default())
            .unwrap();
    }

    fn apply(&mut self, op: Op) {
        self.workbook
            .apply_ops(vec![op], CalculationOptions::default())
            .unwrap();
    }

    fn undo(&mut self) {
        assert!(
            self.workbook
                .undo(CalculationOptions::default())
                .unwrap()
                .applied
        );
    }

    fn redo(&mut self) {
        assert!(
            self.workbook
                .redo(CalculationOptions::default())
                .unwrap()
                .applied
        );
    }
}

fn chart_output(frame: &DisplayList) -> Vec<u8> {
    let commands = frame
        .commands
        .iter()
        .filter(|command| {
            matches!(
                command,
                DrawCmd::Path { .. } | DrawCmd::Text { chart: true, .. }
            )
        })
        .collect::<Vec<_>>();
    serde_json::to_vec(&(&frame.charts, commands)).unwrap()
}

fn rename() -> Op {
    Op::RenameSheet {
        sheet: SheetId(0),
        name: "Renamed".to_owned(),
    }
}

fn number_format() -> Op {
    Op::SetRangeNumberFormat {
        sheet: SheetId(0),
        range: CellRange::new(VALUE, VALUE),
        format: NumberFormatMutation::Custom {
            pattern: "0.00%".to_owned(),
        },
    }
}

#[test]
fn warm_viewport_reuses_chart_plan_and_space() {
    let fixture = Fixture::new();
    let first = fixture.render((1, 1));
    let second = fixture.render((0, 0));
    assert_eq!(
        serde_json::to_vec(&first).unwrap(),
        serde_json::to_vec(&second).unwrap(),
    );
}

#[test]
fn unrelated_edits_reuse_chart_space_on_both_sheets() {
    let mut fixture = Fixture::new();
    let before = chart_output(&fixture.render((1, 1)));
    for sheet in [SheetId(0), SheetId(1)] {
        fixture.edit(sheet, CellRef::new(6, 0), "77");
        assert_eq!(chart_output(&fixture.render((0, 0))), before);
    }
}

#[test]
fn referenced_value_edit_redecodes_chart_space() {
    let mut fixture = Fixture::new();
    let before = chart_output(&fixture.render((1, 1)));
    fixture.edit(SheetId(0), VALUE, "96");
    let after = fixture.render((0, 1));
    assert_ne!(chart_output(&after), before);
    let chart = &fixture.workbook.sheet(SheetId(0)).unwrap().charts[0];
    assert_eq!(
        fixture.reference_space(chart).plot_groups[0].series[0].values[0],
        96.0,
    );
    assert_eq!(chart_output(&fixture.render((0, 0))), chart_output(&after));
}

#[test]
fn referenced_sheet_rename_and_undo_match_uncached_chart() {
    let mut fixture = Fixture::new();
    fixture.edit(SheetId(0), VALUE, "96");
    let before = chart_output(&fixture.render((1, 1)));
    fixture.apply(rename());
    let renamed = chart_output(&fixture.render((0, 1)));
    assert_ne!(renamed, before);
    fixture.undo();
    assert_eq!(chart_output(&fixture.render((0, 1))), before);
}

#[test]
fn referenced_edit_undo_redo_match_uncached_chart() {
    let mut fixture = Fixture::new();
    let before = chart_output(&fixture.render((1, 1)));
    fixture.edit(SheetId(0), VALUE, "96");
    let edited = chart_output(&fixture.render((0, 1)));
    assert_ne!(edited, before);
    fixture.undo();
    assert_eq!(chart_output(&fixture.render((0, 1))), before);
    fixture.redo();
    assert_eq!(chart_output(&fixture.render((0, 1))), edited);
}

#[test]
fn referenced_number_format_edit_reuses_chart_space() {
    let mut fixture = Fixture::new();
    let before = chart_output(&fixture.render((1, 1)));
    fixture.apply(number_format());
    assert_eq!(chart_output(&fixture.render((0, 0))), before);
}

#[test]
fn mixed_edits_match_uncached_chart_after_every_step() {
    enum Step {
        Edit(SheetId, CellRef, &'static str),
        Apply(Op),
        Undo,
        Redo,
    }

    let mut fixture = Fixture::new();
    fixture.render((1, 1));
    for (step, counters) in [
        (Step::Edit(SheetId(0), CellRef::new(6, 0), "77"), (0, 0)),
        (Step::Edit(SheetId(0), VALUE, "96"), (0, 1)),
        (
            Step::Edit(SheetId(1), CellRef::new(0, 0), "Other value"),
            (0, 0),
        ),
        (Step::Apply(number_format()), (0, 0)),
        (Step::Apply(rename()), (0, 1)),
        (Step::Undo, (0, 1)),
        (Step::Undo, (0, 0)),
        (Step::Edit(SheetId(0), CellRef::new(1, 0), "Q5"), (0, 1)),
        (Step::Undo, (0, 1)),
        (Step::Redo, (0, 1)),
    ] {
        match step {
            Step::Edit(sheet, at, input) => fixture.edit(sheet, at, input),
            Step::Apply(op) => fixture.apply(op),
            Step::Undo => fixture.undo(),
            Step::Redo => fixture.redo(),
        }
        fixture.render(counters);
    }
}
