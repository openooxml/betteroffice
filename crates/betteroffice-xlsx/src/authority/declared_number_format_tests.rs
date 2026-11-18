use std::sync::Mutex;

use super::*;
use crate::{CalculationOptions, NumberFormatKind, UpdateEvent, UpdateOrigin, Workbook};
use xlsx_ops::{CapturedFormat, NumberFormatMutation};

const UNAFFECTED_SEED: &str = include_str!("fixtures/previous_version/unaffected_seed.hex");
const CUSTOM_SEED: &str = include_str!("fixtures/previous_version/custom_format_seed.hex");
const CUSTOM_RECOVERY: &str = include_str!("fixtures/previous_version/custom_format_recovery.hex");
const DECLARED_SEED: &str = include_str!("fixtures/previous_version/declared_builtin_seed.hex");
const DECLARED_EDIT: &str = include_str!("fixtures/previous_version/declared_builtin_update.hex");
const DECLARED_EDITED: &str =
    include_str!("fixtures/previous_version/declared_builtin_edited_state.hex");
const PRE_FACET_STATE: &str = include_str!("fixtures/previous_version/pre_facet_state.hex");
const DECLARED_PACKAGE: &str = include_str!("fixtures/declared_builtin_package.hex");
const DECLARED_FINGERPRINT: &str =
    "xf-facets:b9b649ed968b1493edaf8790bc811413a9a24f461aba6c1f589b00ca39f10ae2";
const DECLARED_NUM_FMT: &str = r#"<numFmt numFmtId="2" formatCode="0.000"/>"#;
const MINUS_ACCOUNTING: &str = r#"_-* #,##0_-;\-* #,##0_-;_-* "-"_-;_-@_-"#;
const REDEFINED: [(u16, &str); 5] = [
    (4, "0.000"),
    (10, "0"),
    (14, "0.00"),
    (20, "0.0"),
    (49, "0.00"),
];

fn fixture(hex: &str) -> Vec<u8> {
    let hex = hex.trim();
    (0..hex.len())
        .step_by(2)
        .map(|index| u8::from_str_radix(&hex[index..index + 2], 16).unwrap())
        .collect()
}

fn holds(bytes: &[u8], needle: &[u8]) -> bool {
    bytes.windows(needle.len()).any(|window| window == needle)
}

fn calculation() -> CalculationOptions {
    CalculationOptions {
        now_serial: Some(45_000.25),
    }
}

fn at(a1: &str) -> CellRef {
    CellRef::parse_a1(a1).unwrap()
}

fn text(value: &str) -> CellValue {
    CellValue::Text {
        value: value.into(),
    }
}

fn shown(model: &WorkbookModel, a1: &str) -> String {
    match model.sheets[0].cell(at(a1)) {
        Some(cell) => xlsx_render::display_text(&model.styles, model.date_system, cell),
        None => String::new(),
    }
}

fn value_of(model: &WorkbookModel, a1: &str) -> CellValue {
    let cell = model.sheets[0].cell(at(a1));
    cell.map_or(CellValue::Empty, |cell| cell.value.clone())
}

fn style_of(model: &WorkbookModel, a1: &str) -> Option<u32> {
    model.sheets[0].cell(at(a1))?.style
}

fn format_id(model: &WorkbookModel, a1: &str) -> Option<u16> {
    let style = style_of(model, a1)?;
    model.styles.cell_xfs.get(style as usize)?.num_fmt_id
}

fn schema(book: &Workbook) -> Result<i64, String> {
    book.authority_for_test().schema_version()
}

fn parse(bytes: &[u8]) -> WorkbookModel {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    xlsx_parse::parse_workbook(&parts).unwrap()
}

fn styles_part(bytes: &[u8]) -> String {
    let parts = ooxml_opc::unzip_parts(bytes).unwrap();
    let (_, styles) = parts
        .into_iter()
        .find(|(name, _)| name == "xl/styles.xml")
        .unwrap();
    String::from_utf8(styles).unwrap()
}

fn recovery_checkpoint_model() -> WorkbookModel {
    let mut model = WorkbookModel::default();
    model.sheets.push(Sheet::new("Data"));
    model.sheets.push(Sheet::new("Notes"));
    model.sheets[0].format.default_row_height_pt = Some(16.0);
    model.sheets[0].format.custom_height = true;
    model
}

fn custom_format_model() -> WorkbookModel {
    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    model.styles.num_fmts.push((164, "#,##0.0".into()));
    model
        .styles
        .cell_xfs
        .push(xlsx_model::styles::Xf::default());
    model.styles.cell_xfs.push(xlsx_model::styles::Xf {
        num_fmt_id: Some(164),
        ..xlsx_model::styles::Xf::default()
    });
    for (a1, value) in [("A1", 1234.5), ("B1", -7.25)] {
        model.sheets[0].set_cell(
            at(a1),
            Cell {
                value: CellValue::Number { value },
                style: Some(1),
                ..Cell::default()
            },
        );
    }
    model
}

fn declared_model() -> WorkbookModel {
    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    model.styles.num_fmts.push((2, "0.000".into()));
    model.styles.num_fmts.push((41, MINUS_ACCOUNTING.into()));
    model
        .styles
        .cell_xfs
        .push(xlsx_model::styles::Xf::default());
    for id in [41, 2] {
        model.styles.cell_xfs.push(xlsx_model::styles::Xf {
            num_fmt_id: Some(id),
            ..xlsx_model::styles::Xf::default()
        });
    }
    let cells = [
        ("A1", CellValue::Number { value: -1234.0 }, Some(1)),
        ("A2", CellValue::Number { value: 1.5 }, Some(2)),
        ("A3", text("label"), None),
    ];
    for (a1, value, style) in cells {
        model.sheets[0].set_cell(
            at(a1),
            Cell {
                value,
                style,
                ..Cell::default()
            },
        );
    }
    model
}

fn single_format_model(id: u16, code: &str) -> WorkbookModel {
    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    model.styles.num_fmts.push((id, code.into()));
    model
        .styles
        .cell_xfs
        .push(xlsx_model::styles::Xf::default());
    model.styles.cell_xfs.push(xlsx_model::styles::Xf {
        num_fmt_id: Some(id),
        ..xlsx_model::styles::Xf::default()
    });
    model.sheets[0].set_cell(
        at("A1"),
        Cell {
            value: CellValue::Number { value: 1.0 },
            style: Some(1),
            ..Cell::default()
        },
    );
    model
}

fn styled_number(a1: &str, value: f64) -> Op {
    Op::SetCell {
        sheet: SheetId(0),
        at: at(a1),
        cell: CellState {
            value: CellValue::Number { value },
            style: Some(1),
            ..CellState::default()
        },
    }
}

fn seeded(model: &WorkbookModel) -> WorkbookAuthority {
    WorkbookAuthority::from_source(model, Some(109), &[], None).unwrap()
}

fn adopted(state: &[u8], client_id: u64) -> Workbook {
    let mut book = Workbook::from_model_collaborative(declared_model(), client_id).unwrap();
    let result = book.apply_update_v1(state, calculation());
    assert!(result.as_ref().is_ok_and(|done| done.applied), "{result:?}");
    book
}

fn presets() -> [(&'static str, NumberFormatMutation, u16, f64); 5] {
    [
        ("A1", NumberFormatMutation::Number, 4, 1234.5),
        ("A2", NumberFormatMutation::Percent, 10, 0.125),
        ("A3", NumberFormatMutation::Date, 14, 45_000.0),
        ("A4", NumberFormatMutation::Time, 20, 0.75),
        ("A5", NumberFormatMutation::PlainText, 49, 7.0),
    ]
}

fn formats_model(declared: &[(u16, &str)]) -> WorkbookModel {
    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    for &(id, code) in declared {
        model.styles.num_fmts.push((id, code.into()));
    }
    model
        .styles
        .cell_xfs
        .push(xlsx_model::styles::Xf::default());
    for (a1, _, _, value) in presets() {
        model.sheets[0].set_cell(
            at(a1),
            Cell {
                value: CellValue::Number { value },
                ..Cell::default()
            },
        );
    }
    model
}

fn formatted(declared: &[(u16, &str)]) -> Workbook {
    let mut book = Workbook::from_model(formats_model(declared)).unwrap();
    for (a1, preset, _, _) in presets() {
        let range = CellRange::new(at(a1), at(a1));
        let result = book.set_range_number_format(SheetId(0), range, preset, calculation());
        assert!(result.is_ok(), "{result:?}");
    }
    let range = CellRange::new(at("A6"), at("A6"));
    let plain_text = NumberFormatMutation::PlainText;
    let result = book.set_range_number_format(SheetId(0), range, plain_text, calculation());
    assert!(result.is_ok(), "{result:?}");
    for (a1, input) in [("A6", "0012"), ("A7", "1/15/2024")] {
        let result = book.edit_cell(SheetId(0), at(a1), input, calculation());
        assert!(result.is_ok(), "{result:?}");
    }
    book
}

#[test]
fn unaffected_workbooks_seed_the_previous_versions_bytes() {
    let checkpoint = seeded(&recovery_checkpoint_model());
    assert_eq!(
        checkpoint.encode_state_as_update_v1(),
        fixture(UNAFFECTED_SEED)
    );
    let custom = seeded(&custom_format_model());
    assert_eq!(custom.encode_state_as_update_v1(), fixture(CUSTOM_SEED));
    assert_eq!(
        custom.encode_recovery_state().ok(),
        Some(fixture(CUSTOM_RECOVERY))
    );
}

#[test]
fn a_workbook_declaring_a_builtin_id_seeds_schema_8_under_another_fingerprint() {
    assert!(holds(
        &fixture(DECLARED_SEED),
        DECLARED_FINGERPRINT.as_bytes()
    ));
    let book = Workbook::from_model_collaborative(declared_model(), 1001).unwrap();
    let authority = book.authority_for_test();
    assert_eq!(authority.schema_version(), Ok(8));
    assert_ne!(authority.base.fingerprint, DECLARED_FINGERPRINT);
    let current = authority.base.fingerprint.as_str();
    assert_eq!(authority.state_fingerprint().as_deref(), Some(current));
}

#[test]
fn a_new_replica_adopts_the_previous_versions_state_without_rewriting_it() {
    let mut book = adopted(&fixture(DECLARED_SEED), 2001);
    assert_eq!(schema(&book), Ok(7));
    assert_eq!(shown(book.model(), "A1"), " (1,234)");
    assert_eq!(shown(book.model(), "A2"), "1.50");
    let result = book.apply_update_v1(&fixture(DECLARED_EDIT), calculation());
    assert!(result.as_ref().is_ok_and(|done| done.applied), "{result:?}");
    assert_eq!(shown(book.model(), "A4"), " (5)");
    assert_eq!(schema(&book), Ok(7));
    assert_eq!(book.encode_state_as_update_v1(), fixture(DECLARED_EDITED));
}

#[test]
fn a_model_cloned_from_an_adopted_old_state_seeds_the_workbooks_own_codes() {
    let restored = adopted(&fixture(DECLARED_SEED), 2011);
    let cloned = Workbook::from_model_collaborative(restored.model().clone(), 2012);
    assert!(cloned.is_ok(), "{:?}", cloned.as_ref().err());
    let book = cloned.unwrap();
    let authority = book.authority_for_test();
    assert_eq!(authority.schema_version(), Ok(8));
    let projected = authority.strict_materialize().map(|(model, _)| model);
    assert!(projected.is_ok(), "{:?}", projected.as_ref().err());
    let projected = projected.unwrap();
    let validated = snapshot_validation::validate_for_test(authority, projected.clone());
    assert!(validated.is_ok(), "{validated:?}");
    let recovery = authority.encode_recovery_state();
    assert!(recovery.is_ok(), "{:?}", recovery.as_ref().err());
    let bytes = recovery.unwrap();
    let recovered = WorkbookAuthority::restore_recovery_state(&bytes, 2013, 0, true, true);
    assert!(recovered.is_ok(), "{:?}", recovered.as_ref().err());
    assert_eq!(shown(&projected, "A1"), "-1,234 ");
    assert_eq!(shown(&projected, "A2"), "1.500");
    assert_eq!(shown(book.model(), "A1"), "-1,234 ");
}

#[test]
fn adopting_a_pre_facet_state_of_a_declared_builtin_id_keeps_every_style_index() {
    let mut book = Workbook::open_collaborative(&fixture(DECLARED_PACKAGE), 77).unwrap();
    let result = book.apply_update_v1(&fixture(PRE_FACET_STATE), calculation());
    assert!(result.as_ref().is_ok_and(|done| done.applied), "{result:?}");
    assert_eq!(schema(&book), Ok(7));
    assert_eq!(style_of(book.model(), "A1"), Some(2));
    assert_eq!(style_of(book.model(), "B1"), Some(1));
    assert_eq!(shown(book.model(), "A1"), "7.00");
    assert_eq!(shown(book.model(), "B1"), "8.00");
    let saved = book.save();
    assert!(saved.is_ok(), "{:?}", saved.as_ref().err());
    let saved = saved.unwrap();
    assert!(styles_part(&saved).contains(DECLARED_NUM_FMT));
    let parsed = parse(&saved);
    assert_eq!(parsed.styles.num_fmts, [(2_u16, String::from("0.000"))]);
    assert_eq!(style_of(&parsed, "A1"), Some(2));
    assert_eq!(style_of(&parsed, "B1"), Some(1));
}

#[test]
fn currency_symbols_follow_the_previous_rule_in_every_reading() {
    let restored = adopted(&fixture(DECLARED_SEED), 2021);
    assert_eq!(
        xlsx_calc::currency_symbols(restored.model()),
        Vec::<char>::new()
    );
    assert_eq!(
        xlsx_calc::currency_symbols(&declared_model()),
        Vec::<char>::new()
    );
    let dollars = single_format_model(5, "0.00");
    let legacy = schema_6_authority(&dollars, 2022).materialize().unwrap();
    assert_eq!(xlsx_calc::currency_symbols(&legacy), Vec::<char>::new());
    assert_eq!(xlsx_calc::currency_symbols(&dollars), Vec::<char>::new());
    let euros = single_format_model(0, "[$€-407]0.00");
    assert_eq!(xlsx_calc::currency_symbols(&euros), ['€']);
}

#[test]
fn toolbar_presets_and_typed_entries_mean_the_builtin_codes() {
    let declaring = formatted(&REDEFINED);
    let plain = formatted(&[]);
    for a1 in ["A1", "A2", "A3", "A4", "A5", "A6", "A7"] {
        assert_eq!(
            shown(declaring.model(), a1),
            shown(plain.model(), a1),
            "{a1}"
        );
    }
    assert_eq!(value_of(declaring.model(), "A6"), text("0012"));
    let saved = declaring.save();
    assert!(saved.is_ok(), "{:?}", saved.as_ref().err());
    let saved = parse(&saved.unwrap());
    let authored = presets()
        .into_iter()
        .map(|(a1, _, id, _)| (a1, id))
        .chain([("A6", 49), ("A7", 14)]);
    for (a1, id) in authored {
        assert_ne!(format_id(&saved, a1), Some(id), "{a1}");
        let pattern = xlsx_model::numfmt::builtin_format_code(id).unwrap_or_default();
        let expected = xlsx_model::NumberFormat::Custom {
            pattern: pattern.into(),
        };
        let format = saved.styles.cell_format(style_of(&saved, a1));
        assert_eq!(format.number_format, expected, "{a1}");
    }
}

#[test]
fn a_preset_keeps_a_builtin_id_the_workbook_declares_with_the_builtin_code() {
    let mut book = Workbook::from_model(formats_model(&[(4, "#,##0.00")])).unwrap();
    let range = CellRange::new(at("A1"), at("A1"));
    let number = NumberFormatMutation::Number;
    let result = book.set_range_number_format(SheetId(0), range, number, calculation());
    assert!(result.is_ok(), "{result:?}");
    assert_eq!(format_id(book.model(), "A1"), Some(4));
    assert_eq!(shown(book.model(), "A1"), "1,234.50");
}

#[test]
fn edits_on_an_adopted_old_state_stay_at_schema_7_and_reach_its_peers() {
    let mut book = adopted(&fixture(DECLARED_EDITED), 2031);
    let mut peer = adopted(&fixture(DECLARED_EDITED), 2032);
    let events: Arc<Mutex<Vec<UpdateEvent>>> = Arc::default();
    let sink = events.clone();
    let subscription = book
        .observe_update_v1(move |event| sink.lock().unwrap().push(event))
        .unwrap();
    let edited = book.apply_ops(vec![styled_number("A5", -9.0)], calculation());
    assert!(edited.is_ok(), "{edited:?}");
    let bold = xlsx_ops::StylePatch {
        bold: Some(true),
        ..xlsx_ops::StylePatch::default()
    };
    let range = CellRange::new(at("A1"), at("A2"));
    let patched = book.patch_range_style(SheetId(0), range, bold, calculation());
    assert!(patched.is_ok(), "{patched:?}");
    drop(subscription);
    assert_eq!(schema(&book), Ok(7));
    assert_eq!(shown(book.model(), "A1"), " (1,234)");
    assert_eq!(shown(book.model(), "A2"), "1.50");
    assert_eq!(shown(book.model(), "A5"), " (9)");
    let updates = events
        .lock()
        .unwrap()
        .iter()
        .filter(|event| event.origin == UpdateOrigin::Local && !event.update.is_empty())
        .map(|event| event.update.clone())
        .collect::<Vec<_>>();
    assert!(!updates.is_empty());
    for update in &updates {
        let result = peer.apply_update_v1(update, calculation());
        assert!(result.is_ok(), "{result:?}");
    }
    assert_eq!(schema(&peer), Ok(7));
    assert_eq!(shown(peer.model(), "A1"), " (1,234)");
    assert_eq!(shown(peer.model(), "A2"), "1.50");
    assert_eq!(shown(peer.model(), "A5"), " (9)");
    let model = peer.model();
    let format = model.styles.cell_format(style_of(model, "A1"));
    assert!(format.font.bold);
}

#[test]
fn unaffected_seeds_keep_the_fast_edit_paths() {
    let model = custom_format_model();
    let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 2041).unwrap();
    let count = fast_set_cell_count();
    let edited = authority.apply_ops(&[styled_number("C1", 3.0)], SyncOrigin::User, &model.styles);
    assert!(edited.is_ok(), "{edited:?}");
    assert_eq!(fast_set_cell_count(), count + 1);
    authority.materialize().unwrap();
    let op = Op::SetRangeNumberFormat {
        sheet: SheetId(0),
        range: CellRange::new(at("A1"), at("B1")),
        format: NumberFormatMutation::Percent,
    };
    let ops = [op];
    let prepared = authority.prepare_range_format_sync(&ops, SyncOrigin::User);
    assert!(prepared.is_some());
}

#[test]
fn a_cell_without_a_style_index_shows_the_default_xf_but_calculates_as_general() {
    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        full_precision: false,
        ..WorkbookModel::default()
    };
    model.styles.num_fmts.push((31, "0.0".into()));
    model.styles.cell_xfs.push(xlsx_model::styles::Xf {
        num_fmt_id: Some(31),
        ..xlsx_model::styles::Xf::default()
    });
    model.sheets[0].set_cell(
        at("A1"),
        Cell {
            value: CellValue::Number { value: 45_000.123 },
            ..Cell::default()
        },
    );
    for (a1, formula) in [("B1", "CELL(\"format\",A1)"), ("C1", "A1")] {
        model.sheets[0].set_cell(
            at(a1),
            Cell {
                formula: Some(formula.into()),
                ..Cell::default()
            },
        );
    }
    let mut book = Workbook::from_model(model).unwrap();
    book.recalculate_all(calculation());
    let model = book.model();
    assert_eq!(shown(model, "A1"), "45000.1");
    assert_eq!(value_of(model, "B1"), text("G"));
    assert_eq!(
        value_of(model, "C1"),
        CellValue::Number { value: 45_000.123 }
    );
    assert_eq!(shown(model, "C1"), "45000.1");
}

#[test]
fn undo_and_redo_on_an_adopted_old_state_keep_the_builtin_display() {
    let mut book = adopted(&fixture(DECLARED_EDITED), 2051);
    let edited = book.apply_ops(vec![styled_number("A5", -7.0)], calculation());
    assert!(edited.is_ok(), "{edited:?}");
    assert_eq!(shown(book.model(), "A5"), " (7)");
    let undone = book.undo(calculation());
    assert!(undone.is_ok(), "{undone:?}");
    assert_eq!(shown(book.model(), "A5"), "");
    let redone = book.redo(calculation());
    assert!(redone.is_ok(), "{redone:?}");
    assert_eq!(shown(book.model(), "A5"), " (7)");
    assert_eq!(shown(book.model(), "A4"), " (5)");
    assert_eq!(shown(book.model(), "A1"), " (1,234)");
    let authority = book.authority_for_test();
    assert_eq!(authority.schema_version(), Ok(7));
    let projected = authority.strict_materialize().map(|(model, _)| model);
    assert!(projected.is_ok(), "{:?}", projected.as_ref().err());
    let validated = snapshot_validation::validate_for_test(authority, projected.unwrap());
    assert!(validated.is_ok(), "{validated:?}");
}

#[test]
fn an_unedited_save_of_a_declared_builtin_id_keeps_its_declaration_and_style_indices() {
    let book = Workbook::open_collaborative(&fixture(DECLARED_PACKAGE), 81).unwrap();
    assert_eq!(schema(&book), Ok(8));
    let saved = book.save();
    assert!(saved.is_ok(), "{:?}", saved.as_ref().err());
    let saved = saved.unwrap();
    assert!(styles_part(&saved).contains(DECLARED_NUM_FMT));
    let parsed = parse(&saved);
    assert_eq!(parsed.styles.num_fmts, [(2_u16, String::from("0.000"))]);
    assert_eq!(style_of(&parsed, "A1"), Some(2));
    assert_eq!(style_of(&parsed, "B1"), Some(1));
    assert_eq!(format_id(&parsed, "A1"), Some(2));
    assert_eq!(format_id(&parsed, "B1"), Some(2));
}

#[test]
fn a_same_code_preset_in_a_collaborative_session_saves_under_a_custom_id_as_it_does_now() {
    let model = formats_model(&[(4, "#,##0.00")]);
    let mut book = Workbook::from_model_collaborative(model, 3001).unwrap();
    let range = CellRange::new(at("A1"), at("A1"));
    let number = NumberFormatMutation::Number;
    let result = book.set_range_number_format(SheetId(0), range, number, calculation());
    assert!(result.is_ok(), "{result:?}");
    assert_eq!(shown(book.model(), "A1"), "1,234.50");
    let saved = book.save();
    assert!(saved.is_ok(), "{:?}", saved.as_ref().err());
    let saved = parse(&saved.unwrap());
    assert_eq!(shown(&saved, "A1"), "1,234.50");
    let format = saved.styles.cell_format(style_of(&saved, "A1"));
    let expected = xlsx_model::NumberFormat::Custom {
        pattern: "#,##0.00".into(),
    };
    assert_eq!(format.number_format, expected);
    let id = format_id(&saved, "A1");
    assert!(id.is_some_and(|id| id >= 164), "{id:?}");
    let num_fmts = &saved.styles.num_fmts;
    assert!(num_fmts.contains(&(4, "#,##0.00".into())), "{num_fmts:?}");
    assert!(
        num_fmts.contains(&(id.unwrap(), "#,##0.00".into())),
        "{num_fmts:?}"
    );
}

fn kind_of(book: &Workbook) -> (Option<NumberFormatKind>, Option<String>) {
    let range = CellRange::new(at("A1"), at("A1"));
    let formatting = book.selection_formatting(SheetId(0), range).unwrap();
    (formatting.number_format, formatting.number_format_pattern)
}

#[test]
fn a_redefined_builtin_id_keeps_the_format_kind_of_its_id() {
    let cases = [
        (41, MINUS_ACCOUNTING, NumberFormatKind::Currency),
        (14, "0.00", NumberFormatKind::Date),
    ];
    for (index, (id, code, kind)) in cases.into_iter().enumerate() {
        let expected = (Some(kind), Some(String::from(code)));
        let model = single_format_model(id, code);
        let standalone = Workbook::from_model(model.clone()).unwrap();
        assert_eq!(kind_of(&standalone), expected, "{id}");
        let client_id = 3011 + index as u64;
        let collaborative = Workbook::from_model_collaborative(model, client_id).unwrap();
        assert_eq!(kind_of(&collaborative), expected, "{id}");
    }
}

#[test]
fn a_preset_follows_the_first_of_two_declarations_of_a_builtin_id() {
    let declared = [(4, "#,##0.00"), (4, "0.000")];
    let mut book = Workbook::from_model(formats_model(&declared)).unwrap();
    let range = CellRange::new(at("A1"), at("A1"));
    let number = NumberFormatMutation::Number;
    let result = book.set_range_number_format(SheetId(0), range, number, calculation());
    assert!(result.is_ok(), "{result:?}");
    assert_eq!(format_id(book.model(), "A1"), Some(4));
    assert_eq!(shown(book.model(), "A1"), "1,234.50");
}

fn painter_book(declared: Option<&str>) -> Workbook {
    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    if let Some(code) = declared {
        model.styles.num_fmts.push((31, code.into()));
    }
    model
        .styles
        .cell_xfs
        .push(xlsx_model::styles::Xf::default());
    model.styles.cell_xfs.push(xlsx_model::styles::Xf {
        num_fmt_id: Some(31),
        ..xlsx_model::styles::Xf::default()
    });
    for (a1, style) in [("A1", Some(1)), ("B1", None)] {
        model.sheets[0].set_cell(
            at(a1),
            Cell {
                value: CellValue::Number { value: 45_000.0 },
                style,
                ..Cell::default()
            },
        );
    }
    Workbook::from_model(model).unwrap()
}

fn capture(book: &Workbook) -> CapturedFormat {
    let captured = book.capture_format(SheetId(0), CellRange::new(at("A1"), at("A1")));
    assert!(captured.is_ok(), "{:?}", captured.as_ref().err());
    captured.unwrap()
}

fn paint(book: &mut Workbook, format: CapturedFormat) {
    let range = CellRange::new(at("B1"), at("B1"));
    let result = book.apply_format(SheetId(0), range, format, calculation());
    assert!(result.is_ok(), "{result:?}");
}

#[test]
fn a_painted_undeclared_builtin_id_shows_general_where_the_id_is_declared() {
    let source = painter_book(None);
    let mut target = painter_book(Some("yyyy-mm-dd"));
    assert_eq!(shown(source.model(), "A1"), "45000");
    paint(&mut target, capture(&source));
    assert_eq!(shown(target.model(), "B1"), "45000");
}

#[test]
fn a_painted_declared_builtin_id_shows_the_declared_code_in_its_own_workbook() {
    let mut book = painter_book(Some("yyyy-mm-dd"));
    assert_eq!(shown(book.model(), "A1"), "2023-03-15");
    let format = capture(&book);
    paint(&mut book, format);
    assert_eq!(shown(book.model(), "B1"), "2023-03-15");
}

#[test]
fn a_painted_declared_builtin_id_shows_the_declared_code_in_another_workbook() {
    let source = painter_book(Some("yyyy-mm-dd"));
    let mut target = painter_book(None);
    paint(&mut target, capture(&source));
    assert_eq!(shown(target.model(), "B1"), "2023-03-15");
}

#[test]
fn a_schema_3_state_of_a_declared_builtin_id_upgrades_to_7_and_is_adopted() {
    let mut written = declared_model();
    written.styles.read_builtin_number_formats();
    let update = super::tests::legacy_update(&written, 3, true);
    let authority = super::tests::authority_from_update(&declared_model(), &update, 3021);
    let materialized = authority.strict_materialize().map(|(model, _)| model);
    assert!(materialized.is_ok(), "{:?}", materialized.as_ref().err());
    assert_eq!(shown(&materialized.unwrap(), "A1"), " (1,234)");
    let staged = authority
        .stage_updates_v1(&[Update::EMPTY_V1], None)
        .unwrap();
    assert!(staged.effective);
    authority
        .apply_staged_update_v1(&staged.commit_update)
        .unwrap();
    assert_eq!(authority.schema_version(), Ok(7));
    let upgraded = authority.strict_materialize().map(|(model, _)| model);
    assert!(upgraded.is_ok(), "{:?}", upgraded.as_ref().err());
    assert_eq!(shown(&upgraded.unwrap(), "A1"), " (1,234)");
    let peer = adopted(&authority.encode_state_as_update_v1(), 3022);
    assert_eq!(schema(&peer), Ok(7));
    assert_eq!(shown(peer.model(), "A1"), " (1,234)");
}
