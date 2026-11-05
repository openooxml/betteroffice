use std::fs;
use std::path::Path;

use betteroffice_xlsx::{
    CalculationOptions, CapturedFormat, CellAddress, CellRange, CellRef, ChartMutation,
    MutationResult, Op, SheetId, Workbook,
};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct Calculation {
    now_serial: f64,
    rand_seed: u32,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct Step {
    op: String,
    args: Value,
    targets: Option<Vec<String>>,
    #[serde(default)]
    expect_error: bool,
}

impl Step {
    fn includes(&self, target: &str) -> bool {
        self.targets
            .as_ref()
            .is_none_or(|targets| targets.iter().any(|item| item == target))
    }
}

#[derive(Deserialize)]
struct Script {
    fixture: String,
    calculation: Calculation,
    steps: Vec<Step>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct SheetInfo {
    sheet_ids: Vec<String>,
    sheet_names: Vec<String>,
    active_sheet: u32,
    content_width: f32,
    content_height: f32,
    frozen_rows: u32,
    frozen_cols: u32,
    initial_scroll_x: f32,
    initial_scroll_y: f32,
    auto_filter: Option<String>,
}

#[derive(Serialize)]
#[serde(rename_all = "camelCase")]
struct CellEdit {
    a1: String,
    input: String,
    is_formula: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    is_array_formula: Option<bool>,
}

fn map_strings(value: &Value, transform: &mut impl FnMut(&str) -> String) -> Value {
    match value {
        Value::String(value) => Value::String(transform(value)),
        Value::Array(values) => Value::Array(
            values
                .iter()
                .map(|value| map_strings(value, transform))
                .collect(),
        ),
        Value::Object(values) => {
            let mut keys: Vec<_> = values.keys().collect();
            keys.sort();
            Value::Object(
                keys.into_iter()
                    .map(|key| (key.clone(), map_strings(&values[key], transform)))
                    .collect(),
            )
        }
        value => value.clone(),
    }
}

fn is_version(value: &str) -> bool {
    value.rsplit_once('-').is_some_and(|(nonce, changes)| {
        nonce.len() == 32
            && nonce
                .bytes()
                .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
            && !changes.is_empty()
            && changes.bytes().all(|byte| byte.is_ascii_digit())
    })
}

fn version_reference(value: &str) -> Option<usize> {
    value
        .strip_prefix("<version:")?
        .strip_suffix('>')
        .filter(|value| {
            value.starts_with(|character: char| ('1'..='9').contains(&character))
                && value.bytes().all(|byte| byte.is_ascii_digit())
        })?
        .parse()
        .ok()
}

fn normalize_versions(
    value: &Value,
    expected: Option<&Value>,
    versions: &mut Vec<Option<String>>,
) -> Value {
    match value {
        Value::String(value) if is_version(value) => {
            let index = if let Some(index) = versions
                .iter()
                .position(|item| item.as_ref() == Some(value))
            {
                index
            } else {
                let index = expected
                    .and_then(Value::as_str)
                    .and_then(version_reference)
                    .map_or(versions.len(), |index| index - 1);
                versions.resize(versions.len().max(index + 1), None);
                assert!(
                    versions[index].is_none(),
                    "different versions share a golden token"
                );
                versions[index] = Some(value.clone());
                index
            };
            json!(format!("<version:{}>", index + 1))
        }
        Value::Array(values) => Value::Array(
            values
                .iter()
                .enumerate()
                .map(|(index, value)| {
                    normalize_versions(value, expected.and_then(|value| value.get(index)), versions)
                })
                .collect(),
        ),
        Value::Object(values) => Value::Object(
            values
                .iter()
                .map(|(key, value)| {
                    (
                        key.clone(),
                        normalize_versions(
                            value,
                            expected.and_then(|value| value.get(key)),
                            versions,
                        ),
                    )
                })
                .collect(),
        ),
        value => value.clone(),
    }
}

fn mutation_result(workbook: &Workbook, result: MutationResult) -> Value {
    let addresses = |cells: &[CellAddress]| {
        let mut addresses: Vec<_> = cells
            .iter()
            .map(|address| workbook.format_address(*address))
            .collect();
        addresses.sort();
        addresses
    };
    json!({
        "applied": result.applied,
        "changed": addresses(&result.changed),
        "limitedCells": addresses(&result.limited_cells),
    })
}

#[derive(Deserialize)]
struct OpsArgs {
    ops: Vec<Op>,
}

/// The ops path the TypeScript handle takes: `JSON.stringify({ ops })`, decoded with its
/// `bad ops:` message.
fn apply_ops_json(
    workbook: &mut Workbook,
    ops: &Value,
    calculation: CalculationOptions,
) -> Result<Value, String> {
    let args: OpsArgs = serde_json::from_str(&json!({ "ops": ops }).to_string())
        .map_err(|error| format!("bad ops: {error}"))?;
    let result = workbook
        .apply_ops(args.ops, calculation)
        .map_err(|error| error.to_string())?;
    Ok(mutation_result(workbook, result))
}

fn wire_range(range: &Value) -> Result<CellRange, String> {
    CellRange::parse_a1(range.as_str().unwrap()).map_err(|error| format!("bad range: {error}"))
}

fn sheet_arg(args: &Value) -> SheetId {
    SheetId(serde_json::from_value(args["sheet"].clone()).unwrap())
}

/// The calls that can fail, with the message the TypeScript handle throws; `None` for any
/// other op.
fn attempt(
    workbook: &mut Workbook,
    op: &str,
    args: &Value,
    calculation: CalculationOptions,
) -> Option<Result<Value, String>> {
    let engine = |error: betteroffice_xlsx::Error| error.to_string();
    let outcome = match op {
        "applyOps" => apply_ops_json(workbook, &args["ops"], calculation),
        "moveRange" => (|| {
            let source = wire_range(&args["source"])?;
            let destination = CellRef::parse_a1(args["destination"].as_str().unwrap())
                .map_err(|error| format!("bad move destination: {error}"))?;
            let result = workbook
                .move_range(sheet_arg(args), source, destination, calculation)
                .map_err(engine)?;
            Ok(mutation_result(workbook, result))
        })(),
        "setHyperlink" => {
            let hyperlink = &args["hyperlink"];
            let field = |key: &str| hyperlink.get(key).cloned().unwrap_or(Value::Null);
            let op = json!({
                "type": "setHyperlink",
                "sheet": args["sheet"],
                "hyperlink": {
                    "range": hyperlink["range"],
                    "external_target": field("externalTarget"),
                    "location": field("location"),
                    "tooltip": field("tooltip"),
                    "display": field("display"),
                },
            });
            apply_ops_json(workbook, &json!([op]), calculation)
        }
        "removeHyperlink" => {
            let op = json!({ "type": "removeHyperlink", "sheet": args["sheet"], "range": args["range"] });
            apply_ops_json(workbook, &json!([op]), calculation)
        }
        "editCharts" => (|| {
            let edits: Vec<ChartMutation> = serde_json::from_str(&args["edits"].to_string())
                .map_err(|error| format!("bad chart edits: {error}"))?;
            let result = workbook.edit_charts(edits, calculation).map_err(engine)?;
            Ok(mutation_result(workbook, result))
        })(),
        "moveChart" => (|| {
            let result = workbook
                .move_chart(
                    sheet_arg(args),
                    args["chart"].as_str().unwrap(),
                    serde_json::from_value(args["dx"].clone()).unwrap(),
                    serde_json::from_value(args["dy"].clone()).unwrap(),
                    calculation,
                )
                .map_err(engine)?;
            Ok(mutation_result(workbook, result))
        })(),
        "captureFormat" => (|| {
            let range = wire_range(&args["range"])?;
            let format = workbook
                .capture_format(sheet_arg(args), range)
                .map_err(engine)?;
            Ok(serialized(&format))
        })(),
        "applyFormat" => (|| {
            let format: CapturedFormat = serde_json::from_str(&args["format"].to_string())
                .map_err(|error| format!("bad format args: {error}"))?;
            let range = wire_range(&args["range"])?;
            let result = workbook
                .apply_format(sheet_arg(args), range, format, calculation)
                .map_err(engine)?;
            Ok(mutation_result(workbook, result))
        })(),
        _ => return None,
    };
    Some(outcome)
}

fn serialized(value: &impl Serialize) -> Value {
    serde_json::from_str(&serde_json::to_string(value).unwrap()).unwrap()
}

fn cell_edit(workbook: &Workbook, sheet: SheetId, cell: betteroffice_xlsx::CellEdit) -> CellEdit {
    CellEdit {
        a1: cell.cell.to_a1(),
        input: cell.input,
        is_formula: cell.is_formula,
        is_array_formula: workbook
            .is_array_formula(sheet, cell.cell)
            .unwrap()
            .then_some(true),
    }
}

fn execute(
    workbook: &mut Workbook,
    op: &str,
    args: &Value,
    calculation: CalculationOptions,
) -> Value {
    let wire = |result: betteroffice_xlsx::Result<String>| {
        let result = result.unwrap_or_else(|error| panic!("parity op {op}: {error}"));
        serde_json::from_str(&result).unwrap()
    };
    match op {
        "sheetNames" => json!(workbook.sheet_info().unwrap().sheet_names),
        "sheetInfo" => {
            let info = workbook.sheet_info().unwrap();
            serialized(&SheetInfo {
                sheet_ids: info.sheet_ids,
                sheet_names: info.sheet_names,
                active_sheet: info.active_sheet.0,
                content_width: info.content_width,
                content_height: info.content_height,
                frozen_rows: info.frozen_rows,
                frozen_cols: info.frozen_cols,
                initial_scroll_x: info.initial_scroll_x,
                initial_scroll_y: info.initial_scroll_y,
                auto_filter: workbook.auto_filter(info.active_sheet).unwrap(),
            })
        }
        "listDefinedNames" => serialized(
            &workbook
                .list_defined_names(&serde_json::from_value(args.clone()).unwrap())
                .unwrap(),
        ),
        "conditionalFormats" => serialized(
            &workbook
                .conditional_formats(SheetId(
                    serde_json::from_value(args["sheet"].clone()).unwrap(),
                ))
                .unwrap(),
        ),
        "calculationStatus" => serialized(&workbook.calculation_status()),
        "cell" => {
            let sheet = SheetId(serde_json::from_value(args["sheet"].clone()).unwrap());
            let cell = workbook
                .cell(
                    sheet,
                    CellRef::new(
                        serde_json::from_value(args["row"].clone()).unwrap(),
                        serde_json::from_value(args["col"].clone()).unwrap(),
                    ),
                )
                .unwrap();
            serialized(&cell_edit(workbook, sheet, cell))
        }
        "rangeCells" => {
            let sheet = SheetId(serde_json::from_value(args["sheet"].clone()).unwrap());
            let range = CellRange::parse_a1(args["range"].as_str().unwrap()).unwrap();
            let cells = workbook
                .range_cells(sheet, range)
                .unwrap()
                .into_iter()
                .map(|row| {
                    row.into_iter()
                        .map(|cell| cell_edit(workbook, sheet, cell))
                        .collect::<Vec<_>>()
                })
                .collect::<Vec<_>>();
            serialized(&cells)
        }
        "searchText" => {
            let info = workbook.sheet_info().unwrap();
            let options = &args["options"];
            let matches = workbook
                .search_text(
                    args["query"].as_str().unwrap(),
                    options["caseSensitive"].as_bool().unwrap_or(false),
                    options.get("limit").map(|limit| {
                        serde_json::from_value::<u32>(limit.clone()).unwrap() as usize
                    }),
                )
                .into_iter()
                .map(|result| {
                    let sheet = result.address.sheet.0 as usize;
                    json!({
                        "sheet": sheet,
                        "sheetId": info.sheet_ids[sheet],
                        "sheetName": info.sheet_names[sheet],
                        "row": result.address.cell.row,
                        "col": result.address.cell.col,
                        "a1": result.address.cell.to_a1(),
                        "text": result.text,
                    })
                })
                .collect::<Vec<_>>();
            serialized(&matches)
        }
        "selectionFormatting" => serialized(
            &workbook
                .selection_formatting(
                    SheetId(serde_json::from_value(args["sheet"].clone()).unwrap()),
                    CellRange::parse_a1(args["range"].as_str().unwrap()).unwrap(),
                )
                .unwrap(),
        ),
        "charts" => serialized(
            &workbook
                .charts(SheetId(
                    serde_json::from_value(args["sheet"].clone()).unwrap(),
                ))
                .unwrap(),
        ),
        "version" => json!(workbook.version()),
        "readCells" => wire(workbook.read_cells_json(&args.to_string())),
        "findText" => wire(workbook.find_text_json(&args.to_string())),
        "validateEdits" => wire(workbook.validate_edits_json(&args.to_string())),
        "applyEdits" => wire(workbook.apply_edits_json(&args.to_string())),
        "historyState" => serialized(&workbook.history_state()),
        "undo" => {
            let result = workbook.undo(calculation).unwrap();
            mutation_result(workbook, result)
        }
        "redo" => {
            let result = workbook.redo(calculation).unwrap();
            mutation_result(workbook, result)
        }
        "exportStructured" => wire(workbook.export_structured_json(&args.to_string())),
        "exportMarkdown" => wire(workbook.export_markdown_json(
            &args["options"].to_string(),
            &args["markdownOptions"].to_string(),
        )),
        "save" => {
            let bytes = workbook.save().unwrap();
            json!({ "sha256": format!("{:x}", Sha256::digest(&bytes)), "byteLength": bytes.len() })
        }
        _ => match attempt(workbook, op, args, calculation) {
            Some(result) => result.unwrap_or_else(|error| panic!("parity op {op}: {error}")),
            None => panic!("unknown parity op {op}"),
        },
    }
}

fn equal_json(expected: &Value, actual: &Value) -> bool {
    match (expected, actual) {
        (Value::Number(left), Value::Number(right)) => left.as_f64() == right.as_f64(),
        (Value::Array(left), Value::Array(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|(left, right)| equal_json(left, right))
        }
        (Value::Object(left), Value::Object(right)) => {
            left.len() == right.len()
                && left.iter().all(|(key, value)| {
                    right.get(key).is_some_and(|other| equal_json(value, other))
                })
        }
        _ => expected == actual,
    }
}

#[test]
fn step_targets_default_to_all_and_filter_explicit_lists() {
    let default: Step = serde_json::from_value(json!({ "op": "version", "args": {} })).unwrap();
    assert!(default.includes("rust"));
    let included: Step = serde_json::from_value(json!({
        "op": "cell", "args": {}, "targets": ["ts", "rust"],
    }))
    .unwrap();
    assert!(included.includes("rust"));
    assert!(!included.includes("python"));
}

#[test]
fn version_normalization_preserves_golden_tokens_after_skipped_steps() {
    let value = json!({ "version": "0123456789abcdef0123456789abcdef-0" });
    let expected = json!({ "version": "<version:3>" });
    let mut versions = Vec::new();
    assert_eq!(
        normalize_versions(&value, Some(&expected), &mut versions),
        expected
    );
    assert_eq!(versions[2].as_deref(), value["version"].as_str());
    assert_eq!(
        normalize_versions(&value, Some(&expected), &mut versions),
        expected
    );
}

#[test]
#[should_panic(expected = "unknown parity op unsupported")]
fn included_unsupported_operations_fail() {
    let mut model = betteroffice_xlsx::WorkbookModel::default();
    model.sheets.push(betteroffice_xlsx::Sheet::new("Data"));
    let mut workbook = Workbook::from_model(model).unwrap();
    execute(
        &mut workbook,
        "unsupported",
        &json!({}),
        CalculationOptions::default(),
    );
}

#[test]
fn scripted_api_parity_matches_shared_golden() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let parity = root.join("parity/xlsx");
    let script: Script =
        serde_json::from_str(&fs::read_to_string(parity.join("script.json")).unwrap()).unwrap();
    let expected: Vec<Value> =
        serde_json::from_str(&fs::read_to_string(parity.join("expected.json")).unwrap()).unwrap();
    assert_eq!(
        expected.len(),
        script.steps.len(),
        "golden step count differs"
    );
    let calculation = CalculationOptions {
        now_serial: Some(script.calculation.now_serial),
    };
    let mut workbook = Workbook::open_recalculated_with_seed(
        &fs::read(root.join(script.fixture)).unwrap(),
        calculation,
        Some(script.calculation.rand_seed),
    )
    .unwrap();
    let mut versions = Vec::<Option<String>>::new();
    let mut actual = Vec::new();
    for (index, step) in script.steps.iter().enumerate() {
        if !step.includes("rust") {
            continue;
        }
        let args = map_strings(&step.args, &mut |value| match version_reference(value) {
            Some(reference) => versions
                .get(reference - 1)
                .and_then(Option::as_ref)
                .unwrap_or_else(|| panic!("unknown version reference {value} at step {index}"))
                .clone(),
            None => value.to_owned(),
        });
        let fixture = || fs::read(root.join(args["fixture"].as_str().unwrap())).unwrap();
        let result = if step.op == "open" {
            workbook = Workbook::open_recalculated_with_seed(
                &fixture(),
                calculation,
                Some(script.calculation.rand_seed),
            )
            .unwrap();
            Value::Null
        } else if step.op == "openCollaborative" {
            let client_id = serde_json::from_value(args["clientId"].clone()).unwrap();
            workbook = Workbook::open_collaborative(&fixture(), client_id).unwrap();
            Value::Null
        } else if step.expect_error {
            match attempt(&mut workbook, &step.op, &args, calculation) {
                Some(Err(error)) => json!({ "error": error }),
                Some(Ok(_)) => panic!("parity step {index} ({}) did not fail", step.op),
                None => panic!("unknown parity op {}", step.op),
            }
        } else {
            execute(&mut workbook, &step.op, &args, calculation)
        };
        let result = normalize_versions(&result, expected[index].get("result"), &mut versions);
        actual.push((index, json!({ "op": step.op, "result": result })));
    }
    for (index, actual) in &actual {
        let expected = &expected[*index];
        assert!(
            equal_json(expected, actual),
            "parity step {index} ({}) differs\n- expected: {}\n+ actual: {}",
            script.steps[*index].op,
            serde_json::to_string_pretty(expected).unwrap(),
            serde_json::to_string_pretty(actual).unwrap(),
        );
    }
}
