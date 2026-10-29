use std::fs;
use std::path::Path;

use betteroffice_xlsx::{
    CalculationOptions, CellAddress, CellRef, MutationResult, SheetId, Workbook,
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
struct Step {
    op: String,
    args: Value,
    targets: Option<Vec<String>>,
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

fn serialized(value: &impl Serialize) -> Value {
    serde_json::from_str(&serde_json::to_string(value).unwrap()).unwrap()
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
        "sheetInfo" => serialized(&workbook.sheet_info().unwrap()),
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
        "cell" => serialized(
            &workbook
                .cell(
                    SheetId(serde_json::from_value(args["sheet"].clone()).unwrap()),
                    CellRef::new(
                        serde_json::from_value(args["row"].clone()).unwrap(),
                        serde_json::from_value(args["col"].clone()).unwrap(),
                    ),
                )
                .unwrap(),
        ),
        "version" => json!(workbook.version()),
        "readCells" => wire(workbook.read_cells_json(&args.to_string())),
        "findText" => wire(workbook.find_text_json(&args.to_string())),
        "validateEdits" => wire(workbook.validate_edits_json(&args.to_string())),
        "applyEdits" => wire(workbook.apply_edits_json(&args.to_string())),
        "applyOps" => {
            let result = workbook
                .apply_ops(
                    serde_json::from_value(args["ops"].clone()).unwrap(),
                    calculation,
                )
                .unwrap();
            mutation_result(workbook, result)
        }
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
        _ => panic!("unknown parity op {op}"),
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
        let result = if step.op == "open" {
            workbook = Workbook::open_recalculated_with_seed(
                &fs::read(root.join(args["fixture"].as_str().unwrap())).unwrap(),
                calculation,
                Some(script.calculation.rand_seed),
            )
            .unwrap();
            Value::Null
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
