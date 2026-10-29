use std::fs;
use std::path::Path;

use betteroffice_xlsx::{CalculationOptions, Workbook};
use serde::Deserialize;
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
        "version" => json!(workbook.version()),
        "readCells" => wire(workbook.read_cells_json(&args.to_string())),
        "findText" => wire(workbook.find_text_json(&args.to_string())),
        "validateEdits" => wire(workbook.validate_edits_json(&args.to_string())),
        "applyEdits" => wire(workbook.apply_edits_json(&args.to_string())),
        "historyState" => json!(workbook.history_state()),
        "undo" => json!(workbook.undo(calculation).unwrap().applied),
        "redo" => json!(workbook.redo(calculation).unwrap().applied),
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
fn scripted_api_parity_matches_shared_golden() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let parity = root.join("parity/xlsx");
    let script: Script =
        serde_json::from_str(&fs::read_to_string(parity.join("script.json")).unwrap()).unwrap();
    let calculation = CalculationOptions {
        now_serial: Some(script.calculation.now_serial),
    };
    let mut workbook = Workbook::open_recalculated_with_seed(
        &fs::read(root.join(script.fixture)).unwrap(),
        calculation,
        Some(script.calculation.rand_seed),
    )
    .unwrap();
    let mut versions = Vec::<String>::new();
    let mut actual = Vec::new();
    for (index, step) in script.steps.iter().enumerate() {
        let args = map_strings(&step.args, &mut |value| {
            let reference = value
                .strip_prefix("<version:")
                .and_then(|value| value.strip_suffix('>'))
                .filter(|value| {
                    value.starts_with(|character: char| ('1'..='9').contains(&character))
                        && value.bytes().all(|byte| byte.is_ascii_digit())
                });
            match reference {
                Some(reference) => reference
                    .parse::<usize>()
                    .ok()
                    .and_then(|index| versions.get(index - 1))
                    .unwrap_or_else(|| panic!("unknown version reference {value} at step {index}"))
                    .clone(),
                None => value.to_owned(),
            }
        });
        let result = execute(&mut workbook, &step.op, &args, calculation);
        let result = map_strings(&result, &mut |value| {
            if !is_version(value) {
                return value.to_owned();
            }
            let index = match versions.iter().position(|version| version == value) {
                Some(index) => index,
                None => {
                    versions.push(value.to_owned());
                    versions.len() - 1
                }
            };
            format!("<version:{}>", index + 1)
        });
        actual.push(json!({ "op": step.op, "result": result }));
    }
    let expected_path = parity.join("expected.json");
    if std::env::var("BETTEROFFICE_PARITY_WRITE").as_deref() == Ok("1") {
        fs::write(
            &expected_path,
            format!("{}\n", serde_json::to_string_pretty(&actual).unwrap()),
        )
        .unwrap();
    }
    let expected: Vec<Value> =
        serde_json::from_str(&fs::read_to_string(expected_path).unwrap()).unwrap();
    assert_eq!(expected.len(), actual.len(), "parity step count differs");
    for (index, (expected, actual)) in expected.iter().zip(&actual).enumerate() {
        assert!(
            equal_json(expected, actual),
            "parity step {index} ({}) differs\n- expected: {}\n+ actual: {}",
            script.steps[index].op,
            serde_json::to_string_pretty(expected).unwrap(),
            serde_json::to_string_pretty(actual).unwrap(),
        );
    }
}
