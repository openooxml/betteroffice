use std::fs;
use std::path::Path;

use betteroffice_pptx::{
    EditCtx, EditRefusal, EditRequest, FindRequest, PptxExportOptions, PptxExportResult,
    Presentation, ReadRequest, StorySnapshot, TextParagraphDraft, export_outcome_json,
    outcome_json,
};
use serde::{Deserialize, Serialize};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

#[derive(Deserialize)]
struct Step {
    op: String,
    args: Value,
}

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct Script {
    fixture: String,
    client_id: u64,
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
    value.rsplit_once('-').is_some_and(|(nonce, epoch)| {
        nonce.len() == 16
            && nonce
                .bytes()
                .all(|byte| byte.is_ascii_digit() || (b'a'..=b'f').contains(&byte))
            && !epoch.is_empty()
            && epoch.bytes().all(|byte| byte.is_ascii_digit())
    })
}

fn outcome<T: Serialize>(result: Result<T, EditRefusal>) -> Value {
    serde_json::from_str(&outcome_json(&result).unwrap()).unwrap()
}

fn export_outcome<T: Serialize>(result: PptxExportResult<T>) -> Value {
    serde_json::from_str(&export_outcome_json(&result).unwrap()).unwrap()
}

fn story_result(story: StorySnapshot) -> Value {
    json!({
        "id": story.id,
        "length": story.length,
        "text": story.plain_text(),
        "paragraphs": story.paragraphs.iter().map(|paragraph| json!({
            "id": paragraph.id,
            "alignment": paragraph.alignment,
            "level": paragraph.level,
            "text": paragraph.runs.iter().map(|run| run.text.as_str()).collect::<String>(),
            "runs": paragraph.runs.iter().map(|run| json!({
                "text": run.text,
                "style": {
                    "bold": run.style.bold,
                    "italic": run.style.italic,
                    "fontSizePt": run.style.font_size_pt,
                    "color": run.style.color,
                    "fontFamily": run.style.font_family,
                    "underline": run.style.underline,
                    "spacingPt": run.style.spacing_pt,
                    "baselinePct": run.style.baseline_pct,
                    "kernPt": run.style.kern_pt,
                    "caps": run.style.caps,
                },
            })).collect::<Vec<_>>(),
        })).collect::<Vec<_>>(),
    })
}

fn execute(presentation: &Presentation, op: &str, args: &Value) -> Value {
    let context = EditCtx::local("wasm");
    match op {
        "version" => json!(presentation.version()),
        "readContent" => {
            let request: ReadRequest = serde_json::from_value(args.clone()).unwrap();
            outcome(presentation.read_content(&request).unwrap())
        }
        "findText" => {
            let request: FindRequest = serde_json::from_value(args.clone()).unwrap();
            outcome(presentation.find_text(&request).unwrap())
        }
        "validateEdits" => {
            let request: EditRequest = serde_json::from_value(args.clone()).unwrap();
            outcome(presentation.validate_edits(&request).unwrap())
        }
        "applyEdits" => {
            let request: EditRequest = serde_json::from_value(args.clone()).unwrap();
            outcome(presentation.apply_edits(&request).unwrap())
        }
        "exportStructured" => {
            let options: PptxExportOptions = serde_json::from_value(args.clone()).unwrap();
            export_outcome(presentation.export_structured(&options).unwrap())
        }
        "exportMarkdown" => {
            let options: PptxExportOptions = serde_json::from_value(args.clone()).unwrap();
            export_outcome(presentation.export_markdown(&options).unwrap())
        }
        "historyStep" => {
            let redo = match args["direction"].as_str().unwrap() {
                "undo" => false,
                "redo" => true,
                direction => panic!("unknown history direction {direction}"),
            };
            json!(presentation.history_step(redo).unwrap())
        }
        "duplicateSlide" => json!(
            presentation
                .duplicate_slide(
                    &context,
                    args["slideId"].as_str().unwrap(),
                    serde_json::from_value(args["toIndex"].clone()).unwrap(),
                )
                .unwrap()
        ),
        "bringShapeToFront" => json!(
            presentation
                .bring_shape_to_front(
                    &context,
                    args["slideId"].as_str().unwrap(),
                    args["shapeId"].as_str().unwrap(),
                )
                .unwrap()
        ),
        "addComment" => {
            let comment = &args["comment"];
            json!(
                presentation
                    .add_comment(
                        &context,
                        args["slideId"].as_str().unwrap(),
                        comment["author"].as_str().unwrap(),
                        comment["initials"].as_str().unwrap(),
                        comment["text"].as_str().unwrap(),
                        comment["created"].as_str().unwrap(),
                        comment["xEmu"].as_i64().unwrap(),
                        comment["yEmu"].as_i64().unwrap(),
                    )
                    .unwrap()
            )
        }
        "setCommentPosition" => json!(
            presentation
                .set_comment_position(
                    &context,
                    args["commentId"].as_str().unwrap(),
                    args["position"]["xEmu"].as_i64().unwrap(),
                    args["position"]["yEmu"].as_i64().unwrap(),
                )
                .unwrap()
        ),
        "setStoryParagraphs" => {
            let paragraphs: Vec<TextParagraphDraft> =
                serde_json::from_value(args["paragraphs"].clone()).unwrap();
            story_result(
                presentation
                    .set_story_paragraphs(&context, args["storyId"].as_str().unwrap(), &paragraphs)
                    .unwrap(),
            )
        }
        "readCommandContent" => {
            let request: ReadRequest = serde_json::from_value(args.clone()).unwrap();
            outcome(presentation.read_command_content(&request).unwrap())
        }
        "save" => {
            let bytes = presentation.save().unwrap();
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
    let parity = root.join("parity/pptx");
    let script: Script =
        serde_json::from_str(&fs::read_to_string(parity.join("script.json")).unwrap()).unwrap();
    let presentation = Presentation::open_collaborative(
        &fs::read(root.join(&script.fixture)).unwrap(),
        script.client_id,
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
        let result = execute(&presentation, &step.op, &args);
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
