#![cfg(feature = "raster")]

use std::fs;
use std::path::{Path, PathBuf};

use betteroffice_docx::{
    BlockContent, Document, ExportOptions, HeaderFooter, InlineNode, Paragraph, ParagraphContent,
    PngTarget, Receipt, RenderOptions, Run, RunContent, SaveOptions, SaveReport, Section, Table,
    TableCell, TableRow, get_paragraph_text,
};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};

fn includes(step: &Value, target: &str) -> bool {
    step.get("targets")
        .and_then(Value::as_array)
        .is_none_or(|targets| targets.iter().any(|item| item == target))
}

fn join_text<'a>(parts: impl Iterator<Item = &'a str>) -> String {
    parts
        .filter(|part| !part.is_empty())
        .collect::<Vec<_>>()
        .join("\n")
}

fn text_of(value: &Value) -> &str {
    value["text"].as_str().unwrap()
}

fn texts(paragraphs: &[Value]) -> String {
    join_text(paragraphs.iter().map(text_of))
}

fn run_text(run: &Run, separators: bool) -> String {
    let mut text = String::new();
    for content in &run.content {
        match content {
            RunContent::Text { text: value, .. } => text.push_str(value),
            RunContent::Tab if separators => text.push('\t'),
            RunContent::Break { break_type, .. } if separators => {
                text.push(if break_type.as_deref() == Some("page") {
                    '\u{000c}'
                } else {
                    '\n'
                });
            }
            _ => {}
        }
    }
    text
}

fn paragraph_runs(paragraph: &Paragraph) -> Vec<(&Run, bool)> {
    let mut runs = Vec::new();
    for content in &paragraph.content {
        let ParagraphContent::Inline(node) = content else {
            continue;
        };
        match node {
            InlineNode::Run(run) => runs.push((run, true)),
            InlineNode::Hyperlink(link) => {
                for child in &link.children {
                    if let InlineNode::Run(run) = child {
                        runs.push((run, false));
                    }
                }
            }
            InlineNode::SimpleField(field) => {
                runs.extend(field.content.iter().map(|run| (run, false)));
            }
            InlineNode::ComplexField(field) => {
                runs.extend(field.field_result.iter().map(|run| (run, false)));
            }
            _ => {}
        }
    }
    runs
}

fn run_result(run: &Run, separators: bool) -> Value {
    let formatting = run.formatting.as_ref();
    json!({
        "text": run_text(run, separators),
        "bold": formatting.and_then(|format| format.bold),
        "italic": formatting.and_then(|format| format.italic),
        "underline": formatting
            .and_then(|format| format.underline.as_ref())
            .map(|underline| underline.style.clone()),
        "strike": formatting.and_then(|format| format.strike),
        "fontSize": formatting
            .and_then(|format| format.font_size)
            .map(|size| size / 2.0),
        "fontFamily": formatting
            .and_then(|format| format.font_family.as_ref())
            .and_then(|family| family.ascii.clone().or_else(|| family.h_ansi.clone())),
        "color": formatting
            .and_then(|format| format.color.as_ref())
            .and_then(|color| color.rgb.clone()),
        "styleId": formatting.and_then(|format| format.style_id.clone()),
    })
}

fn paragraph_result(paragraph: &Paragraph) -> Value {
    let formatting = paragraph.formatting.as_ref();
    let mut runs = Vec::new();
    for (run, separators) in paragraph_runs(paragraph) {
        runs.push(run_result(run, separators));
    }
    json!({
        "id": paragraph.para_id,
        "style": formatting.and_then(|format| format.style_id.clone()),
        "alignment": formatting.and_then(|format| format.alignment.clone()),
        "runs": runs,
        "text": get_paragraph_text(paragraph),
    })
}

fn paragraph_results(paragraphs: Vec<&Paragraph>) -> Value {
    Value::Array(paragraphs.into_iter().map(paragraph_result).collect())
}

fn paragraph_ids(paragraphs: Vec<&Paragraph>) -> Value {
    let mut ids = Vec::new();
    for paragraph in paragraphs {
        ids.push(json!(paragraph.para_id));
    }
    Value::Array(ids)
}

fn collect_paragraphs<'a>(blocks: &'a [BlockContent], output: &mut Vec<&'a Paragraph>) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => output.push(paragraph),
            BlockContent::Table(table) => {
                for row in &table.rows {
                    for cell in &row.cells {
                        collect_paragraphs(&cell.content, output);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => collect_paragraphs(&sdt.content, output),
            BlockContent::RawXml(_) => {}
        }
    }
}

fn collect_tables<'a>(blocks: &'a [BlockContent], output: &mut Vec<&'a Table>) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(_) | BlockContent::RawXml(_) => {}
            BlockContent::Table(table) => {
                output.push(table);
                for row in &table.rows {
                    for cell in &row.cells {
                        collect_tables(&cell.content, output);
                    }
                }
            }
            BlockContent::BlockSdt(sdt) => collect_tables(&sdt.content, output),
        }
    }
}

fn story_paragraphs<'a>(blocks: &'a [BlockContent], output: &mut Vec<&'a Paragraph>) {
    for block in blocks {
        match block {
            BlockContent::Paragraph(paragraph) => output.push(paragraph),
            BlockContent::BlockSdt(sdt) => story_paragraphs(&sdt.content, output),
            BlockContent::Table(_) | BlockContent::RawXml(_) => {}
        }
    }
}

fn excepted(paragraph: &Paragraph, ids: &Value) -> bool {
    let Some(id) = paragraph.para_id.as_deref() else {
        return false;
    };
    ids.as_array().is_some_and(|ids| {
        ids.iter()
            .filter_map(Value::as_str)
            .any(|item| item.eq_ignore_ascii_case(id))
    })
}

fn paragraphs_of(blocks: &[BlockContent]) -> Vec<Value> {
    let mut found = Vec::new();
    collect_paragraphs(blocks, &mut found);
    found.into_iter().map(paragraph_result).collect()
}

fn cell_result(cell: &TableCell) -> Value {
    let paragraphs = paragraphs_of(&cell.content);
    let mut tables = Vec::new();
    collect_tables(&cell.content, &mut tables);
    let tables: Vec<Value> = tables.into_iter().map(table_result).collect();
    json!({ "paragraphs": paragraphs, "tables": tables, "text": texts(&paragraphs) })
}

fn row_result(row: &TableRow) -> Value {
    let cells: Vec<Value> = row.cells.iter().map(cell_result).collect();
    let text = cells.iter().map(text_of).collect::<Vec<_>>().join("\t");
    json!({ "cells": cells, "text": text })
}

fn table_result(table: &Table) -> Value {
    let rows: Vec<Value> = table.rows.iter().map(row_result).collect();
    let text = join_text(rows.iter().map(text_of));
    json!({ "rows": rows, "columnWidths": table.column_widths, "text": text })
}

fn run_highlights(paragraphs: Vec<&Paragraph>) -> Value {
    let mut results = Vec::new();
    for paragraph in paragraphs {
        let mut highlights = Vec::new();
        for (run, _) in paragraph_runs(paragraph) {
            let formatting = run.formatting.as_ref();
            highlights.push(formatting.and_then(|format| format.highlight.clone()));
        }
        results.push(json!({ "id": paragraph.para_id, "highlights": highlights }));
    }
    Value::Array(results)
}

fn cell_merges(table: &Table) -> Value {
    let mut rows = Vec::new();
    for row in &table.rows {
        let mut cells = Vec::new();
        for cell in &row.cells {
            let formatting = cell.formatting.as_ref();
            cells.push(json!({
                "gridSpan": formatting.and_then(|format| format.grid_span),
                "vMerge": formatting.and_then(|format| format.v_merge.clone()),
            }));
        }
        rows.push(Value::Array(cells));
    }
    json!({ "rows": rows })
}

fn section_result(section: &Section) -> Value {
    let properties = &section.properties;
    let paragraphs = paragraphs_of(&section.content);
    json!({
        "id": section.id.clone().or_else(|| properties.section_id.clone()),
        "pageWidth": properties.page_width,
        "pageHeight": properties.page_height,
        "orientation": properties.orientation,
        "marginTop": properties.margin_top,
        "marginBottom": properties.margin_bottom,
        "marginLeft": properties.margin_left,
        "marginRight": properties.margin_right,
        "text": texts(&paragraphs),
        "paragraphs": paragraphs,
    })
}

fn header_footer_result((rel_id, story): &(String, HeaderFooter)) -> Value {
    let paragraphs = paragraphs_of(&story.content);
    json!({
        "relId": rel_id,
        "kind": story.hdr_ftr_type,
        "text": texts(&paragraphs),
        "paragraphs": paragraphs,
    })
}

fn header_footer_results(entries: &[(String, HeaderFooter)]) -> Value {
    Value::Array(entries.iter().map(header_footer_result).collect())
}

fn document_text(document: &Document) -> String {
    let mut texts = Vec::new();
    for paragraph in document.paragraphs() {
        texts.push(get_paragraph_text(paragraph));
    }
    join_text(texts.iter().map(String::as_str))
}

fn receipt_result(receipt: Receipt) -> Value {
    let range = receipt.range;
    json!({
        "story": range.as_ref().map(|range| range.start.story.clone()),
        "start": range.as_ref().map(|range| range.start.offset),
        "end": range.as_ref().map(|range| range.end.offset),
        "revisionIds": receipt.revision_ids,
    })
}

fn report_result(report: &SaveReport) -> Value {
    json!({
        "keptUnchanged": report.kept_unchanged.as_ref().map(|kept| json!({
            "structure": kept.structure,
            "paragraphs": kept.paragraphs,
        })),
        "structureDropped": report.structure_dropped,
        "parts": parts_of(&report.bytes),
    })
}

fn parts_of(bytes: &[u8]) -> Vec<Value> {
    let mut parts = ooxml_opc::unzip_parts(bytes).unwrap();
    parts.sort_by(|left, right| left.0.cmp(&right.0));
    parts
        .iter()
        .map(|(name, data)| {
            json!({
                "name": name,
                "sha256": format!("{:x}", Sha256::digest(data)),
                "byteLength": data.len(),
            })
        })
        .collect()
}

fn read_golden(path: &Path, script_sha256: &str) -> Vec<Value> {
    let golden: Value = serde_json::from_str(&fs::read_to_string(path).unwrap()).unwrap();
    assert_eq!(
        golden["scriptSha256"].as_str(),
        Some(script_sha256),
        "{} was written for another script.json",
        path.display()
    );
    golden["results"].as_array().unwrap().clone()
}

fn png_size(bytes: &[u8]) -> (u32, u32) {
    (
        u32::from_be_bytes(bytes[16..20].try_into().unwrap()),
        u32::from_be_bytes(bytes[20..24].try_into().unwrap()),
    )
}

struct Runner {
    root: PathBuf,
    now: String,
    document: Document,
    saved: Option<Vec<u8>>,
}

impl Runner {
    fn reopen(&mut self, bytes: &[u8]) -> Value {
        self.document = Document::open(bytes).unwrap();
        self.saved = None;
        Value::Null
    }

    fn save_options(&self) -> SaveOptions {
        SaveOptions {
            now: self.now.clone(),
            update_modified_date: false,
            modified_by: None,
        }
    }

    fn execute(&mut self, op: &str, args: &Value) -> Value {
        let document = &self.document;
        match op {
            "open" => {
                let path = self.root.join(args["fixture"].as_str().unwrap());
                let bytes = fs::read(path).unwrap();
                self.reopen(&bytes)
            }
            "openSaved" => {
                let bytes = self.saved.clone().expect("openSaved needs a save first");
                self.reopen(&bytes)
            }
            "text" => json!(document_text(document)),
            "structure" => {
                let structure = document.structure();
                json!({
                    "bodyParagraphs": structure.body_paragraphs,
                    "bodyTables": structure.body_tables,
                    "sections": structure.sections,
                    "headers": structure.headers,
                    "footers": structure.footers,
                    "footnotes": structure.footnotes,
                    "endnotes": structure.endnotes,
                })
            }
            "paragraphIds" => paragraph_ids(document.paragraphs()),
            "paragraphs" => paragraph_results(document.paragraphs()),
            "paragraph" => match args["index"].as_u64() {
                Some(index) => paragraph_result(document.paragraphs()[index as usize]),
                None => document
                    .paragraph(args["id"].as_str().unwrap())
                    .map_or(Value::Null, paragraph_result),
            },
            "runHighlights" => run_highlights(document.paragraphs()),
            "tables" => Value::Array(document.tables().into_iter().map(table_result).collect()),
            "cellMerges" => Value::Array(document.tables().into_iter().map(cell_merges).collect()),
            "sections" => Value::Array(document.sections().iter().map(section_result).collect()),
            "headers" => header_footer_results(document.headers()),
            "footers" => header_footer_results(document.footers()),
            "warnings" => json!(document.model().warnings),
            "listContentControls" => {
                let options = serde_json::from_value(args.clone()).unwrap();
                json!(document.list_content_controls(&options).unwrap())
            }
            "findContentControls" => {
                let query = serde_json::from_value(args["query"].clone()).unwrap();
                let options = serde_json::from_value(args["options"].clone()).unwrap();
                json!(document.find_content_controls(&query, &options).unwrap())
            }
            "exportStructured" => {
                let options: ExportOptions = serde_json::from_value(args.clone()).unwrap();
                json!(document.export_structured(&options).unwrap())
            }
            "exportMarkdown" => {
                let options: ExportOptions = serde_json::from_value(args.clone()).unwrap();
                json!(document.export_markdown(&options).unwrap())
            }
            "renderPng" => {
                let target = PngTarget {
                    page: args["page"].as_u64().unwrap() as usize,
                };
                let page = document
                    .render_png(target, &RenderOptions::default())
                    .unwrap();
                let (width, height) = png_size(&page.bytes);
                json!({ "width": width, "height": height, "skippedPictures": page.skipped_images })
            }
            "replaceParagraphText" => {
                let para_id = args["paraId"].as_str().unwrap();
                let text = args["text"].as_str().unwrap();
                let receipt = self.document.replace_paragraph_text(para_id, text);
                receipt_result(receipt.unwrap())
            }
            "replaceParagraphTextAt" => {
                let index = args["index"].as_u64().unwrap() as usize;
                let text = args["text"].as_str().unwrap();
                let receipt = self.document.replace_paragraph_text_at(index, text);
                receipt_result(receipt.unwrap())
            }
            "readEdited" => {
                let mut story = Vec::new();
                story_paragraphs(&document.body().content, &mut story);
                let texts: Vec<String> = story
                    .into_iter()
                    .filter(|paragraph| !excepted(paragraph, &args["exceptParaIds"]))
                    .map(get_paragraph_text)
                    .collect();
                json!({
                    "paragraph": document
                        .paragraph(args["paraId"].as_str().unwrap())
                        .map(get_paragraph_text),
                    "story": join_text(texts.iter().map(String::as_str)),
                })
            }
            "save" => {
                let bytes = document.save_with_options(self.save_options()).unwrap();
                let result = json!({ "parts": parts_of(&bytes) });
                self.saved = Some(bytes);
                result
            }
            "saveReport" => {
                let report = document
                    .save_report_with_options(self.save_options())
                    .unwrap();
                let result = report_result(&report);
                self.saved = Some(report.bytes);
                result
            }
            _ => panic!("unknown parity op {op}"),
        }
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
    assert!(includes(&json!({ "op": "text" }), "rust"));
    let listed = json!({ "op": "text", "targets": ["ts", "rust"] });
    assert!(includes(&listed, "rust"));
    assert!(!includes(&listed, "python"));
}

#[test]
fn shared_and_engine_only_scripted_steps_match_their_goldens() {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../..");
    let parity = root.join("parity/docx");
    let script_bytes = fs::read(parity.join("script.json")).unwrap();
    let script_sha256 = format!("{:x}", Sha256::digest(&script_bytes));
    let script: Value = serde_json::from_slice(&script_bytes).unwrap();
    let expected = read_golden(&parity.join("expected.json"), &script_sha256);
    let steps = script["steps"].as_array().unwrap();
    assert_eq!(expected.len(), steps.len(), "golden step count differs");
    let fixture = fs::read(root.join(script["fixture"].as_str().unwrap())).unwrap();
    let mut runner = Runner {
        document: Document::open(&fixture).unwrap(),
        now: script["now"].as_str().unwrap().to_owned(),
        saved: None,
        root,
    };
    let mut shared = Vec::new();
    for (index, step) in steps.iter().enumerate() {
        if !includes(step, "rust") {
            continue;
        }
        let op = step["op"].as_str().unwrap();
        let result = runner.execute(op, &step["args"]);
        shared.push((index, json!({ "op": op, "result": result })));
    }
    let engine: Vec<Value> = script["engineSteps"]
        .as_array()
        .unwrap()
        .iter()
        .map(|step| {
            let op = step["op"].as_str().unwrap();
            json!({ "op": op, "result": runner.execute(op, &step["args"]) })
        })
        .collect();
    let engine_path = parity.join("engine-expected.json");
    if std::env::var("BETTEROFFICE_PARITY_WRITE").as_deref() == Ok("1") {
        assert!(
            std::env::var_os("CI").is_none(),
            "refusing to write a parity golden while CI is set"
        );
        let golden = json!({ "scriptSha256": script_sha256, "results": engine });
        let mut text = serde_json::to_string_pretty(&golden).unwrap();
        text.push('\n');
        fs::write(&engine_path, text).unwrap();
    }
    for (index, actual) in &shared {
        let op = actual["op"].as_str().unwrap();
        assert!(
            equal_json(&expected[*index], actual),
            "parity step {index} ({op}) differs\n- expected: {}\n+ actual: {}",
            serde_json::to_string_pretty(&expected[*index]).unwrap(),
            serde_json::to_string_pretty(actual).unwrap(),
        );
    }
    let engine_expected = read_golden(&engine_path, &script_sha256);
    assert_eq!(
        engine_expected.len(),
        engine.len(),
        "engine golden step count differs"
    );
    for (index, (expected, actual)) in engine_expected.iter().zip(&engine).enumerate() {
        assert!(
            equal_json(expected, actual),
            "engine step {index} differs\n- expected: {}\n+ actual: {}",
            serde_json::to_string_pretty(expected).unwrap(),
            serde_json::to_string_pretty(actual).unwrap(),
        );
    }
}
