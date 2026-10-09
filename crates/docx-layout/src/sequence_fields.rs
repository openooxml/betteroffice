//! SEQ field numbering.
//!
//! A SEQ field counts: it names a sequence and shows that sequence's next
//! number, the way Word numbers figure and table captions. The result Word
//! cached goes stale as soon as a caption is added or moved, so
//! [`number_sequence_fields`] recomputes it from the fields before it in
//! reading order. It reads nothing but the instructions of the fields it is
//! given, and evaluates no other field.
//!
//! Only instructions it understands in full are numbered. A sequence with a
//! chapter reset (`\s`), a bookmark argument, an unusable reset, an unknown
//! switch or a field nested in another field keeps all of its cached results,
//! and so does a single field with a format other than `ARABIC`, `ROMAN` or
//! `ALPHABETIC` among its switches, a numeric or date picture, or a
//! `w:fldLock` (a locked field still counts).

use std::collections::{HashMap, HashSet};

use crate::types::{FieldRun, LayoutBlock, ParagraphBlock, Run, ShapeBlock};

/// Largest number spelled in letters or Roman numerals, whose length grows
/// with the number.
const MAX_SPELLED: i64 = 32_767;

/// Bytes of results one numbering pass generates; fields past it keep their
/// cached results, so spelled numbers cannot grow the document's text without
/// bound.
const MAX_GENERATED_BYTES: usize = 1 << 20;

#[derive(Clone, Copy)]
enum Step {
    Next,
    Repeat,
    Reset(i64),
}

#[derive(Clone, Copy)]
enum Format {
    Arabic,
    Roman { upper: bool },
    Alphabetic { upper: bool },
}

enum Parsed {
    /// Not a SEQ field, or one without a sequence name.
    Skip,
    /// A field that leaves its sequence with the cached results.
    Opaque(String),
    Seq(Seq),
}

struct Seq {
    name: String,
    step: Step,
    hidden: bool,
    format: Format,
    /// A numbering format switch, which shows a hidden (`\h`) result.
    named_format: bool,
    keeps_cached: bool,
}

/// Recomputes the result of every SEQ field in `blocks` in reading order:
/// block by block, table cells row by row, and a text box or shape where its
/// block stands.
pub fn number_sequence_fields(blocks: &mut [LayoutBlock]) {
    number_sequence_fields_with_opaque(blocks, std::iter::empty());
}

/// Numbers SEQ fields while keeping opaque sequences' cached results.
pub fn number_sequence_fields_with_opaque(
    blocks: &mut [LayoutBlock],
    opaque_sequences: impl IntoIterator<Item = String>,
) {
    let mut fields = Vec::new();
    // Sequences with a field nested in another field, which Word counts and
    // the blocks don't show.
    let mut nested = opaque_sequences.into_iter().collect();
    collect_blocks(blocks, &mut fields, &mut nested);
    let parsed: Vec<_> = fields.iter().map(|field| parse(field)).collect();
    let opaque: HashSet<&str> = parsed
        .iter()
        .filter_map(|parsed| match parsed {
            Parsed::Opaque(name) => Some(name.as_str()),
            _ => None,
        })
        .chain(nested.iter().map(String::as_str))
        .collect();
    let mut counters = HashMap::<&str, i64>::new();
    let mut results = Vec::with_capacity(parsed.len());
    let mut generated = 0_usize;
    for parsed in &parsed {
        let Parsed::Seq(seq) = parsed else {
            results.push(None);
            continue;
        };
        let counter = counters.entry(seq.name.as_str()).or_default();
        *counter = match seq.step {
            Step::Next => counter.saturating_add(1),
            Step::Repeat => *counter,
            Step::Reset(value) => value,
        };
        if opaque.contains(seq.name.as_str()) {
            results.push(None);
        } else {
            results.push(result_text(seq, *counter).filter(|text| {
                generated = generated.saturating_add(text.len());
                generated <= MAX_GENERATED_BYTES
            }));
        }
    }
    for (field, result) in fields.into_iter().zip(results) {
        if let Some(result) = result
            && !field.locked
        {
            field.fallback = Some(result);
        }
    }
}

/// Whether numbering `blocks` reads any of their fields: a SEQ field, or a
/// field nesting a sequence.
pub fn reads_sequence_fields(blocks: &[LayoutBlock]) -> bool {
    let mut blocks = blocks.to_vec();
    let (mut fields, mut nested) = (Vec::new(), HashSet::new());
    collect_blocks(&mut blocks, &mut fields, &mut nested);
    !fields.is_empty() || !nested.is_empty()
}

fn collect_blocks<'a>(
    blocks: &'a mut [LayoutBlock],
    fields: &mut Vec<&'a mut FieldRun>,
    nested: &mut HashSet<String>,
) {
    for block in blocks {
        match block {
            LayoutBlock::Paragraph(paragraph) => collect_paragraph(paragraph, fields, nested),
            LayoutBlock::Table(table) => {
                for cell in table.rows.iter_mut().flat_map(|row| &mut row.cells) {
                    collect_blocks(&mut cell.blocks, fields, nested);
                }
            }
            LayoutBlock::TextBox(text_box) => {
                for paragraph in &mut text_box.content {
                    collect_paragraph(paragraph, fields, nested);
                }
            }
            LayoutBlock::Shape(shape) => collect_shape(shape, fields, nested),
            _ => {}
        }
    }
}

fn collect_shape<'a>(
    shape: &'a mut ShapeBlock,
    fields: &mut Vec<&'a mut FieldRun>,
    nested: &mut HashSet<String>,
) {
    nested.extend(shape.nested_sequences.iter().cloned());
    for paragraph in shape.inner_text.iter_mut().flatten() {
        collect_paragraph(paragraph, fields, nested);
    }
    for child in &mut shape.children {
        collect_shape(child, fields, nested);
    }
}

fn collect_paragraph<'a>(
    paragraph: &'a mut ParagraphBlock,
    fields: &mut Vec<&'a mut FieldRun>,
    nested: &mut HashSet<String>,
) {
    for run in &mut paragraph.runs {
        let Run::Field(field) = run else { continue };
        nested.extend(field.nested_sequences.iter().cloned());
        if field.raw_type.as_deref() == Some("SEQ") {
            fields.push(field);
        }
    }
}

/// The sequence a SEQ instruction counts, in lower case (Word matches names
/// without regard to case); `None` for another field or a SEQ without a name.
pub fn sequence_name(instruction: &str) -> Option<String> {
    seq_name(&mut tokens(instruction))
}

fn seq_name<'a>(tokens: &mut impl Iterator<Item = &'a str>) -> Option<String> {
    tokens
        .next()
        .filter(|token| token.eq_ignore_ascii_case("SEQ"))?;
    tokens
        .next()
        .filter(|token| !token.starts_with('\\'))
        .map(str::to_lowercase)
}

fn parse(field: &FieldRun) -> Parsed {
    let Some(instruction) = field.instruction.as_deref() else {
        return Parsed::Skip;
    };
    let mut tokens = tokens(instruction);
    let Some(name) = seq_name(&mut tokens) else {
        return Parsed::Skip;
    };
    let mut step = None;
    let mut seq = Seq {
        name: String::new(),
        step: Step::Next,
        hidden: false,
        format: Format::Arabic,
        named_format: false,
        keeps_cached: false,
    };
    while let Some(token) = tokens.next() {
        let Some(switch) = token.strip_prefix('\\') else {
            return Parsed::Opaque(name);
        };
        let mut chars = switch.chars();
        let key = chars.next().map(|key| key.to_ascii_lowercase());
        let attached = chars.as_str();
        let mut argument = || {
            if attached.is_empty() {
                tokens.next()
            } else {
                Some(attached)
            }
        };
        // The first of \c, \n and \r decides the step; Word ignores the others.
        match key {
            Some('c') if attached.is_empty() => {
                step.get_or_insert(Step::Repeat);
            }
            Some('n') if attached.is_empty() => {
                step.get_or_insert(Step::Next);
            }
            Some('h') if attached.is_empty() => seq.hidden = true,
            Some('r') => {
                let Some(value) = argument()
                    .filter(|value| value.bytes().all(|byte| byte.is_ascii_digit()))
                    .and_then(|value| value.parse::<u32>().ok())
                else {
                    return Parsed::Opaque(name);
                };
                step.get_or_insert(Step::Reset(i64::from(value)));
            }
            Some('*') => {
                let Some(format) = argument() else {
                    return Parsed::Opaque(name);
                };
                let upper = format.starts_with(|first: char| first.is_ascii_uppercase());
                match format.to_ascii_lowercase().as_str() {
                    "arabic" => seq.format = Format::Arabic,
                    "roman" => seq.format = Format::Roman { upper },
                    "alphabetic" => seq.format = Format::Alphabetic { upper },
                    "mergeformat" | "charformat" => continue,
                    _ => {
                        seq.keeps_cached = true;
                        continue;
                    }
                }
                seq.named_format = true;
            }
            Some('#' | '@') => {
                if argument().is_none() {
                    return Parsed::Opaque(name);
                }
                seq.keeps_cached = true;
            }
            _ => return Parsed::Opaque(name),
        }
    }
    seq.name = name;
    seq.step = step.unwrap_or(Step::Next);
    Parsed::Seq(seq)
}

/// Whitespace-separated instruction tokens, a quoted token without its quotes.
fn tokens(instruction: &str) -> impl Iterator<Item = &str> {
    let mut rest = instruction;
    std::iter::from_fn(move || {
        rest = rest.trim_start();
        if rest.is_empty() {
            return None;
        }
        if let Some(quoted) = rest.strip_prefix('"') {
            let end = quoted.find('"').unwrap_or(quoted.len());
            rest = quoted.get(end + 1..).unwrap_or("");
            return Some(&quoted[..end]);
        }
        let end = rest.find(char::is_whitespace).unwrap_or(rest.len());
        let (token, tail) = rest.split_at(end);
        rest = tail;
        Some(token)
    })
}

fn result_text(seq: &Seq, value: i64) -> Option<String> {
    if seq.keeps_cached {
        return None;
    }
    if seq.hidden && !seq.named_format {
        return Some(String::new());
    }
    match seq.format {
        Format::Arabic => Some(value.to_string()),
        Format::Roman { upper } => spelled(value, upper, roman),
        Format::Alphabetic { upper } => spelled(value, upper, letters),
    }
}

fn spelled(value: i64, upper: bool, spell: fn(i64) -> String) -> Option<String> {
    (value <= MAX_SPELLED).then(|| {
        let text = spell(value);
        if upper {
            text.to_ascii_uppercase()
        } else {
            text
        }
    })
}

/// Lower-case Roman numerals, thousands repeated past 3999; empty for 0.
fn roman(mut value: i64) -> String {
    let mut output = String::new();
    for (step, numeral) in [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ] {
        while value >= step {
            output.push_str(numeral);
            value -= step;
        }
    }
    output
}

/// Word's letter numbering: a to z, then aa to zz, then aaa; empty for 0.
fn letters(value: i64) -> String {
    if value <= 0 {
        return String::new();
    }
    let letter = char::from(b'a' + ((value - 1) % 26) as u8);
    std::iter::repeat_n(letter, ((value - 1) / 26 + 1) as usize).collect()
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;

    #[test]
    fn spelled_results_stop_at_the_generated_text_budget() {
        let field = json!({
            "kind": "field", "fieldType": "other", "rawType": "SEQ",
            "instruction": " SEQ Figure \\r 32767 \\* alphabetic ", "fallback": "7",
        });
        let spelled = "g".repeat(1261);
        let fit = MAX_GENERATED_BYTES / spelled.len();
        let mut blocks: Vec<LayoutBlock> = (0..fit + 3)
            .map(|id| {
                serde_json::from_value(json!({ "kind": "paragraph", "id": id, "runs": [field] }))
                    .unwrap()
            })
            .collect();
        number_sequence_fields(&mut blocks);
        let results: Vec<&str> = blocks
            .iter()
            .map(|block| match block {
                LayoutBlock::Paragraph(ParagraphBlock { runs, .. }) => match &runs[0] {
                    Run::Field(field) => field.fallback.as_deref().unwrap(),
                    _ => unreachable!(),
                },
                _ => unreachable!(),
            })
            .collect();
        assert!(results[..fit].iter().all(|result| *result == spelled));
        assert!(results[fit..].iter().all(|result| *result == "7"));
    }
}
