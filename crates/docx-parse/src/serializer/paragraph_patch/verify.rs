use std::collections::{BTreeMap, BTreeSet};

use serde_json::{Value, json};

use crate::paragraph_patch::{Attrs, PatchErrorCode, TextEditV1, UnitV1};

use super::lower::{Cell, Lowered, WORD, decoded_scalars};
use super::{Failure, flatten, unit_attrs};

fn mismatch(key: &str) -> Failure {
    Failure::structure(format!(
        "ambiguous: source property {key} differs from seed"
    ))
}

fn agrees(direct: &Value, seeded: Option<&Value>) -> bool {
    if direct.is_null() {
        return true;
    }
    if let Some(object) = direct.as_object() {
        return object
            .iter()
            .all(|(key, value)| agrees(value, seeded.and_then(|v| v.get(key))));
    }
    seeded.is_some_and(|seeded| {
        direct == seeded
            || (direct.is_number() && seeded.is_number() && direct.as_f64() == seeded.as_f64())
    })
}

fn same_marker_value(left: &Value, right: &Value) -> bool {
    match (left, right) {
        (Value::Array(left), Value::Array(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|(left, right)| same_marker_value(left, right))
        }
        (Value::Object(left), Value::Object(right)) => {
            left.len() == right.len()
                && left.iter().all(|(key, left)| {
                    right
                        .get(key)
                        .is_some_and(|right| same_marker_value(left, right))
                })
        }
        _ => agrees(left, Some(right)),
    }
}

pub(super) fn markers(
    source: &str,
    lowered: &Lowered,
    properties: &Attrs,
    code: PatchErrorCode,
) -> Result<(), Failure> {
    let mut bookmarks = Vec::new();
    let mut comments = Vec::new();
    let mut drawing_end = 0;
    for node in &lowered.xml.nodes {
        #[cfg(test)]
        super::work::add(1);
        if node.span.start < drawing_end {
            continue;
        }
        if node.word("drawing") {
            drawing_end = node.span.end;
            continue;
        }
        if node.namespace != WORD
            || !matches!(
                node.local.as_str(),
                "bookmarkStart" | "bookmarkEnd" | "commentRangeStart" | "commentRangeEnd"
            )
        {
            continue;
        }
        let before = lowered
            .cells
            .partition_point(|cell| cell.raw.end <= node.span.start);
        if lowered.cells.get(before).is_some_and(|cell| {
            cell.retained
                && matches!(cell.unit, UnitV1::Atom { .. })
                && cell.raw.start <= node.span.start
                && node.span.end <= cell.raw.end
        }) {
            continue;
        }
        let offset = before.checked_sub(1).map_or(0, |at| lowered.cells[at].to);
        let number = |name| {
            node.attr(source, name)
                .and_then(|value| crate::xml::parse_javascript_integer_prefix(&value))
        };
        if node.word("bookmarkStart") || node.word("bookmarkEnd") {
            let start = node.word("bookmarkStart");
            let mut bookmark = json!({
                "id": number("w:id").unwrap_or_default(),
                "kind": if start { "start" } else { "end" },
                "offset": offset,
            });
            if start {
                bookmark["name"] = json!(node.attr(source, "w:name").unwrap_or_default());
                for (attribute, key) in [("w:colFirst", "colFirst"), ("w:colLast", "colLast")] {
                    if let Some(value) = number(attribute) {
                        bookmark[key] = json!(value);
                    }
                }
            }
            bookmarks.push(bookmark);
        } else if let Some(id) = number("w:id") {
            comments.push((offset, id.to_string(), node.word("commentRangeEnd")));
        }
    }
    comments.sort_unstable();
    for (key, actual) in [
        ("bookmarks", json!(bookmarks)),
        ("markers", json!(comments)),
    ] {
        if let Some(expected) = properties.get(key)
            && !same_marker_value(&actual, expected)
        {
            return Err(Failure::new(code, format!("source {key} differ from plan")));
        }
    }
    Ok(())
}

fn check(key: &str, direct: &Value, seeded: Option<&Value>) -> Result<(), Failure> {
    if agrees(direct, seeded) {
        Ok(())
    } else {
        Err(mismatch(key))
    }
}

fn flag(key: &str, direct: &Value, seeded: &Attrs) -> Result<(), Failure> {
    let enabled = seeded.get(key) == Some(&Value::Bool(true));
    if direct.as_bool() == Some(enabled) {
        Ok(())
    } else {
        Err(mismatch(key))
    }
}

pub(super) fn formatting(source: &str, paragraph: bool) -> Result<Value, Failure> {
    let limits = crate::xml::ParseLimits::default();
    let mut budget = crate::xml::ParseBudget::new(&limits);
    let xml = crate::xml::parse_xml(source.as_bytes(), "", &mut budget)
        .map_err(|_| Failure::structure("ambiguous: invalid source properties"))?;
    let value = if paragraph {
        serde_json::to_value(crate::formatting::parse_paragraph_properties(
            xml.root(),
            None,
        ))
    } else {
        serde_json::to_value(crate::formatting::parse_run_properties(xml.root(), None))
    };
    value.map_err(|_| Failure::structure("ambiguous: invalid source properties"))
}

pub(super) fn seed(
    source: &str,
    lowered: &Lowered,
    units: &[UnitV1],
    properties: &Attrs,
) -> Result<(), Failure> {
    let runs: Vec<_> = lowered
        .runs
        .iter()
        .map(|run| {
            run.properties
                .map(|at| formatting(&source[lowered.xml.nodes[at].span.clone()], false))
                .transpose()
        })
        .collect::<Result<_, _>>()?;
    for (cell, unit) in lowered.cells.iter().zip(flatten(units)) {
        if cell.retained && matches!(cell.unit, UnitV1::Atom { .. }) {
            continue;
        }
        let attrs = unit_attrs(&unit);
        let node = &lowered.xml.nodes[cell.node];
        if node.word("drawing") {
            continue;
        }
        if node.word("commentReference") {
            if !attrs.is_empty() {
                return Err(mismatch("commentReference"));
            }
            continue;
        }
        if let Some(direct) = &runs[cell.run] {
            run(direct, attrs, node.word("sym"))?;
        } else {
            language(None, attrs.get("language"))?;
        }
        if node.word("sym")
            && let Some(font) = node.attr(source, "w:font")
        {
            for slot in ["ascii", "hAnsi", "eastAsia", "cs"] {
                check(
                    "fontFamily",
                    &json!(font),
                    attrs.get("fontFamily").and_then(|v| v.get(slot)),
                )?;
            }
        }
    }
    if let Some(at) = lowered.ppr {
        let direct = formatting(&source[lowered.xml.nodes[at].span.clone()], true)?;
        paragraph(&direct, properties)?;
    }
    Ok(())
}

fn language(direct: Option<&Value>, live: Option<&Value>) -> Result<(), Failure> {
    if let Some(live) = live.and_then(Value::as_object) {
        for (key, expected) in live {
            let key = if key == "val" { "latin" } else { key.as_str() };
            let value = direct
                .and_then(|value| value.get(key))
                .filter(|value| !value.is_null());
            if value != (!expected.is_null()).then_some(expected) {
                return Err(mismatch("language"));
            }
        }
    }
    Ok(())
}

fn run(direct: &Value, attrs: &Attrs, symbol: bool) -> Result<(), Failure> {
    language(direct.get("language"), attrs.get("language"))?;
    for key in ["bold", "italic", "allCaps", "smallCaps", "hidden"] {
        if let Some(value) = direct.get(key) {
            flag(key, value, attrs)?;
        }
    }
    for key in ["boldCs", "italicCs"] {
        if attrs.contains_key(key)
            && let Some(value) = direct.get(key)
        {
            flag(key, value, attrs)?;
        }
    }
    for (source, key, field) in [
        ("fontSize", "fontSize", "size"),
        ("fontSizeCs", "fontSize", "sizeCs"),
        ("styleId", "runStyle", "styleId"),
        ("spacing", "characterSpacing", "spacing"),
        ("position", "characterSpacing", "position"),
        ("scale", "characterSpacing", "scale"),
        ("kerning", "characterSpacing", "kerning"),
        ("ligatures", "typography", "ligatures"),
        ("contextualAlternates", "typography", "contextualAlternates"),
    ] {
        if let Some(value) = direct.get(source) {
            check(key, value, attrs.get(key).and_then(|v| v.get(field)))?;
        }
    }
    if !symbol && let Some(font) = direct.get("fontFamily") {
        for key in [
            "ascii",
            "hAnsi",
            "eastAsia",
            "cs",
            "asciiTheme",
            "hAnsiTheme",
            "eastAsiaTheme",
            "csTheme",
        ] {
            if let Some(value) = font.get(key) {
                check(
                    "fontFamily",
                    value,
                    attrs.get("fontFamily").and_then(|v| v.get(key)),
                )?;
            }
        }
    }
    if let Some(color) = direct.get("color") {
        let seeded = attrs.get("textColor").filter(|value| !value.is_null());
        if color.get("auto") != Some(&Value::Bool(true)) {
            for key in ["rgb", "themeColor", "themeTint", "themeShade"] {
                let value = color.get(key).filter(|value| !value.is_null());
                let seeded = seeded
                    .and_then(|value| value.get(key))
                    .filter(|value| !value.is_null());
                match (value, seeded) {
                    (Some(value), seeded) => check("textColor", value, seeded)?,
                    (None, Some(_)) => return Err(mismatch("textColor")),
                    (None, None) => {}
                }
            }
        }
    }
    if let Some(underline) = direct.get("underline") {
        if underline.get("style").and_then(Value::as_str) == Some("none") {
            if attrs.get("underline").is_some_and(|v| !v.is_null()) {
                return Err(mismatch("underline"));
            }
        } else {
            check("underline", underline, attrs.get("underline"))?;
        }
    }
    let highlight = direct
        .get("highlight")
        .filter(|v| v.as_str() != Some("none"));
    let shading = direct.get("shading");
    let fill = shading
        .filter(|s| s.get("pattern").is_none_or(|v| v.as_str() == Some("clear")))
        .and_then(|s| s.get("fill"));
    let highlight = highlight.cloned().or_else(|| {
        fill.filter(|v| v.get("auto") != Some(&Value::Bool(true)))
            .and_then(|v| v.get("rgb"))
            .and_then(Value::as_str)
            .map(|rgb| json!(format!("#{}", rgb.trim_start_matches('#'))))
    });
    if let Some(highlight) = highlight {
        let value = attrs.get("highlight").map(|v| v.get("color").unwrap_or(v));
        check("highlight", &highlight, value)?;
    }
    if let Some(align) = direct.get("vertAlign").and_then(Value::as_str) {
        for key in ["superscript", "subscript"] {
            flag(key, &json!(align == key), attrs)?;
        }
    }
    if direct.get("strike") == Some(&Value::Bool(true))
        || direct.get("doubleStrike") == Some(&Value::Bool(true))
    {
        let strike = attrs.get("strike").ok_or_else(|| mismatch("strike"))?;
        if let Some(double) = direct.get("doubleStrike") {
            check("strike.double", double, strike.get("double"))?;
        }
    } else if direct.get("strike") == Some(&Value::Bool(false))
        && direct.get("doubleStrike") == Some(&Value::Bool(false))
        && attrs.get("strike").is_some_and(|v| !v.is_null())
    {
        return Err(mismatch("strike"));
    }
    Ok(())
}

fn paragraph(direct: &Value, attrs: &Attrs) -> Result<(), Failure> {
    for key in [
        "alignment",
        "outlineLevel",
        "spaceBefore",
        "spaceAfter",
        "spaceBeforeLines",
        "spaceAfterLines",
        "beforeAutospacing",
        "afterAutospacing",
        "lineSpacing",
        "lineSpacingRule",
        "indentLeft",
        "indentRight",
        "indentFirstLine",
        "hangingIndent",
        "indentLeftChars",
        "indentRightChars",
        "indentFirstLineChars",
        "hangingIndentChars",
        "keepNext",
        "keepLines",
        "widowControl",
        "pageBreakBefore",
        "contextualSpacing",
        "numPr",
    ] {
        if let Some(value) = direct.get(key) {
            let seeded = attrs.get(key).or_else(|| {
                attrs
                    .get("_originalFormatting")
                    .and_then(|formatting| formatting.get(key))
            });
            check(key, value, seeded)?;
        }
    }
    if let Some(value) = direct.get("styleId") {
        check(
            "pStyle",
            value,
            attrs.get("pStyle").or_else(|| attrs.get("styleId")),
        )?;
    }
    language(
        direct
            .get("runProperties")
            .and_then(|value| value.get("language")),
        attrs.get("language"),
    )?;
    if let Some(run) = direct.get("runProperties").and_then(Value::as_object) {
        for (key, value) in run {
            if key == "language" && attrs.contains_key("language") {
                continue;
            }
            if key == "color" && value.get("auto") == Some(&Value::Bool(true)) {
                continue;
            }
            check(
                "defaultTextFormatting",
                value,
                attrs.get("defaultTextFormatting").and_then(|v| v.get(key)),
            )?;
        }
    }
    Ok(())
}

fn visible(source: &str, lowered: &Lowered) -> Result<Vec<(String, Option<u32>)>, Failure> {
    let runs: BTreeSet<_> = lowered.runs.iter().map(|run| run.node).collect();
    let mut cells: BTreeMap<usize, Vec<&Cell>> = BTreeMap::new();
    for cell in &lowered.cells {
        if matches!(cell.unit, UnitV1::Text { .. }) {
            cells.entry(cell.node).or_default().push(cell);
        }
    }
    let mut result = Vec::new();
    for (at, node) in lowered.xml.nodes.iter().enumerate() {
        if let Some(cells) = cells.get(&at) {
            for cell in cells {
                if let UnitV1::Text { text, .. } = &cell.unit {
                    result.push((text.clone(), Some(cell.from)));
                }
            }
        } else if node.word("t")
            && !node.empty
            && node.parent.is_some_and(|run| runs.contains(&run))
        {
            for (text, _) in decoded_scalars(source, node.open.end..node.close.start)? {
                result.push((text, None));
            }
        }
    }
    Ok(result)
}

fn inserted(edit: &TextEditV1) -> String {
    edit.insert
        .iter()
        .map(|insert| insert.text.as_str())
        .collect()
}

pub(super) fn visible_text(
    source: &str,
    lowered: &Lowered,
    edits: &[TextEditV1],
    output: &str,
    written: &Lowered,
) -> Result<(), Failure> {
    let mut edits: Vec<_> = edits.iter().collect();
    edits.sort_by_key(|edit| edit.from);
    let mut next = 0;
    let mut expected = String::new();
    for (text, offset) in visible(source, lowered)? {
        if let Some(offset) = offset {
            while let Some(edit) = edits.get(next).filter(|edit| edit.from <= offset) {
                expected.push_str(&inserted(edit));
                next += 1;
            }
            if edits
                .iter()
                .any(|edit| edit.from <= offset && offset < edit.to)
            {
                continue;
            }
        }
        expected.push_str(&text);
    }
    for edit in &edits[next..] {
        expected.push_str(&inserted(edit));
    }
    let actual: String = visible(output, written)?
        .into_iter()
        .map(|(text, _)| text)
        .collect();
    if actual != expected {
        return Err(Failure::structure("visible text differs outside the edit"));
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::super::PatchContext;
    use super::super::lower::lower;
    use super::*;
    use crate::paragraph_patch::{DonorV1, InsertedTextV1};

    fn paragraph(text: &str, result: &str) -> String {
        format!(
            "<w:p><w:r><w:t>{text}</w:t></w:r><w:r><w:fldChar w:fldCharType=\"begin\"/></w:r><w:r><w:instrText xml:space=\"preserve\"> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType=\"separate\"/></w:r><w:r><w:t>{result}</w:t></w:r><w:r><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>"
        )
    }

    fn units(text: &str) -> Vec<UnitV1> {
        let field = json!({
            "fieldType": "PAGE", "instruction": "PAGE", "displayText": "12",
            "fieldKind": "complex", "modelKind": "field"
        });
        vec![
            UnitV1::Text {
                text: text.to_owned(),
                attrs: Attrs::new(),
            },
            UnitV1::Atom {
                atom_kind: "field".to_owned(),
                payload: serde_json::from_value(field).unwrap(),
                attrs: Attrs::new(),
            },
        ]
    }

    fn check_typing_x(patched: &str) -> Result<(), Failure> {
        let context = PatchContext::word_document();
        let source = paragraph("abc", "12");
        let edit = TextEditV1 {
            from: 1,
            to: 1,
            insert: vec![InsertedTextV1 {
                text: "X".to_owned(),
                attrs: Attrs::new(),
                inherit: Attrs::new(),
            }],
            donor: DonorV1::At { at: 1 },
        };
        let (Ok(lowered), Ok(written)) = (
            lower(&source, &context, &units("abc"), false),
            lower(patched, &context, &units("aXbc"), false),
        ) else {
            panic!("lowering failed");
        };
        visible_text(&source, &lowered, &[edit], patched, &written)
    }

    #[test]
    fn a_patch_keeping_the_field_result_text_is_accepted() {
        assert!(check_typing_x(&paragraph("aXbc", "12")).is_ok());
    }

    #[test]
    fn a_patch_dropping_a_field_result_character_is_declined() {
        let Err(failure) = check_typing_x(&paragraph("aXbc", "1")) else {
            panic!("dropped result text was accepted");
        };
        assert_eq!(failure.code, PatchErrorCode::UnsupportedStructure);
    }
}
