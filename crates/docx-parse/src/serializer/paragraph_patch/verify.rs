use serde_json::{Value, json};

use crate::paragraph_patch::{Attrs, PatchErrorCode, UnitV1};

use super::lower::{Lowered, WORD};
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
    for node in &lowered.xml.nodes {
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
        let attrs = unit_attrs(&unit);
        let node = &lowered.xml.nodes[cell.node];
        if node.word("commentReference") {
            if !attrs.is_empty() {
                return Err(mismatch("commentReference"));
            }
            continue;
        }
        if let Some(direct) = &runs[cell.run] {
            run(direct, attrs, node.word("sym"))?;
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

fn run(direct: &Value, attrs: &Attrs, symbol: bool) -> Result<(), Failure> {
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
        if color.get("auto") == Some(&Value::Bool(true)) {
            if seeded.is_some() {
                return Err(mismatch("textColor"));
            }
        } else {
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
    if let Some(run) = direct.get("runProperties").and_then(Value::as_object) {
        for (key, value) in run {
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
