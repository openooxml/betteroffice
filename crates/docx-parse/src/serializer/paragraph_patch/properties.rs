use std::collections::{BTreeMap, BTreeSet};

use serde_json::Value;

use crate::paragraph_patch::{AttrDeltaV1, Attrs, PatchErrorCode, SlotV1};
use crate::serializer::xml_writer::{XmlWriter, escape_xml, int_attr};

use super::lower::{Node, WORD, Xml};
use super::{Failure, apply_bytes};

const RANK_R: &[&str] = &[
    "rStyle",
    "rFonts",
    "b",
    "bCs",
    "i",
    "iCs",
    "caps",
    "smallCaps",
    "strike",
    "dstrike",
    "outline",
    "shadow",
    "emboss",
    "imprint",
    "noProof",
    "snapToGrid",
    "vanish",
    "webHidden",
    "color",
    "spacing",
    "w",
    "kern",
    "position",
    "sz",
    "szCs",
    "highlight",
    "u",
    "effect",
    "bdr",
    "shd",
    "fitText",
    "vertAlign",
    "rtl",
    "cs",
    "em",
    "lang",
    "eastAsianLayout",
    "specVanish",
    "oMath",
    "rPrChange",
];
const RANK_P: &[&str] = &[
    "pStyle",
    "keepNext",
    "keepLines",
    "pageBreakBefore",
    "framePr",
    "widowControl",
    "numPr",
    "suppressLineNumbers",
    "pBdr",
    "shd",
    "tabs",
    "suppressAutoHyphens",
    "kinsoku",
    "wordWrap",
    "overflowPunct",
    "topLinePunct",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "adjustRightInd",
    "snapToGrid",
    "spacing",
    "ind",
    "contextualSpacing",
    "mirrorIndents",
    "suppressOverlap",
    "jc",
    "textDirection",
    "textAlignment",
    "textboxTightWrap",
    "outlineLvl",
    "divId",
    "cnfStyle",
    "rPr",
    "sectPr",
    "pPrChange",
];

#[derive(Default)]
struct Change {
    remove: bool,
    attributes: BTreeMap<&'static str, Option<String>>,
    children: BTreeMap<&'static str, Change>,
}

pub(super) fn run_properties(
    source: &str,
    deltas: &[AttrDeltaV1],
    inherit: &Attrs,
    target: &Attrs,
) -> Result<String, Failure> {
    if deltas.is_empty() {
        return Ok(source.to_owned());
    }
    let mut changes = Change::default();
    for delta in deltas {
        #[cfg(test)]
        super::work::add(1);
        let key = delta.path.first().map(String::as_str).unwrap_or_default();
        if matches!(key, "ins" | "del" | "insertion" | "deletion" | "hyperlink") {
            let kind = if key == "hyperlink" {
                "hyperlink"
            } else {
                "tc"
            };
            return Err(Failure::structure(format!("{kind}: live {key} mutation")));
        }
        if !matches!(
            key,
            "bold"
                | "italic"
                | "boldCs"
                | "italicCs"
                | "underline"
                | "strike"
                | "doubleStrike"
                | "textColor"
                | "highlight"
                | "fontSize"
                | "fontFamily"
                | "superscript"
                | "subscript"
                | "allCaps"
                | "smallCaps"
                | "hidden"
        ) {
            return Err(unsupported(&delta.path));
        }
        expand(delta, &mut |path, after| {
            run_leaf(path, after, inherit, target, &mut changes)
        })?;
    }
    if changes.children.contains_key("color")
        && let Some(color) = target.get("textColor")
        && ["themeColor", "themeTint", "themeShade"]
            .iter()
            .any(|key| color.get(key).is_some_and(|value| !value.is_null()))
    {
        let value = string(color.get("rgb").filter(|value| !value.is_null()))?
            .filter(|value| !value.is_empty())
            .unwrap_or_else(|| "auto".to_owned());
        set(&mut changes, "color", "w:val", Some(value));
    }
    let highlight_changed = deltas
        .iter()
        .any(|delta| delta.path.first().is_some_and(|key| key == "highlight"));
    let hex = target
        .get("highlight")
        .and_then(|value| value.get("color").unwrap_or(value).as_str())
        .is_some_and(|value| value.starts_with('#'));
    if highlight_changed && !hex && !source.is_empty() {
        let xml = property_xml(source)?;
        if let Some(at) = xml.child(0, "shd")?
            && xml.nodes[at]
                .attr(source, "w:val")
                .is_none_or(|value| value == "clear")
        {
            set(&mut changes, "shd", "w:fill", None);
            set(&mut changes, "shd", "w:themeFill", None);
            set(&mut changes, "shd", "w:themeFillTint", None);
            set(&mut changes, "shd", "w:themeFillShade", None);
            if xml.nodes[at].attr(source, "w:val").is_none() {
                set(&mut changes, "shd", "w:val", Some("clear".to_owned()));
            }
        }
    }
    write_container(source, "rPr", &changes)
}

pub(super) fn paragraph_properties(
    source: &str,
    deltas: &[AttrDeltaV1],
    inherit: &Attrs,
    target: &Attrs,
) -> Result<String, Failure> {
    if deltas.is_empty() {
        return Ok(source.to_owned());
    }
    let mut changes = Change::default();
    for delta in deltas {
        let key = delta.path.first().map(String::as_str).unwrap_or_default();
        if !matches!(
            key,
            "alignment"
                | "outlineLevel"
                | "spaceBefore"
                | "spaceAfter"
                | "spaceBeforeLines"
                | "spaceAfterLines"
                | "beforeAutospacing"
                | "afterAutospacing"
                | "lineSpacing"
                | "lineSpacingRule"
                | "indentLeft"
                | "indentRight"
                | "indentFirstLine"
                | "hangingIndent"
                | "indentLeftChars"
                | "indentRightChars"
                | "indentFirstLineChars"
                | "hangingIndentChars"
                | "keepNext"
                | "keepLines"
                | "widowControl"
                | "pageBreakBefore"
                | "contextualSpacing"
        ) {
            return Err(unsupported(&delta.path));
        }
        expand(delta, &mut |path, after| {
            paragraph_leaf(path, after, inherit, target, &mut changes)
        })?;
    }
    write_container(source, "pPr", &changes)
}

pub(super) fn expand(
    delta: &AttrDeltaV1,
    emit: &mut impl FnMut(&[String], Option<&Value>) -> Result<(), Failure>,
) -> Result<(), Failure> {
    fn walk(
        path: &mut Vec<String>,
        before: Option<&Value>,
        after: Option<&Value>,
        emit: &mut impl FnMut(&[String], Option<&Value>) -> Result<(), Failure>,
    ) -> Result<(), Failure> {
        if before == after {
            return Ok(());
        }
        if path.len() == 2 && path[0] == "underline" && path[1] == "color" {
            let normalize = |value: Option<&Value>| {
                value.map(|value| {
                    if value.is_string() {
                        serde_json::json!({"rgb": value})
                    } else {
                        value.clone()
                    }
                })
            };
            let before = normalize(before);
            let after = normalize(after);
            return walk_object(path, before.as_ref(), after.as_ref(), emit);
        }
        if before.is_some_and(Value::is_object) || after.is_some_and(Value::is_object) {
            walk_object(path, before, after, emit)
        } else {
            emit(path, after.filter(|value| !value.is_null()))
        }
    }
    fn walk_object(
        path: &mut Vec<String>,
        before: Option<&Value>,
        after: Option<&Value>,
        emit: &mut impl FnMut(&[String], Option<&Value>) -> Result<(), Failure>,
    ) -> Result<(), Failure> {
        let mut keys = BTreeSet::new();
        for value in [before, after].into_iter().flatten() {
            if let Some(object) = value.as_object() {
                keys.extend(object.keys());
            } else if !value.is_null() {
                return Err(conflict("object changed type"));
            }
        }
        for key in keys {
            path.push(key.clone());
            walk(
                path,
                before.and_then(|value| value.get(key)),
                after.and_then(|value| value.get(key)),
                emit,
            )?;
            path.pop();
        }
        Ok(())
    }
    if delta.path.first().is_some_and(|key| key == "highlight") && delta.path.len() == 1 {
        for value in [slot(&delta.before), slot(&delta.after)]
            .into_iter()
            .flatten()
        {
            if let Some(object) = value.as_object()
                && object.keys().any(|key| key != "color")
            {
                return Err(unsupported(&delta.path));
            }
        }
        let after = slot(&delta.after)
            .map(|value| value.get("color").unwrap_or(value))
            .filter(|value| !value.is_null());
        return emit(&["highlight".to_owned(), "color".to_owned()], after);
    }
    walk(
        &mut delta.path.clone(),
        slot(&delta.before),
        slot(&delta.after),
        emit,
    )
}

pub(super) fn slot(slot: &SlotV1) -> Option<&Value> {
    match slot {
        SlotV1::Missing => None,
        SlotV1::Present { value } => Some(value),
    }
}

fn conflict(detail: impl Into<String>) -> Failure {
    Failure::new(PatchErrorCode::PropertyConflict, detail)
}

fn unsupported(path: &[String]) -> Failure {
    Failure::new(
        PatchErrorCode::UnsupportedKey,
        format!("key:{} unsupported property", path.join(".")),
    )
}

fn string(value: Option<&Value>) -> Result<Option<String>, Failure> {
    value
        .map(|value| {
            value
                .as_str()
                .filter(|value| super::valid_xml_text(value))
                .map(str::to_owned)
                .ok_or_else(|| conflict("expected XML string property"))
        })
        .transpose()
}

fn number(value: Option<&Value>) -> Result<Option<String>, Failure> {
    value
        .map(|value| {
            let number = value
                .as_f64()
                .ok_or_else(|| conflict("expected number property"))?;
            if number.fract() != 0.0 {
                return Err(conflict("property needs an integer value"));
            }
            Ok(int_attr(Some(number)))
        })
        .transpose()
}

fn set(changes: &mut Change, child: &'static str, attr: &'static str, value: Option<String>) {
    changes
        .children
        .entry(child)
        .or_default()
        .attributes
        .insert(attr, value);
}

fn flag(
    changes: &mut Change,
    child: &'static str,
    value: Option<&Value>,
    inherited: Option<&Value>,
) -> Result<(), Failure> {
    if value.is_some_and(|value| !value.is_boolean()) {
        return Err(conflict("expected boolean property"));
    }
    let change = changes.children.entry(child).or_default();
    if value == Some(&Value::Bool(true)) {
        change.attributes.insert("w:val", None);
    } else if inherited == Some(&Value::Bool(true)) || value == Some(&Value::Bool(false)) {
        change.attributes.insert("w:val", Some("0".to_owned()));
    } else {
        change.remove = true;
    }
    Ok(())
}

fn run_leaf(
    path: &[String],
    value: Option<&Value>,
    inherit: &Attrs,
    target: &Attrs,
    changes: &mut Change,
) -> Result<(), Failure> {
    let key = path.first().map(String::as_str).unwrap_or_default();
    let flag_child = match key {
        "bold" => Some("b"),
        "italic" => Some("i"),
        "boldCs" => Some("bCs"),
        "italicCs" => Some("iCs"),
        "allCaps" => Some("caps"),
        "smallCaps" => Some("smallCaps"),
        "hidden" => Some("vanish"),
        _ => None,
    };
    if let Some(child) = flag_child {
        if path.len() != 1 {
            return Err(unsupported(path));
        }
        return flag(changes, child, value, inherit.get(key));
    }
    if value.is_none() && matches!(key, "fontFamily" | "fontSize" | "textColor") {
        let mut inherited = inherit.get(key);
        for field in &path[1..] {
            inherited = inherited.and_then(|value| value.get(field));
        }
        if inherited.is_some_and(|value| !value.is_null()) {
            return Err(conflict(format!(
                "unrepresentable reset {}",
                path.join(".")
            )));
        }
    }
    match key {
        "fontFamily" if path.len() == 2 => {
            let (attr, theme) = match path[1].as_str() {
                "ascii" => ("w:ascii", Some("w:asciiTheme")),
                "hAnsi" => ("w:hAnsi", Some("w:hAnsiTheme")),
                "eastAsia" => ("w:eastAsia", Some("w:eastAsiaTheme")),
                "cs" => ("w:cs", Some("w:cstheme")),
                "asciiTheme" => ("w:asciiTheme", None),
                "hAnsiTheme" => ("w:hAnsiTheme", None),
                "eastAsiaTheme" => ("w:eastAsiaTheme", None),
                "csTheme" => ("w:cstheme", None),
                _ => return Err(unsupported(path)),
            };
            set(changes, "rFonts", attr, string(value)?);
            if value.is_some()
                && let Some(theme) = theme
            {
                set(changes, "rFonts", theme, None);
            }
        }
        "fontSize" if path.len() == 2 => {
            let child = match path[1].as_str() {
                "size" => "sz",
                "sizeCs" => "szCs",
                _ => return Err(unsupported(path)),
            };
            set(changes, child, "w:val", number(value)?);
        }
        "textColor" if path.len() == 2 => {
            let attr = match path[1].as_str() {
                "rgb" => "w:val",
                "themeColor" => "w:themeColor",
                "themeTint" => "w:themeTint",
                "themeShade" => "w:themeShade",
                _ => return Err(unsupported(path)),
            };
            set(changes, "color", attr, string(value)?);
        }
        "underline" if path.len() >= 2 => match path[1].as_str() {
            "style" if path.len() == 2 => {
                let value = string(value)?.or_else(|| Some("none".to_owned()));
                set(changes, "u", "w:val", value);
            }
            "color" if path.len() == 2 => set(changes, "u", "w:color", string(value)?),
            "color" if path.len() == 3 => {
                let attr = match path[2].as_str() {
                    "rgb" => "w:color",
                    "themeColor" => "w:themeColor",
                    "themeTint" => "w:themeTint",
                    "themeShade" => "w:themeShade",
                    _ => return Err(unsupported(path)),
                };
                set(changes, "u", attr, string(value)?);
            }
            _ => return Err(unsupported(path)),
        },
        "strike" if path.len() == 1 || (path.len() == 2 && path[1] == "double") => {
            if value.is_some_and(|value| !value.is_boolean()) {
                return Err(unsupported(path));
            }
            let strike = target.get("strike");
            let double = strike
                .and_then(|value| value.get("double"))
                .and_then(Value::as_bool)
                .unwrap_or(false);
            let enabled =
                strike.is_some_and(|value| !value.is_null() && value != &Value::Bool(false));
            flag(
                changes,
                "strike",
                Some(&Value::Bool(enabled && !double)),
                inherit.get(key),
            )?;
            flag(
                changes,
                "dstrike",
                Some(&Value::Bool(enabled && double)),
                inherit.get(key),
            )?;
        }
        "doubleStrike" if path.len() == 1 => flag(changes, "dstrike", value, inherit.get(key))?,
        "superscript" | "subscript" if path.len() == 1 => {
            if value.is_some_and(|value| !value.is_boolean()) {
                return Err(unsupported(path));
            }
            let super_on = target.get("superscript") == Some(&Value::Bool(true));
            let sub_on = target.get("subscript") == Some(&Value::Bool(true));
            let value = match (super_on, sub_on) {
                (true, true) => return Err(conflict("conflicting vertical alignment")),
                (true, false) => "superscript",
                (false, true) => "subscript",
                (false, false) => "baseline",
            };
            set(changes, "vertAlign", "w:val", Some(value.to_owned()));
        }
        "highlight" if path.len() == 2 && path[1] == "color" => {
            let color = string(value)?;
            if let Some(hex) = color.as_deref().and_then(|value| value.strip_prefix('#')) {
                if hex.len() != 6 || !hex.bytes().all(|byte| byte.is_ascii_hexdigit()) {
                    return Err(unsupported(path));
                }
                set(changes, "shd", "w:val", Some("clear".to_owned()));
                set(changes, "shd", "w:fill", Some(hex.to_owned()));
                set(changes, "shd", "w:themeFill", None);
                set(changes, "shd", "w:themeFillTint", None);
                set(changes, "shd", "w:themeFillShade", None);
                changes.children.entry("highlight").or_default().remove = true;
            } else {
                if let Some(color) = color.as_deref()
                    && !super::super::run::VALID_HIGHLIGHT_COLORS.contains(&color)
                    && color != "none"
                {
                    return Err(unsupported(path));
                }
                set(
                    changes,
                    "highlight",
                    "w:val",
                    color.or_else(|| Some("none".to_owned())),
                );
            }
        }
        _ => return Err(unsupported(path)),
    }
    Ok(())
}

fn paragraph_leaf(
    path: &[String],
    value: Option<&Value>,
    inherit: &Attrs,
    target: &Attrs,
    changes: &mut Change,
) -> Result<(), Failure> {
    let key = path.first().map(String::as_str).unwrap_or_default();
    if path.len() == 1 {
        let child = match key {
            "keepNext" => Some("keepNext"),
            "keepLines" => Some("keepLines"),
            "widowControl" => Some("widowControl"),
            "pageBreakBefore" => Some("pageBreakBefore"),
            "contextualSpacing" => Some("contextualSpacing"),
            _ => None,
        };
        if let Some(child) = child {
            return flag(changes, child, value, inherit.get(key));
        }
        if matches!(
            key,
            "indentFirstLine" | "hangingIndent" | "indentFirstLineChars" | "hangingIndentChars"
        ) {
            return first_line(key, value, inherit, target, changes);
        }
        let (child, attr) = match key {
            "alignment" => ("jc", "w:val"),
            "outlineLevel" => ("outlineLvl", "w:val"),
            "spaceBefore" => ("spacing", "w:before"),
            "spaceAfter" => ("spacing", "w:after"),
            "spaceBeforeLines" => ("spacing", "w:beforeLines"),
            "spaceAfterLines" => ("spacing", "w:afterLines"),
            "beforeAutospacing" => ("spacing", "w:beforeAutospacing"),
            "afterAutospacing" => ("spacing", "w:afterAutospacing"),
            "lineSpacing" => ("spacing", "w:line"),
            "lineSpacingRule" => ("spacing", "w:lineRule"),
            "indentLeft" => ("ind", "w:left"),
            "indentRight" => ("ind", "w:right"),
            "indentLeftChars" => ("ind", "w:leftChars"),
            "indentRightChars" => ("ind", "w:rightChars"),
            _ => return Err(unsupported(path)),
        };
        let mut emitted = if matches!(key, "alignment" | "lineSpacingRule") {
            string(value)?
        } else if matches!(key, "beforeAutospacing" | "afterAutospacing") {
            value
                .map(|value| {
                    value
                        .as_bool()
                        .map(|value| if value { "1" } else { "0" }.to_owned())
                        .ok_or_else(|| conflict("expected autospacing boolean"))
                })
                .transpose()?
        } else {
            number(value)?
        };
        if emitted.is_none() && inherit.get(key).is_some_and(|value| !value.is_null()) {
            emitted = match key {
                "outlineLevel" => Some("9".to_owned()),
                "beforeAutospacing" | "afterAutospacing" => Some("0".to_owned()),
                _ => return Err(conflict(format!("unrepresentable reset {key}"))),
            };
        }
        set(changes, child, attr, emitted);
    } else {
        return Err(unsupported(path));
    }
    Ok(())
}

fn first_line(
    key: &str,
    value: Option<&Value>,
    inherit: &Attrs,
    target: &Attrs,
    changes: &mut Change,
) -> Result<(), Failure> {
    let chars = key.ends_with("Chars");
    let (amount, hanging, normal, inverse) = if chars {
        (
            "indentFirstLineChars",
            "hangingIndentChars",
            "w:firstLineChars",
            "w:hangingChars",
        )
    } else {
        (
            "indentFirstLine",
            "hangingIndent",
            "w:firstLine",
            "w:hanging",
        )
    };
    let is_hanging = target
        .get(hanging)
        .and_then(Value::as_bool)
        .unwrap_or(false);
    if let Some(number) = target.get(amount).and_then(Value::as_f64) {
        if !is_hanging && number < 0.0 {
            return Err(Failure::new(
                PatchErrorCode::AmbiguousAlignment,
                format!("ambiguous: indent kind {amount}"),
            ));
        }
        let (selected, other) = if is_hanging {
            (inverse, normal)
        } else {
            (normal, inverse)
        };
        let number = if is_hanging { number.abs() } else { number };
        set(changes, "ind", selected, Some(int_attr(Some(number))));
        set(changes, "ind", other, None);
    } else if value.is_none() {
        if is_hanging {
            return Err(conflict("hanging indent needs its amount"));
        }
        if inherit.get(amount).is_some_and(|value| !value.is_null())
            || inherit.get(hanging) == Some(&Value::Bool(true))
        {
            return Err(conflict(format!("unrepresentable reset {amount}")));
        }
        set(changes, "ind", normal, None);
        set(changes, "ind", inverse, None);
    } else {
        return Err(conflict("hanging indent needs its amount"));
    }
    Ok(())
}

fn property_xml(source: &str) -> Result<Xml, Failure> {
    Xml::parse(source, &BTreeMap::from([("w".to_owned(), WORD.to_owned())]))
}

fn write_container(source: &str, name: &'static str, change: &Change) -> Result<String, Failure> {
    let created = format!("<w:{name}></w:{name}>");
    let mut output = if source.is_empty() {
        created.clone()
    } else {
        source.to_owned()
    };
    for (&child, change) in &change.children {
        output = write_child(&output, name, child, change)?;
    }
    if source.is_empty() && output == created {
        Ok(String::new())
    } else {
        Ok(output)
    }
}

fn rank(parent: &str, child: &str) -> Option<usize> {
    let table = match parent {
        "rPr" => RANK_R,
        "pPr" => RANK_P,
        _ => return None,
    };
    table.iter().position(|name| *name == child)
}

fn boolean(child: &str) -> bool {
    matches!(
        child,
        "b" | "bCs"
            | "i"
            | "iCs"
            | "caps"
            | "smallCaps"
            | "vanish"
            | "strike"
            | "dstrike"
            | "keepNext"
            | "keepLines"
            | "widowControl"
            | "pageBreakBefore"
            | "contextualSpacing"
    )
}

fn write_child(
    source: &str,
    parent: &str,
    child: &'static str,
    change: &Change,
) -> Result<String, Failure> {
    let xml = property_xml(source)?;
    if let Some(at) = xml.child(0, child)? {
        let node = &xml.nodes[at];
        if change.remove {
            if node
                .attributes
                .iter()
                .any(|attribute| attribute.name != "w:val")
                || !node.children.is_empty()
            {
                return Err(conflict(format!("cannot remove opaque {child}")));
            }
            return apply_bytes(source, vec![(node.span.clone(), String::new())]);
        }
        if change.attributes.get("w:val") == Some(&None)
            && matches!(
                child,
                "sz" | "szCs" | "jc" | "outlineLvl" | "vertAlign" | "highlight" | "color" | "shd"
            )
        {
            if node.attributes.iter().any(|attribute| {
                attribute.name != "w:val"
                    && !change.attributes.contains_key(attribute.name.as_str())
            }) || !node.children.is_empty()
            {
                return Err(conflict("reset retains opaque attributes"));
            }
            if change.attributes.values().all(Option::is_none) {
                return apply_bytes(source, vec![(node.span.clone(), String::new())]);
            }
            return Err(conflict(format!(
                "reset removes required value from {child}"
            )));
        }
        let mut attributes = change.attributes.clone();
        if child == "ind" {
            for (normal, alias) in [("w:left", "w:start"), ("w:right", "w:end")] {
                if let Some(value) = attributes.get(normal).cloned()
                    && node.attr(source, alias).is_some()
                {
                    if node.attr(source, normal).is_none() {
                        attributes.remove(normal);
                    }
                    attributes.insert(alias, value);
                }
            }
        }
        let replacement = patch_attributes(source, node, &attributes)?;
        return apply_bytes(source, vec![(node.span.clone(), replacement)]);
    }
    if change.remove {
        return Ok(source.to_owned());
    }
    if change.attributes.values().all(Option::is_none)
        && !(change.attributes.contains_key("w:val") && boolean(child))
    {
        return Ok(source.to_owned());
    }
    let fragment = if let Some(fragment) = sparse_fragment(parent, child, &change.attributes)? {
        fragment
    } else {
        let qualified = qualified(child);
        let mut writer = XmlWriter::with_capacity(64);
        if change.attributes.len() == 1 && change.attributes.contains_key("w:val") {
            if let Some(value) = &change.attributes["w:val"] {
                if parent == "pPr" {
                    super::super::paragraph::empty_attr(&mut writer, qualified, "w:val", value);
                } else {
                    super::super::run::empty_attr(&mut writer, qualified, "w:val", value);
                }
            } else {
                writer.start_element(qualified).end_element();
            }
        } else {
            writer.start_element(qualified);
            let mut attributes: Vec<_> = change.attributes.iter().collect();
            attributes.sort_by_key(|(attr, _)| attribute_rank(child, attr));
            for (attr, value) in attributes {
                if let Some(value) = value {
                    writer.attribute(attr, value);
                }
            }
            writer.end_element();
        }
        writer.finish()
    };
    let root = &xml.nodes[0];
    let wanted = rank(parent, child).ok_or_else(|| conflict("unknown property rank"))?;
    let mut previous = None;
    let mut insertion = None;
    for &at in &root.children {
        let node = &xml.nodes[at];
        if node.namespace != WORD {
            continue;
        }
        if let Some(current) = rank(parent, &node.local) {
            if previous.is_some_and(|previous| current <= previous) {
                return Err(conflict("ambiguous property order"));
            }
            previous = Some(current);
            if current > wanted && insertion.is_none() {
                insertion = Some(node.span.start);
            }
        }
    }
    if root.empty {
        let opener = source[root.open.clone()]
            .trim_end_matches('>')
            .trim_end_matches('/');
        return Ok(format!("{opener}>{fragment}</{}>", root.name));
    }
    let at = insertion.unwrap_or(root.close.start);
    apply_bytes(source, vec![(at..at, fragment)])
}

fn sparse_fragment(
    parent: &str,
    child: &str,
    attributes: &BTreeMap<&str, Option<String>>,
) -> Result<Option<String>, Failure> {
    let mut fields = serde_json::Map::new();
    let value = |name: &str| attributes.get(name).and_then(Option::as_deref);
    let scalar = |name: &str, numeric: bool| -> Result<Value, Failure> {
        match value(name) {
            Some(value) if numeric => value
                .parse::<f64>()
                .map(|value| serde_json::json!(value))
                .map_err(|_| conflict("invalid numeric property")),
            Some(value) => Ok(Value::String(value.to_owned())),
            None => Ok(Value::Null),
        }
    };
    let generated = if parent == "rPr" {
        let flag = match child {
            "b" => Some("bold"),
            "bCs" => Some("boldCs"),
            "i" => Some("italic"),
            "iCs" => Some("italicCs"),
            "caps" => Some("allCaps"),
            "smallCaps" => Some("smallCaps"),
            "vanish" => Some("hidden"),
            "strike" => Some("strike"),
            "dstrike" => Some("doubleStrike"),
            _ => None,
        };
        if let Some(flag) = flag {
            fields.insert(flag.to_owned(), Value::Bool(value("w:val") != Some("0")));
        } else {
            match child {
                "sz" | "szCs" => {
                    let key = if child == "sz" {
                        "fontSize"
                    } else {
                        "fontSizeCs"
                    };
                    fields.insert(key.to_owned(), scalar("w:val", true)?);
                }
                "vertAlign" | "highlight" => {
                    fields.insert(child.to_owned(), scalar("w:val", false)?);
                }
                "rFonts" => {
                    let mut fonts = serde_json::Map::new();
                    for (name, value) in attributes {
                        if let Some(value) = value {
                            let key = name.trim_start_matches("w:");
                            let key = if key == "cstheme" { "csTheme" } else { key };
                            fonts.insert(key.to_owned(), Value::String(value.clone()));
                        }
                    }
                    fields.insert("fontFamily".to_owned(), Value::Object(fonts));
                }
                "color" | "u" => {
                    let mut color = serde_json::Map::new();
                    for (attribute, key) in [
                        (if child == "u" { "w:color" } else { "w:val" }, "rgb"),
                        ("w:themeColor", "themeColor"),
                        ("w:themeTint", "themeTint"),
                        ("w:themeShade", "themeShade"),
                    ] {
                        if let Some(value) = value(attribute) {
                            color.insert(key.to_owned(), Value::String(value.to_owned()));
                        }
                    }
                    if child == "u" {
                        fields.insert(
                            "underline".to_owned(),
                            serde_json::json!({
                                "style": value("w:val").unwrap_or("single"), "color": color,
                            }),
                        );
                    } else {
                        fields.insert("color".to_owned(), Value::Object(color));
                    }
                }
                _ => return Ok(None),
            }
        }
        let formatting =
            serde_json::from_value::<crate::formatting::TextFormatting>(Value::Object(fields))
                .map_err(|error| conflict(error.to_string()))?;
        super::super::run::serialize_text_formatting(Some(&formatting))
    } else {
        match child {
            "jc" | "outlineLvl" => {
                let key = if child == "jc" {
                    "alignment"
                } else {
                    "outlineLevel"
                };
                fields.insert(key.to_owned(), scalar("w:val", child == "outlineLvl")?);
            }
            "keepNext" | "keepLines" | "widowControl" | "pageBreakBefore" | "contextualSpacing" => {
                fields.insert(child.to_owned(), Value::Bool(value("w:val") != Some("0")));
            }
            "spacing" => {
                for (attribute, key) in [
                    ("w:before", "spaceBefore"),
                    ("w:after", "spaceAfter"),
                    ("w:beforeLines", "spaceBeforeLines"),
                    ("w:afterLines", "spaceAfterLines"),
                    ("w:line", "lineSpacing"),
                    ("w:lineRule", "lineSpacingRule"),
                    ("w:beforeAutospacing", "beforeAutospacing"),
                    ("w:afterAutospacing", "afterAutospacing"),
                ] {
                    if value(attribute).is_none() {
                        continue;
                    }
                    let emitted = if key.ends_with("Autospacing") {
                        Value::Bool(value(attribute) == Some("1"))
                    } else {
                        scalar(attribute, key != "lineSpacingRule")?
                    };
                    fields.insert(key.to_owned(), emitted);
                }
            }
            "ind" => {
                for (attribute, key) in [
                    ("w:left", "indentLeft"),
                    ("w:right", "indentRight"),
                    ("w:leftChars", "indentLeftChars"),
                    ("w:rightChars", "indentRightChars"),
                    ("w:firstLine", "indentFirstLine"),
                    ("w:hanging", "indentFirstLine"),
                    ("w:firstLineChars", "indentFirstLineChars"),
                    ("w:hangingChars", "indentFirstLineChars"),
                ] {
                    if value(attribute).is_none() {
                        continue;
                    }
                    fields.insert(key.to_owned(), scalar(attribute, true)?);
                    if attribute.starts_with("w:hanging") {
                        let flag = if attribute.ends_with("Chars") {
                            "hangingIndentChars"
                        } else {
                            "hangingIndent"
                        };
                        fields.insert(flag.to_owned(), Value::Bool(true));
                    }
                }
            }
            _ => return Ok(None),
        }
        let formatting =
            serde_json::from_value::<crate::formatting::ParagraphFormatting>(Value::Object(fields))
                .map_err(|error| conflict(error.to_string()))?;
        super::super::paragraph::serialize_paragraph_formatting(
            Some(&formatting),
            None,
            None,
            None,
            true,
            None,
        )
        .map_err(|error| conflict(error.to_string()))?
    };
    if generated.is_empty() {
        return Ok(None);
    }
    let xml = property_xml(&generated)?;
    let Some(at) = xml.child(0, child)? else {
        return Ok(None);
    };
    Ok(Some(generated[xml.nodes[at].span.clone()].to_owned()))
}

fn attribute_rank(child: &str, attribute: &str) -> usize {
    let table: &[&str] = match child {
        "rFonts" => &[
            "w:ascii",
            "w:hAnsi",
            "w:eastAsia",
            "w:cs",
            "w:asciiTheme",
            "w:hAnsiTheme",
            "w:eastAsiaTheme",
            "w:cstheme",
            "w:hint",
        ],
        "spacing" => &[
            "w:before",
            "w:after",
            "w:beforeLines",
            "w:afterLines",
            "w:line",
            "w:lineRule",
            "w:beforeAutospacing",
            "w:afterAutospacing",
        ],
        "ind" => &[
            "w:left",
            "w:leftChars",
            "w:right",
            "w:rightChars",
            "w:firstLine",
            "w:hanging",
            "w:firstLineChars",
            "w:hangingChars",
        ],
        _ => &[
            "w:val",
            "w:color",
            "w:fill",
            "w:themeColor",
            "w:themeTint",
            "w:themeShade",
        ],
    };
    table
        .iter()
        .position(|name| *name == attribute)
        .unwrap_or(usize::MAX)
}

fn patch_attributes(
    source: &str,
    node: &Node,
    changes: &BTreeMap<&str, Option<String>>,
) -> Result<String, Failure> {
    let mut edits = Vec::new();
    let mut ordered: Vec<_> = changes.iter().collect();
    ordered.sort_by_key(|(name, _)| attribute_rank(&node.local, name));
    for (name, value) in ordered {
        let matching: Vec<_> = node
            .attributes
            .iter()
            .filter(|attribute| attribute.name == *name)
            .collect();
        if matching.len() > 1 {
            return Err(conflict("duplicate attribute"));
        }
        if let Some(attribute) = matching.first() {
            edits.push(match value {
                Some(value) => (attribute.value.clone(), escape_xml(value)),
                None => (attribute.token.clone(), String::new()),
            });
        } else if let Some(value) = value {
            let rank = attribute_rank(&node.local, name);
            let at = node
                .attributes
                .iter()
                .find(|attribute| attribute_rank(&node.local, &attribute.name) > rank)
                .map_or(
                    node.open.end - if node.empty { 2 } else { 1 },
                    |attribute| attribute.token.start,
                );
            edits.push((at..at, format!(" {name}=\"{}\"", escape_xml(value))));
        }
    }
    let whole = apply_bytes(source, edits)?;
    let end = whole.len() - (source.len() - node.span.end);
    Ok(whole[node.span.start..end].to_owned())
}

fn qualified(child: &str) -> &'static str {
    match child {
        "rPr" => "w:rPr",
        "pPr" => "w:pPr",
        "rFonts" => "w:rFonts",
        "b" => "w:b",
        "bCs" => "w:bCs",
        "i" => "w:i",
        "iCs" => "w:iCs",
        "caps" => "w:caps",
        "smallCaps" => "w:smallCaps",
        "vanish" => "w:vanish",
        "strike" => "w:strike",
        "dstrike" => "w:dstrike",
        "color" => "w:color",
        "sz" => "w:sz",
        "szCs" => "w:szCs",
        "u" => "w:u",
        "highlight" => "w:highlight",
        "shd" => "w:shd",
        "vertAlign" => "w:vertAlign",
        "jc" => "w:jc",
        "outlineLvl" => "w:outlineLvl",
        "spacing" => "w:spacing",
        "ind" => "w:ind",
        "keepNext" => "w:keepNext",
        "keepLines" => "w:keepLines",
        "widowControl" => "w:widowControl",
        "pageBreakBefore" => "w:pageBreakBefore",
        "contextualSpacing" => "w:contextualSpacing",
        _ => unreachable!("checked property name"),
    }
}
