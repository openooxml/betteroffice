use std::collections::BTreeSet;

use serde_json::{Value, json};

use crate::paragraph_patch::{AttrDeltaV1, Attrs, SlotV1};

use super::{Failure, properties, verify};

const FONT_SLOTS: [(&str, &str); 4] = [
    ("ascii", "asciiTheme"),
    ("hAnsi", "hAnsiTheme"),
    ("eastAsia", "eastAsiaTheme"),
    ("cs", "csTheme"),
];

pub(super) fn retained(
    source: &str,
    deltas: &[AttrDeltaV1],
    inherit: &Attrs,
    target: &Attrs,
    paragraph: bool,
) -> Result<Vec<AttrDeltaV1>, Failure> {
    let direct = if source.is_empty() {
        Value::Null
    } else {
        verify::formatting(source, paragraph)?
    };
    let mut paths = Vec::new();
    for delta in deltas {
        properties::expand(delta, &mut |path, _| {
            paths.push(path.to_vec());
            Ok(())
        })?;
    }
    let mut effective = inherit.clone();
    let direct = direct_attrs(&direct, inherit, paragraph);
    for (key, value) in &direct {
        if key == "fontFamily" {
            merge_fonts(
                effective.entry(key.clone()).or_insert(Value::Null),
                value,
                target,
            );
        } else if key == "textColor" {
            effective.insert(key.clone(), value.clone());
        } else {
            overlay(effective.entry(key.clone()).or_insert(Value::Null), value);
        }
    }
    if let Some(font) = effective
        .get_mut("fontFamily")
        .and_then(Value::as_object_mut)
    {
        for (explicit, theme) in FONT_SLOTS {
            let target = target.get("fontFamily");
            if target
                .and_then(|value| value.get(explicit))
                .is_some_and(|value| !value.is_null())
                && target
                    .and_then(|value| value.get(theme))
                    .is_none_or(Value::is_null)
                && font.get(theme).is_some_and(|value| !value.is_null())
            {
                font.remove(explicit);
                font.remove(theme);
            }
        }
    }
    let mut result = Vec::new();
    for key in inherit.keys().chain(target.keys()).collect::<BTreeSet<_>>() {
        if matches!(
            key.as_str(),
            "ins"
                | "del"
                | "insertion"
                | "deletion"
                | "hyperlink"
                | "bookmarks"
                | "markers"
                | "defaultTextFormatting"
        ) {
            continue;
        }
        collect(
            &mut vec![key.clone()],
            effective.get(key),
            target.get(key),
            inherit.get(key),
            &paths,
            &mut result,
        );
    }
    Ok(result)
}

fn direct_attrs(direct: &Value, inherit: &Attrs, paragraph: bool) -> Attrs {
    let mut attrs: Attrs = direct
        .as_object()
        .into_iter()
        .flatten()
        .map(|(key, value)| (key.clone(), value.clone()))
        .collect();
    if paragraph {
        if let Some(style) = attrs.get("styleId").cloned() {
            attrs.insert("pStyle".to_owned(), style);
        }
        for (amount, hanging) in [
            ("indentFirstLine", "hangingIndent"),
            ("indentFirstLineChars", "hangingIndentChars"),
        ] {
            if attrs.contains_key(amount) && !attrs.contains_key(hanging) {
                attrs.insert(hanging.to_owned(), json!(false));
            }
        }
        return attrs;
    }
    let size = attrs.remove("fontSize");
    let size_cs = attrs.remove("fontSizeCs");
    if size.is_some() || size_cs.is_some() {
        let mut sizes = serde_json::Map::new();
        if let Some(size) = size {
            sizes.insert("size".to_owned(), size);
        }
        if let Some(size) = size_cs {
            sizes.insert("sizeCs".to_owned(), size);
        }
        attrs.insert("fontSize".to_owned(), Value::Object(sizes));
    }
    if let Some(style) = attrs.remove("styleId") {
        attrs.insert("runStyle".to_owned(), json!({"styleId": style}));
    }
    if let Some(color) = attrs.remove("color")
        && color.get("auto") != Some(&Value::Bool(true))
    {
        attrs.insert("textColor".to_owned(), color);
    }
    if attrs
        .get("underline")
        .and_then(|value| value.get("style"))
        .and_then(Value::as_str)
        == Some("none")
    {
        attrs.insert("underline".to_owned(), Value::Null);
    }
    let highlight = attrs.remove("highlight");
    let shading = attrs.get("shading");
    let fill = shading
        .filter(|value| {
            value
                .get("pattern")
                .is_none_or(|value| value.as_str() == Some("clear"))
        })
        .and_then(|value| value.get("fill"))
        .filter(|value| value.get("auto") != Some(&Value::Bool(true)))
        .and_then(|value| value.get("rgb"))
        .and_then(Value::as_str);
    if let Some(highlight) = highlight.filter(|value| value.as_str() != Some("none")) {
        attrs.insert("highlight".to_owned(), json!({"color": highlight}));
    } else if let Some(fill) = fill {
        let highlight = json!({"color": format!("#{}", fill.trim_start_matches('#'))});
        attrs.insert("highlight".to_owned(), highlight);
    } else if direct.get("highlight").is_some() {
        attrs.insert("highlight".to_owned(), Value::Null);
    }
    if let Some(align) = attrs.remove("vertAlign") {
        attrs.insert("superscript".to_owned(), json!(align == "superscript"));
        attrs.insert("subscript".to_owned(), json!(align == "subscript"));
    }
    let strike = attrs.remove("strike");
    let double = attrs.remove("doubleStrike");
    if strike.is_some() || double.is_some() {
        let inherited = inherit.get("strike").filter(|value| !value.is_null());
        let inherited_double =
            inherited.and_then(|value| value.get("double")) == Some(&Value::Bool(true));
        let single = strike
            .and_then(|value| value.as_bool())
            .unwrap_or(inherited.is_some() && !inherited_double);
        let double = double
            .and_then(|value| value.as_bool())
            .unwrap_or(inherited_double);
        attrs.insert(
            "strike".to_owned(),
            if single || double {
                json!({"double": double})
            } else {
                Value::Null
            },
        );
    }
    for (source, target) in [("shadow", "textShadow"), ("outline", "textOutline")] {
        if let Some(value) = attrs.remove(source) {
            attrs.insert(target.to_owned(), value);
        }
    }
    attrs
}

fn overlay(target: &mut Value, direct: &Value) {
    if let Some(target) = target.as_object_mut()
        && let Some(direct) = direct.as_object()
    {
        for (key, value) in direct {
            overlay(target.entry(key.clone()).or_insert(Value::Null), value);
        }
    } else {
        *target = direct.clone();
    }
}

fn merge_fonts(effective: &mut Value, direct: &Value, target: &Attrs) {
    if !effective.is_object() {
        *effective = json!({});
    }
    let effective = effective.as_object_mut().unwrap();
    for (explicit, theme) in FONT_SLOTS {
        if direct.get(explicit).is_some() || direct.get(theme).is_some() {
            effective.remove(explicit);
            effective.remove(theme);
            for key in [explicit, theme] {
                if let Some(value) = direct.get(key) {
                    effective.insert(key.to_owned(), value.clone());
                }
            }
            let target = target.get("fontFamily");
            if direct.get(explicit).is_none()
                && let Some(reference) = direct.get(theme)
                && target.and_then(|value| value.get(theme)) == Some(reference)
                && let Some(value) = target.and_then(|value| value.get(explicit))
            {
                effective.insert(explicit.to_owned(), value.clone());
            }
        }
    }
}

fn collect(
    path: &mut Vec<String>,
    actual: Option<&Value>,
    target: Option<&Value>,
    inherited: Option<&Value>,
    paths: &[Vec<String>],
    result: &mut Vec<AttrDeltaV1>,
) {
    if paths.iter().any(|changed| path.starts_with(changed)) {
        return;
    }
    let target = target.filter(|value| !value.is_null());
    let inherited = inherited.filter(|value| !value.is_null());
    if target.is_some_and(Value::is_object) || inherited.is_some_and(Value::is_object) {
        let keys: BTreeSet<_> = [target, inherited]
            .into_iter()
            .flatten()
            .filter_map(Value::as_object)
            .flat_map(|value| value.keys())
            .collect();
        for key in keys {
            path.push(key.clone());
            collect(
                path,
                actual.and_then(|value| value.get(key)),
                target.and_then(|value| value.get(key)),
                inherited.and_then(|value| value.get(key)),
                paths,
                result,
            );
            path.pop();
        }
        return;
    }
    let actual = present(actual, path);
    let target = present(target, path);
    if actual == target
        || actual.zip(target).is_some_and(|(actual, target)| {
            actual.is_number() && target.is_number() && actual.as_f64() == target.as_f64()
        })
    {
        return;
    }
    let slot = |value: Option<&Value>| {
        value.map_or(SlotV1::Missing, |value| SlotV1::Present {
            value: value.clone(),
        })
    };
    result.push(AttrDeltaV1 {
        path: path.clone(),
        before: slot(actual),
        after: slot(target),
    });
}

fn present<'a>(value: Option<&'a Value>, path: &[String]) -> Option<&'a Value> {
    let off_is_absent = path.len() == 1
        && matches!(
            path[0].as_str(),
            "bold"
                | "italic"
                | "boldCs"
                | "italicCs"
                | "strike"
                | "doubleStrike"
                | "superscript"
                | "subscript"
                | "allCaps"
                | "smallCaps"
                | "hidden"
                | "keepNext"
                | "keepLines"
                | "pageBreakBefore"
                | "contextualSpacing"
                | "beforeAutospacing"
                | "afterAutospacing"
        );
    value.filter(|value| !value.is_null() && (!off_is_absent || *value != &Value::Bool(false)))
}
