use docx_parse::formatting::ParagraphFormatting;
use docx_parse::paragraph::{Paragraph, ParagraphContent};
use serde_json::{Map, Value};

use super::{
    Attrs, ProjectionContext, ProjectionError, StoryState, attr_string, decode, json_value, runs,
};
use crate::ParagraphProperties;

const TRUTHY_PROPERTIES: &[&str] = &[
    "alignment",
    "bidi",
    "borders",
    "cnfStyle",
    "contextualSpacing",
    "indentFirstLine",
    "indentLeft",
    "indentRight",
    "lineSpacing",
    "pageBreakBefore",
    "shading",
    "styleId",
    "tabs",
];
const PRESENT_PROPERTIES: &[&str] = &[
    "adjustRightInd",
    "afterAutospacing",
    "autoSpaceDE",
    "autoSpaceDN",
    "beforeAutospacing",
    "outlineLevel",
    "spaceAfter",
    "spaceAfterLines",
    "spaceBefore",
    "spaceBeforeLines",
    "widowControl",
];
const AUTHORED_OVERRIDES: &[&str] = &[
    "adjustRightInd",
    "alignment",
    "autoSpaceDE",
    "autoSpaceDN",
    "bidi",
    "pageBreakBefore",
    "styleId",
    "widowControl",
];

pub(super) fn new(
    source: Option<&Paragraph>,
    saved_id: Option<String>,
    formatting: Option<ParagraphFormatting>,
    content: Vec<ParagraphContent>,
) -> Paragraph {
    let repeated_para_id = source
        .filter(|source| source.para_id == saved_id)
        .and_then(|source| source.repeated_para_id);
    Paragraph {
        node_type: "paragraph".to_owned(),
        para_id: saved_id,
        repeated_para_id,
        para_id_attribute: source.and_then(|source| source.para_id_attribute.clone()),
        source_ordinal: source.and_then(|source| source.source_ordinal),
        text_id: source.and_then(|source| source.text_id.clone()),
        extra_attributes: source.map_or_else(Vec::new, |source| source.extra_attributes.clone()),
        formatting,
        property_changes: source.and_then(|source| source.property_changes.clone()),
        p_pr_ins: source.and_then(|source| source.p_pr_ins.clone()),
        p_pr_del: source.and_then(|source| source.p_pr_del.clone()),
        content,
        list_rendering: None,
        rendered_page_break_before: source.and_then(|source| source.rendered_page_break_before),
        section_properties: source.and_then(|source| source.section_properties.clone()),
    }
}

pub(super) fn formatting(
    context: &mut ProjectionContext<'_, '_>,
    story_id: &str,
    story: &StoryState,
    properties: &ParagraphProperties,
    source: Option<&Paragraph>,
) -> Result<Option<ParagraphFormatting>, ProjectionError> {
    let mut attrs = runs::attributes(&properties.values)?;
    if let Some(style) = attrs.remove("pStyle").filter(|style| !style.is_null()) {
        attrs.insert("styleId".to_owned(), style);
    } else {
        attrs.remove("styleId");
    }
    let numbering = attrs.get("numPr").filter(|value| !value.is_null());
    let mut inherited = context
        .inheritance
        .paragraph_formatting(
            attr_string(&properties.values, "pStyle"),
            story.table_paragraph_formatting.clone(),
        )
        .and_then(|value| value.as_object().cloned());
    if let Some(resolved) = &mut inherited {
        if numbering.is_none() && resolved.contains_key("numPr") {
            inherited = None;
        } else if numbering.is_some() || crate::structured::source::story_root(story_id) != story_id
        {
            resolved.retain(|key, _| {
                ["autoSpaceDE", "autoSpaceDN", "adjustRightInd"].contains(&key.as_str())
            });
        } else {
            resolved.remove("indentLeft");
            resolved.remove("indentRight");
        }
    }
    let original = non_null(&attrs, "_originalFormatting")
        .map(|value| {
            value.as_object().ok_or_else(|| {
                ProjectionError::InvalidInput(
                    "original paragraph formatting must be an object".to_owned(),
                )
            })
        })
        .transpose()?;
    let style_numbering = attrs
        .get("numPrFromStyle")
        .or_else(|| original.and_then(|original| original.get("numPrFromStyle")));
    let style_sourced = numbering
        .is_some_and(|numbering| style_numbering.is_some_and(|style| same(numbering, style)));
    if let Some(original) = original {
        let inherits =
            |key| original.get(key).is_none() && inherits_value(&attrs, inherited.as_ref(), key);
        let mut result = original.clone();
        for key in AUTHORED_OVERRIDES {
            if inherits(key) {
                continue;
            }
            let value = non_null(&attrs, key).filter(|value| {
                !matches!(*key, "bidi" | "pageBreakBefore")
                    || original.contains_key(*key)
                    || good(value)
            });
            set(&mut result, key, value);
        }
        if style_sourced || !same_option(numbering, original.get("numPr")) {
            result.remove("numPrFromStyle");
            set(
                &mut result,
                "numPr",
                if style_sourced { None } else { numbering },
            );
        }
        if let Some(cnf) = non_null(&attrs, "cnfStyle") {
            set(&mut result, "cnfStyle", Some(cnf));
        }
        paragraph_mark(&mut result, &properties.values, source)?;
        return decode(Some(&Value::Object(result)));
    }

    let mut result: Map<_, _> = TRUTHY_PROPERTIES
        .iter()
        .filter_map(|key| {
            attrs
                .get(*key)
                .filter(|value| good(value))
                .map(|value| ((*key).to_owned(), value.clone()))
        })
        .chain(PRESENT_PROPERTIES.iter().filter_map(|key| {
            non_null(&attrs, key).map(|value| ((*key).to_owned(), value.clone()))
        }))
        .collect();
    if result.is_empty() && numbering.is_none() {
        return Ok(None);
    }
    if !style_sourced {
        set(&mut result, "numPr", numbering);
    }
    for key in ["lineSpacingRule", "hangingIndent"] {
        set(&mut result, key, attrs.get(key).filter(|value| good(value)));
    }
    if let Some(inherited) = inherited {
        if inherited.contains_key("lineSpacing")
            && same_option(result.get("lineSpacing"), inherited.get("lineSpacing"))
            && result
                .get("lineSpacingRule")
                .and_then(Value::as_str)
                .unwrap_or("auto")
                == inherited
                    .get("lineSpacingRule")
                    .and_then(Value::as_str)
                    .unwrap_or("auto")
        {
            result.remove("lineSpacing");
            result.remove("lineSpacingRule");
        }
        result.retain(|key, value| {
            matches!(
                key.as_str(),
                "styleId"
                    | "numPr"
                    | "lineSpacing"
                    | "lineSpacingRule"
                    | "indentFirstLine"
                    | "hangingIndent"
            ) || !same_option(Some(value), inherited.get(key))
        });
        if result.is_empty() {
            return Ok(None);
        }
    }
    decode(Some(&Value::Object(result)))
}

fn paragraph_mark(
    result: &mut Map<String, Value>,
    attrs: &Attrs,
    source: Option<&Paragraph>,
) -> Result<(), ProjectionError> {
    let mark = source.map(|source| {
        source
            .formatting
            .as_ref()
            .and_then(|format| format.run_properties.as_ref())
    });
    if !result.contains_key("runProperties")
        && mark.flatten().is_none()
        && !attrs.contains_key("language")
    {
        return Ok(());
    }
    let mut run = match result.remove("runProperties") {
        None | Some(Value::Null) => Map::new(),
        Some(Value::Object(run)) => run,
        Some(_) => {
            return Err(ProjectionError::InvalidInput(
                "paragraph mark formatting must be an object".to_owned(),
            ));
        }
    };
    if let Some(mark) = mark {
        set(
            &mut run,
            "fontFamily",
            mark.and_then(|mark| mark.font_family.as_ref())
                .map(json_value)
                .transpose()?
                .as_ref(),
        );
        set(
            &mut run,
            "language",
            mark.and_then(|mark| mark.language.as_ref())
                .map(json_value)
                .transpose()?
                .as_ref(),
        );
    }
    if let Some(language) = attrs.get("language") {
        let patch = json_value(language)?;
        let language = runs::live_language(run.get("language"), &patch);
        set(&mut run, "language", language.as_ref());
    }
    result.insert("runProperties".to_owned(), Value::Object(run));
    Ok(())
}

fn inherits_value(
    attrs: &Map<String, Value>,
    inherited: Option<&Map<String, Value>>,
    key: &str,
) -> bool {
    inherited
        .and_then(|inherited| non_null(inherited, key))
        .is_some_and(|inherited| non_null(attrs, key).is_some_and(|value| same(value, inherited)))
}

pub(super) fn non_null<'a>(values: &'a Map<String, Value>, key: &str) -> Option<&'a Value> {
    values.get(key).filter(|value| !value.is_null())
}

pub(super) fn set(result: &mut Map<String, Value>, key: &str, value: Option<&Value>) {
    if let Some(value) = value {
        result.insert(key.to_owned(), value.clone());
    } else {
        result.remove(key);
    }
}

pub(super) fn good(value: &Value) -> bool {
    match value {
        Value::Null => false,
        Value::Bool(value) => *value,
        Value::Number(value) => value.as_f64().is_some_and(|value| value != 0.0),
        Value::String(value) => !value.is_empty(),
        _ => true,
    }
}

pub(super) fn same_option(left: Option<&Value>, right: Option<&Value>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => same(left, right),
        (None, None) => true,
        _ => false,
    }
}

pub(super) fn same(left: &Value, right: &Value) -> bool {
    let mut pending = vec![(left, right)];
    while let Some((left, right)) = pending.pop() {
        match (left, right) {
            (Value::Number(left), Value::Number(right)) if left.as_f64() == right.as_f64() => {}
            (Value::Array(left), Value::Array(right)) if left.len() == right.len() => {
                pending.extend(left.iter().zip(right));
            }
            (Value::Object(left), Value::Object(right)) if left.len() == right.len() => {
                for (key, value) in left {
                    let Some(other) = right.get(key) else {
                        return false;
                    };
                    pending.push((value, other));
                }
            }
            (Value::Array(_) | Value::Object(_), _) => return false,
            _ if left == right => {}
            _ => return false,
        }
    }
    true
}
