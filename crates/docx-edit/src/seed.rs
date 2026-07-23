use std::collections::{BTreeMap, BTreeSet, HashMap};
use std::fmt;
use std::sync::Arc;

use serde::Deserialize;
use serde::de::{MapAccess, SeqAccess, Visitor};
use serde_json::{Map, Value, json};
use yrs::Any;
use yrs::types::Attrs;

use crate::{EditCtx, EditingDoc, RawOp};

type JsonObject = BTreeMap<String, Value>;

#[derive(Clone)]
struct Mark {
    name: String,
    attrs: Vec<(String, Value)>,
}

#[derive(Clone)]
enum UnitContent {
    Text(String),
    Embed { kind: String, payload: JsonObject },
}

#[derive(Clone)]
struct InlineUnit {
    content: UnitContent,
    attrs: JsonObject,
    pm_size: u32,
    comment_id: Option<String>,
    marks: Vec<Mark>,
}

struct StoryPlan {
    story_id: String,
    units: Vec<InlineUnit>,
    comment_coverage: Vec<(String, Vec<(u32, u32)>)>,
}

struct ProjectedCell {
    attrs: JsonObject,
    content: Vec<Value>,
}

struct ProjectedRow {
    attrs: JsonObject,
    cells: Vec<ProjectedCell>,
}

struct ProjectedTable {
    attrs: JsonObject,
    rows: Vec<ProjectedRow>,
}

#[derive(Clone, Copy)]
struct StoryOptions {
    include_page_breaks: bool,
    append_body_tail: bool,
    seed_comments: bool,
}

struct LoweringContext {
    styles: StyleResolver,
    theme: Option<Value>,
    source_json: Arc<BTreeMap<String, String>>,
    plans: Vec<StoryPlan>,
}

enum OrderedValue {
    Null,
    Bool(bool),
    Number(serde_json::Number),
    String(String),
    Array(Vec<OrderedValue>),
    Object(Vec<(String, OrderedValue)>),
}

struct OrderedValueVisitor;

impl<'de> Visitor<'de> for OrderedValueVisitor {
    type Value = OrderedValue;

    fn expecting(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("a JSON value")
    }

    fn visit_unit<E>(self) -> Result<Self::Value, E> {
        Ok(OrderedValue::Null)
    }

    fn visit_none<E>(self) -> Result<Self::Value, E> {
        Ok(OrderedValue::Null)
    }

    fn visit_bool<E>(self, value: bool) -> Result<Self::Value, E> {
        Ok(OrderedValue::Bool(value))
    }

    fn visit_i64<E>(self, value: i64) -> Result<Self::Value, E> {
        Ok(OrderedValue::Number(value.into()))
    }

    fn visit_u64<E>(self, value: u64) -> Result<Self::Value, E> {
        Ok(OrderedValue::Number(value.into()))
    }

    fn visit_f64<E>(self, value: f64) -> Result<Self::Value, E>
    where
        E: serde::de::Error,
    {
        serde_json::Number::from_f64(value)
            .map(OrderedValue::Number)
            .ok_or_else(|| E::custom("non-finite JSON number"))
    }

    fn visit_str<E>(self, value: &str) -> Result<Self::Value, E> {
        Ok(OrderedValue::String(value.to_owned()))
    }

    fn visit_string<E>(self, value: String) -> Result<Self::Value, E> {
        Ok(OrderedValue::String(value))
    }

    fn visit_seq<A>(self, mut sequence: A) -> Result<Self::Value, A::Error>
    where
        A: SeqAccess<'de>,
    {
        let mut values = Vec::new();
        while let Some(value) = sequence.next_element()? {
            values.push(value);
        }
        Ok(OrderedValue::Array(values))
    }

    fn visit_map<A>(self, mut object: A) -> Result<Self::Value, A::Error>
    where
        A: MapAccess<'de>,
    {
        let mut entries = Vec::new();
        while let Some(entry) = object.next_entry()? {
            entries.push(entry);
        }
        Ok(OrderedValue::Object(entries))
    }
}

impl<'de> Deserialize<'de> for OrderedValue {
    fn deserialize<D>(deserializer: D) -> Result<Self, D::Error>
    where
        D: serde::Deserializer<'de>,
    {
        deserializer.deserialize_any(OrderedValueVisitor)
    }
}

#[derive(Default)]
struct StyleResolver {
    enabled: bool,
    styles: BTreeMap<String, Value>,
    doc_defaults: Option<Value>,
    default_paragraph: Option<String>,
    default_table: Option<String>,
    default_character: Option<String>,
}

fn object(value: Option<&Value>) -> Option<&Map<String, Value>> {
    value?.as_object()
}

fn array(value: Option<&Value>) -> &[Value] {
    value
        .and_then(Value::as_array)
        .map(Vec::as_slice)
        .unwrap_or(&[])
}

fn field<'a>(value: Option<&'a Value>, key: &str) -> Option<&'a Value> {
    object(value)?.get(key)
}

fn string(value: Option<&Value>) -> Option<&str> {
    value?.as_str()
}

fn number(value: Option<&Value>) -> Option<f64> {
    value?.as_f64()
}

fn boolean(value: Option<&Value>) -> Option<bool> {
    value?.as_bool()
}

fn truthy(value: Option<&Value>) -> bool {
    match value {
        None | Some(Value::Null) => false,
        Some(Value::Bool(value)) => *value,
        Some(Value::Number(value)) => value.as_f64().is_some_and(|value| value != 0.0),
        Some(Value::String(value)) => !value.is_empty(),
        Some(Value::Array(_) | Value::Object(_)) => true,
    }
}

fn nullish(value: Option<&Value>) -> Value {
    value.cloned().unwrap_or(Value::Null)
}

fn js_string(value: &Value) -> String {
    match value {
        Value::String(value) => value.clone(),
        Value::Number(value) => {
            if let Some(number) = value.as_f64()
                && number.fract() == 0.0
                && number.abs() <= 9_007_199_254_740_991.0
            {
                return format!("{number:.0}");
            }
            value.to_string()
        }
        Value::Bool(value) => value.to_string(),
        Value::Null => "null".to_owned(),
        _ => serde_json::to_string(value).unwrap_or_default(),
    }
}

fn utf16_len(value: &str) -> u32 {
    value.encode_utf16().count() as u32
}

fn ordered_object(
    entries: impl IntoIterator<Item = (impl Into<String>, Value)>,
) -> Vec<(String, Value)> {
    entries
        .into_iter()
        .map(|(key, value)| (key.into(), value))
        .collect()
}

fn ordered_json(entries: &[(String, Value)]) -> String {
    let mut output = String::from("{");
    for (index, (key, value)) in entries.iter().enumerate() {
        if index > 0 {
            output.push(',');
        }
        output.push_str(&serde_json::to_string(key).unwrap());
        output.push(':');
        output.push_str(&js_json(value));
    }
    output.push('}');
    output
}

fn js_json(value: &Value) -> String {
    match value {
        Value::Null => "null".to_owned(),
        Value::Bool(value) => value.to_string(),
        Value::Number(value) => {
            let value = value.as_f64().unwrap_or_default();
            if value == 0.0 {
                "0".to_owned()
            } else {
                ryu_js::Buffer::new().format(value).to_owned()
            }
        }
        Value::String(value) => serde_json::to_string(value).unwrap(),
        Value::Array(values) => format!(
            "[{}]",
            values.iter().map(js_json).collect::<Vec<_>>().join(",")
        ),
        Value::Object(values) => format!(
            "{{{}}}",
            values
                .iter()
                .map(|(key, value)| format!(
                    "{}:{}",
                    serde_json::to_string(key).unwrap(),
                    js_json(value)
                ))
                .collect::<Vec<_>>()
                .join(",")
        ),
    }
}

impl OrderedValue {
    fn value(&self) -> Value {
        match self {
            Self::Null => Value::Null,
            Self::Bool(value) => Value::Bool(*value),
            Self::Number(value) => Value::Number(value.clone()),
            Self::String(value) => Value::String(value.clone()),
            Self::Array(values) => Value::Array(values.iter().map(Self::value).collect()),
            Self::Object(entries) => Value::Object(
                entries
                    .iter()
                    .map(|(key, value)| (key.clone(), value.value()))
                    .collect(),
            ),
        }
    }

    fn js_json(&self) -> String {
        match self {
            Self::Null => "null".to_owned(),
            Self::Bool(value) => value.to_string(),
            Self::Number(value) => {
                let value = value.as_f64().unwrap_or_default();
                if value == 0.0 {
                    "0".to_owned()
                } else {
                    ryu_js::Buffer::new().format(value).to_owned()
                }
            }
            Self::String(value) => serde_json::to_string(value).unwrap(),
            Self::Array(values) => format!(
                "[{}]",
                values
                    .iter()
                    .map(Self::js_json)
                    .collect::<Vec<_>>()
                    .join(",")
            ),
            Self::Object(entries) => format!(
                "{{{}}}",
                entries
                    .iter()
                    .map(|(key, value)| format!(
                        "{}:{}",
                        serde_json::to_string(key).unwrap(),
                        value.js_json()
                    ))
                    .collect::<Vec<_>>()
                    .join(",")
            ),
        }
    }

    fn insert_source_json(&self, output: &mut BTreeMap<String, String>) {
        output
            .entry(serde_json::to_string(&self.value()).unwrap())
            .or_insert_with(|| self.js_json());
    }

    fn collect_source_json(&self, output: &mut BTreeMap<String, String>) {
        match self {
            Self::Array(values) => {
                for value in values {
                    value.collect_source_json(output);
                }
            }
            Self::Object(entries) => {
                let node_type = entries.iter().find_map(|(key, value)| {
                    (key == "type")
                        .then_some(value)
                        .and_then(|value| match value {
                            Self::String(value) => Some(value.as_str()),
                            _ => None,
                        })
                });
                if matches!(
                    node_type,
                    Some("simpleField" | "complexField" | "shape" | "chart")
                ) {
                    self.insert_source_json(output);
                }
                if matches!(node_type, Some("inlineSdt" | "blockSdt"))
                    && let Some((_, properties)) =
                        entries.iter().find(|(key, _)| key == "properties")
                {
                    properties.insert_source_json(output);
                    if let Self::Object(properties) = properties {
                        for (_, value) in properties
                            .iter()
                            .filter(|(key, _)| matches!(key.as_str(), "listItems" | "dataBinding"))
                        {
                            value.insert_source_json(output);
                        }
                    }
                }
                for (_, value) in entries {
                    value.collect_source_json(output);
                }
            }
            _ => {}
        }
    }
}

fn needs_source_json(value: &Value) -> bool {
    match value {
        Value::Array(values) => values.iter().any(needs_source_json),
        Value::Object(values) => {
            matches!(
                string(values.get("type")),
                Some("simpleField" | "complexField" | "shape" | "chart" | "inlineSdt" | "blockSdt")
            ) || values.values().any(needs_source_json)
        }
        _ => false,
    }
}

fn source_json(value: &Value, values: &BTreeMap<String, String>) -> String {
    serde_json::to_string(value)
        .ok()
        .and_then(|key| values.get(&key).cloned())
        .unwrap_or_else(|| js_json(value))
}

fn drop_nulls(value: Value) -> Value {
    match value {
        Value::Array(values) => Value::Array(values.into_iter().map(drop_nulls).collect()),
        Value::Object(values) => Value::Object(
            values
                .into_iter()
                .filter(|(_, value)| !value.is_null())
                .map(|(key, value)| (key, drop_nulls(value)))
                .collect(),
        ),
        value => value,
    }
}

fn map_from_value(value: Value) -> JsonObject {
    match drop_nulls(value) {
        Value::Object(value) => value.into_iter().collect(),
        _ => JsonObject::new(),
    }
}

fn value_from_map(value: &JsonObject) -> Value {
    Value::Object(
        value
            .iter()
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect(),
    )
}

fn any_from_value(value: Value) -> Result<Any, String> {
    match value {
        Value::Null => Ok(Any::Null),
        Value::Bool(value) => Ok(Any::Bool(value)),
        Value::Number(value) if value.is_i64() => Ok(Any::from(value.as_i64().unwrap())),
        Value::Number(value) if value.is_u64() => Any::try_from(value.as_u64().unwrap())
            .map_err(|value| format!("JSON number {value} exceeds the yrs integer range")),
        Value::Number(value) => Ok(Any::Number(
            value
                .as_f64()
                .ok_or_else(|| format!("invalid JSON number {value}"))?,
        )),
        Value::String(value) => Ok(Any::String(Arc::from(value))),
        Value::Array(values) => values
            .into_iter()
            .map(any_from_value)
            .collect::<Result<Vec<_>, _>>()
            .map(Arc::from)
            .map(Any::Array),
        Value::Object(values) => values
            .into_iter()
            .map(|(key, value)| Ok((key, any_from_value(value)?)))
            .collect::<Result<HashMap<_, _>, _>>()
            .map(Arc::new)
            .map(Any::Map),
    }
}

fn yrs_attrs(values: JsonObject) -> Result<Attrs, String> {
    values
        .into_iter()
        .map(|(key, value)| Ok((Arc::<str>::from(key), any_from_value(value)?)))
        .collect()
}

fn payload(values: JsonObject) -> Result<Vec<(String, Any)>, String> {
    values
        .into_iter()
        .map(|(key, value)| Ok((key, any_from_value(value)?)))
        .collect()
}

fn merge_plain(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
    match (object(target), object(source)) {
        (None, None) => None,
        (Some(target), None) => Some(Value::Object(target.clone())),
        (None, Some(source)) => Some(Value::Object(source.clone())),
        (Some(target), Some(source)) => {
            let mut result = target.clone();
            for (key, value) in source {
                result.insert(key.clone(), value.clone());
            }
            Some(Value::Object(result))
        }
    }
}

fn merge_font_family(target: Option<&Value>, source: &Value) -> Value {
    let mut result = object(target).cloned().unwrap_or_default();
    let source = source.as_object().cloned().unwrap_or_default();
    for (explicit, theme) in [
        ("ascii", "asciiTheme"),
        ("hAnsi", "hAnsiTheme"),
        ("eastAsia", "eastAsiaTheme"),
        ("cs", "csTheme"),
    ] {
        if source.contains_key(explicit) || source.contains_key(theme) {
            result.remove(explicit);
            result.remove(theme);
            if let Some(value) = source.get(explicit) {
                result.insert(explicit.to_owned(), value.clone());
            }
            if let Some(value) = source.get(theme) {
                result.insert(theme.to_owned(), value.clone());
            }
        }
    }
    for (key, value) in source {
        if ![
            "ascii",
            "asciiTheme",
            "hAnsi",
            "hAnsiTheme",
            "eastAsia",
            "eastAsiaTheme",
            "cs",
            "csTheme",
        ]
        .contains(&key.as_str())
        {
            result.insert(key, value);
        }
    }
    Value::Object(result)
}

fn merge_text_formatting(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
    let target_object = object(target);
    let source_object = object(source);
    if source_object.is_none() {
        return target.cloned();
    }
    if target_object.is_none() {
        return source.cloned();
    }
    let target_object = target_object.unwrap();
    let mut result = target_object.clone();
    for (key, value) in source_object.unwrap() {
        if key == "fontFamily" && value.is_object() {
            result.insert(
                key.clone(),
                merge_font_family(target_object.get(key), value),
            );
        } else if key == "color" && value.is_object() {
            let explicit = truthy(field(Some(value), "rgb"))
                || truthy(field(Some(value), "themeColor"))
                || truthy(field(Some(value), "themeTint"))
                || truthy(field(Some(value), "themeShade"));
            if !truthy(field(Some(value), "auto")) || explicit {
                result.insert(key.clone(), value.clone());
            }
        } else if value.is_object() {
            result.insert(
                key.clone(),
                merge_plain(target_object.get(key), Some(value)).unwrap(),
            );
        } else {
            result.insert(key.clone(), value.clone());
        }
    }
    Some(Value::Object(result))
}

fn merge_paragraph_formatting(target: Option<&Value>, source: Option<&Value>) -> Option<Value> {
    let Some(source) = object(source) else {
        return target.cloned();
    };
    let mut result = object(target).cloned().unwrap_or_default();
    for (key, value) in source {
        if key == "runProperties" {
            if let Some(merged) = merge_text_formatting(result.get(key), Some(value)) {
                result.insert(key.clone(), merged);
            }
        } else if ["borders", "numPr", "frame"].contains(&key.as_str()) {
            result.insert(
                key.clone(),
                merge_plain(result.get(key), Some(value)).unwrap_or_else(|| value.clone()),
            );
        } else {
            result.insert(key.clone(), value.clone());
        }
    }
    Some(Value::Object(result))
}

impl StyleResolver {
    fn new(definitions: Option<&Value>) -> Self {
        let Some(definitions) = object(definitions) else {
            return Self::default();
        };
        let mut resolver = Self {
            enabled: true,
            doc_defaults: definitions.get("docDefaults").cloned(),
            ..Self::default()
        };
        for style in array(definitions.get("styles")) {
            let Some(style_id) = string(field(Some(style), "styleId")) else {
                continue;
            };
            resolver.styles.insert(style_id.to_owned(), style.clone());
        }
        resolver.default_paragraph = resolver.find_default("paragraph").or_else(|| {
            resolver
                .styles
                .contains_key("Normal")
                .then(|| "Normal".to_owned())
        });
        resolver.default_table = resolver.find_default("table");
        resolver.default_character = resolver.find_default("character");
        resolver
    }

    fn find_default(&self, style_type: &str) -> Option<String> {
        self.styles.iter().find_map(|(id, style)| {
            (string(field(Some(style), "type")) == Some(style_type)
                && truthy(field(Some(style), "default")))
            .then(|| id.clone())
        })
    }

    fn style(&self, style_id: &str) -> Option<&Value> {
        self.styles.get(style_id)
    }

    fn default_style(&self, style_type: &str) -> Option<&Value> {
        let id = match style_type {
            "paragraph" => self.default_paragraph.as_deref(),
            "table" => self.default_table.as_deref(),
            "character" => self.default_character.as_deref(),
            _ => None,
        };
        id.and_then(|id| self.style(id))
    }

    fn resolve_paragraph_style(&self, style_id: Option<&str>) -> (Option<Value>, Option<Value>) {
        let mut paragraph = field(self.doc_defaults.as_ref(), "pPr").cloned();
        let mut run = field(self.doc_defaults.as_ref(), "rPr").cloned();
        let style = style_id
            .and_then(|id| self.style(id))
            .or_else(|| self.default_style("paragraph"));
        if let Some(style) = style {
            paragraph = merge_paragraph_formatting(paragraph.as_ref(), field(Some(style), "pPr"));
            run = merge_text_formatting(run.as_ref(), field(Some(style), "rPr"));
        }
        if style_id.is_some() && style.is_none() {
            if let Some(style) = self.default_style("paragraph") {
                paragraph =
                    merge_paragraph_formatting(paragraph.as_ref(), field(Some(style), "pPr"));
                run = merge_text_formatting(run.as_ref(), field(Some(style), "rPr"));
            }
        }
        if style_id.is_none() && style.is_none() && self.doc_defaults.is_none() {
            paragraph = Some(json!({
                "spaceAfter": 160,
                "lineSpacing": 259,
                "lineSpacingRule": "auto"
            }));
        }
        (paragraph, run)
    }

    fn resolve_run_style(&self, style_id: Option<&str>) -> Option<Value> {
        let mut result = field(self.doc_defaults.as_ref(), "rPr").cloned();
        result = merge_text_formatting(
            result.as_ref(),
            self.default_style("character")
                .and_then(|style| field(Some(style), "rPr")),
        );
        if let Some(style) = style_id.and_then(|id| self.style(id)) {
            result = merge_text_formatting(result.as_ref(), field(Some(style), "rPr"));
        }
        result
    }

    fn run_style_own(&self, style_id: Option<&str>) -> Option<Value> {
        style_id
            .and_then(|id| self.style(id))
            .and_then(|style| field(Some(style), "rPr"))
            .cloned()
    }
}

fn mark(name: &str, attrs: Vec<(String, Value)>) -> Mark {
    Mark {
        name: name.to_owned(),
        attrs,
    }
}

fn formatting_to_marks(formatting: Option<&Value>) -> Vec<Mark> {
    let mut marks = Vec::new();
    let Some(formatting) = object(formatting) else {
        return marks;
    };
    if truthy(formatting.get("bold")) {
        marks.push(mark("bold", vec![]));
    }
    if truthy(formatting.get("italic")) {
        marks.push(mark("italic", vec![]));
    }
    if let Some(underline) = object(formatting.get("underline"))
        && string(underline.get("style")) != Some("none")
    {
        marks.push(mark(
            "underline",
            ordered_object([
                (
                    "style",
                    underline.get("style").cloned().unwrap_or(Value::Null),
                ),
                ("color", nullish(underline.get("color"))),
            ]),
        ));
    }
    if truthy(formatting.get("strike")) || truthy(formatting.get("doubleStrike")) {
        marks.push(mark(
            "strike",
            ordered_object([(
                "double",
                Value::Bool(truthy(formatting.get("doubleStrike"))),
            )]),
        ));
    }
    if let Some(color) = object(formatting.get("color"))
        && !truthy(color.get("auto"))
    {
        marks.push(mark(
            "textColor",
            ordered_object([
                ("rgb", nullish(color.get("rgb"))),
                ("themeColor", nullish(color.get("themeColor"))),
                ("themeTint", nullish(color.get("themeTint"))),
                ("themeShade", nullish(color.get("themeShade"))),
            ]),
        ));
    }
    let shading_fill =
        object(formatting.get("shading")).and_then(|shading| object(shading.get("fill")));
    let shading_highlight = shading_fill.and_then(|fill| {
        let pattern =
            string(object(formatting.get("shading")).and_then(|shading| shading.get("pattern")));
        (pattern.is_none() || pattern == Some("clear"))
            .then(|| string(fill.get("rgb")))
            .flatten()
            .filter(|_| !truthy(fill.get("auto")))
            .map(|rgb| {
                if rgb.starts_with('#') {
                    rgb.to_owned()
                } else {
                    format!("#{rgb}")
                }
            })
    });
    let highlight = string(formatting.get("highlight"))
        .filter(|value| *value != "none")
        .map(str::to_owned)
        .or(shading_highlight);
    if let Some(highlight) = highlight {
        marks.push(mark(
            "highlight",
            ordered_object([("color", Value::String(highlight))]),
        ));
    }
    if formatting.contains_key("fontSize") || formatting.contains_key("fontSizeCs") {
        marks.push(mark(
            "fontSize",
            ordered_object([
                ("size", nullish(formatting.get("fontSize"))),
                ("sizeCs", nullish(formatting.get("fontSizeCs"))),
            ]),
        ));
    }
    if let Some(font) = object(formatting.get("fontFamily")) {
        marks.push(mark(
            "fontFamily",
            ordered_object([
                ("ascii", nullish(font.get("ascii"))),
                ("hAnsi", nullish(font.get("hAnsi"))),
                ("eastAsia", nullish(font.get("eastAsia"))),
                ("cs", nullish(font.get("cs"))),
                ("asciiTheme", nullish(font.get("asciiTheme"))),
                ("hAnsiTheme", nullish(font.get("hAnsiTheme"))),
                ("eastAsiaTheme", nullish(font.get("eastAsiaTheme"))),
                ("csTheme", nullish(font.get("csTheme"))),
            ]),
        ));
    }
    match string(formatting.get("vertAlign")) {
        Some("superscript") => marks.push(mark("superscript", vec![])),
        Some("subscript") => marks.push(mark("subscript", vec![])),
        _ => {}
    }
    for (key, name) in [
        ("allCaps", "allCaps"),
        ("smallCaps", "smallCaps"),
        ("emboss", "emboss"),
        ("imprint", "imprint"),
        ("shadow", "textShadow"),
        ("outline", "textOutline"),
        ("hidden", "hidden"),
        ("rtl", "rtl"),
    ] {
        if truthy(formatting.get(key)) {
            marks.push(mark(name, vec![]));
        }
    }
    if ["spacing", "position", "scale", "kerning"]
        .iter()
        .any(|key| formatting.contains_key(*key))
    {
        marks.push(mark(
            "characterSpacing",
            ordered_object([
                ("spacing", nullish(formatting.get("spacing"))),
                ("position", nullish(formatting.get("position"))),
                ("scale", nullish(formatting.get("scale"))),
                ("kerning", nullish(formatting.get("kerning"))),
            ]),
        ));
    }
    if let Some(value) = string(formatting.get("emphasisMark")).filter(|value| *value != "none") {
        marks.push(mark(
            "emphasisMark",
            ordered_object([("type", Value::String(value.to_owned()))]),
        ));
    }
    if let Some(value) = string(formatting.get("effect")).filter(|value| *value != "none") {
        marks.push(mark(
            "textEffect",
            ordered_object([("effect", Value::String(value.to_owned()))]),
        ));
    }
    if let Some(value) = formatting.get("modernEffects") {
        marks.push(mark(
            "modernTextEffects",
            ordered_object([("effects", value.clone())]),
        ));
    }
    if let Some(value) = formatting.get("styleId") {
        marks.push(mark(
            "runStyle",
            ordered_object([("styleId", value.clone())]),
        ));
    }
    marks
}

fn mark_attrs(mark: &Mark) -> Value {
    Value::Object(mark.attrs.iter().cloned().collect())
}

fn marks_to_attrs(marks: &[Mark]) -> JsonObject {
    let boolean_marks = [
        "bold",
        "italic",
        "superscript",
        "subscript",
        "allCaps",
        "smallCaps",
        "emboss",
        "imprint",
        "textShadow",
        "textOutline",
        "hidden",
        "rtl",
    ];
    let mut attrs = JsonObject::new();
    for mark in marks {
        if mark.name == "comment" || mark.name == "footnoteRef" {
            continue;
        }
        if boolean_marks.contains(&mark.name.as_str()) {
            attrs.insert(mark.name.clone(), Value::Bool(true));
        } else if mark.name == "highlight" {
            attrs.insert(
                "highlight".to_owned(),
                mark.attrs
                    .iter()
                    .find(|(key, _)| key == "color")
                    .map(|(_, value)| value.clone())
                    .unwrap_or(Value::Null),
            );
        } else if mark.name == "insertion" || mark.name == "deletion" {
            let get = |name: &str| {
                mark.attrs
                    .iter()
                    .find(|(key, _)| key == name)
                    .map(|(_, value)| value.clone())
                    .unwrap_or(Value::Null)
            };
            attrs.insert(
                if mark.name == "insertion" {
                    "ins".to_owned()
                } else {
                    "del".to_owned()
                },
                drop_nulls(json!({
                    "id": get("revisionId"),
                    "author": get("author"),
                    "date": get("date")
                })),
            );
        } else {
            attrs.insert(mark.name.clone(), drop_nulls(mark_attrs(mark)));
        }
    }
    attrs
}

fn marks_key(marks: &[Mark]) -> String {
    let mut values: Vec<String> = marks
        .iter()
        .filter(|mark| mark.name != "hyperlink" && mark.name != "comment")
        .map(|mark| format!("{}:{}", mark.name, ordered_json(&mark.attrs)))
        .collect();
    values.sort();
    values.join("|")
}

fn with_mark(marks: &[Mark], next: Mark) -> Vec<Mark> {
    let name = next.name.clone();
    marks
        .iter()
        .filter(|mark| mark.name != name)
        .cloned()
        .chain(std::iter::once(next))
        .collect()
}

fn text_unit(text: String, marks: &[Mark], comment_id: Option<String>) -> InlineUnit {
    InlineUnit {
        pm_size: utf16_len(&text),
        content: UnitContent::Text(text),
        attrs: marks_to_attrs(marks),
        comment_id,
        marks: marks.to_vec(),
    }
}

fn embed_unit(
    kind: &str,
    payload: JsonObject,
    marks: &[Mark],
    comment_id: Option<String>,
    pm_size: u32,
) -> InlineUnit {
    InlineUnit {
        content: UnitContent::Embed {
            kind: kind.to_owned(),
            payload,
        },
        attrs: marks_to_attrs(marks),
        pm_size,
        comment_id,
        marks: marks.to_vec(),
    }
}

fn run_marks(run: &Value, style_formatting: Option<&Value>, styles: &StyleResolver) -> Vec<Mark> {
    let formatting = field(Some(run), "formatting");
    let run_style = styles.run_style_own(
        formatting.and_then(|formatting| string(field(Some(formatting), "styleId"))),
    );
    let inherited = merge_text_formatting(style_formatting, run_style.as_ref());
    let merged = merge_text_formatting(inherited.as_ref(), formatting);
    formatting_to_marks(merged.as_ref())
}

fn emu_to_pixels(value: f64) -> f64 {
    value / 914_400.0 * 96.0
}

fn image_payload(image: &Value) -> JsonObject {
    let size = field(Some(image), "size");
    let wrap = field(Some(image), "wrap");
    let position = field(Some(image), "position");
    let transform = field(Some(image), "transform");
    let outline = field(Some(image), "outline");
    let wrap_type = string(field(wrap, "type")).unwrap_or_default();
    let wrap_text = string(field(wrap, "wrapText"));
    let horizontal = field(position, "horizontal");
    let vertical = field(position, "vertical");
    let alignment = string(field(horizontal, "alignment"));
    let css_float = if wrap_type == "inline" || wrap_type == "topAndBottom" {
        "none"
    } else if ["square", "tight", "through"].contains(&wrap_type) {
        if wrap_text == Some("left") {
            "right"
        } else if wrap_text == Some("right") {
            "left"
        } else if matches!(alignment, Some("left" | "right")) {
            alignment.unwrap()
        } else {
            "none"
        }
    } else {
        "none"
    };
    let display_mode = if wrap_type == "inline" {
        "inline"
    } else if wrap_type == "topAndBottom" {
        "block"
    } else if matches!(wrap_type, "behind" | "inFront") || css_float != "none" {
        "float"
    } else {
        "block"
    };
    let mut transforms = Vec::new();
    if let Some(rotation) = number(field(transform, "rotation")).filter(|value| *value != 0.0) {
        transforms.push(format!("rotate({rotation}deg)"));
    }
    if truthy(field(transform, "flipH")) {
        transforms.push("scaleX(-1)".to_owned());
    }
    if truthy(field(transform, "flipV")) {
        transforms.push("scaleY(-1)".to_owned());
    }
    let outline_width = number(field(outline, "width")).filter(|value| *value != 0.0);
    let border_width =
        outline_width.map(|value| (value / 914_400.0 * 96.0 * 100.0).round() / 100.0);
    let border_color =
        string(field(field(outline, "color"), "rgb")).map(|value| format!("#{value}"));
    let border_style = outline_width.map(|_| match string(field(outline, "style")) {
        Some("dot" | "sysDot") => "dotted",
        Some(
            "dash" | "lgDash" | "dashDot" | "lgDashDot" | "lgDashDotDot" | "sysDash" | "sysDashDot"
            | "sysDashDotDot",
        ) => "dashed",
        _ => "solid",
    });
    let axis = |axis: Option<&Value>| {
        axis.map(|axis| {
            json!({
                "relativeTo": nullish(field(Some(axis), "relativeTo")),
                "posOffset": nullish(field(Some(axis), "posOffset")),
                "align": nullish(field(Some(axis), "alignment"))
            })
        })
    };
    map_from_value(json!({
        "src": string(field(Some(image), "src")).unwrap_or_default(),
        "alt": nullish(field(Some(image), "alt")),
        "title": nullish(field(Some(image), "title")),
        "width": number(field(size, "width")).filter(|value| *value != 0.0).map(emu_to_pixels),
        "height": number(field(size, "height")).filter(|value| *value != 0.0).map(emu_to_pixels),
        "rId": nullish(field(Some(image), "rId")),
        "wrapType": wrap_type,
        "displayMode": display_mode,
        "cssFloat": css_float,
        "transform": (!transforms.is_empty()).then(|| transforms.join(" ")),
        "distTop": number(field(wrap, "distT")).map(emu_to_pixels),
        "distBottom": number(field(wrap, "distB")).map(emu_to_pixels),
        "distLeft": number(field(wrap, "distL")).map(emu_to_pixels),
        "distRight": number(field(wrap, "distR")).map(emu_to_pixels),
        "position": position.map(|_| json!({
            "horizontal": axis(horizontal),
            "vertical": axis(vertical)
        })),
        "borderWidth": border_width,
        "borderColor": border_color,
        "borderStyle": border_style,
        "wrapText": wrap_text,
        "hlinkHref": nullish(field(Some(image), "hlinkHref")),
        "cropTop": nullish(field(field(Some(image), "crop"), "top")),
        "cropRight": nullish(field(field(Some(image), "crop"), "right")),
        "cropBottom": nullish(field(field(Some(image), "crop"), "bottom")),
        "cropLeft": nullish(field(field(Some(image), "crop"), "left")),
        "opacity": nullish(field(Some(image), "opacity")),
        "effectExtentTop": number(field(field(Some(image), "padding"), "top"))
            .filter(|value| *value != 0.0)
            .map(emu_to_pixels),
        "effectExtentBottom": number(field(field(Some(image), "padding"), "bottom"))
            .filter(|value| *value != 0.0)
            .map(emu_to_pixels),
        "effectExtentLeft": number(field(field(Some(image), "padding"), "left"))
            .filter(|value| *value != 0.0)
            .map(emu_to_pixels),
        "effectExtentRight": number(field(field(Some(image), "padding"), "right"))
            .filter(|value| *value != 0.0)
            .map(emu_to_pixels),
        "layoutInCell": nullish(field(Some(image), "layoutInCell")),
        "allowOverlap": nullish(field(Some(image), "allowOverlap"))
    }))
}

fn shape_payload(shape: &Value, source: &BTreeMap<String, String>) -> JsonObject {
    map_from_value(json!({ "shapeJson": source_json(shape, source) }))
}

fn chart_payload(chart: &Value, source: &BTreeMap<String, String>) -> JsonObject {
    let size = field(Some(chart), "size");
    map_from_value(json!({
        "chartJson": source_json(chart, source),
        "chartType": nullish(field(Some(chart), "chartType")),
        "title": nullish(field(Some(chart), "title")),
        "width": number(field(size, "width")).filter(|value| *value != 0.0).map(emu_to_pixels).unwrap_or(320.0),
        "height": number(field(size, "height")).filter(|value| *value != 0.0).map(emu_to_pixels).unwrap_or(220.0),
        "rId": nullish(field(Some(chart), "rId")),
        "path": nullish(field(Some(chart), "path"))
    }))
}

fn field_payload(
    field_value: &Value,
    style_formatting: Option<&Value>,
    source: &BTreeMap<String, String>,
) -> (JsonObject, Vec<Mark>) {
    let kind = string(field(Some(field_value), "type")).unwrap_or_default();
    let runs = if kind == "simpleField" {
        array(field(Some(field_value), "content"))
    } else {
        array(field(Some(field_value), "fieldResult"))
    };
    let mut display_text = String::new();
    let mut field_formatting = None;
    for child in runs {
        if string(field(Some(child), "type")) != Some("run") {
            continue;
        }
        for content in array(field(Some(child), "content")) {
            if string(field(Some(content), "type")) == Some("text") {
                display_text.push_str(string(field(Some(content), "text")).unwrap_or_default());
            }
        }
        if field_formatting.is_none() {
            field_formatting = field(Some(child), "formatting");
        }
    }
    let formatting = field_formatting.or_else(|| {
        (kind == "complexField")
            .then(|| field(Some(field_value), "formatting"))
            .flatten()
    });
    let merged = merge_text_formatting(style_formatting, formatting);
    (
        map_from_value(json!({
            "fieldType": nullish(field(Some(field_value), "fieldType")),
            "instruction": nullish(field(Some(field_value), "instruction")),
            "displayText": display_text,
            "fieldKind": if kind == "simpleField" { "simple" } else { "complex" },
            "fldLock": boolean(field(Some(field_value), "fldLock")).unwrap_or(false),
            "dirty": boolean(field(Some(field_value), "dirty")).unwrap_or(false),
            "displayMode": string(field(field(Some(field_value), "fieldTree"), "displayMode")).unwrap_or("result"),
            "hasCachedResult": !display_text.is_empty(),
            "fieldData": source_json(field_value, source),
            "modelKind": "field"
        })),
        formatting_to_marks(merged.as_ref()),
    )
}

fn math_payload(math: &Value) -> JsonObject {
    map_from_value(json!({
        "display": nullish(field(Some(math), "display")),
        "ommlXml": nullish(field(Some(math), "ommlXml")),
        "plainText": string(field(Some(math), "plainText")).unwrap_or_default()
    }))
}

fn hyperlink_mark(hyperlink: &Value) -> Mark {
    let href = string(field(Some(hyperlink), "href"))
        .filter(|value| !value.is_empty())
        .map(str::to_owned)
        .or_else(|| {
            string(field(Some(hyperlink), "anchor"))
                .filter(|value| !value.is_empty())
                .map(|value| format!("#{value}"))
        })
        .unwrap_or_default();
    mark(
        "hyperlink",
        ordered_object([
            ("href", Value::String(href)),
            ("tooltip", nullish(field(Some(hyperlink), "tooltip"))),
            ("rId", nullish(field(Some(hyperlink), "rId"))),
        ]),
    )
}

fn note_ref_unit(
    id: &Value,
    note_type: &str,
    marks: &[Mark],
    comment_id: Option<String>,
) -> InlineUnit {
    let note_mark = mark(
        "footnoteRef",
        ordered_object([
            ("id", Value::String(js_string(id))),
            ("noteType", Value::String(note_type.to_owned())),
        ]),
    );
    let all_marks: Vec<Mark> = marks
        .iter()
        .cloned()
        .chain(std::iter::once(note_mark))
        .collect();
    embed_unit(
        "noteRef",
        map_from_value(if note_type == "endnote" {
            json!({ "endnoteRefId": id })
        } else {
            json!({ "footnoteRefId": id })
        }),
        &all_marks,
        comment_id,
        utf16_len(&js_string(id)),
    )
}

fn run_content_to_units(
    content: &Value,
    marks: &[Mark],
    comment_id: Option<String>,
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    match string(field(Some(content), "type")).unwrap_or_default() {
        "text" => string(field(Some(content), "text"))
            .filter(|text| !text.is_empty())
            .map(|text| vec![text_unit(text.to_owned(), marks, comment_id)])
            .unwrap_or_default(),
        "tab" => vec![text_unit("\t".to_owned(), marks, comment_id)],
        "break"
            if string(field(Some(content), "breakType"))
                .is_none_or(|kind| kind == "textWrapping") =>
        {
            vec![embed_unit("break", JsonObject::new(), marks, comment_id, 1)]
        }
        "softHyphen" => vec![text_unit("\u{00ad}".to_owned(), marks, comment_id)],
        "noBreakHyphen" => vec![text_unit("\u{2011}".to_owned(), marks, comment_id)],
        "symbol" => {
            let Some(codepoint) = string(field(Some(content), "char"))
                .and_then(|value| u32::from_str_radix(value, 16).ok())
                .and_then(char::from_u32)
            else {
                return vec![];
            };
            let font = string(field(Some(content), "font"))
                .filter(|value| !value.is_empty())
                .map(|value| Value::String(value.to_owned()))
                .unwrap_or(Value::Null);
            let symbol_mark = mark(
                "fontFamily",
                ordered_object([
                    ("ascii", font.clone()),
                    ("hAnsi", font.clone()),
                    ("eastAsia", font.clone()),
                    ("cs", font),
                    ("asciiTheme", Value::Null),
                    ("hAnsiTheme", Value::Null),
                    ("eastAsiaTheme", Value::Null),
                    ("csTheme", Value::Null),
                ]),
            );
            vec![text_unit(
                codepoint.to_string(),
                &with_mark(marks, symbol_mark),
                comment_id,
            )]
        }
        "commentReference" => {
            let mut value = json!({
                "fieldType": "COMMENT",
                "instruction": "",
                "displayText": "",
                "fieldKind": "simple",
                "fldLock": false,
                "dirty": false,
                "displayMode": "result",
                "hasCachedResult": false,
                "modelKind": "commentReference"
            });
            if let Some(id) = field(Some(content), "id") {
                value
                    .as_object_mut()
                    .unwrap()
                    .insert("commentId".to_owned(), id.clone());
            }
            vec![embed_unit("field", map_from_value(value), &[], None, 1)]
        }
        "drawing" => vec![embed_unit(
            "image",
            image_payload(field(Some(content), "image").unwrap_or(&Value::Null)),
            &[],
            None,
            1,
        )],
        "shape" => vec![embed_unit(
            "shape",
            shape_payload(
                field(Some(content), "shape").unwrap_or(&Value::Null),
                source,
            ),
            &[],
            None,
            1,
        )],
        "chart" => vec![embed_unit(
            "chart",
            chart_payload(
                field(Some(content), "chart").unwrap_or(&Value::Null),
                source,
            ),
            &[],
            None,
            1,
        )],
        "footnoteRef" => field(Some(content), "id")
            .map(|id| note_ref_unit(id, "footnote", marks, comment_id))
            .into_iter()
            .collect(),
        "endnoteRef" => field(Some(content), "id")
            .map(|id| note_ref_unit(id, "endnote", marks, comment_id))
            .into_iter()
            .collect(),
        _ => vec![],
    }
}

fn run_to_units(
    run: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    comment_id: Option<String>,
    extra_marks: &[Mark],
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    let marks: Vec<Mark> = run_marks(run, style_formatting, styles)
        .into_iter()
        .chain(extra_marks.iter().cloned())
        .collect();
    array(field(Some(run), "content"))
        .iter()
        .flat_map(|content| run_content_to_units(content, &marks, comment_id.clone(), source))
        .collect()
}

fn hyperlink_to_units(
    hyperlink: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    extra_marks: &[Mark],
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    let mut units = Vec::new();
    let link = hyperlink_mark(hyperlink);
    let children =
        field(Some(hyperlink), "structuredChildren").or_else(|| field(Some(hyperlink), "children"));
    for child in array(children) {
        match string(field(Some(child), "type")).unwrap_or_default() {
            "run" => {
                let marks: Vec<Mark> = run_marks(child, style_formatting, styles)
                    .into_iter()
                    .chain(extra_marks.iter().cloned())
                    .chain(std::iter::once(link.clone()))
                    .collect();
                for content in array(field(Some(child), "content")) {
                    units.extend(run_content_to_units(content, &marks, None, source));
                }
            }
            "simpleField" | "complexField" => {
                let (payload, marks) = field_payload(child, style_formatting, source);
                let marks: Vec<Mark> = marks
                    .into_iter()
                    .chain(extra_marks.iter().cloned())
                    .chain(std::iter::once(link.clone()))
                    .collect();
                units.push(embed_unit("field", payload, &marks, None, 1));
            }
            "mathEquation" => {
                let marks: Vec<Mark> = extra_marks
                    .iter()
                    .cloned()
                    .chain(std::iter::once(link.clone()))
                    .collect();
                units.push(embed_unit("math", math_payload(child), &marks, None, 1));
            }
            _ => {}
        }
    }
    units
}

fn tracked_mark(info: &Value, kind: &str, is_move_pair: bool) -> Mark {
    mark(
        kind,
        ordered_object([
            ("revisionId", nullish(field(Some(info), "id"))),
            ("author", nullish(field(Some(info), "author"))),
            ("date", nullish(field(Some(info), "date"))),
            ("isMovePair", Value::Bool(is_move_pair)),
        ]),
    )
}

fn tracked_to_units(
    content: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    comment_id: Option<String>,
    source: &BTreeMap<String, String>,
) -> Vec<InlineUnit> {
    let content_type = string(field(Some(content), "type")).unwrap_or_default();
    let kind = if matches!(content_type, "insertion" | "moveTo") {
        "insertion"
    } else {
        "deletion"
    };
    let marker = tracked_mark(
        field(Some(content), "info").unwrap_or(&Value::Null),
        kind,
        matches!(content_type, "moveFrom" | "moveTo"),
    );
    let mut units = Vec::new();
    for child in array(field(Some(content), "content")) {
        if string(field(Some(child), "type")) == Some("run") {
            units.extend(run_to_units(
                child,
                style_formatting,
                styles,
                comment_id.clone(),
                std::slice::from_ref(&marker),
                source,
            ));
        } else {
            let mut linked = hyperlink_to_units(
                child,
                style_formatting,
                styles,
                std::slice::from_ref(&marker),
                source,
            );
            if let Some(comment_id) = &comment_id {
                for unit in &mut linked {
                    unit.comment_id = Some(comment_id.clone());
                }
            }
            units.extend(linked);
        }
    }
    units
}

fn sdt_properties_attrs(properties: &Value, source: &BTreeMap<String, String>) -> JsonObject {
    map_from_value(json!({
        "sdtType": nullish(field(Some(properties), "sdtType")),
        "id": nullish(field(Some(properties), "id")),
        "alias": nullish(field(Some(properties), "alias")),
        "tag": nullish(field(Some(properties), "tag")),
        "lock": nullish(field(Some(properties), "lock")),
        "placeholder": nullish(field(Some(properties), "placeholder")),
        "showingPlaceholder": boolean(field(Some(properties), "showingPlaceholder")).unwrap_or(false),
        "dateFormat": nullish(field(Some(properties), "dateFormat")),
        "listItems": field(Some(properties), "listItems").map(|value| source_json(value, source)),
        "checked": nullish(field(Some(properties), "checked")),
        "dataBinding": field(Some(properties), "dataBinding").map(|value| source_json(value, source)),
        "rawPropertiesXml": nullish(field(Some(properties), "rawPropertiesXml")),
        "rawEndPropertiesXml": nullish(field(Some(properties), "rawEndPropertiesXml"))
    }))
}

fn sdt_payload(
    sdt: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
) -> JsonObject {
    let mut content = Vec::new();
    let append = |content: &mut Vec<Value>, unit: InlineUnit| match unit.content {
        UnitContent::Text(text) if text == "\t" => {
            content.push(json!({ "kind": "tab", "attrs": value_from_map(&unit.attrs) }));
        }
        UnitContent::Text(text) => {
            if let Some(previous) = content.last_mut()
                && string(field(Some(&*previous), "kind")) == Some("text")
                && field(Some(&*previous), "attrs") == Some(&value_from_map(&unit.attrs))
            {
                let previous_text = previous
                    .as_object_mut()
                    .and_then(|value| value.get_mut("text"))
                    .and_then(|value| value.as_str())
                    .unwrap_or_default()
                    .to_owned();
                previous.as_object_mut().unwrap().insert(
                    "text".to_owned(),
                    Value::String(format!("{previous_text}{text}")),
                );
            } else {
                content.push(json!({
                    "kind": "text",
                    "text": text,
                    "attrs": value_from_map(&unit.attrs)
                }));
            }
        }
        UnitContent::Embed { kind, payload } => {
            content.push(json!({
                "kind": kind,
                "payload": value_from_map(&payload),
                "attrs": value_from_map(&unit.attrs)
            }));
        }
    };
    for child in array(field(Some(sdt), "content")) {
        match string(field(Some(child), "type")).unwrap_or_default() {
            "run" => {
                for unit in run_to_units(child, style_formatting, styles, None, &[], source) {
                    append(&mut content, unit);
                }
            }
            "hyperlink" => {
                for unit in hyperlink_to_units(child, style_formatting, styles, &[], source) {
                    append(&mut content, unit);
                }
            }
            "simpleField" | "complexField" => {
                let (payload, marks) = field_payload(child, style_formatting, source);
                append(&mut content, embed_unit("field", payload, &marks, None, 1));
            }
            "inlineSdt" => append(
                &mut content,
                embed_unit(
                    "sdt",
                    sdt_payload(child, style_formatting, styles, source),
                    &[],
                    None,
                    1,
                ),
            ),
            "mathEquation" => append(
                &mut content,
                embed_unit("math", math_payload(child), &[], None, 1),
            ),
            _ => {}
        }
    }
    let properties = field(Some(sdt), "properties").unwrap_or(&Value::Null);
    let mut result = sdt_properties_attrs(properties, source);
    result.insert(
        "propertiesJson".to_owned(),
        Value::String(source_json(properties, source)),
    );
    result.insert("content".to_owned(), Value::Array(content));
    result
}

fn paragraph_style_formatting(
    paragraph: &Value,
    styles: &StyleResolver,
    extra: Option<&Value>,
) -> Option<Value> {
    let style_id = string(field(field(Some(paragraph), "formatting"), "styleId"));
    let style = styles
        .enabled
        .then(|| styles.resolve_paragraph_style(style_id).1)
        .flatten();
    merge_text_formatting(style.as_ref(), extra)
}

fn run_boundary(
    run: &Value,
    style_formatting: Option<&Value>,
    styles: &StyleResolver,
    source: &BTreeMap<String, String>,
) -> Option<Value> {
    let units = run_to_units(run, style_formatting, styles, None, &[], source);
    if units
        .iter()
        .any(|unit| matches!(&unit.content, UnitContent::Embed { kind, .. } if kind != "noteRef"))
    {
        return None;
    }
    let keys: Vec<_> = units.iter().map(|unit| marks_key(&unit.marks)).collect();
    if keys
        .first()
        .is_some_and(|first| keys.iter().any(|key| key != first))
    {
        return None;
    }
    let text = units
        .iter()
        .map(|unit| match &unit.content {
            UnitContent::Text(text) => text.clone(),
            UnitContent::Embed { payload, .. } => payload
                .get("footnoteRefId")
                .or_else(|| payload.get("endnoteRefId"))
                .map(js_string)
                .unwrap_or_default(),
        })
        .collect::<String>();
    let mut boundary = Map::new();
    boundary.insert("text".to_owned(), Value::String(text));
    if let Some(key) = keys.first() {
        boundary.insert("marksKey".to_owned(), Value::String(key.clone()));
    }
    if let Some(formatting) = field(Some(run), "formatting") {
        boundary.insert("formatting".to_owned(), formatting.clone());
    }
    if let Some(changes) = field(Some(run), "propertyChanges") {
        boundary.insert("propertyChanges".to_owned(), changes.clone());
    }
    Some(Value::Object(boundary))
}

fn resolved_text_formatting(formatting: Option<&Value>, styles: &StyleResolver) -> Option<Value> {
    let style = formatting
        .and_then(|formatting| string(field(Some(formatting), "styleId")))
        .and_then(|style_id| styles.resolve_run_style(Some(style_id)));
    merge_text_formatting(style.as_ref(), formatting)
}

fn paragraph_attrs(
    paragraph: &Value,
    styles: &StyleResolver,
    units: &[InlineUnit],
    unit_counts: &[usize],
    run_boundaries: Option<Vec<Value>>,
) -> JsonObject {
    let formatting = field(Some(paragraph), "formatting");
    let style_id = string(field(formatting, "styleId"));
    let list = field(Some(paragraph), "listRendering");
    let mut attrs = map_from_value(json!({
        "paraId": nullish(field(Some(paragraph), "paraId")),
        "textId": nullish(field(Some(paragraph), "textId")),
        "styleId": style_id,
        "numPr": nullish(field(formatting, "numPr")),
        "numPrFromStyle": nullish(field(formatting, "numPrFromStyle")),
        "listNumFmt": nullish(field(list, "numFmt")),
        "listIsBullet": nullish(field(list, "isBullet")),
        "listMarker": nullish(field(list, "marker")),
        "listMarkerHidden": truthy(field(list, "markerHidden")).then(|| field(list, "markerHidden").cloned()).flatten(),
        "listMarkerFontFamily": string(field(list, "markerFontFamily")).filter(|value| !value.is_empty()),
        "listMarkerFontSize": number(field(list, "markerFontSize")).filter(|value| *value != 0.0),
        "listMarkerSuffix": string(field(list, "markerSuffix")).filter(|value| !value.is_empty()),
        "listLevelNumFmts": truthy(field(list, "levelNumFmts")).then(|| field(list, "levelNumFmts").cloned()).flatten(),
        "listAbstractNumId": nullish(field(list, "abstractNumId")),
        "listStartOverride": nullish(field(list, "startOverride")),
        "_originalFormatting": nullish(formatting)
    }));
    if styles.enabled {
        let (style_ppr, resolved_run) = styles.resolve_paragraph_style(style_id);
        let style_ppr_ref = style_ppr.as_ref();
        for key in [
            "alignment",
            "spaceBefore",
            "spaceAfter",
            "lineSpacing",
            "lineSpacingRule",
            "indentLeft",
            "indentRight",
            "borders",
            "shading",
            "tabs",
            "pageBreakBefore",
            "keepNext",
            "keepLines",
            "contextualSpacing",
            "outlineLevel",
            "bidi",
        ] {
            attrs.insert(
                key.to_owned(),
                field(formatting, key)
                    .or_else(|| field(style_ppr_ref, key))
                    .cloned()
                    .unwrap_or(Value::Null),
            );
        }
        attrs.insert(
            "spacingExplicit".to_owned(),
            truthy(field(formatting, "spacingExplicit"))
                .then(|| field(formatting, "spacingExplicit").cloned())
                .flatten()
                .unwrap_or(Value::Null),
        );
        let numbering_removed = number(field(field(formatting, "numPr"), "numId")) == Some(0.0)
            && field(style_ppr_ref, "numPr").is_some()
            && number(field(field(style_ppr_ref, "numPr"), "numId")) != Some(0.0);
        attrs.insert(
            "indentFirstLine".to_owned(),
            field(formatting, "indentFirstLine")
                .or_else(|| {
                    (!numbering_removed)
                        .then(|| field(style_ppr_ref, "indentFirstLine"))
                        .flatten()
                })
                .cloned()
                .unwrap_or(Value::Null),
        );
        attrs.insert(
            "hangingIndent".to_owned(),
            field(formatting, "hangingIndent")
                .or_else(|| {
                    (!numbering_removed)
                        .then(|| field(style_ppr_ref, "hangingIndent"))
                        .flatten()
                })
                .cloned()
                .unwrap_or(Value::Bool(false)),
        );
        let default_character = styles
            .default_style("character")
            .and_then(|style| field(Some(style), "rPr"));
        let style_rpr = if default_character.is_some() {
            merge_text_formatting(resolved_run.as_ref(), default_character)
        } else {
            resolved_run
        };
        let direct = resolved_text_formatting(field(formatting, "runProperties"), styles);
        attrs.insert(
            "defaultTextFormatting".to_owned(),
            merge_text_formatting(style_rpr.as_ref(), direct.as_ref()).unwrap_or(Value::Null),
        );
        if field(formatting, "numPr").is_none()
            && field(style_ppr_ref, "numPr").is_some()
            && number(field(field(style_ppr_ref, "numPr"), "numId")) != Some(0.0)
        {
            let num_pr = field(style_ppr_ref, "numPr").unwrap().clone();
            attrs.insert("numPr".to_owned(), num_pr.clone());
            attrs.insert("numPrFromStyle".to_owned(), num_pr);
        }
    } else {
        for key in [
            "alignment",
            "spaceBefore",
            "spaceAfter",
            "lineSpacing",
            "lineSpacingRule",
            "indentLeft",
            "indentRight",
            "indentFirstLine",
            "borders",
            "shading",
            "tabs",
            "pageBreakBefore",
            "keepNext",
            "keepLines",
            "outlineLevel",
            "bidi",
        ] {
            attrs.insert(
                key.to_owned(),
                field(formatting, key).cloned().unwrap_or(Value::Null),
            );
        }
        attrs.insert(
            "spacingExplicit".to_owned(),
            truthy(field(formatting, "spacingExplicit"))
                .then(|| field(formatting, "spacingExplicit").cloned())
                .flatten()
                .unwrap_or(Value::Null),
        );
        attrs.insert(
            "hangingIndent".to_owned(),
            field(formatting, "hangingIndent")
                .cloned()
                .unwrap_or(Value::Bool(false)),
        );
        attrs.insert(
            "defaultTextFormatting".to_owned(),
            field(formatting, "runProperties")
                .cloned()
                .unwrap_or(Value::Null),
        );
    }
    if let Some(section) = field(Some(paragraph), "sectionProperties") {
        attrs.insert("_sectionProperties".to_owned(), section.clone());
        if let Some(start @ ("nextPage" | "continuous" | "oddPage" | "evenPage")) =
            string(field(Some(section), "sectionStart"))
        {
            attrs.insert(
                "sectionBreakType".to_owned(),
                Value::String(start.to_owned()),
            );
        }
    }
    if truthy(field(Some(paragraph), "renderedPageBreakBefore")) {
        attrs.insert("renderedPageBreakBefore".to_owned(), Value::Bool(true));
    }
    if paragraph_starts_with_page_break(paragraph) {
        attrs.insert("pageBreakBefore".to_owned(), Value::Bool(true));
    }
    for (source, target) in [("pPrIns", "pPrIns"), ("pPrDel", "pPrDel")] {
        if let Some(info) = field(Some(paragraph), source) {
            attrs.insert(
                target.to_owned(),
                json!({
                    "revisionId": nullish(field(Some(info), "id")),
                    "author": nullish(field(Some(info), "author")),
                    "date": nullish(field(Some(info), "date"))
                }),
            );
        }
    }
    if array(field(Some(paragraph), "propertyChanges")).len() > 0 {
        attrs.insert(
            "pPrChange".to_owned(),
            field(Some(paragraph), "propertyChanges").unwrap().clone(),
        );
    }
    let mut bookmarks = Vec::new();
    let mut unit_index = 0usize;
    let mut pm_offset = 0u32;
    for (content_index, content) in array(field(Some(paragraph), "content")).iter().enumerate() {
        match string(field(Some(content), "type")).unwrap_or_default() {
            "bookmarkStart" => {
                let mut bookmark = json!({
                    "id": nullish(field(Some(content), "id")),
                    "name": nullish(field(Some(content), "name")),
                    "kind": "start",
                    "offset": pm_offset
                });
                for key in ["colFirst", "colLast"] {
                    if let Some(value) = field(Some(content), key) {
                        bookmark
                            .as_object_mut()
                            .unwrap()
                            .insert(key.to_owned(), value.clone());
                    }
                }
                bookmarks.push(bookmark);
            }
            "bookmarkEnd" => bookmarks.push(json!({
                "id": nullish(field(Some(content), "id")),
                "kind": "end",
                "offset": pm_offset
            })),
            _ => {
                for _ in 0..unit_counts.get(content_index).copied().unwrap_or_default() {
                    pm_offset += units.get(unit_index).map(|unit| unit.pm_size).unwrap_or(0);
                    unit_index += 1;
                }
            }
        }
    }
    if !bookmarks.is_empty() {
        attrs.insert("bookmarks".to_owned(), Value::Array(bookmarks));
    }
    if let Some(boundaries) = run_boundaries.filter(|boundaries| !boundaries.is_empty()) {
        attrs.insert(
            "_originalRunBoundaries".to_owned(),
            Value::Array(boundaries),
        );
    }
    attrs
}

fn para_attrs_to_ppr(attrs: JsonObject) -> JsonObject {
    attrs
        .into_iter()
        .filter(|(key, value)| {
            ![
                "paraId",
                "textId",
                "renderedPageBreakBefore",
                "numPrFromStyle",
            ]
            .contains(&key.as_str())
                && !value.is_null()
        })
        .map(|(key, value)| {
            (
                match key.as_str() {
                    "styleId" => "pStyle".to_owned(),
                    "_sectionProperties" => "sectPr".to_owned(),
                    _ => key,
                },
                drop_nulls(value),
            )
        })
        .collect()
}

fn paragraph_units(
    paragraph: &Value,
    styles: &StyleResolver,
    extra_run_formatting: Option<&Value>,
    source: &BTreeMap<String, String>,
) -> (Vec<InlineUnit>, JsonObject) {
    let mut units = Vec::new();
    let mut active_comments: Vec<String> = Vec::new();
    let mut boundaries = Some(Vec::new());
    let mut unit_counts = Vec::new();
    let style_formatting = paragraph_style_formatting(paragraph, styles, extra_run_formatting);
    for content in array(field(Some(paragraph), "content")) {
        let start = units.len();
        let comment_id = active_comments.first().cloned();
        match string(field(Some(content), "type")).unwrap_or_default() {
            "commentRangeStart" => {
                if let Some(id) = field(Some(content), "id") {
                    let id = js_string(id);
                    if !active_comments.contains(&id) {
                        active_comments.push(id);
                    }
                }
            }
            "commentRangeEnd" => {
                if let Some(id) = field(Some(content), "id") {
                    let id = js_string(id);
                    active_comments.retain(|candidate| candidate != &id);
                }
            }
            "run" => {
                let boundary = run_boundary(content, style_formatting.as_ref(), styles, source);
                if let (Some(boundaries), Some(boundary)) = (&mut boundaries, boundary) {
                    boundaries.push(boundary);
                } else {
                    boundaries = None;
                }
                units.extend(run_to_units(
                    content,
                    style_formatting.as_ref(),
                    styles,
                    comment_id,
                    &[],
                    source,
                ));
            }
            "hyperlink" => {
                boundaries = None;
                units.extend(hyperlink_to_units(
                    content,
                    style_formatting.as_ref(),
                    styles,
                    &[],
                    source,
                ));
            }
            "simpleField" | "complexField" => {
                boundaries = None;
                let (payload, marks) = field_payload(content, style_formatting.as_ref(), source);
                units.push(embed_unit("field", payload, &marks, None, 1));
            }
            "inlineSdt" => {
                boundaries = None;
                units.push(embed_unit(
                    "sdt",
                    sdt_payload(content, style_formatting.as_ref(), styles, source),
                    &[],
                    None,
                    2,
                ));
            }
            "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                boundaries = None;
                units.extend(tracked_to_units(
                    content,
                    style_formatting.as_ref(),
                    styles,
                    comment_id,
                    source,
                ));
            }
            "mathEquation" => {
                boundaries = None;
                units.push(embed_unit("math", math_payload(content), &[], None, 1));
            }
            "bookmarkStart" | "bookmarkEnd" => {}
            _ => boundaries = None,
        }
        unit_counts.push(units.len() - start);
    }
    let attrs = paragraph_attrs(paragraph, styles, &units, &unit_counts, boundaries);
    (units, para_attrs_to_ppr(attrs))
}

fn run_tokens(run: &Value, tokens: &mut Vec<&'static str>) {
    for content in array(field(Some(run), "content")) {
        if string(field(Some(content), "type")) == Some("break")
            && string(field(Some(content), "breakType")) == Some("page")
        {
            tokens.push("pageBreak");
        } else if string(field(Some(content), "type")) != Some("text")
            || !string(field(Some(content), "text"))
                .unwrap_or_default()
                .is_empty()
        {
            tokens.push("visible");
        }
    }
}

fn inline_tokens(content: &[Value], tokens: &mut Vec<&'static str>) {
    for item in content {
        match string(field(Some(item), "type")).unwrap_or_default() {
            "run" => run_tokens(item, tokens),
            "hyperlink" => {
                for child in array(field(Some(item), "children")) {
                    if string(field(Some(child), "type")) == Some("run") {
                        run_tokens(child, tokens);
                    }
                }
            }
            "simpleField" => {
                for child in array(field(Some(item), "content")) {
                    if string(field(Some(child), "type")) == Some("run") {
                        run_tokens(child, tokens);
                    }
                }
            }
            "complexField" => {
                for key in ["fieldCode", "fieldResult"] {
                    for child in array(field(Some(item), key)) {
                        run_tokens(child, tokens);
                    }
                }
            }
            "inlineSdt" => inline_tokens(array(field(Some(item), "content")), tokens),
            "insertion" | "deletion" | "moveFrom" | "moveTo" => {
                for child in array(field(Some(item), "content")) {
                    if string(field(Some(child), "type")) == Some("run") {
                        run_tokens(child, tokens);
                    }
                }
            }
            "mathEquation" => tokens.push("visible"),
            _ => {}
        }
    }
}

fn paragraph_starts_with_page_break(paragraph: &Value) -> bool {
    let mut tokens = Vec::new();
    inline_tokens(array(field(Some(paragraph), "content")), &mut tokens);
    tokens.first() == Some(&"pageBreak")
}

fn paragraph_has_non_leading_page_break(paragraph: &Value) -> bool {
    let mut tokens = Vec::new();
    inline_tokens(array(field(Some(paragraph), "content")), &mut tokens);
    let mut leading = false;
    let mut visible = false;
    for token in tokens {
        if token == "pageBreak" {
            if visible || leading {
                return true;
            }
            leading = true;
        } else {
            visible = true;
        }
    }
    false
}

fn modifier(value: &str) -> f64 {
    let prefix: String = value
        .chars()
        .take_while(|character| character.is_ascii_hexdigit())
        .collect();
    u8::from_str_radix(&prefix, 16)
        .map(|value| f64::from(value) / 255.0)
        .unwrap_or(1.0)
}

fn rgb_channels(value: &str) -> [u8; 3] {
    let mut normalized = value.trim_start_matches('#').to_owned();
    while normalized.len() < 6 {
        normalized.insert(0, '0');
    }
    normalized.truncate(6);
    [
        u8::from_str_radix(&normalized[0..2], 16).unwrap_or(0),
        u8::from_str_radix(&normalized[2..4], 16).unwrap_or(0),
        u8::from_str_radix(&normalized[4..6], 16).unwrap_or(0),
    ]
}

fn resolve_color_to_hex(color: Option<&Value>, theme: Option<&Value>) -> Option<String> {
    let color = object(color)?;
    if truthy(color.get("auto")) {
        return None;
    }
    if let Some(theme_color) = string(color.get("themeColor"))
        && let Some(theme) = theme
    {
        let slot = match theme_color {
            "dark1" | "text1" | "tx1" => "dk1",
            "light1" | "background1" | "bg1" => "lt1",
            "dark2" | "text2" | "tx2" => "dk2",
            "light2" | "background2" | "bg2" => "lt2",
            "hyperlink" => "hlink",
            "followedHyperlink" => "folHlink",
            value => value,
        };
        let known = [
            "dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5",
            "accent6", "hlink", "folHlink",
        ];
        let mut hex = if known.contains(&slot) {
            string(field(field(Some(theme), "colorScheme"), slot))
                .or_else(|| string(color.get("rgb")))
                .unwrap_or("000000")
                .to_owned()
        } else {
            string(color.get("rgb")).unwrap_or("000000").to_owned()
        };
        let mut channels = rgb_channels(&hex);
        if let Some(tint) = string(color.get("themeTint")) {
            let tint = modifier(tint);
            channels = channels
                .map(|channel| (f64::from(channel) * tint + 255.0 * (1.0 - tint)).round() as u8);
        } else if let Some(shade) = string(color.get("themeShade")) {
            let shade = modifier(shade);
            channels = channels.map(|channel| (f64::from(channel) * shade).round() as u8);
        }
        hex = format!("{:02X}{:02X}{:02X}", channels[0], channels[1], channels[2]);
        return Some(hex);
    }
    string(color.get("rgb"))
        .filter(|value| *value != "auto")
        .map(|value| value.trim_start_matches('#').to_ascii_uppercase())
}

fn calculate_row_spans(table: &Value) -> BTreeMap<(usize, usize), (usize, bool)> {
    let mut result = BTreeMap::new();
    let mut active = BTreeMap::<usize, usize>::new();
    for (row_index, row) in array(field(Some(table), "rows")).iter().enumerate() {
        let mut column = 0usize;
        let cells: Vec<_> = array(field(Some(row), "cells"))
            .iter()
            .map(|cell| {
                let current = column;
                column += number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0)
                    as usize;
                (
                    current,
                    string(field(field(Some(cell), "formatting"), "vMerge")),
                )
            })
            .collect();
        let empty = !cells.is_empty()
            && cells
                .iter()
                .all(|(column, merge)| *merge == Some("continue") && active.contains_key(column));
        if empty {
            for (column, _) in cells {
                active.remove(&column);
                result.insert((row_index, column), (1, false));
            }
            continue;
        }
        for (column, merge) in cells {
            match merge {
                Some("restart") => {
                    active.insert(column, row_index);
                    result.insert((row_index, column), (1, false));
                }
                Some("continue") => {
                    if let Some(start) = active.get(&column).copied() {
                        if let Some(owner) = result.get_mut(&(start, column)) {
                            owner.0 += 1;
                        }
                        result.insert((row_index, column), (1, true));
                    } else {
                        result.insert((row_index, column), (1, false));
                    }
                }
                _ => {
                    active.remove(&column);
                    result.insert((row_index, column), (1, false));
                }
            }
        }
    }
    result
}

fn revision_attrs(info: &Value) -> Value {
    json!({
        "revisionId": nullish(field(Some(info), "id")),
        "author": nullish(field(Some(info), "author")),
        "date": nullish(field(Some(info), "date"))
    })
}

fn cell_borders(
    formatting: Option<&Value>,
    table_borders: Option<&Value>,
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
) -> Option<Value> {
    let inherited = object(table_borders).map(|borders| {
        json!({
            "top": nullish(if first_row { borders.get("top") } else { borders.get("insideH") }),
            "bottom": nullish(if last_row { borders.get("bottom") } else { borders.get("insideH") }),
            "left": nullish(if first_column { borders.get("left") } else { borders.get("insideV") }),
            "right": nullish(if last_column { borders.get("right") } else { borders.get("insideV") })
        })
    });
    let direct = field(formatting, "borders");
    if inherited.is_none() && direct.is_none() {
        None
    } else {
        merge_plain(inherited.as_ref(), direct)
    }
}

struct CellOptions<'a> {
    is_header: bool,
    rowspan: usize,
    grid_width: Option<f64>,
    first_row: bool,
    last_row: bool,
    first_column: bool,
    last_column: bool,
    table_borders: Option<&'a Value>,
    default_margins: Option<&'a Value>,
    theme: Option<&'a Value>,
    table_bidi: bool,
}

fn structural_attrs(attrs: JsonObject, skipped: &[&str]) -> JsonObject {
    attrs
        .into_iter()
        .filter(|(key, value)| !skipped.contains(&key.as_str()) && !value.is_null())
        .map(|(key, value)| (key, drop_nulls(value)))
        .collect()
}

fn project_cell(cell: &Value, options: CellOptions<'_>) -> ProjectedCell {
    let formatting = field(Some(cell), "formatting");
    let background =
        resolve_color_to_hex(field(field(formatting, "shading"), "fill"), options.theme);
    let width = field(field(formatting, "width"), "value")
        .cloned()
        .or_else(|| options.grid_width.map(|value| json!(value)));
    let width_type = field(field(formatting, "width"), "type")
        .cloned()
        .or_else(|| options.grid_width.map(|_| Value::String("pct".to_owned())));
    let margins = if let Some(margins) = field(formatting, "margins") {
        Some(json!({
            "top": nullish(field(field(Some(margins), "top"), "value")),
            "bottom": nullish(field(field(Some(margins), "bottom"), "value")),
            "left": nullish(
                field(field(Some(margins), "left"), "value").or_else(|| {
                    field(
                        field(Some(margins), if options.table_bidi { "end" } else { "start" }),
                        "value",
                    )
                })
            ),
            "right": nullish(
                field(field(Some(margins), "right"), "value").or_else(|| {
                    field(
                        field(Some(margins), if options.table_bidi { "start" } else { "end" }),
                        "value",
                    )
                })
            )
        }))
    } else {
        options.default_margins.cloned()
    };
    let mut attrs = map_from_value(json!({
        "colspan": number(field(formatting, "gridSpan")).unwrap_or(1.0),
        "rowspan": options.rowspan,
        "width": width,
        "widthType": width_type,
        "verticalAlign": nullish(field(formatting, "verticalAlign")),
        "backgroundColor": background,
        "borders": cell_borders(
            formatting,
            options.table_borders,
            options.first_row,
            options.last_row,
            options.first_column,
            options.last_column,
        ),
        "margins": margins,
        "textDirection": nullish(field(formatting, "textDirection")),
        "noWrap": boolean(field(formatting, "noWrap")).unwrap_or(false),
        "_originalFormatting": nullish(formatting),
        "_originalResolvedFill": background
    }));
    if let Some(change) = field(Some(cell), "structuralChange") {
        let info = revision_attrs(field(Some(change), "info").unwrap_or(&Value::Null));
        match string(field(Some(change), "type")).unwrap_or_default() {
            "tableCellInsertion" => {
                attrs.insert(
                    "cellMarker".to_owned(),
                    json!({ "kind": "ins", "info": info }),
                );
            }
            "tableCellDeletion" => {
                attrs.insert(
                    "cellMarker".to_owned(),
                    json!({ "kind": "del", "info": info }),
                );
            }
            "tableCellMerge" => {
                let mut marker = json!({
                    "kind": "merge",
                    "info": info,
                    "vMerge": string(field(Some(change), "vMerge")).unwrap_or("cont")
                });
                if let Some(value) =
                    string(field(Some(change), "vMergeOrig")).filter(|value| !value.is_empty())
                {
                    marker
                        .as_object_mut()
                        .unwrap()
                        .insert("vMergeOrig".to_owned(), Value::String(value.to_owned()));
                }
                attrs.insert("cellMarker".to_owned(), marker);
            }
            _ => {}
        }
    }
    if !array(field(Some(cell), "propertyChanges")).is_empty() {
        attrs.insert(
            "tcPrChange".to_owned(),
            field(Some(cell), "propertyChanges").unwrap().clone(),
        );
    }
    let mut attrs = structural_attrs(attrs, &[]);
    if options.is_header {
        attrs.insert("header".to_owned(), Value::Bool(true));
    }
    let content = array(field(Some(cell), "content"));
    ProjectedCell {
        attrs,
        content: if content.is_empty() {
            vec![json!({ "type": "paragraph", "content": [] })]
        } else {
            content.to_vec()
        },
    }
}

fn project_row(
    row: &Value,
    table: &Value,
    row_index: usize,
    row_spans: &BTreeMap<(usize, usize), (usize, bool)>,
    table_borders: Option<&Value>,
    default_margins: Option<&Value>,
    theme: Option<&Value>,
) -> ProjectedRow {
    let formatting = field(Some(row), "formatting");
    let mut attrs = map_from_value(json!({
        "height": nullish(field(field(formatting, "height"), "value")),
        "heightRule": nullish(field(formatting, "heightRule")),
        "isHeader": truthy(field(formatting, "header")),
        "_originalFormatting": nullish(formatting)
    }));
    if let Some(change) = field(Some(row), "structuralChange") {
        let value = revision_attrs(field(Some(change), "info").unwrap_or(&Value::Null));
        match string(field(Some(change), "type")).unwrap_or_default() {
            "tableRowInsertion" => {
                attrs.insert("trIns".to_owned(), value);
            }
            "tableRowDeletion" => {
                attrs.insert("trDel".to_owned(), value);
            }
            _ => {}
        }
    }
    if !array(field(Some(row), "propertyChanges")).is_empty() {
        attrs.insert(
            "trPrChange".to_owned(),
            field(Some(row), "propertyChanges").unwrap().clone(),
        );
    }
    let widths = array(field(Some(table), "columnWidths"));
    let total_width: f64 = widths.iter().filter_map(Value::as_f64).sum();
    let rows = array(field(Some(table), "rows"));
    let total_columns = if !widths.is_empty() {
        widths.len()
    } else {
        rows.iter()
            .map(|row| {
                array(field(Some(row), "cells"))
                    .iter()
                    .map(|cell| {
                        number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0)
                            as usize
                    })
                    .sum()
            })
            .max()
            .unwrap_or(0)
    };
    let cells_source = array(field(Some(row), "cells"));
    let mut column = 0usize;
    let mut cells = Vec::new();
    for cell in cells_source {
        let colspan =
            number(field(field(Some(cell), "formatting"), "gridSpan")).unwrap_or(1.0) as usize;
        let start_column = column;
        let span = row_spans.get(&(row_index, start_column));
        let grid_width = (!widths.is_empty() && total_width > 0.0).then(|| {
            let cell_width: f64 = widths
                .iter()
                .skip(start_column)
                .take(colspan)
                .filter_map(Value::as_f64)
                .sum();
            (cell_width / total_width * 100.0).round()
        });
        column += colspan;
        if span.is_some_and(|(_, skip)| *skip) {
            continue;
        }
        cells.push(project_cell(
            cell,
            CellOptions {
                is_header: row_index == 0
                    && truthy(field(
                        field(field(Some(table), "formatting"), "look"),
                        "firstRow",
                    )),
                rowspan: span.map(|(rowspan, _)| *rowspan).unwrap_or(1),
                grid_width,
                first_row: row_index == 0,
                last_row: row_index + 1 == rows.len(),
                first_column: start_column == 0,
                last_column: column == total_columns,
                table_borders,
                default_margins,
                theme,
                table_bidi: truthy(field(field(Some(table), "formatting"), "bidi")),
            },
        ));
    }
    if cells.is_empty() {
        let synthetic = if total_columns > 1 {
            json!({
                "type": "tableCell",
                "formatting": { "gridSpan": total_columns },
                "content": [{ "type": "paragraph", "content": [] }]
            })
        } else {
            json!({
                "type": "tableCell",
                "content": [{ "type": "paragraph", "content": [] }]
            })
        };
        cells.push(project_cell(
            &synthetic,
            CellOptions {
                is_header: row_index == 0
                    && truthy(field(
                        field(field(Some(table), "formatting"), "look"),
                        "firstRow",
                    )),
                rowspan: 1,
                grid_width: (total_width > 0.0).then_some(100.0),
                first_row: row_index == 0,
                last_row: row_index + 1 == rows.len(),
                first_column: true,
                last_column: true,
                table_borders,
                default_margins,
                theme,
                table_bidi: truthy(field(field(Some(table), "formatting"), "bidi")),
            },
        ));
    }
    ProjectedRow {
        attrs: structural_attrs(attrs, &[]),
        cells,
    }
}

fn project_table(table: &Value, styles: &StyleResolver, theme: Option<&Value>) -> ProjectedTable {
    let formatting = field(Some(table), "formatting");
    let default_style = styles.default_style("table");
    let style_id = string(field(formatting, "styleId"));
    let effective_style_id =
        style_id.or_else(|| default_style.and_then(|style| string(field(Some(style), "styleId"))));
    let table_style = effective_style_id.and_then(|id| styles.style(id));
    let borders = field(formatting, "borders")
        .or_else(|| field(field(table_style, "tblPr"), "borders"))
        .or_else(|| field(field(default_style, "tblPr"), "borders"));
    let margins = field(formatting, "cellMargins")
        .or_else(|| field(field(table_style, "tblPr"), "cellMargins"))
        .or_else(|| field(field(default_style, "tblPr"), "cellMargins"));
    let logical_left = field(
        margins,
        if truthy(field(formatting, "bidi")) {
            "end"
        } else {
            "start"
        },
    );
    let logical_right = field(
        margins,
        if truthy(field(formatting, "bidi")) {
            "start"
        } else {
            "end"
        },
    );
    let default_margins = margins.map(|margins| {
        drop_nulls(json!({
            "top": nullish(field(field(Some(margins), "top"), "value")),
            "bottom": nullish(field(field(Some(margins), "bottom"), "value")),
            "left": nullish(
                field(field(Some(margins), "left"), "value")
                    .or_else(|| field(logical_left, "value"))
            ),
            "right": nullish(
                field(field(Some(margins), "right"), "value")
                    .or_else(|| field(logical_right, "value"))
            )
        }))
    });
    let mut based_on = Vec::new();
    let mut visited = BTreeSet::new();
    let mut inherited = table_style;
    while let Some(parent_id) = inherited.and_then(|style| string(field(Some(style), "basedOn"))) {
        if based_on.len() >= 32 || !visited.insert(parent_id.to_owned()) {
            break;
        }
        based_on.insert(0, Value::String(parent_id.to_owned()));
        inherited = styles.style(parent_id);
    }
    let mut original_formatting = object(formatting).cloned().unwrap_or_default();
    original_formatting.insert(
        "styleCascade".to_owned(),
        drop_nulls(json!({
            "selectedStyleId": style_id.filter(|value| !value.is_empty()),
            "defaultStyleId": default_style
                .and_then(|style| string(field(Some(style), "styleId")))
                .filter(|value| !value.is_empty()),
            "basedOnStyleIds": (!based_on.is_empty()).then_some(based_on)
        })),
    );
    let mut attrs = map_from_value(json!({
        "styleId": style_id,
        "width": nullish(field(field(formatting, "width"), "value")),
        "widthType": nullish(field(field(formatting, "width"), "type")),
        "justification": nullish(field(formatting, "justification")),
        "columnWidths": nullish(field(Some(table), "columnWidths")),
        "tableLayout": nullish(field(formatting, "layout")),
        "floating": nullish(field(formatting, "floating")),
        "cellMargins": default_margins,
        "look": nullish(field(formatting, "look")),
        "bidi": truthy(field(formatting, "bidi")).then_some(true),
        "_originalFormatting": Value::Object(original_formatting)
    }));
    if !array(field(Some(table), "propertyChanges")).is_empty() {
        attrs.insert(
            "tblPrChange".to_owned(),
            field(Some(table), "propertyChanges").unwrap().clone(),
        );
    }
    let row_spans = calculate_row_spans(table);
    let rows = array(field(Some(table), "rows"))
        .iter()
        .enumerate()
        .map(|(row_index, row)| {
            project_row(
                row,
                table,
                row_index,
                &row_spans,
                borders,
                default_margins.as_ref(),
                theme,
            )
        })
        .collect();
    ProjectedTable { attrs, rows }
}

fn table_cell_story_id(parent: &str, table: usize, row: usize, cell: usize) -> String {
    format!("{parent}:t{table}:r{row}c{cell}")
}

fn add_comment_coverage(plan: &mut StoryPlan) {
    let mut offset = 0u32;
    for unit in &plan.units {
        let width = match &unit.content {
            UnitContent::Text(text) => utf16_len(text),
            UnitContent::Embed { .. } => 1,
        };
        if let Some(comment_id) = &unit.comment_id
            && comment_id != "0"
        {
            let index = plan
                .comment_coverage
                .iter()
                .position(|(id, _)| id == comment_id);
            if let Some(index) = index {
                let ranges = &mut plan.comment_coverage[index].1;
                if let Some(previous) = ranges.last_mut()
                    && previous.1 == offset
                {
                    previous.1 = offset + width;
                } else {
                    ranges.push((offset, offset + width));
                }
            } else {
                plan.comment_coverage
                    .push((comment_id.clone(), vec![(offset, offset + width)]));
            }
        }
        offset += width;
    }
}

fn visit_story(
    context: &mut LoweringContext,
    story_id: String,
    source_blocks: &[Value],
    options: StoryOptions,
) {
    let plan_index = context.plans.len();
    context.plans.push(StoryPlan {
        story_id: story_id.clone(),
        units: Vec::new(),
        comment_coverage: Vec::new(),
    });
    let blocks = if source_blocks.is_empty() {
        vec![json!({ "type": "paragraph", "content": [] })]
    } else {
        source_blocks.to_vec()
    };
    let mut table_index = 0usize;
    let mut sdt_index = 0usize;
    let mut paragraph_index = 0usize;
    let mut last_kind = None;
    for block in blocks {
        match string(field(Some(&block), "type")).unwrap_or_default() {
            "paragraph" => {
                let (units, mut ppr) =
                    paragraph_units(&block, &context.styles, None, &context.source_json);
                let fallback = format!("{story_id}:p{paragraph_index}");
                ppr.insert(
                    "paraId".to_owned(),
                    Value::String(
                        string(field(Some(&block), "paraId"))
                            .filter(|value| !value.is_empty())
                            .unwrap_or(&fallback)
                            .to_owned(),
                    ),
                );
                context.plans[plan_index].units.extend(units);
                context.plans[plan_index]
                    .units
                    .push(embed_unit("pilcrow", ppr, &[], None, 1));
                paragraph_index += 1;
                if options.include_page_breaks && paragraph_has_non_leading_page_break(&block) {
                    context.plans[plan_index].units.push(embed_unit(
                        "pageBreak",
                        JsonObject::new(),
                        &[],
                        None,
                        1,
                    ));
                }
                last_kind = Some("paragraph");
            }
            "table" => {
                let current_table = table_index;
                table_index += 1;
                let table = project_table(&block, &context.styles, context.theme.as_ref());
                let rows: Vec<Value> = table
                    .rows
                    .iter()
                    .enumerate()
                    .map(|(row_index, row)| {
                        json!({
                            "trPr": value_from_map(&row.attrs),
                            "cells": row.cells.iter().enumerate().map(|(cell_index, cell)| {
                                json!({
                                    "tcPr": value_from_map(&cell.attrs),
                                    "story": table_cell_story_id(
                                        &story_id,
                                        current_table,
                                        row_index,
                                        cell_index,
                                    )
                                })
                            }).collect::<Vec<_>>()
                        })
                    })
                    .collect();
                let tbl_pr = structural_attrs(table.attrs.clone(), &["columnWidths"]);
                let grid = table
                    .attrs
                    .get("columnWidths")
                    .and_then(Value::as_array)
                    .cloned()
                    .unwrap_or_default()
                    .into_iter()
                    .map(drop_nulls)
                    .collect::<Vec<_>>();
                context.plans[plan_index].units.push(embed_unit(
                    "table",
                    map_from_value(json!({
                        "tblPr": value_from_map(&tbl_pr),
                        "grid": grid,
                        "rows": rows
                    })),
                    &[],
                    None,
                    1,
                ));
                for (row_index, row) in table.rows.into_iter().enumerate() {
                    for (cell_index, cell) in row.cells.into_iter().enumerate() {
                        visit_story(
                            context,
                            table_cell_story_id(&story_id, current_table, row_index, cell_index),
                            &cell.content,
                            StoryOptions {
                                include_page_breaks: false,
                                append_body_tail: false,
                                seed_comments: false,
                            },
                        );
                    }
                }
                last_kind = Some("table");
            }
            _ => {
                let current_sdt = sdt_index;
                sdt_index += 1;
                let child_story = format!("{story_id}:sdt{current_sdt}");
                let mut properties = sdt_properties_attrs(
                    field(Some(&block), "properties").unwrap_or(&Value::Null),
                    &context.source_json,
                );
                properties.insert("story".to_owned(), Value::String(child_story.clone()));
                context.plans[plan_index].units.push(embed_unit(
                    "blockSdt",
                    properties,
                    &[],
                    None,
                    1,
                ));
                visit_story(
                    context,
                    child_story,
                    array(field(Some(&block), "content")),
                    StoryOptions {
                        include_page_breaks: options.include_page_breaks,
                        append_body_tail: false,
                        seed_comments: false,
                    },
                );
                last_kind = Some("blockSdt");
            }
        }
    }
    if options.append_body_tail && matches!(last_kind, Some("table" | "blockSdt")) {
        context.plans[plan_index].units.push(embed_unit(
            "pilcrow",
            map_from_value(json!({
                "hangingIndent": false,
                "paraId": format!("{story_id}:p{paragraph_index}")
            })),
            &[],
            None,
            1,
        ));
    }
    if options.seed_comments {
        add_comment_coverage(&mut context.plans[plan_index]);
    }
}

fn units_to_raw_ops(units: Vec<InlineUnit>) -> Result<Vec<RawOp>, String> {
    let mut ops = vec![RawOp::Delete { index: 0, len: 1 }];
    let mut index = 0u32;
    let mut text = String::new();
    let mut attrs = JsonObject::new();
    let flush = |ops: &mut Vec<RawOp>,
                 index: &mut u32,
                 text: &mut String,
                 attrs: &mut JsonObject|
     -> Result<(), String> {
        if text.is_empty() {
            return Ok(());
        }
        let inserted = std::mem::take(text);
        let len = utf16_len(&inserted);
        ops.push(RawOp::Insert {
            index: *index,
            text: inserted,
            attrs: yrs_attrs(std::mem::take(attrs))?,
        });
        *index += len;
        Ok(())
    };
    for unit in units {
        match unit.content {
            UnitContent::Text(value) => {
                if text.is_empty() {
                    attrs = unit.attrs;
                } else if unit.attrs != attrs {
                    flush(&mut ops, &mut index, &mut text, &mut attrs)?;
                    attrs = unit.attrs;
                }
                text.push_str(&value);
            }
            UnitContent::Embed {
                kind,
                payload: values,
            } => {
                flush(&mut ops, &mut index, &mut text, &mut attrs)?;
                ops.push(RawOp::InsertEmbed {
                    index,
                    kind,
                    payload: payload(values)?,
                    attrs: yrs_attrs(unit.attrs)?,
                });
                index += 1;
            }
        }
    }
    flush(&mut ops, &mut index, &mut text, &mut attrs)?;
    Ok(ops)
}

fn seed_plan(plan: StoryPlan) -> Result<(String, Vec<RawOp>), String> {
    let StoryPlan {
        story_id,
        units,
        comment_coverage,
    } = plan;
    let mut ops = units_to_raw_ops(units)?;
    if !comment_coverage.is_empty() {
        ops.extend(
            comment_coverage
                .into_iter()
                .map(|(id, ranges)| RawOp::SetComment {
                    id,
                    ranges,
                    author: String::new(),
                    date: String::new(),
                    body: Any::Null,
                }),
        );
    }
    Ok((story_id, ops))
}

fn entry_parts(entry: &Value) -> Option<(&str, &Value)> {
    let entry = entry.as_array()?;
    Some((entry.first()?.as_str()?, entry.get(1)?))
}

pub fn seed_from_docx(document: &EditingDoc, bytes: &[u8]) -> Result<(), String> {
    let envelope = docx_parse::parse_docx_s9_wire(bytes, docx_parse::S9ParseOptions::default())
        .map_err(|error| error.to_string())?;
    let parsed = serde_json::to_value(&envelope.document).map_err(|error| error.to_string())?;
    let source_json = if needs_source_json(&parsed) {
        let serialized =
            serde_json::to_string(&envelope.document).map_err(|error| error.to_string())?;
        let ordered: OrderedValue =
            serde_json::from_str(&serialized).map_err(|error| error.to_string())?;
        let mut values = BTreeMap::new();
        ordered.collect_source_json(&mut values);
        values
    } else {
        BTreeMap::new()
    };
    drop(envelope);
    let package =
        field(Some(&parsed), "package").ok_or_else(|| "parsed DOCX has no package".to_owned())?;
    let mut context = LoweringContext {
        styles: StyleResolver::new(field(Some(package), "styles")),
        theme: field(Some(package), "theme").cloned(),
        source_json: Arc::new(source_json),
        plans: Vec::new(),
    };
    visit_story(
        &mut context,
        "body".to_owned(),
        array(field(field(Some(package), "document"), "content")),
        StoryOptions {
            include_page_breaks: true,
            append_body_tail: true,
            seed_comments: true,
        },
    );
    for entry in array(field(Some(package), "headerEntries")) {
        let Some((relationship_id, part)) = entry_parts(entry) else {
            continue;
        };
        visit_story(
            &mut context,
            format!("hf:{relationship_id}"),
            array(field(Some(part), "content")),
            StoryOptions {
                include_page_breaks: false,
                append_body_tail: false,
                seed_comments: true,
            },
        );
    }
    for entry in array(field(Some(package), "footerEntries")) {
        let Some((relationship_id, part)) = entry_parts(entry) else {
            continue;
        };
        let story_id = format!("hf:{relationship_id}");
        if context.plans.iter().any(|plan| plan.story_id == story_id) {
            continue;
        }
        visit_story(
            &mut context,
            story_id,
            array(field(Some(part), "content")),
            StoryOptions {
                include_page_breaks: false,
                append_body_tail: false,
                seed_comments: true,
            },
        );
    }
    for (key, prefix) in [("footnotes", "fn"), ("endnotes", "en")] {
        for note in array(field(Some(package), key)) {
            let Some(id) = field(Some(note), "id") else {
                continue;
            };
            visit_story(
                &mut context,
                format!("{prefix}:{}", js_string(id)),
                array(field(Some(note), "content")),
                StoryOptions {
                    include_page_breaks: false,
                    append_body_tail: false,
                    seed_comments: true,
                },
            );
        }
    }
    drop(parsed);
    document
        .create_empty_stories(
            &context
                .plans
                .iter()
                .map(|plan| plan.story_id.clone())
                .collect::<Vec<_>>(),
        )
        .map_err(|error| error.to_string())?;
    let batches = context
        .plans
        .into_iter()
        .map(seed_plan)
        .collect::<Result<Vec<_>, _>>()?;
    document
        .apply_raw_story_batches(batches, &EditCtx::local(String::new(), String::new()))
        .map_err(|error| error.to_string())?;
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn source_json_preserves_wire_order_with_js_number_formatting() {
        let ordered: OrderedValue =
            serde_json::from_str(r#"{"type":"shape","z":1.0,"nested":{"b":2,"a":3}}"#).unwrap();
        let value = ordered.value();
        let mut source = BTreeMap::new();
        ordered.collect_source_json(&mut source);

        assert_eq!(
            source_json(&value, &source),
            r#"{"type":"shape","z":1,"nested":{"b":2,"a":3}}"#
        );
    }
}
