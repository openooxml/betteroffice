use std::ops::Range;

use docx_parse::formatting::{ParagraphFormatting, TextFormatting};
use docx_parse::inline::{InlineNode, Run, RunContent, RunType};
use docx_parse::paragraph::{Paragraph, ParagraphContent};
use docx_parse::theme::{Theme, resolve_theme_font_ref};
use ooxml_diff::{DiffKind, DiffLimits, LimitFallback, diff_tokens};
use serde::Deserialize;
use serde_json::{Map, Value, json};

use super::paragraph::{good, same_option, set};
use super::{Attrs, ProjectionContext, ProjectionError, StoryState, decode, json_value};

const ON_OFF: &[&str] = &[
    "bold", "italic", "allCaps", "smallCaps", "strike", "doubleStrike", "outline", "shadow",
    "emboss", "imprint", "hidden", "rtl", "boldCs", "italicCs",
];
const SOURCE_PROPERTIES: &[&str] = &[
    "bold", "italic", "fontSize", "boldCs", "italicCs", "fontSizeCs", "language", "fontFamily",
    "snapToGrid",
];
const BOOLEAN_MARKS: &[&str] = &[
    "bold", "italic", "superscript", "subscript", "allCaps", "smallCaps", "emboss", "imprint",
    "textShadow", "textOutline", "hidden", "rtl",
];

type Properties = Map<String, Value>;

#[derive(Deserialize)]
#[serde(rename_all = "camelCase")]
struct OriginalBoundary {
    text: String,
    marks_key: Option<String>,
    formatting: Option<TextFormatting>,
    breaks: Option<Vec<Value>>,
    note_marks: Option<Vec<Value>>,
    property_changes: Option<Vec<docx_parse::inline::RunPropertyChange>>,
}

pub(super) fn lower_plain(
    context: &mut ProjectionContext<'_, '_>,
    state: &StoryState,
    paragraph_attrs: &Attrs,
    items: &[(&str, &Attrs)],
    source: Option<&Paragraph>,
    allow_mark_fallback: bool,
) -> Result<Vec<ParagraphContent>, ProjectionError> {
    let boundaries: Vec<OriginalBoundary> = decode(
        paragraph_attrs.get("_originalRunBoundaries").map(json_value).transpose()?.as_ref(),
    )?.unwrap_or_default();
    let live: String = items.iter().map(|(text, _)| *text).collect();
    let mark_fallback = if allow_mark_fallback && source.is_none() && boundaries.is_empty() {
        let original = paragraph_attrs.get("_originalFormatting").map(json_value).transpose()?;
        decode::<Option<ParagraphFormatting>>(original.as_ref())?.flatten()
            .and_then(|formatting| formatting.run_properties)
            .map(|mark| TextFormatting {
                font_family: mark.font_family,
                language: mark.language,
                ..TextFormatting::default()
            })
    } else {
        None
    };
    let mut sources = source.and_then(plain_source)
        .unwrap_or_else(|| boundaries.iter().map(|boundary| {
            (boundary.text.clone(), boundary.formatting.as_ref())
        }).collect());
    if let Some(mark) = &mark_fallback {
        sources.push((live.clone(), Some(mark)));
    }
    let live_attrs: Vec<_> = items.iter().map(|(_, attrs)| attributes(attrs)).collect::<Result<_, _>>()?;
    if let Some(content) = restore(&boundaries, &sources, &live, items, &live_attrs)? {
        return Ok(content);
    }

    let source_text: String = sources.iter().map(|(text, _)| text.as_str()).collect();
    let live_chars: Vec<_> = live.chars().collect();
    let original_chars: Vec<_> = source_text.chars().collect();
    let spans = source_spans(&sources, &original_chars, &live_chars);
    let source_is_plain = source.is_some_and(|paragraph| paragraph.content.iter().all(|content| {
        matches!(content, ParagraphContent::Inline(InlineNode::Run(run))
            if run.property_changes.is_none() && run.content.iter().all(is_plain))
    }));
    let coalesce = source_is_plain && live != source_text;
    let mut projected = Vec::new();
    let mut current: Option<(Run, Properties, Option<usize>)> = None;
    let mut offset = 0;
    let mut span_index = 0;
    for ((text, _), attrs) in items.iter().zip(live_attrs) {
        let end = offset + text.chars().count();
        while offset < end {
            while spans.get(span_index).is_some_and(|span| span.range.end <= offset) {
                span_index += 1;
            }
            let span = spans.get(span_index);
            let next = span.map_or(end, |span| span.range.end.min(end));
            let text: String = live_chars[offset..next].iter().collect();
            let source = span.and_then(|span| sources[span.index].1);
            let run = make_run(context, state, paragraph_attrs, &text, &attrs, source)?;
            let mut key = attrs.clone();
            for name in ["hyperlink", "ins", "del"] {
                key.remove(name);
            }
            if let Some(source) = source {
                let mut authored = properties(source)?;
                authored.retain(|name, _| SOURCE_PROPERTIES.contains(&name.as_str()));
                key.insert("_sourceRunFormatting".to_owned(), Value::Object(authored));
            } else if !sources.is_empty() {
                key.insert("_sourceRunFormatting".to_owned(), json!({}));
            }
            let index = if coalesce { span.map(|span| span.index) } else { None };
            let merge = current.as_ref().is_some_and(|(current_run, current_key, previous)| {
                let changed_neighbours = coalesce && span.is_some_and(|span| {
                    span.changed && span.previous_changed == *previous && previous.is_some()
                });
                let compatible = if index.is_some() {
                    (index == *previous || changed_neighbours)
                        && run.content.iter().all(is_plain)
                        && current_run.content.iter().all(is_plain)
                } else {
                    previous.is_none()
                };
                let same_format = if coalesce && span.is_some_and(|span| span.changed) {
                    current_run.formatting == run.formatting
                        && current_run.direction_scopes == run.direction_scopes
                } else {
                    properties_equal(current_key, &key)
                };
                compatible && same_format
            });
            if merge {
                let (previous, _, previous_index) = current.as_mut().expect("matching run");
                append(previous, run);
                *previous_index = index;
            } else {
                if let Some((run, _, _)) = current.take() {
                    projected.push(inline(run));
                }
                current = Some((run, key, index));
            }
            offset = next;
        }
    }
    if let Some((run, _, _)) = current {
        projected.push(inline(run));
    }
    Ok(projected)
}

pub(super) fn make_run(
    context: &mut ProjectionContext<'_, '_>,
    state: &StoryState,
    paragraph_attrs: &Attrs,
    text: &str,
    attrs: &Properties,
    source: Option<&TextFormatting>,
) -> Result<Run, ProjectionError> {
    let style = attrs.get("runStyle").and_then(|style| style.get("styleId")).and_then(Value::as_str);
    let paragraph_style = super::attr_string(paragraph_attrs, "pStyle");
    let inherited = match &state.table_run_formatting {
        Some(table) => context.inheritance.table_run_attrs(paragraph_style, style, table.clone()),
        None => context.inheritance.run_attrs(paragraph_style, style),
    };
    let inherited = direct(&inherited.into_iter().collect());
    let live = direct(attrs);
    let formatting = relative(attrs, &live, &inherited, source, context.theme)?;
    Ok(Run {
        node_type: RunType::Run,
        formatting,
        property_changes: None,
        content: run_content(text, &live),
        direction_scopes: decode(attrs.get("directionScopes"))?.unwrap_or_default(),
    })
}

pub(super) fn inline(run: Run) -> ParagraphContent {
    ParagraphContent::Inline(InlineNode::Run(run))
}

pub(super) fn attributes(attrs: &Attrs) -> Result<Properties, ProjectionError> {
    properties(attrs)
}

fn properties<T: serde::Serialize>(input: &T) -> Result<Properties, ProjectionError> {
    match json_value(input)? {
        Value::Object(values) => Ok(values),
        _ => Err(ProjectionError::InvalidInput("expected a formatting object".to_owned())),
    }
}

fn direct(attrs: &Properties) -> Properties {
    let mut result = Properties::new();
    for (mark, property) in [
        ("bold", "bold"), ("italic", "italic"), ("boldCs", "boldCs"), ("italicCs", "italicCs"),
        ("allCaps", "allCaps"), ("smallCaps", "smallCaps"), ("emboss", "emboss"),
        ("imprint", "imprint"), ("textShadow", "shadow"), ("textOutline", "outline"),
        ("hidden", "hidden"), ("rtl", "rtl"),
    ] {
        if attrs.get(mark).is_some_and(good) {
            result.insert(property.to_owned(), Value::Bool(true));
        }
    }
    let object = |name| attrs.get(name).and_then(Value::as_object);
    if let Some(strike) = object("strike") {
        let property = if strike.get("double").is_some_and(good) { "doubleStrike" } else { "strike" };
        result.insert(property.to_owned(), Value::Bool(true));
    }
    if let Some(underline) = object("underline").filter(|_| !lent(attrs, "underline")) {
        let style = underline.get("style").and_then(Value::as_str).filter(|style| !style.is_empty()).unwrap_or("single");
        let mut underline_value = json!({ "style": style });
        if let Some(color) = underline.get("color").filter(|color| color.is_object()) {
            underline_value["color"] = color.clone();
        }
        result.insert("underline".to_owned(), underline_value);
    }
    for (mark, property, keys) in [
        ("textColor", "color", &["rgb", "themeColor", "themeTint", "themeShade"][..]),
        ("fontFamily", "fontFamily", &["ascii", "hAnsi", "eastAsia", "cs", "asciiTheme", "hAnsiTheme", "eastAsiaTheme", "csTheme", "hint"]),
    ] {
        if let Some(values) = object(mark).filter(|_| !lent(attrs, mark)) {
            let typed = keys.iter().filter_map(|key| {
                values.get(*key).filter(|value| value.is_string()).map(|value| ((*key).to_owned(), value.clone()))
            }).collect();
            result.insert(property.to_owned(), Value::Object(typed));
        }
    }
    if let Some(alignment) = ["subscript", "superscript"].into_iter().find(|key| attrs.get(*key).is_some_and(good)) {
        result.insert("vertAlign".to_owned(), Value::String(alignment.to_owned()));
    }
    for (mark, field, property) in [
        ("fontSize", "size", "fontSize"), ("fontSize", "sizeCs", "fontSizeCs"),
        ("characterSpacing", "spacing", "spacing"), ("characterSpacing", "position", "position"),
        ("characterSpacing", "scale", "scale"), ("characterSpacing", "kerning", "kerning"),
    ] {
        set(&mut result, property, object(mark).and_then(|value| value.get(field)).filter(|value| {
            value.as_f64().is_some_and(f64::is_finite)
        }));
    }
    for (mark, field, property, fallback) in [
        ("emphasisMark", "type", "emphasisMark", "dot"),
        ("textEffect", "effect", "effect", "blinkBackground"),
        ("runStyle", "styleId", "styleId", ""),
    ] {
        if let Some(mark) = object(mark) {
            let value = mark.get(field).and_then(Value::as_str).filter(|text| !text.is_empty()).unwrap_or(fallback);
            if !value.is_empty() {
                result.insert(property.to_owned(), Value::String(value.to_owned()));
            }
        }
    }
    let highlight = attrs.get("highlight").and_then(|value| value.as_str().or_else(|| value.get("color")?.as_str()));
    if let Some(highlight) = highlight {
        result.insert("highlight".to_owned(), Value::String(highlight.to_owned()));
    }
    if let Some(typography) = object("typography") {
        set(&mut result, "ligatures", typography.get("ligatures").filter(|value| value.is_string()));
        set(&mut result, "contextualAlternates", typography.get("contextualAlternates").filter(|value| value.is_boolean()));
    }
    set(&mut result, "modernEffects", object("modernTextEffects").and_then(|mark| mark.get("effects")).filter(|value| good(value)));
    if let Some(language) = attrs.get("language") {
        set(&mut result, "language", live_language(None, language).as_ref());
    }
    result
}

fn lent(attrs: &Properties, name: &str) -> bool {
    attrs.get(name).and_then(|mark| mark.get("inheritedHyperlink")).and_then(Value::as_bool) == Some(true)
}

pub(super) fn live_language(source: Option<&Value>, live: &Value) -> Option<Value> {
    let mut result = source.and_then(Value::as_object).cloned().unwrap_or_default();
    if let Some(patch) = live.as_object() {
        for slot in ["latin", "eastAsia", "bidi"] {
            let value = patch.get(slot).or_else(|| (slot == "latin").then(|| patch.get("val")).flatten());
            if let Some(value) = value {
                set(&mut result, slot, value.is_string().then_some(value));
            }
        }
    }
    (!result.is_empty()).then_some(Value::Object(result))
}

fn relative(
    attrs: &Properties,
    live: &Properties,
    inherited: &Properties,
    source: Option<&TextFormatting>,
    theme: &Theme,
) -> Result<Option<TextFormatting>, ProjectionError> {
    let source = source.map(properties).transpose()?.unwrap_or_default();
    let mut result = Properties::new();
    for key in live.keys().chain(inherited.keys()) {
        let key = key.as_str();
        if matches!(key, "fontFamily" | "language")
            || (matches!(key, "boldCs" | "italicCs") && !attrs.get(key).is_some_and(Value::is_boolean))
        {
            continue;
        }
        let value = live.get(key);
        if same_option(value, inherited.get(key)) {
            continue;
        }
        if value.is_some() {
            set(&mut result, key, value);
        } else if ON_OFF.contains(&key) {
            result.insert(key.to_owned(), Value::Bool(false));
        } else if key == "color" && !lent(attrs, "textColor") {
            result.insert(key.to_owned(), json!({ "auto": true }));
        } else if key == "underline" && !lent(attrs, "underline") {
            result.insert(key.to_owned(), json!({ "style": "none" }));
        }
    }
    let fonts = relative_fonts(live.get("fontFamily"), inherited.get("fontFamily"), source.get("fontFamily"), theme);
    set(&mut result, "fontFamily", fonts.as_ref());
    for key in SOURCE_PROPERTIES {
        let Some(value) = source.get(*key) else { continue };
        if *key == "fontFamily" || (*key == "snapToGrid" && value.as_bool() != Some(false)) {
            continue;
        }
        let off = Value::Bool(false);
        let live_value = if *key == "snapToGrid" {
            attrs.get(*key)
        } else {
            live.get(*key).or_else(|| ON_OFF.contains(key).then_some(&off))
        };
        if same_option(Some(value), live_value)
            || (matches!(*key, "language" | "boldCs" | "italicCs") && !attrs.contains_key(*key))
        {
            set(&mut result, key, Some(value));
        }
    }
    if let Some(language) = attrs.get("language") {
        set(&mut result, "language", live_language(source.get("language"), language).as_ref());
    }
    if result.is_empty() { Ok(None) } else { decode(Some(&Value::Object(result))) }
}

fn relative_fonts(live: Option<&Value>, inherited: Option<&Value>, source: Option<&Value>, theme: &Theme) -> Option<Value> {
    let mut result = Properties::new();
    for (slot, reference) in [
        ("ascii", "asciiTheme"), ("hAnsi", "hAnsiTheme"), ("eastAsia", "eastAsiaTheme"), ("cs", "csTheme"),
    ] {
        let canonical = |fonts: Option<&Value>| {
            let font_ref = fonts.and_then(|value| value.get(reference)).and_then(Value::as_str);
            let name = fonts.and_then(|value| value.get(slot)).and_then(Value::as_str);
            let explicit = name.filter(|font| font_ref.is_none_or(|reference| {
                resolve_theme_font_ref(Some(theme), reference) != *font
            }));
            (font_ref.map(str::to_owned), explicit.map(str::to_owned))
        };
        let source_has_slot = source.is_some_and(|value| value.get(slot).is_some() || value.get(reference).is_some());
        let expected = if source_has_slot { source } else { inherited };
        let values = if canonical(live) == canonical(expected) { source } else { live };
        for name in [slot, reference] {
            set(&mut result, name, values.and_then(|values| values.get(name)).filter(|value| !value.is_null()));
        }
    }
    set(&mut result, "hint", source.and_then(|value| value.get("hint")).or_else(|| live.and_then(|value| value.get("hint"))));
    (!result.is_empty()).then_some(Value::Object(result))
}

fn is_plain(content: &RunContent) -> bool {
    matches!(content, RunContent::Text { .. })
}

fn plain_source(paragraph: &Paragraph) -> Option<Vec<(String, Option<&TextFormatting>)>> {
    paragraph.content.iter().map(|content| {
        let ParagraphContent::Inline(InlineNode::Run(run)) = content else { return None };
        let mut text = String::new();
        for entry in &run.content {
            match entry {
                RunContent::Text { text: entry, .. } => text.push_str(entry),
                RunContent::SoftHyphen => text.push('\u{ad}'),
                RunContent::NoBreakHyphen => text.push('\u{2011}'),
                _ => return None,
            }
        }
        Some((text, run.formatting.as_ref()))
    }).collect()
}

fn marks_key(key: Option<&str>) -> Result<Properties, ProjectionError> {
    let mut attrs = Properties::new();
    let mut remaining = key.unwrap_or_default().trim_start_matches('|');
    while !remaining.is_empty() {
        let (name, json) = remaining.split_once(':')
            .filter(|(name, _)| !name.is_empty() && !name.contains('|'))
            .ok_or_else(|| ProjectionError::InvalidInput("invalid original run mark".to_owned()))?;
        let (value, rest) = crate::json_tree::parse_prefix(json).map_err(ProjectionError::InvalidInput)?;
        let Value::Object(value) = value.into_value().map_err(ProjectionError::InvalidInput)? else {
            return Err(ProjectionError::InvalidInput("original run mark must be an object".to_owned()));
        };
        let rest = rest.trim_start_matches([' ', '\t', '\r', '\n']);
        remaining = if rest.is_empty() {
            rest
        } else {
            rest.strip_prefix('|').map(|rest| rest.trim_start_matches('|'))
                .ok_or_else(|| ProjectionError::InvalidInput("invalid original run mark separator".to_owned()))?
        };
        match name {
            "comment" | "footnoteRef" => continue,
            name if BOOLEAN_MARKS.contains(&name) => set(&mut attrs, name, Some(&Value::Bool(true))),
            "boldCs" | "italicCs" => set(&mut attrs, name, value.get("value")),
            "highlight" => set(&mut attrs, name, value.get("color")),
            "rPrChange" => set(&mut attrs, name, value.get("changes")),
            "directionScopes" => set(&mut attrs, name, value.get("scopes")),
            "insertion" | "deletion" => continue,
            _ => {
                let value = value.into_iter().filter(|(_, value)| !value.is_null()).collect();
                attrs.insert(name.to_owned(), Value::Object(value));
            }
        }
    }
    Ok(attrs)
}

fn properties_equal(left: &Properties, right: &Properties) -> bool {
    left.len() == right.len() && left.iter().all(|(key, value)| same_option(Some(value), right.get(key)))
}

fn restore(
    boundaries: &[OriginalBoundary],
    sources: &[(String, Option<&TextFormatting>)],
    text: &str,
    items: &[(&str, &Attrs)],
    attrs: &[Properties],
) -> Result<Option<Vec<ParagraphContent>>, ProjectionError> {
    if boundaries.is_empty() || boundaries.iter().any(|boundary| {
        boundary.breaks.as_ref().is_some_and(|breaks| !breaks.is_empty())
            || boundary.note_marks.as_ref().is_some_and(|marks| !marks.is_empty())
    }) || text != boundaries.iter().map(|boundary| boundary.text.as_str()).collect::<String>() {
        return Ok(None);
    }
    let source_matches = sources.len() == boundaries.len()
        && sources.iter().zip(boundaries).all(|((text, _), boundary)| text == &boundary.text);
    let mut current = items.iter().zip(attrs).peekable();
    let mut consumed = 0;
    let mut result = Vec::with_capacity(boundaries.len());
    for (index, boundary) in boundaries.iter().enumerate() {
        let expected = marks_key(boundary.marks_key.as_deref())?;
        let mut remaining = boundary.text.encode_utf16().count();
        while remaining != 0 {
            let Some(((text, _), attrs)) = current.peek() else { return Ok(None) };
            let filtered: Properties = attrs.iter().filter(|(key, _)| {
                !matches!(key.as_str(), "hyperlink" | "ins" | "del")
            }).map(|(key, value)| (key.clone(), value.clone())).collect();
            if !properties_equal(&expected, &filtered) {
                return Ok(None);
            }
            let len = text.encode_utf16().count();
            let used = remaining.min(len - consumed);
            remaining -= used;
            consumed += used;
            if consumed == len {
                current.next();
                consumed = 0;
            }
        }
        let mut formatting = boundary.formatting.clone().unwrap_or_default();
        if source_matches {
            let source = sources[index].1;
            formatting.font_family = source.and_then(|source| source.font_family.clone());
            formatting.language = source.and_then(|source| source.language.clone());
        }
        result.push(inline(Run {
            node_type: RunType::Run,
            formatting: (formatting != TextFormatting::default()).then_some(formatting),
            property_changes: boundary.property_changes.clone(),
            content: run_content(&boundary.text, &direct(&expected)),
            direction_scopes: decode(expected.get("directionScopes"))?.unwrap_or_default(),
        }));
    }
    Ok(current.next().is_none().then_some(result))
}

struct SourceSpan {
    range: Range<usize>,
    index: usize,
    changed: bool,
    previous_changed: Option<usize>,
}

fn source_spans(
    sources: &[(String, Option<&TextFormatting>)],
    original: &[char],
    live: &[char],
) -> Vec<SourceSpan> {
    if sources.is_empty() {
        return Vec::new();
    }
    let diff = diff_tokens(original, live, DiffLimits::new(usize::MAX, 1 << 20, LimitFallback::ReplaceMiddle))
        .expect("replace fallback handles the budget");
    let mut matches = diff.hunks.iter().filter(|hunk| hunk.kind == DiffKind::Equal).peekable();
    let mut original_end = 0;
    let mut live_end = 0;
    let mut mapped = |offset| {
        if offset == 0 { return 0; }
        while let Some(region) = matches.peek() {
            if offset < region.old.start {
                return if offset == original_end { live_end } else { region.new.start };
            }
            if offset <= region.old.end {
                return region.new.start + offset - region.old.start;
            }
            original_end = region.old.end;
            live_end = region.new.end;
            matches.next();
        }
        if offset == original_end { live_end } else { live.len() }
    };

    let mut old_start = 0;
    let mut start = 0;
    let mut previous_changed = None;
    let mut spans = Vec::new();
    for (index, (text, _)) in sources.iter().enumerate() {
        let old_end = old_start + text.chars().count();
        let end = if index + 1 == sources.len() { live.len() } else { mapped(old_end) };
        let changed = original[old_start..old_end] != live[start..end];
        if start != end {
            spans.push(SourceSpan { range: start..end, index, changed, previous_changed });
        }
        if !changed || start != end {
            previous_changed = changed.then_some(index);
        }
        old_start = old_end;
        start = end;
    }
    spans
}

fn run_content(text: &str, formatting: &Properties) -> Vec<RunContent> {
    let fonts: Vec<_> = ["ascii", "hAnsi", "eastAsia", "cs"].into_iter().filter_map(|slot| {
        formatting.get("fontFamily")?.get(slot)?.as_str().filter(|font| !font.is_empty())
    }).collect();
    let symbol_font = fonts.first().filter(|first| fonts.iter().all(|font| font == *first));
    let mut pending = String::new();
    let mut content = Vec::new();
    for char in text.chars() {
        let node = match char {
            '\u{ad}' => RunContent::SoftHyphen,
            '\u{2011}' => RunContent::NoBreakHyphen,
            '\u{f000}'..='\u{f8ff}' if symbol_font.is_some() => RunContent::Symbol {
                font: symbol_font.expect("symbol font exists").to_string(),
                char: format!("{:04X}", u32::from(char)),
            },
            _ => {
                pending.push(char);
                continue;
            }
        };
        if !pending.is_empty() {
            content.push(RunContent::Text { text: std::mem::take(&mut pending), preserve_space: None });
        }
        content.push(node);
    }
    if !pending.is_empty() {
        content.push(RunContent::Text { text: pending, preserve_space: None });
    }
    content
}

fn append(target: &mut Run, source: Run) {
    for entry in source.content {
        match (target.content.last_mut(), entry) {
            (Some(RunContent::Text { text, .. }), RunContent::Text { text: next, .. }) => text.push_str(&next),
            (_, entry) => target.content.push(entry),
        }
    }
}
