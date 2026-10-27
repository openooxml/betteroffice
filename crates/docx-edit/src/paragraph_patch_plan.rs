use std::collections::{BTreeMap, BTreeSet};
use std::rc::Rc;

use docx_parse::paragraph_patch::{
    AttrDeltaV1, Attrs, DonorV1, InheritanceV1, InsertedTextV1, LegacyParagraphV1,
    ParagraphPatchV1, PatchErrorCode, RunEditV1, SeedV1, SlotV1, TextEditV1, UnitV1,
    canonical_json_v1, result_sha256, unit_sha256,
};
use serde_json::{Value, json};
use yrs::Any;

use crate::{SegmentContent, StorySegment};

pub(crate) struct PatchContext<'a> {
    pub ordinal: u32,
    pub session_key: &'a str,
    pub story: &'a str,
    pub metadata: Option<&'a crate::seed::SourceMetadata>,
    pub seeded_markers: &'a [(u32, String, bool)],
    pub live_markers: &'a [(u32, String, bool)],
}

impl PatchContext<'_> {
    fn legacy(&self, code: PatchErrorCode, detail: impl Into<String>) -> LegacyParagraphV1 {
        LegacyParagraphV1 {
            ordinal: self.ordinal,
            code,
            detail: detail.into(),
        }
    }

    fn restyle(&self, detail: &str) -> LegacyParagraphV1 {
        self.legacy(
            PatchErrorCode::UnsupportedStructure,
            format!("restyle: {detail}"),
        )
    }

    fn inheritance(&self, properties: &Attrs, attrs: &Attrs) -> (Attrs, Attrs) {
        let style = properties.get("pStyle").and_then(Value::as_str);
        let run_style = attrs
            .get("runStyle")
            .and_then(|value| value.get("styleId"))
            .and_then(Value::as_str);
        let (properties, run) = self.metadata.map_or_else(
            || (Attrs::new(), Attrs::new()),
            |metadata| metadata.patch_inheritance(self.story, style, run_style),
        );
        let normalize = |attrs: Attrs| {
            attrs
                .into_iter()
                .map(|(key, value)| (key, wire_value(value)))
                .collect()
        };
        (normalize(properties), normalize(run))
    }
}

const CARRIERS: &[&str] = &[
    "_originalFormatting",
    "_originalRunBoundaries",
    "defaultTextFormatting",
    "spacingExplicit",
    "bookmarks",
];

struct Snapshot {
    units: Vec<UnitV1>,
    properties: Attrs,
    carriers: Attrs,
    carried_properties: BTreeSet<String>,
    run_carriers: Vec<Rc<Attrs>>,
}

fn attrs(values: &BTreeMap<String, Any>) -> Result<Attrs, String> {
    fn finite(value: &Any) -> bool {
        match value {
            Any::Number(value) => value.is_finite(),
            Any::Array(values) => values.iter().all(finite),
            Any::Map(values) => values.values().all(finite),
            Any::Undefined => false,
            _ => true,
        }
    }
    values
        .iter()
        .map(|(key, value)| {
            if !finite(value) {
                return Err(format!("key:{key} non-JSON value"));
            }
            serde_json::to_value(value)
                .map(|value| (key.clone(), wire_value(value)))
                .map_err(|error| format!("key:{key} {error}"))
        })
        .collect()
}

fn wire_value(value: Value) -> Value {
    match value {
        Value::Number(number) => {
            // Hash the number representation that survives the JavaScript transport.
            let number = number.as_f64().expect("finite JSON number");
            serde_json::from_str(ryu_js::Buffer::new().format(number)).expect("finite JSON number")
        }
        Value::Array(values) => Value::Array(values.into_iter().map(wire_value).collect()),
        Value::Object(values) => Value::Object(
            values
                .into_iter()
                .map(|(key, value)| (key, wire_value(value)))
                .collect(),
        ),
        value => value,
    }
}

fn snapshot(segments: &[StorySegment]) -> Result<Snapshot, String> {
    let mut units = Vec::new();
    let mut properties = Attrs::new();
    let mut carriers = Attrs::new();
    let mut carried_properties = BTreeSet::new();
    let mut run_carriers = Vec::new();
    for segment in segments {
        let mut attributes = attrs(&segment.attributes)?;
        let mut derived = Attrs::new();
        for key in ["textColor", "underline"] {
            if let Some(value) = attributes.get_mut(key).and_then(Value::as_object_mut)
                && let Some(flag) = value.remove("inheritedHyperlink")
            {
                if !flag.is_boolean() {
                    return Err(format!("key:{key}.inheritedHyperlink"));
                }
                derived.insert(key.to_owned(), flag);
            }
        }
        let count = match &segment.content {
            SegmentContent::Text(text) => text.chars().count(),
            SegmentContent::OtherEmbed { .. } => 1,
            SegmentContent::Pilcrow(_) => 0,
        };
        run_carriers.extend(std::iter::repeat_n(Rc::new(derived), count));
        match &segment.content {
            SegmentContent::Text(text) => push_text(&mut units, text, &attributes),
            SegmentContent::OtherEmbed { kind, payload } => units.push(UnitV1::Atom {
                atom_kind: kind.clone(),
                payload: attrs(payload)?,
                attrs: attributes,
            }),
            SegmentContent::Pilcrow(pilcrow) => {
                properties = attrs(&pilcrow.values)?;
                if let Some(original) = properties
                    .get("_originalFormatting")
                    .and_then(Value::as_object)
                {
                    let carried: Vec<_> = original
                        .iter()
                        .filter(|(key, _)| {
                            para_key(std::slice::from_ref(*key))
                                || matches!(key.as_str(), "pStyle" | "numPr")
                        })
                        .map(|(key, value)| (key.clone(), value.clone()))
                        .collect();
                    for (key, value) in carried {
                        properties.entry(key.clone()).or_insert_with(|| {
                            carried_properties.insert(key);
                            value
                        });
                    }
                }
                for &key in CARRIERS {
                    if let Some(value) = properties.remove(key) {
                        carriers.insert(key.to_owned(), value);
                    }
                }
                if !attributes.is_empty() {
                    carriers.insert("pilcrowAttrs".to_owned(), json!(attributes));
                }
            }
        }
    }
    Ok(Snapshot {
        units,
        properties,
        carriers,
        carried_properties,
        run_carriers,
    })
}

fn push_text(units: &mut Vec<UnitV1>, text: &str, attrs: &Attrs) {
    if text.is_empty() {
        return;
    }
    if let Some(UnitV1::Text {
        text: previous,
        attrs: previous_attrs,
    }) = units.last_mut()
        && previous_attrs == attrs
    {
        previous.push_str(text);
    } else {
        units.push(UnitV1::Text {
            text: text.to_owned(),
            attrs: attrs.clone(),
        });
    }
}

fn run_key(path: &[String]) -> bool {
    let Some(key) = path.first().map(String::as_str) else {
        return false;
    };
    let children: Option<&[&str]> = match key {
        "bold" | "italic" | "boldCs" | "italicCs" | "superscript" | "subscript" | "allCaps"
        | "smallCaps" | "hidden" | "doubleStrike" => None,
        "underline" => Some(&["style", "color"]),
        "strike" => Some(&["double"]),
        "textColor" => Some(&["rgb", "themeColor", "themeTint", "themeShade"]),
        "highlight" => Some(&["color"]),
        "fontSize" => Some(&["size", "sizeCs"]),
        "fontFamily" => Some(&[
            "ascii",
            "hAnsi",
            "eastAsia",
            "cs",
            "asciiTheme",
            "hAnsiTheme",
            "eastAsiaTheme",
            "csTheme",
        ]),
        _ => return false,
    };
    path.len() == 1
        || (path.len() == 2 && children.is_some_and(|keys| keys.contains(&path[1].as_str())))
}

fn para_key(path: &[String]) -> bool {
    let Some(key) = path.first().map(String::as_str) else {
        return false;
    };
    match key {
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
        | "indentLeftChars"
        | "indentRightChars"
        | "indentFirstLineChars"
        | "hangingIndent"
        | "hangingIndentChars"
        | "keepNext"
        | "keepLines"
        | "widowControl"
        | "pageBreakBefore"
        | "contextualSpacing" => path.len() == 1,
        _ => false,
    }
}

fn deltas(before: &Attrs, after: &Attrs) -> Vec<AttrDeltaV1> {
    if std::ptr::eq(before, after) || before == after {
        return Vec::new();
    }
    let mut result = Vec::new();
    for key in before.keys().chain(after.keys()).collect::<BTreeSet<_>>() {
        delta_at(
            vec![key.clone()],
            before.get(key),
            after.get(key),
            &mut result,
        );
    }
    result
}

fn delta_at(
    path: Vec<String>,
    before: Option<&Value>,
    after: Option<&Value>,
    result: &mut Vec<AttrDeltaV1>,
) {
    if before == after {
        return;
    }
    let old = before.and_then(Value::as_object);
    let new = after.and_then(Value::as_object);
    if (old.is_some() || before.is_none())
        && (new.is_some() || after.is_none())
        && (old.is_some() || new.is_some())
        && !old.is_some_and(|map| map.is_empty())
        && !new.is_some_and(|map| map.is_empty())
    {
        let keys: BTreeSet<_> = old
            .into_iter()
            .flat_map(|map| map.keys())
            .chain(new.into_iter().flat_map(|map| map.keys()))
            .collect();
        if !keys.is_empty() {
            for key in keys {
                let mut child = path.clone();
                child.push(key.clone());
                delta_at(
                    child,
                    old.and_then(|map| map.get(key)),
                    new.and_then(|map| map.get(key)),
                    result,
                );
            }
            return;
        }
    }
    let slot = |value: Option<&Value>| {
        value.map_or(SlotV1::Missing, |value| SlotV1::Present {
            value: value.clone(),
        })
    };
    result.push(AttrDeltaV1 {
        path,
        before: slot(before),
        after: slot(after),
    });
}

fn verify_changes(before: &Attrs, changes: &[AttrDeltaV1], after: &Attrs) -> bool {
    if changes.is_empty() {
        return std::ptr::eq(before, after) || before == after;
    }
    fn lookup<'a>(value: &'a Value, path: &[String]) -> Option<&'a Value> {
        path.iter().try_fold(value, |value, key| value.get(key))
    }
    fn apply(value: &mut Value, path: &[String], slot: &SlotV1, target: Option<&Value>) -> bool {
        let Some(map) = value.as_object_mut() else {
            return false;
        };
        let key = &path[0];
        if path.len() == 1 {
            match slot {
                SlotV1::Missing => {
                    map.remove(key);
                }
                SlotV1::Present { value } => {
                    map.insert(key.clone(), value.clone());
                }
            }
            return true;
        }
        let child = map.entry(key.clone()).or_insert_with(|| json!({}));
        let target = target.and_then(|value| value.get(key));
        if !apply(child, &path[1..], slot, target) {
            return false;
        }
        if target.is_none() && child.as_object().is_some_and(|map| map.is_empty()) {
            map.remove(key);
        }
        true
    }
    let target = json!(after);
    let mut replay = json!(before);
    for delta in changes {
        let expected = match &delta.before {
            SlotV1::Missing => None,
            SlotV1::Present { value } => Some(value),
        };
        if lookup(&replay, &delta.path) != expected
            || !apply(&mut replay, &delta.path, &delta.after, Some(&target))
        {
            return false;
        }
    }
    replay == target
}

#[derive(Clone, Debug, PartialEq)]
enum Content {
    Text(char),
    Atom(String, Attrs),
}

#[derive(Clone, Debug)]
struct Token {
    content: Content,
    attrs: Rc<Attrs>,
}

fn same_attrs(left: &Rc<Attrs>, right: &Rc<Attrs>) -> bool {
    Rc::ptr_eq(left, right) || left == right
}

impl PartialEq for Token {
    fn eq(&self, other: &Self) -> bool {
        self.content == other.content && same_attrs(&self.attrs, &other.attrs)
    }
}

impl Token {
    fn width(&self) -> u32 {
        match self.content {
            Content::Text(value) => value.len_utf16() as u32,
            Content::Atom(..) => 1,
        }
    }

    fn is_text(&self) -> bool {
        matches!(self.content, Content::Text(_))
    }
}

fn tokens(units: &[UnitV1]) -> Vec<Token> {
    let mut result = Vec::new();
    for unit in units {
        match unit {
            UnitV1::Text { text, attrs } => {
                let attrs = Rc::new(attrs.clone());
                result.extend(text.chars().map(|value| Token {
                    content: Content::Text(value),
                    attrs: Rc::clone(&attrs),
                }));
            }
            UnitV1::Atom {
                atom_kind,
                payload,
                attrs,
            } => result.push(Token {
                content: Content::Atom(atom_kind.clone(), payload.clone()),
                attrs: Rc::new(attrs.clone()),
            }),
        }
    }
    result
}

fn units(tokens: &[Token]) -> Vec<UnitV1> {
    let mut result = Vec::new();
    let mut previous: Option<&Token> = None;
    for token in tokens {
        match &token.content {
            Content::Text(value) => {
                if previous.is_some_and(|old| same_attrs(&old.attrs, &token.attrs))
                    && let Some(UnitV1::Text { text, .. }) = result.last_mut()
                {
                    text.push(*value);
                } else {
                    push_text(&mut result, &value.to_string(), &token.attrs);
                }
            }
            Content::Atom(kind, payload) => result.push(UnitV1::Atom {
                atom_kind: kind.clone(),
                payload: payload.clone(),
                attrs: token.attrs.as_ref().clone(),
            }),
        }
        previous = Some(token);
    }
    result
}

struct Match {
    pair: (usize, usize),
    previous: Option<Rc<Match>>,
}

#[derive(Clone, Default)]
struct Alignment {
    length: usize,
    path: Option<Rc<Match>>,
    ambiguous: bool,
}

impl Alignment {
    fn merge(left: &Self, right: &Self) -> Self {
        if left.length > right.length {
            return left.clone();
        }
        if right.length > left.length {
            return right.clone();
        }
        let same = match (&left.path, &right.path) {
            (None, None) => true,
            (Some(left), Some(right)) => Rc::ptr_eq(left, right),
            _ => false,
        };
        if left.ambiguous || right.ambiguous || !same {
            Self {
                length: left.length,
                path: None,
                ambiguous: true,
            }
        } else {
            left.clone()
        }
    }
}

fn alignment(before: &[Token], after: &[Token]) -> Result<Vec<(usize, usize)>, &'static str> {
    const WORK_LIMIT: usize = 4_000_000;
    let prefix = before
        .iter()
        .zip(after)
        .take_while(|(old, new)| old.content == new.content)
        .count();
    if prefix == before.len() && prefix == after.len() {
        return Ok((0..prefix).map(|at| (at, at)).collect());
    }
    let suffix = before[prefix..]
        .iter()
        .rev()
        .zip(after[prefix..].iter().rev())
        .take_while(|(old, new)| old.content == new.content)
        .count();
    // Keep one boundary token on each side to detect ambiguous edge matches.
    let prefix = prefix.saturating_sub(1);
    let suffix = suffix.saturating_sub(1);
    let old_end = before.len() - suffix;
    let new_end = after.len() - suffix;
    let old_middle = &before[prefix..old_end];
    let new_middle = &after[prefix..new_end];
    if (old_middle.len() + 1)
        .checked_mul(new_middle.len() + 1)
        .is_none_or(|cells| cells > WORK_LIMIT)
    {
        return Err("ambiguous: alignment too large");
    }
    let middle =
        middle_alignment(old_middle, new_middle).ok_or("ambiguous: nonunique content alignment")?;
    Ok((0..prefix)
        .map(|at| (at, at))
        .chain(middle.into_iter().map(|(i, j)| (prefix + i, prefix + j)))
        .chain((0..suffix).map(|at| (old_end + at, new_end + at)))
        .collect())
}

fn middle_alignment(before: &[Token], after: &[Token]) -> Option<Vec<(usize, usize)>> {
    let mut row = vec![Alignment::default(); after.len() + 1];
    for (i, old) in before.iter().enumerate() {
        let mut diagonal = Alignment::default();
        for (j, new) in after.iter().enumerate() {
            let above = row[j + 1].clone();
            let mut best = Alignment::merge(&row[j], &above);
            if old.content == new.content {
                let length = diagonal.length + 1;
                if length > best.length {
                    let path = (!diagonal.ambiguous).then(|| {
                        Rc::new(Match {
                            pair: (i, j),
                            previous: diagonal.path.clone(),
                        })
                    });
                    best = Alignment {
                        length,
                        path,
                        ambiguous: diagonal.ambiguous,
                    };
                } else if length == best.length {
                    best = Alignment {
                        length,
                        path: None,
                        ambiguous: true,
                    };
                }
            }
            row[j + 1] = best;
            diagonal = above;
        }
    }
    let result = row.pop()?;
    drop(row);
    if result.ambiguous {
        return None;
    }
    let mut pairs = Vec::with_capacity(result.length);
    let mut path = result.path;
    while let Some(node) = path {
        pairs.push(node.pair);
        path = node.previous.clone();
    }
    pairs.reverse();
    Some(pairs)
}

fn inherited_runs(
    context: &PatchContext<'_>,
    properties: &Attrs,
    tokens: &[Token],
) -> Vec<Rc<Attrs>> {
    let mut cache = BTreeMap::new();
    let mut result: Vec<Rc<Attrs>> = Vec::with_capacity(tokens.len());
    for (at, token) in tokens.iter().enumerate() {
        if at > 0 && Rc::ptr_eq(&tokens[at - 1].attrs, &token.attrs) {
            result.push(Rc::clone(result.last().expect("preceding token")));
            continue;
        }
        let key = token.attrs.get("runStyle").map(canonical_json_v1);
        let inherited = cache
            .entry(key)
            .or_insert_with(|| Rc::new(run_inheritance(context, properties, &token.attrs)));
        result.push(Rc::clone(inherited));
    }
    result
}

fn run_inheritance(context: &PatchContext<'_>, properties: &Attrs, attrs: &Attrs) -> Attrs {
    let mut inherited = context.inheritance(properties, attrs).1;
    for key in ["textColor", "underline"] {
        if let Some(value) = inherited.get_mut(key).and_then(Value::as_object_mut) {
            value.remove("inheritedHyperlink");
        }
    }
    inherited
}

fn mark_attrs(snapshot: &Snapshot) -> Attrs {
    let style = snapshot
        .carriers
        .get("defaultTextFormatting")
        .and_then(|formatting| formatting.get("styleId"));
    let mut attrs = Attrs::new();
    if let Some(style) = style {
        attrs.insert("runStyle".to_owned(), json!({ "styleId": style }));
    }
    attrs
}

fn mark_inheritance(context: &PatchContext<'_>, snapshot: &Snapshot) -> Attrs {
    run_inheritance(context, &snapshot.properties, &mark_attrs(snapshot))
}

fn unsupported(snapshot: &Snapshot) -> Option<String> {
    for unit in &snapshot.units {
        let attributes = match unit {
            UnitV1::Text { attrs, .. } => attrs,
            UnitV1::Atom {
                atom_kind,
                payload,
                attrs,
            } => {
                let comment = atom_kind == "field"
                    && payload.get("modelKind").and_then(Value::as_str) == Some("commentReference");
                if !matches!(atom_kind.as_str(), "break" | "noteRef") && !comment {
                    let kind = match atom_kind.as_str() {
                        "field" => "field",
                        "sdt" | "blockSdt" => "sdt",
                        "math" => "math",
                        "image" | "shape" | "shapeGroup" | "chart" | "opaqueDrawing"
                        | "horizontalRule" => "drawing",
                        _ => return Some(format!("key:{atom_kind} unsupported atom")),
                    };
                    return Some(format!("{kind}: atom {atom_kind}"));
                }
                attrs
            }
        };
        if attributes
            .keys()
            .any(|key| matches!(key.as_str(), "ins" | "del"))
        {
            return Some("tc: revision marks".to_owned());
        }
        if attributes
            .keys()
            .any(|key| matches!(key.as_str(), "fieldResult" | "fieldCode"))
        {
            return Some("field: field marks".to_owned());
        }
    }
    if ["pPrIns", "pPrDel", "pPrChange"]
        .iter()
        .any(|key| snapshot.properties.contains_key(*key))
    {
        return Some("tc: paragraph revision".to_owned());
    }
    if snapshot
        .carriers
        .get("_originalRunBoundaries")
        .and_then(Value::as_array)
        .is_some_and(|runs| {
            runs.iter().any(|run| {
                run.get("propertyChanges")
                    .and_then(Value::as_array)
                    .is_some_and(|changes| !changes.is_empty())
            })
        })
    {
        return Some("tc: run property revision".to_owned());
    }
    None
}

fn validate_deltas(
    changes: &[AttrDeltaV1],
    accepts: fn(&[String]) -> bool,
    context: &PatchContext<'_>,
) -> Result<(), LegacyParagraphV1> {
    fn unknown_leaf(
        path: &[String],
        value: &Value,
        accepts: fn(&[String]) -> bool,
    ) -> Option<Vec<String>> {
        if let Value::Object(values) = value {
            for (key, value) in values {
                let mut child = path.to_vec();
                child.push(key.clone());
                if !accepts(&child) {
                    return Some(child);
                }
                if let Some(path) = unknown_leaf(&child, value, accepts) {
                    return Some(path);
                }
            }
        }
        None
    }
    for delta in changes {
        let rejected = if accepts(&delta.path) {
            [&delta.before, &delta.after]
                .into_iter()
                .find_map(|slot| match slot {
                    SlotV1::Present { value } => unknown_leaf(&delta.path, value, accepts),
                    SlotV1::Missing => None,
                })
        } else {
            Some(delta.path.clone())
        };
        if let Some(path) = rejected {
            let code = if path.first().is_some_and(|key| key == "hyperlink") {
                PatchErrorCode::UnsupportedStructure
            } else {
                PatchErrorCode::UnsupportedKey
            };
            let prefix = if code == PatchErrorCode::UnsupportedStructure {
                "hyperlink:"
            } else {
                "key:"
            };
            return Err(context.legacy(code, format!("{prefix}{}", path.join("."))));
        }
    }
    Ok(())
}

fn attr_at<'a>(attrs: &'a Attrs, path: &[String]) -> Option<&'a Value> {
    let value = attrs.get(&path[0])?;
    path[1..]
        .iter()
        .try_fold(value, |value, key| value.get(key))
}

fn validate_indent_kinds(
    before: &Snapshot,
    after: &Snapshot,
    context: &PatchContext<'_>,
) -> Result<(), LegacyParagraphV1> {
    let source = before.carriers.get("_originalFormatting");
    let live = after.carriers.get("_originalFormatting");
    let ambiguous = |key| {
        context.legacy(
            PatchErrorCode::AmbiguousAlignment,
            format!("ambiguous: indent kind {key}"),
        )
    };
    for flag in ["hangingIndent", "hangingIndentChars"] {
        if before.properties.get(flag) != after.properties.get(flag)
            || source.and_then(|value| value.get(flag)) != live.and_then(|value| value.get(flag))
        {
            return Err(ambiguous(flag));
        }
    }
    for (key, flag) in [
        ("indentFirstLine", "hangingIndent"),
        ("indentFirstLineChars", "hangingIndentChars"),
    ] {
        for (before, after, hanging) in [
            (
                before.properties.get(key),
                after.properties.get(key),
                after.properties.get(flag),
            ),
            (
                source.and_then(|value| value.get(key)),
                live.and_then(|value| value.get(key)),
                live.and_then(|value| value.get(flag)),
            ),
        ] {
            if before == after {
                continue;
            }
            let before = before.and_then(Value::as_f64);
            let after = after.and_then(Value::as_f64);
            let sign_changed = before
                .zip(after)
                .is_some_and(|(old, new)| old.is_sign_negative() != new.is_sign_negative());
            let negative_first_line =
                after.is_some_and(|value| value < 0.0) && hanging != Some(&Value::Bool(true));
            if sign_changed || negative_first_line {
                return Err(ambiguous(key));
            }
        }
    }
    Ok(())
}

fn validate_paragraph_clears(
    before: &Snapshot,
    after: &Snapshot,
    changes: &[AttrDeltaV1],
    context: &PatchContext<'_>,
) -> Result<(), LegacyParagraphV1> {
    let cleared = changes.iter().find_map(|delta| {
        let reset = match &delta.after {
            SlotV1::Missing => true,
            SlotV1::Present { value } => value.is_null(),
        };
        (para_key(&delta.path) && reset).then_some(delta.path[0].as_str())
    });
    let source = before
        .carriers
        .get("_originalFormatting")
        .and_then(Value::as_object);
    let live = after
        .carriers
        .get("_originalFormatting")
        .and_then(Value::as_object);
    let cleared = cleared.or_else(|| {
        source?.iter().find_map(|(key, value)| {
            let live = live.and_then(|values| values.get(key));
            (para_key(std::slice::from_ref(key))
                && !value.is_null()
                && live.is_none_or(Value::is_null))
            .then_some(key.as_str())
        })
    });
    if let Some(key) = cleared {
        return Err(context.legacy(
            PatchErrorCode::AmbiguousAlignment,
            format!("ambiguous: paragraph property clear {key}"),
        ));
    }
    Ok(())
}

fn validate_theme_fonts(
    before: &Attrs,
    changes: &[AttrDeltaV1],
    context: &PatchContext<'_>,
) -> Result<(), LegacyParagraphV1> {
    let themes = [
        "asciiTheme",
        "hAnsiTheme",
        "eastAsiaTheme",
        "csTheme",
        "cstheme",
    ];
    let themed = before.get("fontFamily").is_some_and(|fonts| {
        themes
            .iter()
            .any(|key| fonts.get(*key).is_some_and(|value| !value.is_null()))
    });
    if themed
        && let Some(delta) = changes
            .iter()
            .find(|delta| delta.path.first().is_some_and(|key| key == "fontFamily"))
    {
        let slot = delta.path.get(1).map_or("fontFamily", String::as_str);
        return Err(context.legacy(
            PatchErrorCode::AmbiguousAlignment,
            format!("ambiguous: theme font {slot}"),
        ));
    }
    Ok(())
}

fn validate_carriers(
    before: &Snapshot,
    after: &Snapshot,
    inherited: &Attrs,
    context: &PatchContext<'_>,
) -> Result<(), LegacyParagraphV1> {
    for change in deltas(&before.carriers, &after.carriers) {
        let key = change.path[0].as_str();
        let mirrored = match key {
            "_originalFormatting" if change.path.len() > 1 => {
                let mut path = change.path[1..].to_vec();
                if path[0] == "styleId" {
                    path[0] = "pStyle".to_owned();
                }
                let target = attr_at(&after.properties, &path);
                if para_key(&path) && attr_at(&before.properties, &path) == target {
                    return Err(context.legacy(
                        PatchErrorCode::UnsupportedKey,
                        format!("key:{} direct property change", path.join(".")),
                    ));
                }
                para_key(&path)
                    && match &change.after {
                        SlotV1::Missing => target.is_none() || target == attr_at(inherited, &path),
                        SlotV1::Present { value } => target == Some(value),
                    }
            }
            "bookmarks" => true,
            "spacingExplicit" => {
                deltas(&before.properties, &after.properties)
                    .iter()
                    .any(|delta| {
                        delta.path.first().is_some_and(|key| {
                            key.starts_with("space") || key.starts_with("lineSpacing")
                        })
                    })
            }
            _ => false,
        };
        if !mirrored {
            let code = if key == "bookmarks" || key == "pilcrowAttrs" {
                PatchErrorCode::UnsupportedStructure
            } else {
                PatchErrorCode::UnsupportedKey
            };
            let prefix = if code == PatchErrorCode::UnsupportedStructure {
                "marker:"
            } else {
                "key:"
            };
            return Err(context.legacy(code, format!("{prefix}{}", change.path.join("."))));
        }
    }
    Ok(())
}

struct DonorIndex {
    offsets: Vec<u32>,
    previous: Vec<Option<usize>>,
    next: Vec<Option<usize>>,
}

impl DonorIndex {
    fn new(tokens: &[Token]) -> Self {
        let len = tokens.len();
        let mut offsets = vec![0; len + 1];
        let mut previous = vec![None; len + 1];
        let mut next = vec![None; len + 1];
        for i in 0..len {
            offsets[i + 1] = offsets[i] + tokens[i].width();
            previous[i + 1] = if tokens[i].is_text() {
                Some(i)
            } else {
                previous[i]
            };
            let j = len - 1 - i;
            next[j] = if tokens[j].is_text() {
                Some(j)
            } else {
                next[j + 1]
            };
        }
        Self {
            offsets,
            previous,
            next,
        }
    }

    fn text_donor(&self, from: usize, to: usize) -> Option<usize> {
        if from != to {
            self.next[from].filter(|&at| at < to)
        } else if from == 0 {
            self.next[0]
        } else {
            self.previous[from]
        }
    }

    fn donor(
        &self,
        from: usize,
        to: usize,
        context: &PatchContext<'_>,
    ) -> Result<DonorV1, LegacyParagraphV1> {
        if let Some(index) = self.text_donor(from, to) {
            Ok(DonorV1::At {
                at: self.offsets[index],
            })
        } else if self.offsets.len() == 1 {
            Ok(DonorV1::Pilcrow)
        } else {
            Err(context.legacy(
                PatchErrorCode::AmbiguousDonor,
                "ambiguous: no adjacent text donor",
            ))
        }
    }
}

pub(crate) fn diff_paragraph_v1(
    seeded: &[StorySegment],
    live: &[StorySegment],
    context: PatchContext<'_>,
) -> Result<ParagraphPatchV1, LegacyParagraphV1> {
    let before = snapshot(seeded)
        .map_err(|detail| context.legacy(PatchErrorCode::UnsupportedKey, detail))?;
    let mut after =
        snapshot(live).map_err(|detail| context.legacy(PatchErrorCode::UnsupportedKey, detail))?;
    for key in after
        .carried_properties
        .difference(&before.carried_properties)
    {
        if before.properties.contains_key(key) {
            after.properties.remove(key);
        }
    }
    for key in ["pStyle", "styleId", "numPr"] {
        if before.properties.get(key) != after.properties.get(key) {
            return Err(context.restyle(key));
        }
    }
    let (seeded_properties, _) = context.inheritance(&before.properties, &Attrs::new());
    let (para_inherit, _) = context.inheritance(&after.properties, &Attrs::new());
    let seeded_mark = mark_inheritance(&context, &before);
    let mark_inherit = mark_inheritance(&context, &after);
    if seeded_properties != para_inherit
        || seeded_mark != mark_inherit
        || mark_attrs(&before) != mark_attrs(&after)
    {
        return Err(context.restyle("paragraph inheritance"));
    }
    for path in [
        &["styleId"][..],
        &["numPr"][..],
        &["runProperties", "styleId"][..],
    ] {
        let source_value = |snapshot: &Snapshot| {
            snapshot
                .carriers
                .get("_originalFormatting")
                .and_then(|value| path.iter().try_fold(value, |value, key| value.get(*key)))
                .cloned()
        };
        if source_value(&before) != source_value(&after) {
            return Err(context.restyle("source style or numbering"));
        }
    }
    if let Some(detail) = unsupported(&before).or_else(|| unsupported(&after)) {
        return Err(context.legacy(PatchErrorCode::UnsupportedStructure, detail));
    }
    if context.metadata.is_some_and(|metadata| {
        metadata.patch_has_run_revisions(context.story, context.session_key)
    }) {
        return Err(context.legacy(
            PatchErrorCode::UnsupportedStructure,
            "tc: run property revision",
        ));
    }
    let para_deltas = deltas(&before.properties, &after.properties);
    validate_indent_kinds(&before, &after, &context)?;
    validate_paragraph_clears(&before, &after, &para_deltas, &context)?;
    validate_carriers(&before, &after, &para_inherit, &context)?;
    validate_deltas(&para_deltas, para_key, &context)?;
    if !verify_changes(&before.properties, &para_deltas, &after.properties) {
        return Err(context.legacy(
            PatchErrorCode::UnsupportedKey,
            "key:_originalFormatting paragraph property replay",
        ));
    }
    let old = tokens(&before.units);
    let new = tokens(&after.units);
    let old_inherit = inherited_runs(&context, &before.properties, &old);
    let new_inherit = inherited_runs(&context, &after.properties, &new);
    let retained = alignment(&old, &new)
        .map_err(|detail| context.legacy(PatchErrorCode::AmbiguousAlignment, detail))?;
    let donors = DonorIndex::new(&old);
    let offsets = &donors.offsets;
    let mut text_edits = Vec::new();
    let mut run_edits: Vec<RunEditV1> = Vec::new();
    let mut replay = Vec::new();
    let mut previous_run: Option<(usize, usize)> = None;
    let mut run_changes = Vec::new();
    let mut previous = (0, 0);
    let end = (old.len(), new.len());
    for &(i, j) in retained.iter().chain(std::iter::once(&end)) {
        if i != previous.0 || j != previous.1 {
            let removed = &old[previous.0..i];
            let inserted = &new[previous.1..j];
            if removed.iter().chain(inserted).any(|token| !token.is_text()) {
                return Err(context.legacy(
                    PatchErrorCode::UnsupportedStructure,
                    "marker: edit touches an atom",
                ));
            }
            let donor = donors.donor(previous.0, i, &context)?;
            let mark_attrs = mark_attrs(&before);
            let donor_attrs = donors
                .text_donor(previous.0, i)
                .map_or(&mark_attrs, |at| old[at].attrs.as_ref());
            if removed
                .iter()
                .any(|token| token.attrs.get("hyperlink") != donor_attrs.get("hyperlink"))
            {
                return Err(context.legacy(
                    PatchErrorCode::UnsupportedStructure,
                    "hyperlink: edit crosses wrappers",
                ));
            }
            if i == previous.0
                && i > 0
                && i < old.len()
                && old[i - 1].attrs.get("hyperlink") != old[i].attrs.get("hyperlink")
            {
                return Err(context.legacy(
                    PatchErrorCode::AmbiguousDonor,
                    "hyperlink: insertion at wrapper edge",
                ));
            }
            let mut insert: Vec<InsertedTextV1> = Vec::new();
            let mut previous_insert: Option<usize> = None;
            for (at, token) in inserted.iter().enumerate() {
                let inherit = &new_inherit[previous.1 + at];
                let same_run = previous_insert
                    .is_some_and(|previous| Rc::ptr_eq(&inserted[previous].attrs, &token.attrs));
                if !same_run {
                    let donor_inherit = donors
                        .text_donor(previous.0, i)
                        .map_or(&seeded_mark, |at| old_inherit[at].as_ref());
                    if token.attrs.get("runStyle") != donor_attrs.get("runStyle")
                        || inherit.as_ref() != donor_inherit
                    {
                        return Err(context.restyle("inserted run inheritance"));
                    }
                    let changes = deltas(donor_attrs, &token.attrs);
                    validate_theme_fonts(donor_attrs, &changes, &context)?;
                    validate_deltas(&changes, run_key, &context)?;
                    if token.attrs.get("hyperlink") != donor_attrs.get("hyperlink") {
                        let code = PatchErrorCode::UnsupportedStructure;
                        return Err(context.legacy(code, "hyperlink: mutation"));
                    }
                }
                let Content::Text(value) = token.content else {
                    unreachable!()
                };
                if let Some(last) = insert.last_mut()
                    && (same_run || (last.attrs == *token.attrs && last.inherit == **inherit))
                {
                    last.text.push(value);
                } else {
                    insert.push(InsertedTextV1 {
                        text: value.to_string(),
                        attrs: token.attrs.as_ref().clone(),
                        inherit: inherit.as_ref().clone(),
                    });
                }
                previous_insert = Some(at);
            }
            text_edits.push(TextEditV1 {
                from: offsets[previous.0],
                to: offsets[i],
                insert,
                donor,
            });
            replay.extend_from_slice(inserted);
        }
        if i < old.len() {
            if old[i].attrs.get("runStyle") != new[j].attrs.get("runStyle")
                || !same_attrs(&old_inherit[i], &new_inherit[j])
            {
                return Err(context.restyle("run inheritance"));
            }
            if marker_atom(&old[i]) && !same_attrs(&old[i].attrs, &new[j].attrs) {
                return Err(context.legacy(
                    PatchErrorCode::UnsupportedStructure,
                    "marker: formatting touches a marker atom",
                ));
            }
            let same_run = previous_run.is_some_and(|(old_at, new_at)| {
                Rc::ptr_eq(&old[old_at].attrs, &old[i].attrs)
                    && Rc::ptr_eq(&new[new_at].attrs, &new[j].attrs)
                    && Rc::ptr_eq(&before.run_carriers[old_at], &before.run_carriers[i])
                    && Rc::ptr_eq(&after.run_carriers[new_at], &after.run_carriers[j])
                    && Rc::ptr_eq(&old_inherit[old_at], &old_inherit[i])
                    && Rc::ptr_eq(&new_inherit[new_at], &new_inherit[j])
            });
            if !same_run {
                for change in deltas(&before.run_carriers[i], &after.run_carriers[j]) {
                    let key = &change.path[0];
                    if old[i].attrs.get(key) == new[j].attrs.get(key)
                        && old[i].attrs.get("runStyle") == new[j].attrs.get("runStyle")
                        && before.properties.get("pStyle") == after.properties.get("pStyle")
                    {
                        return Err(context.legacy(
                            PatchErrorCode::UnsupportedKey,
                            format!("key:{key}.inheritedHyperlink"),
                        ));
                    }
                }
                run_changes = deltas(&old[i].attrs, &new[j].attrs);
                validate_theme_fonts(&old[i].attrs, &run_changes, &context)?;
                validate_deltas(&run_changes, run_key, &context)?;
                if !verify_changes(&old[i].attrs, &run_changes, &new[j].attrs) {
                    return Err(context.legacy(
                        PatchErrorCode::UnsupportedKey,
                        "key:runStyle run property replay",
                    ));
                }
            }
            let inherit = &new_inherit[j];
            if !run_changes.is_empty() && !marker_atom(&old[i]) {
                if let Some(last) = run_edits.last_mut()
                    && last.to == offsets[i]
                    && (same_run
                        || (last.deltas == run_changes
                            && last.target == *new[j].attrs
                            && last.inherit == **inherit))
                {
                    last.to = offsets[i + 1];
                } else {
                    run_edits.push(RunEditV1 {
                        from: offsets[i],
                        to: offsets[i + 1],
                        deltas: run_changes.clone(),
                        target: new[j].attrs.as_ref().clone(),
                        inherit: inherit.as_ref().clone(),
                    });
                }
            }
            previous_run = Some((i, j));
            replay.push(Token {
                content: old[i].content.clone(),
                attrs: new[j].attrs.clone(),
            });
        }
        previous = (i + 1, j + 1);
    }
    if replay != new {
        return Err(context.legacy(
            PatchErrorCode::AmbiguousAlignment,
            "ambiguous: content replay differs",
        ));
    }
    validate_markers(&before, &after, &text_edits, &context)?;
    let inherit = InheritanceV1 {
        paragraph: seeded_properties,
        runs: before
            .units
            .iter()
            .map(|unit| {
                let attrs = match unit {
                    UnitV1::Text { attrs, .. } | UnitV1::Atom { attrs, .. } => attrs,
                };
                run_inheritance(&context, &before.properties, attrs)
            })
            .collect(),
        paragraph_mark: seeded_mark,
    };
    let seed = SeedV1 {
        inherit,
        unit_sha256: unit_sha256(&before.units),
        properties: verification_properties(&before, context.seeded_markers),
        units: before.units,
    };
    Ok(ParagraphPatchV1 {
        ordinal: context.ordinal,
        session_key: context.session_key.to_owned(),
        seed,
        text_edits,
        run_edits,
        para_deltas,
        para_inherit,
        mark_inherit,
        result_sha256: result_sha256(
            &units(&replay),
            &verification_properties(&after, context.live_markers),
        ),
    })
}

fn verification_properties(snapshot: &Snapshot, markers: &[(u32, String, bool)]) -> Attrs {
    let mut properties = snapshot.properties.clone();
    if let Some(formatting) = snapshot
        .carriers
        .get("_originalFormatting")
        .and_then(|value| value.get("runProperties"))
        .filter(|value| !value.is_null())
    {
        properties.insert("defaultTextFormatting".to_owned(), formatting.clone());
    }
    if let Some(bookmarks) = snapshot.carriers.get("bookmarks") {
        properties.insert("bookmarks".to_owned(), bookmarks.clone());
    }
    if !markers.is_empty() {
        properties.insert("markers".to_owned(), json!(markers));
    }
    properties
}

fn marker_atom(token: &Token) -> bool {
    matches!(
        &token.content,
        Content::Atom(kind, payload)
            if kind == "field"
                && payload.get("modelKind").and_then(Value::as_str) == Some("commentReference")
    )
}

fn shifted_marker(at: u32, edits: &[TextEditV1]) -> Option<u32> {
    let mut shifted = i64::from(at);
    for edit in edits {
        if edit.from <= at && at <= edit.to {
            return None;
        }
        if edit.to < at {
            let inserted: usize = edit
                .insert
                .iter()
                .map(|run| run.text.encode_utf16().count())
                .sum();
            shifted += inserted as i64 - i64::from(edit.to - edit.from);
        }
    }
    u32::try_from(shifted).ok()
}

fn validate_markers(
    before: &Snapshot,
    after: &Snapshot,
    edits: &[TextEditV1],
    context: &PatchContext<'_>,
) -> Result<(), LegacyParagraphV1> {
    let failure = || {
        context.legacy(
            PatchErrorCode::UnsupportedStructure,
            "marker: mutation or touch",
        )
    };
    let moved: Option<Vec<_>> = context
        .seeded_markers
        .iter()
        .map(|(at, key, end)| Some((shifted_marker(*at, edits)?, key.clone(), *end)))
        .collect();
    if moved.as_deref() != Some(context.live_markers) {
        return Err(failure());
    }
    let mut offset = 0;
    for token in tokens(&before.units) {
        let end = offset + token.width();
        if marker_atom(&token)
            && edits
                .iter()
                .any(|edit| edit.from <= end && offset <= edit.to)
        {
            return Err(failure());
        }
        offset = end;
    }
    let mut bookmarks = before.carriers.get("bookmarks").cloned();
    if let Some(Value::Array(markers)) = &mut bookmarks {
        for marker in markers {
            let at = marker
                .get("offset")
                .and_then(Value::as_u64)
                .and_then(|at| u32::try_from(at).ok())
                .ok_or_else(failure)?;
            let moved = shifted_marker(at, edits).ok_or_else(failure)?;
            marker["offset"] = json!(moved);
        }
    }
    if bookmarks.as_ref() != after.carriers.get("bookmarks") {
        return Err(failure());
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn paragraph(text: &str, properties: Value) -> Vec<StorySegment> {
        let properties = properties
            .as_object()
            .unwrap()
            .iter()
            .map(|(key, value)| {
                let value = crate::seed::any_from_value(value.clone()).unwrap();
                (key.clone(), value)
            })
            .collect();
        vec![
            StorySegment {
                content: SegmentContent::Text(text.to_owned()),
                attributes: BTreeMap::new(),
            },
            StorySegment {
                content: SegmentContent::Pilcrow(crate::ParagraphProperties {
                    para_id: "p1".to_owned(),
                    values: properties,
                }),
                attributes: BTreeMap::new(),
            },
        ]
    }

    fn context() -> PatchContext<'static> {
        PatchContext {
            ordinal: 0,
            session_key: "p1",
            story: "body",
            metadata: None,
            seeded_markers: &[],
            live_markers: &[],
        }
    }

    fn bookmarks(start: u32, end: u32) -> Value {
        json!({ "bookmarks": [
            { "id": 1, "name": "mark", "kind": "start", "offset": start },
            { "id": 1, "kind": "end", "offset": end }
        ] })
    }

    #[test]
    fn cached_bookmarks_must_match_the_script_transform() {
        let before = paragraph("abcd", bookmarks(1, 3));
        let after = paragraph("Xabcd", bookmarks(1, 3));
        let legacy = diff_paragraph_v1(&before, &after, context()).unwrap_err();
        assert_eq!(legacy.code, PatchErrorCode::UnsupportedStructure);
        assert!(legacy.detail.starts_with("marker:"));
    }

    #[test]
    fn result_hash_covers_transformed_bookmarks() {
        let before = paragraph("abcd", bookmarks(1, 3));
        let after = paragraph("Xabcd", bookmarks(2, 4));
        let patch = diff_paragraph_v1(&before, &after, context()).unwrap();
        let target = vec![UnitV1::Text {
            text: "Xabcd".to_owned(),
            attrs: Attrs::new(),
        }];
        let properties = serde_json::from_value(bookmarks(2, 4)).unwrap();
        let expected_hash = result_sha256(&target, &properties);
        assert_eq!(patch.result_sha256, expected_hash);
        assert_eq!(
            patch.seed.properties,
            serde_json::from_value::<Attrs>(bookmarks(1, 3)).unwrap()
        );
    }

    #[test]
    fn result_hash_covers_live_comment_marker_positions() {
        let before = paragraph("abcd", json!({}));
        let after = paragraph("Xabcd", json!({}));
        let seeded = [
            (1, "comment:1".to_owned(), false),
            (3, "comment:1".to_owned(), true),
        ];
        let live = [
            (2, "comment:1".to_owned(), false),
            (4, "comment:1".to_owned(), true),
        ];
        let patch = diff_paragraph_v1(
            &before,
            &after,
            PatchContext {
                seeded_markers: &seeded,
                live_markers: &live,
                ..context()
            },
        )
        .unwrap();
        let target = vec![UnitV1::Text {
            text: "Xabcd".to_owned(),
            attrs: Attrs::new(),
        }];
        let properties = Attrs::from([("markers".to_owned(), json!(live))]);
        let expected_hash = result_sha256(&target, &properties);
        assert_eq!(patch.result_sha256, expected_hash);
    }

    #[test]
    fn empty_object_deltas_check_leaves_in_both_slots() {
        for (key, leaf, paragraph) in [
            ("fontFamily", "hint", false),
            ("fontSize", "unknown", false),
            ("underline", "unknown", false),
            ("alignment", "unknown", true),
        ] {
            let accepts: fn(&[String]) -> bool = if paragraph { para_key } else { run_key };
            let empty = Attrs::from([(key.to_owned(), json!({}))]);
            let unknown = Attrs::from([(key.to_owned(), json!({ (leaf): "value" }))]);
            for (before, after) in [(&empty, &unknown), (&unknown, &empty)] {
                let changes = deltas(before, after);
                let legacy = validate_deltas(&changes, accepts, &context()).unwrap_err();
                assert_eq!(legacy.code, PatchErrorCode::UnsupportedKey);
                assert_eq!(legacy.detail, format!("key:{key}.{leaf}"));
            }
        }
    }

    #[test]
    fn inserted_and_split_run_styles_request_restyle_replace() {
        for middle in ["b", "X"] {
            let before = paragraph("abc", json!({}));
            let mut after = paragraph("a", json!({}));
            let mark = after.pop().unwrap();
            after.push(StorySegment {
                content: SegmentContent::Text(middle.to_owned()),
                attributes: BTreeMap::from([(
                    "runStyle".to_owned(),
                    crate::seed::any_from_value(json!({"styleId": "C"})).unwrap(),
                )]),
            });
            after.push(StorySegment {
                content: SegmentContent::Text(if middle == "b" { "c" } else { "bc" }.to_owned()),
                attributes: BTreeMap::new(),
            });
            after.push(mark);
            let legacy = diff_paragraph_v1(&before, &after, context()).unwrap_err();
            assert_eq!(legacy.code, PatchErrorCode::UnsupportedStructure);
            assert!(legacy.detail.starts_with("restyle:"));
        }
    }

    #[test]
    fn scalar_tokens_share_their_runs_formatting() {
        use std::borrow::Borrow;

        let formatting = Attrs::from([
            (
                "fontFamily".to_owned(),
                json!({ "ascii": "Serif", "hAnsi": "Serif" }),
            ),
            ("textColor".to_owned(), json!({ "rgb": "123456" })),
        ]);
        let source = vec![UnitV1::Text {
            text: "a".repeat(50_000),
            attrs: formatting.clone(),
        }];
        let tokens = tokens(&source);
        let first: &Attrs = tokens[0].attrs.borrow();
        let last: &Attrs = tokens[tokens.len() - 1].attrs.borrow();
        assert!(std::ptr::eq(first, last));
        let replay = tokens.clone();
        let replay_attrs: &Attrs = replay[0].attrs.borrow();
        assert!(std::ptr::eq(first, replay_attrs));
        assert_eq!(units(&replay), source);

        let mut before = paragraph("abcd", json!({}));
        before[0].attributes = formatting
            .into_iter()
            .map(|(key, value)| (key, crate::seed::any_from_value(value).unwrap()))
            .collect();
        let before = snapshot(&before).unwrap();
        let tokens = super::tokens(&before.units);
        let inherited = inherited_runs(&context(), &before.properties, &tokens);
        let first: &Attrs = inherited[0].borrow();
        let last: &Attrs = inherited[inherited.len() - 1].borrow();
        assert!(std::ptr::eq(first, last));
    }

    #[test]
    fn donor_index_precomputes_utf16_offsets_and_text_neighbors() {
        let units = vec![
            UnitV1::Atom {
                atom_kind: "noteRef".to_owned(),
                payload: Attrs::new(),
                attrs: Attrs::new(),
            },
            UnitV1::Text {
                text: "😀x".to_owned(),
                attrs: Attrs::new(),
            },
            UnitV1::Atom {
                atom_kind: "break".to_owned(),
                payload: Attrs::new(),
                attrs: Attrs::new(),
            },
        ];
        let tokens = tokens(&units);
        let index = DonorIndex::new(&tokens);
        assert_eq!(index.offsets, [0, 1, 3, 4, 5]);
        assert_eq!(index.previous, [None, None, Some(1), Some(2), Some(2)]);
        assert_eq!(index.next, [Some(1), Some(1), Some(2), None, None]);
        for (from, to, at) in [(0, 0, 1), (1, 3, 1), (2, 2, 1), (4, 4, 3)] {
            assert_eq!(
                index.donor(from, to, &context()).unwrap(),
                DonorV1::At { at }
            );
        }
    }

    #[test]
    fn alignment_limits_only_the_changed_middle() {
        let text = "a".repeat(50_000);
        let before = paragraph(&text, json!({}));
        let after = paragraph(&format!("{text}X"), json!({}));
        let patch = diff_paragraph_v1(&before, &after, context()).unwrap();
        assert_eq!(patch.text_edits.len(), 1);
        let edit = &patch.text_edits[0];
        assert_eq!((edit.from, edit.to), (50_000, 50_000));
        assert_eq!(patch.text_edits[0].donor, DonorV1::At { at: 49_999 });
    }

    #[test]
    fn oversized_alignment_middle_is_legacy() {
        let before = paragraph(&"a".repeat(2_001), json!({}));
        let after = paragraph(&"b".repeat(2_001), json!({}));
        let legacy = diff_paragraph_v1(&before, &after, context()).unwrap_err();
        assert_eq!(legacy.code, PatchErrorCode::AmbiguousAlignment);
        assert_eq!(legacy.detail, "ambiguous: alignment too large");
    }

    #[test]
    fn trimmed_alignment_still_refuses_repeated_text_ambiguity() {
        for (before, after) in [("aaaa", "aaa"), ("abc", "aabc"), ("abca", "abcca")] {
            let before = paragraph(before, json!({}));
            let after = paragraph(after, json!({}));
            let legacy = diff_paragraph_v1(&before, &after, context()).unwrap_err();
            assert_eq!(legacy.code, PatchErrorCode::AmbiguousAlignment);
        }
    }
}
