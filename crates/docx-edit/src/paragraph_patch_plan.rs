use std::collections::{BTreeMap, BTreeSet};
use std::rc::Rc;

use docx_parse::paragraph_patch::{
    AttrDeltaV1, Attrs, DonorV1, InsertedTextV1, LegacyParagraphV1, ParagraphPatchV1,
    PatchErrorCode, RunEditV1, SeedV1, SlotV1, TextEditV1, UnitV1,
};
use serde_json::{Value, json};
use sha2::{Digest, Sha256};
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
    run_carriers: Vec<Attrs>,
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
                return Err(format!("non-JSON value at {key}"));
            }
            serde_json::to_value(value)
                .map(|value| (key.clone(), wire_value(value)))
                .map_err(|error| format!("{key}: {error}"))
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
    let mut run_carriers = Vec::new();
    for segment in segments {
        let mut attributes = attrs(&segment.attributes)?;
        let mut derived = Attrs::new();
        for key in ["textColor", "underline"] {
            if let Some(value) = attributes.get_mut(key).and_then(Value::as_object_mut)
                && let Some(flag) = value.remove("inheritedHyperlink")
            {
                if !flag.is_boolean() {
                    return Err(format!("{key}.inheritedHyperlink"));
                }
                derived.insert(key.to_owned(), flag);
            }
        }
        let count = match &segment.content {
            SegmentContent::Text(text) => text.chars().count(),
            SegmentContent::OtherEmbed { .. } => 1,
            SegmentContent::Pilcrow(_) => 0,
        };
        run_carriers.extend(std::iter::repeat_n(derived, count));
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
                        .filter(|(key, _)| para_key(std::slice::from_ref(*key)))
                        .map(|(key, value)| (key.clone(), value.clone()))
                        .collect();
                    for (key, value) in carried {
                        properties.entry(key).or_insert(value);
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
        "runStyle" => Some(&["styleId"]),
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
        "pStyle" | "alignment" | "outlineLevel" | "spaceBefore" | "spaceAfter"
        | "spaceBeforeLines" | "spaceAfterLines" | "beforeAutospacing" | "afterAutospacing"
        | "lineSpacing" | "lineSpacingRule" | "indentLeft" | "indentRight" | "indentFirstLine"
        | "indentLeftChars" | "indentRightChars" | "indentFirstLineChars" | "hangingIndent"
        | "hangingIndentChars" | "keepNext" | "keepLines" | "widowControl" | "pageBreakBefore"
        | "contextualSpacing" => path.len() == 1,
        "numPr" => {
            path.len() == 1 || (path.len() == 2 && ["numId", "ilvl"].contains(&path[1].as_str()))
        }
        _ => false,
    }
}

fn deltas(before: &Attrs, after: &Attrs) -> Vec<AttrDeltaV1> {
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

#[derive(Clone, Debug, PartialEq)]
struct Token {
    content: Content,
    attrs: Attrs,
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
            UnitV1::Text { text, attrs } => result.extend(text.chars().map(|value| Token {
                content: Content::Text(value),
                attrs: attrs.clone(),
            })),
            UnitV1::Atom {
                atom_kind,
                payload,
                attrs,
            } => result.push(Token {
                content: Content::Atom(atom_kind.clone(), payload.clone()),
                attrs: attrs.clone(),
            }),
        }
    }
    result
}

fn units(tokens: &[Token]) -> Vec<UnitV1> {
    let mut result = Vec::new();
    for token in tokens {
        match &token.content {
            Content::Text(value) => push_text(&mut result, &value.to_string(), &token.attrs),
            Content::Atom(kind, payload) => result.push(UnitV1::Atom {
                atom_kind: kind.clone(),
                payload: payload.clone(),
                attrs: token.attrs.clone(),
            }),
        }
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

fn alignment(before: &[Token], after: &[Token]) -> Option<Vec<(usize, usize)>> {
    if before.len() == after.len()
        && before.iter().zip(after).all(|(old, new)| old.content == new.content)
    {
        return Some((0..before.len()).map(|at| (at, at)).collect());
    }
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

fn inherited_runs(context: &PatchContext<'_>, properties: &Attrs, tokens: &[Token]) -> Vec<Attrs> {
    let mut cache = BTreeMap::new();
    tokens
        .iter()
        .map(|token| {
            let key = token.attrs.get("runStyle").map(canonical_json_v1);
            cache
                .entry(key)
                .or_insert_with(|| {
                    let mut inherited = context.inheritance(properties, &token.attrs).1;
                    for key in ["textColor", "underline"] {
                        if let Some(value) = inherited.get_mut(key).and_then(Value::as_object_mut) {
                            value.remove("inheritedHyperlink");
                        }
                    }
                    inherited
                })
                .clone()
        })
        .collect()
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
                    return Some(format!("atom {atom_kind}"));
                }
                attrs
            }
        };
        if attributes
            .keys()
            .any(|key| matches!(key.as_str(), "ins" | "del" | "fieldResult" | "fieldCode"))
        {
            return Some("revision or field marks".to_owned());
        }
    }
    if ["pPrIns", "pPrDel", "pPrChange"]
        .iter()
        .any(|key| snapshot.properties.contains_key(*key))
    {
        return Some("paragraph revision".to_owned());
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
        return Some("run property revision".to_owned());
    }
    None
}

fn validate_deltas(
    changes: &[AttrDeltaV1],
    accepts: fn(&[String]) -> bool,
    context: &PatchContext<'_>,
) -> Result<(), LegacyParagraphV1> {
    for delta in changes {
        if !accepts(&delta.path) {
            let code = if delta.path.first().is_some_and(|key| key == "hyperlink") {
                PatchErrorCode::UnsupportedStructure
            } else {
                PatchErrorCode::UnsupportedKey
            };
            return Err(context.legacy(code, delta.path.join(".")));
        }
    }
    Ok(())
}

fn attr_at<'a>(attrs: &'a Attrs, path: &[String]) -> Option<&'a Value> {
    let value = attrs.get(&path[0])?;
    path[1..].iter().try_fold(value, |value, key| value.get(key))
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
                para_key(&path)
                    && match &change.after {
                        SlotV1::Missing => target.is_none() || target == attr_at(inherited, &path),
                        SlotV1::Present { value } => target == Some(value),
                    }
            }
            "defaultTextFormatting" => {
                before.properties.get("pStyle") != after.properties.get("pStyle")
                    && context
                        .inheritance(&after.properties, &Attrs::new())
                        .0
                        .get(key)
                        == after.carriers.get(key)
            }
            "bookmarks" => true,
            "spacingExplicit" => {
                before.properties.get("pStyle") != after.properties.get("pStyle")
                    || deltas(&before.properties, &after.properties).iter().any(|delta| {
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
            return Err(context.legacy(code, change.path.join(".")));
        }
    }
    Ok(())
}

fn donor(
    before: &[Token],
    from: usize,
    to: usize,
    context: &PatchContext<'_>,
) -> Result<DonorV1, LegacyParagraphV1> {
    let index = if from != to {
        (from..to).find(|&at| before[at].is_text())
    } else if from == 0 {
        before.iter().position(Token::is_text)
    } else {
        before[..from].iter().rposition(Token::is_text)
    };
    if let Some(index) = index {
        let offset = before[..index].iter().map(Token::width).sum();
        Ok(DonorV1::At { at: offset })
    } else if before.is_empty() {
        Ok(DonorV1::Pilcrow)
    } else {
        Err(context.legacy(PatchErrorCode::AmbiguousDonor, "no adjacent text donor"))
    }
}

pub(crate) fn diff_paragraph_v1(
    seeded: &[StorySegment],
    live: &[StorySegment],
    context: PatchContext<'_>,
) -> Result<ParagraphPatchV1, LegacyParagraphV1> {
    let before = snapshot(seeded)
        .map_err(|detail| context.legacy(PatchErrorCode::UnsupportedKey, detail))?;
    let after = snapshot(live)
        .map_err(|detail| context.legacy(PatchErrorCode::UnsupportedKey, detail))?;
    if let Some(detail) = unsupported(&before).or_else(|| unsupported(&after)) {
        return Err(context.legacy(PatchErrorCode::UnsupportedStructure, detail));
    }
    if context.metadata.is_some_and(|metadata| {
        metadata.patch_has_run_revisions(context.story, context.session_key)
    }) {
        return Err(context.legacy(PatchErrorCode::UnsupportedStructure, "run property revision"));
    }
    let (inherited_properties, _) = context.inheritance(&after.properties, &Attrs::new());
    validate_carriers(&before, &after, &inherited_properties, &context)?;
    let para_inherit = inherited_properties
        .into_iter()
        .filter(|(key, _)| para_key(std::slice::from_ref(key)))
        .collect();
    let para_deltas = deltas(&before.properties, &after.properties);
    validate_deltas(&para_deltas, para_key, &context)?;
    if !verify_changes(&before.properties, &para_deltas, &after.properties) {
        return Err(context.legacy(PatchErrorCode::UnsupportedKey, "paragraph property replay"));
    }
    let old = tokens(&before.units);
    let new = tokens(&after.units);
    let old_inherit = inherited_runs(&context, &before.properties, &old);
    let new_inherit = inherited_runs(&context, &after.properties, &new);
    let retained = alignment(&old, &new).ok_or_else(|| {
        context.legacy(PatchErrorCode::AmbiguousAlignment, "nonunique content alignment")
    })?;
    let mut offsets = vec![0];
    for token in &old {
        offsets.push(offsets.last().copied().unwrap_or_default() + token.width());
    }
    let mut text_edits = Vec::new();
    let mut run_edits: Vec<RunEditV1> = Vec::new();
    let mut replay = Vec::new();
    let mut previous = (0, 0);
    let end = (old.len(), new.len());
    for &(i, j) in retained.iter().chain(std::iter::once(&end)) {
        if i != previous.0 || j != previous.1 {
            let removed = &old[previous.0..i];
            let inserted = &new[previous.1..j];
            if removed.iter().chain(inserted).any(|token| !token.is_text()) {
                return Err(context.legacy(
                    PatchErrorCode::UnsupportedStructure,
                    "edit touches an atom",
                ));
            }
            let donor = donor(&old, previous.0, i, &context)?;
            let empty_attrs = Attrs::new();
            let donor_attrs = match donor {
                DonorV1::At { at } => {
                    &old[offsets.binary_search(&at).expect("donor boundary")].attrs
                }
                DonorV1::Pilcrow => &empty_attrs,
            };
            if removed
                .iter()
                .any(|token| token.attrs.get("hyperlink") != donor_attrs.get("hyperlink"))
            {
                return Err(context.legacy(
                    PatchErrorCode::UnsupportedStructure,
                    "edit crosses hyperlinks",
                ));
            }
            if i == previous.0
                && i > 0
                && i < old.len()
                && old[i - 1].attrs.get("hyperlink") != old[i].attrs.get("hyperlink")
            {
                return Err(context.legacy(
                    PatchErrorCode::AmbiguousDonor,
                    "insertion at hyperlink edge",
                ));
            }
            let mut insert: Vec<InsertedTextV1> = Vec::new();
            for (token, inherit) in inserted.iter().zip(&new_inherit[previous.1..j]) {
                validate_deltas(&deltas(donor_attrs, &token.attrs), run_key, &context)?;
                if token.attrs.get("hyperlink") != donor_attrs.get("hyperlink") {
                    return Err(context.legacy(
                        PatchErrorCode::UnsupportedStructure,
                        "hyperlink mutation",
                    ));
                }
                let Content::Text(value) = token.content else {
                    unreachable!()
                };
                if let Some(last) = insert.last_mut()
                    && last.attrs == token.attrs
                    && last.inherit == *inherit
                {
                    last.text.push(value);
                } else {
                    insert.push(InsertedTextV1 {
                        text: value.to_string(),
                        attrs: token.attrs.clone(),
                        inherit: inherit.clone(),
                    });
                }
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
            if marker_atom(&old[i]) && old[i].attrs != new[j].attrs {
                return Err(context.legacy(
                    PatchErrorCode::UnsupportedStructure,
                    "formatting touches a marker atom",
                ));
            }
            for change in deltas(&before.run_carriers[i], &after.run_carriers[j]) {
                let key = &change.path[0];
                if old[i].attrs.get(key) == new[j].attrs.get(key)
                    && old[i].attrs.get("runStyle") == new[j].attrs.get("runStyle")
                    && before.properties.get("pStyle") == after.properties.get("pStyle")
                {
                    return Err(context.legacy(
                        PatchErrorCode::UnsupportedKey,
                        format!("{key}.inheritedHyperlink"),
                    ));
                }
            }
            let changes = deltas(&old[i].attrs, &new[j].attrs);
            validate_deltas(&changes, run_key, &context)?;
            if !verify_changes(&old[i].attrs, &changes, &new[j].attrs) {
                return Err(context.legacy(PatchErrorCode::UnsupportedKey, "run property replay"));
            }
            let inherit = &new_inherit[j];
            if (!changes.is_empty() || *inherit != old_inherit[i]) && !marker_atom(&old[i]) {
                if let Some(last) = run_edits.last_mut()
                    && last.to == offsets[i]
                    && last.deltas == changes
                    && last.target == new[j].attrs
                    && last.inherit == *inherit
                {
                    last.to = offsets[i + 1];
                } else {
                    run_edits.push(RunEditV1 {
                        from: offsets[i],
                        to: offsets[i + 1],
                        deltas: changes,
                        target: new[j].attrs.clone(),
                        inherit: inherit.clone(),
                    });
                }
            }
            replay.push(Token {
                content: old[i].content.clone(),
                attrs: new[j].attrs.clone(),
            });
        }
        previous = (i + 1, j + 1);
    }
    if replay != new {
        return Err(context.legacy(PatchErrorCode::AmbiguousAlignment, "content replay differs"));
    }
    validate_markers(&before, &after, &text_edits, &context)?;
    let seed = SeedV1 {
        unit_sha256: unit_sha256(&before.units),
        units: before.units,
        properties: before.properties,
    };
    Ok(ParagraphPatchV1 {
        ordinal: context.ordinal,
        session_key: context.session_key.to_owned(),
        seed,
        text_edits,
        run_edits,
        para_deltas,
        para_inherit,
        result_sha256: result_sha256(&units(&replay), &after.properties),
    })
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
            "marker mutation or touch",
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
            && edits.iter().any(|edit| edit.from <= end && offset <= edit.to)
        {
            return Err(failure());
        }
        offset = end;
    }
    let mut bookmarks = before.carriers.get("bookmarks").cloned();
    let cached = bookmarks.as_ref() == after.carriers.get("bookmarks");
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
    if !cached && bookmarks.as_ref() != after.carriers.get("bookmarks") {
        return Err(failure());
    }
    Ok(())
}

fn canonical_json_v1(value: &Value) -> String {
    fn ordered(value: &Value) -> Value {
        match value {
            Value::Object(map) => {
                let sorted: BTreeMap<_, _> = map
                    .iter()
                    .map(|(key, value)| (key, ordered(value)))
                    .collect();
                Value::Object(
                    sorted
                        .into_iter()
                        .map(|(key, value)| (key.clone(), value))
                        .collect(),
                )
            }
            Value::Array(values) => Value::Array(values.iter().map(ordered).collect()),
            value => value.clone(),
        }
    }
    serde_json::to_string(&ordered(value)).expect("JSON value")
}

fn digest(value: &Value) -> String {
    format!("{:x}", Sha256::digest(canonical_json_v1(value).as_bytes()))
}

fn unit_sha256(units: &[UnitV1]) -> String {
    let bare: Vec<_> = units
        .iter()
        .map(|unit| match unit {
            UnitV1::Text { text, .. } => UnitV1::Text {
                text: text.clone(),
                attrs: Attrs::new(),
            },
            UnitV1::Atom {
                atom_kind, payload, ..
            } => UnitV1::Atom {
                atom_kind: atom_kind.clone(),
                payload: payload.clone(),
                attrs: Attrs::new(),
            },
        })
        .collect();
    let bare = self::units(&tokens(&bare));
    let values: Vec<_> = bare
        .iter()
        .map(|unit| match unit {
            UnitV1::Text { text, .. } => json!({ "kind": "text", "text": text }),
            UnitV1::Atom {
                atom_kind, payload, ..
            } => json!({ "kind": "atom", "atomKind": atom_kind, "payload": payload }),
        })
        .collect();
    digest(&json!(values))
}

fn result_sha256(units: &[UnitV1], properties: &Attrs) -> String {
    digest(&json!({ "units": self::units(&tokens(units)), "properties": properties }))
}
