mod lower;
mod properties;
mod text;
mod verify;

#[cfg(test)]
mod tests;

#[cfg(test)]
mod work {
    use std::cell::Cell;

    thread_local! {
        static WORK: Cell<u64> = const { Cell::new(0) };
    }

    pub(super) fn add(amount: usize) {
        WORK.with(|work| work.set(work.get() + amount as u64));
    }

    pub(super) fn reset() {
        WORK.with(|work| work.set(0));
    }

    pub(super) fn get() -> u64 {
        WORK.with(Cell::get)
    }
}

use std::collections::{BTreeMap, BTreeSet};
use std::ops::Range;

use serde_json::Value;

use crate::paragraph_identity::tags;
use crate::paragraph_patch::{
    AttrDeltaV1, Attrs, ParagraphPatchError, ParagraphPatchV1, PatchErrorCode, SlotV1, UnitV1,
    result_sha256, unit_sha256,
};

pub(crate) struct PatchContext {
    pub part: String,
    pub namespaces: BTreeMap<String, String>,
    pub in_sdt: bool,
}

impl PatchContext {
    pub(crate) fn at(part: &str, source: &str, offset: usize) -> Self {
        let mut scopes = vec![BTreeMap::new()];
        let mut ancestors = Vec::new();
        if let Some(tags) = tags(source) {
            for tag in tags {
                if tag.range.start >= offset {
                    break;
                }
                if tag.end {
                    scopes.pop();
                    ancestors.pop();
                } else {
                    let mut scope = scopes.last().cloned().unwrap_or_default();
                    for (name, range) in tag.attributes {
                        if let Some(prefix) = name.strip_prefix("xmlns:")
                            && let Ok(value) = lower::decode(&source[range])
                        {
                            scope.insert(prefix.to_owned(), value);
                        }
                    }
                    if !tag.empty {
                        scopes.push(scope);
                        ancestors.push(tag.name.to_owned());
                    }
                }
            }
        }
        Self {
            part: part.to_owned(),
            namespaces: scopes.pop().unwrap_or_default(),
            in_sdt: ancestors
                .iter()
                .any(|name| name == "w:sdt" || name == "w:sdtContent"),
        }
    }

    #[cfg(test)]
    fn word_document() -> Self {
        Self {
            part: "word/document.xml".to_owned(),
            namespaces: BTreeMap::from([("w".to_owned(), lower::WORD.to_owned())]),
            in_sdt: false,
        }
    }
}

pub(super) struct Failure {
    code: PatchErrorCode,
    detail: String,
}

impl Failure {
    fn new(code: PatchErrorCode, detail: impl Into<String>) -> Self {
        let mut detail = detail.into();
        if code.is_unsupported()
            && ![
                "restyle:",
                "field:",
                "tc:",
                "drawing:",
                "sdt:",
                "math:",
                "key:",
                "hyperlink:",
                "marker:",
                "ambiguous:",
            ]
            .iter()
            .any(|prefix| detail.starts_with(prefix))
        {
            detail = format!("ambiguous: {detail}");
        }
        Self { code, detail }
    }

    fn structure(detail: impl Into<String>) -> Self {
        Self::new(PatchErrorCode::UnsupportedStructure, detail)
    }
}

pub(crate) fn patch_paragraph_v1(
    source_paragraph: &str,
    patch: &ParagraphPatchV1,
    context: &PatchContext,
) -> Result<String, ParagraphPatchError> {
    patch_inner(source_paragraph, patch, context).map_err(|failure| ParagraphPatchError {
        part: context.part.clone(),
        ordinal: Some(patch.ordinal),
        code: failure.code,
        detail: failure.detail,
    })
}

fn patch_inner(
    source: &str,
    patch: &ParagraphPatchV1,
    context: &PatchContext,
) -> Result<String, Failure> {
    let mut lowered = lower::lower(source, context)?;
    let source_units: Vec<_> = lowered.cells.iter().map(|cell| cell.unit.clone()).collect();
    if !same_content(&source_units, &patch.seed.units)
        || unit_sha256(&patch.seed.units) != patch.seed.unit_sha256
    {
        return Err(Failure::new(
            PatchErrorCode::SeedMismatch,
            "source units differ from seed",
        ));
    }
    verify::markers(
        source,
        &lowered,
        &patch.seed.properties,
        PatchErrorCode::SeedMismatch,
    )?;
    unchanged_inheritance(patch)?;
    verify::seed(source, &lowered, &patch.seed.units, &patch.seed.properties)?;
    for (cell, unit) in lowered.cells.iter_mut().zip(flatten(&patch.seed.units)) {
        cell.unit = unit;
    }
    let (mut output, units) = text::write(source, &lowered, patch)?;
    let mut properties = patch.seed.properties.clone();
    apply_deltas(&mut properties, &patch.para_deltas)?;
    shift_markers(&mut properties, &patch.text_edits)?;
    if !patch.para_deltas.is_empty() {
        let xml = lower::Xml::parse(&output, &context.namespaces)?;
        let ppr = xml.child(0, "pPr")?;
        let original = ppr.map_or("", |at| &output[xml.nodes[at].span.clone()]);
        let replacement = properties::paragraph_properties(
            original,
            &patch.para_deltas,
            &patch.para_inherit,
            &properties,
        )?;
        output = if let Some(at) = ppr {
            apply_bytes(&output, vec![(xml.nodes[at].span.clone(), replacement)])?
        } else {
            insert_into_root(&output, &xml.nodes[0], &replacement, true)?
        };
    }
    if result_sha256(&units, &properties) != patch.result_sha256 {
        return Err(Failure::new(
            PatchErrorCode::ResultMismatch,
            "result digest differs from plan",
        ));
    }
    let verified = lower::lower(&output, context)?;
    verify::markers(
        &output,
        &verified,
        &properties,
        PatchErrorCode::ResultMismatch,
    )?;
    let actual: Vec<_> = verified.cells.into_iter().map(|cell| cell.unit).collect();
    if !same_content(&actual, &units) {
        return Err(Failure::new(
            PatchErrorCode::ResultMismatch,
            "written units differ from result",
        ));
    }
    Ok(output)
}

fn unchanged_inheritance(patch: &ParagraphPatchV1) -> Result<(), Failure> {
    let restyle = || Failure::structure("restyle: style or inheritance changed");
    let style_delta = |deltas: &[AttrDeltaV1], keys: &[&str]| {
        deltas.iter().any(|delta| {
            delta
                .path
                .first()
                .is_some_and(|key| keys.contains(&key.as_str()))
        })
    };
    if style_delta(&patch.para_deltas, &["pStyle", "styleId", "numPr"])
        || patch.para_inherit != patch.seed.inherit.paragraph
        || patch.mark_inherit != patch.seed.inherit.paragraph_mark
        || (!patch.seed.inherit.runs.is_empty()
            && patch.seed.inherit.runs.len() != patch.seed.units.len())
    {
        return Err(restyle());
    }
    let empty = Attrs::new();
    let mut offset = 0;
    let runs: Vec<_> = patch
        .seed
        .units
        .iter()
        .enumerate()
        .map(|(at, unit)| {
            let from = offset;
            offset += match unit {
                UnitV1::Text { text, .. } => text.encode_utf16().count() as u32,
                UnitV1::Atom { .. } => 1,
            };
            let inherit = patch.seed.inherit.runs.get(at).unwrap_or(&empty);
            (from, offset, unit_attrs(unit), inherit)
        })
        .collect();
    for edit in &patch.run_edits {
        if style_delta(&edit.deltas, &["runStyle", "styleId"])
            || runs.iter().any(|(from, to, attrs, inherit)| {
                *from < edit.to
                    && edit.from < *to
                    && (edit.inherit != **inherit
                        || edit.target.get("runStyle") != attrs.get("runStyle"))
            })
        {
            return Err(restyle());
        }
    }
    let mut mark_attrs = Attrs::new();
    if let Some(style) = patch
        .seed
        .properties
        .get("defaultTextFormatting")
        .or_else(|| patch.seed.inherit.paragraph.get("defaultTextFormatting"))
        .and_then(|formatting| formatting.get("styleId"))
    {
        mark_attrs.insert(
            "runStyle".to_owned(),
            serde_json::json!({ "styleId": style }),
        );
    }
    for edit in &patch.text_edits {
        let (attrs, inherit) = match edit.donor {
            crate::paragraph_patch::DonorV1::At { at } => runs
                .iter()
                .find(|(from, to, _, _)| *from <= at && at < *to)
                .map_or((&empty, &empty), |(_, _, attrs, inherit)| {
                    (*attrs, *inherit)
                }),
            crate::paragraph_patch::DonorV1::Pilcrow => {
                (&mark_attrs, &patch.seed.inherit.paragraph_mark)
            }
        };
        if edit.insert.iter().any(|insert| {
            insert.inherit != *inherit || insert.attrs.get("runStyle") != attrs.get("runStyle")
        }) {
            return Err(restyle());
        }
    }
    Ok(())
}

fn shift_markers(
    properties: &mut Attrs,
    edits: &[crate::paragraph_patch::TextEditV1],
) -> Result<(), Failure> {
    for key in ["bookmarks", "markers"] {
        let Some(markers) = properties.get_mut(key) else {
            continue;
        };
        let markers = markers
            .as_array_mut()
            .ok_or_else(|| Failure::new(PatchErrorCode::InvalidPlan, format!("invalid {key}")))?;
        for marker in markers {
            let offset = if key == "bookmarks" {
                marker.get_mut("offset")
            } else {
                marker.get_mut(0)
            }
            .ok_or_else(|| Failure::new(PatchErrorCode::InvalidPlan, "missing marker offset"))?;
            let at = offset
                .as_u64()
                .and_then(|at| u32::try_from(at).ok())
                .ok_or_else(|| {
                    Failure::new(PatchErrorCode::InvalidPlan, "invalid marker offset")
                })?;
            let mut shifted = i64::from(at);
            for edit in edits {
                if edit.from <= at && at <= edit.to {
                    return Err(Failure::structure("marker: text edit touches a marker"));
                }
                if edit.to < at {
                    let inserted = edit
                        .insert
                        .iter()
                        .map(|run| run.text.encode_utf16().count())
                        .sum::<usize>();
                    let inserted = i64::try_from(inserted).map_err(|_| {
                        Failure::new(PatchErrorCode::InvalidPlan, "marker offset overflow")
                    })?;
                    shifted = shifted
                        .checked_add(inserted - i64::from(edit.to - edit.from))
                        .ok_or_else(|| {
                            Failure::new(PatchErrorCode::InvalidPlan, "marker offset overflow")
                        })?;
                }
            }
            *offset = serde_json::json!(u32::try_from(shifted).map_err(|_| {
                Failure::new(PatchErrorCode::InvalidPlan, "marker offset overflow")
            })?);
        }
    }
    Ok(())
}

fn same_content(left: &[UnitV1], right: &[UnitV1]) -> bool {
    let left = flatten(left);
    let right = flatten(right);
    left.len() == right.len()
        && left
            .iter()
            .zip(&right)
            .all(|(left, right)| match (left, right) {
                (UnitV1::Text { text: left, .. }, UnitV1::Text { text: right, .. }) => {
                    left == right
                }
                (
                    UnitV1::Atom {
                        atom_kind: left_kind,
                        payload: left,
                        ..
                    },
                    UnitV1::Atom {
                        atom_kind: right_kind,
                        payload: right,
                        ..
                    },
                ) => {
                    left_kind == right_kind
                        && left.len() == right.len()
                        && left.iter().all(|(key, value)| {
                            right.get(key).is_some_and(|right| {
                                value == right
                                    || (value.is_number()
                                        && right.is_number()
                                        && value.as_f64() == right.as_f64())
                            })
                        })
                }
                _ => false,
            })
}

fn flatten(units: &[UnitV1]) -> Vec<UnitV1> {
    let mut result = Vec::new();
    for unit in units {
        match unit {
            UnitV1::Text { text, attrs } => {
                result.extend(text.chars().map(|scalar| UnitV1::Text {
                    text: scalar.to_string(),
                    attrs: attrs.clone(),
                }));
            }
            _ => result.push(unit.clone()),
        }
    }
    result
}

fn unit_attrs(unit: &UnitV1) -> &Attrs {
    match unit {
        UnitV1::Text { attrs, .. } | UnitV1::Atom { attrs, .. } => attrs,
    }
}

fn unit_attrs_mut(unit: &mut UnitV1) -> &mut Attrs {
    match unit {
        UnitV1::Text { attrs, .. } | UnitV1::Atom { attrs, .. } => attrs,
    }
}

fn apply_deltas(attrs: &mut Attrs, deltas: &[AttrDeltaV1]) -> Result<(), Failure> {
    let mut seen: Vec<&[String]> = Vec::new();
    for delta in deltas {
        #[cfg(test)]
        work::add(1);
        if delta.path.is_empty()
            || seen
                .iter()
                .any(|path| path.starts_with(&delta.path) || delta.path.starts_with(path))
        {
            return Err(Failure::new(
                PatchErrorCode::PatchConflict,
                "overlapping property paths",
            ));
        }
        seen.push(&delta.path);
        if path_value(attrs, &delta.path) != properties::slot(&delta.before) {
            return Err(Failure::new(
                PatchErrorCode::PropertyConflict,
                format!("stale property {}", delta.path.join(".")),
            ));
        }
        let first = &delta.path[0];
        if delta.path.len() == 1 {
            match &delta.after {
                SlotV1::Missing => {
                    attrs.remove(first);
                }
                SlotV1::Present { value } => {
                    attrs.insert(first.clone(), value.clone());
                }
            }
        } else {
            let mut value = attrs
                .entry(first.clone())
                .or_insert_with(|| serde_json::json!({}));
            for key in &delta.path[1..delta.path.len() - 1] {
                value = object_mut(value)?
                    .entry(key.clone())
                    .or_insert_with(|| serde_json::json!({}));
            }
            let object = object_mut(value)?;
            let key = delta.path.last().unwrap();
            match &delta.after {
                SlotV1::Missing => {
                    object.remove(key);
                }
                SlotV1::Present { value } => {
                    object.insert(key.clone(), value.clone());
                }
            }
        }
    }
    Ok(())
}

fn object_mut(value: &mut Value) -> Result<&mut serde_json::Map<String, Value>, Failure> {
    value
        .as_object_mut()
        .ok_or_else(|| Failure::new(PatchErrorCode::PropertyConflict, "nonobject property path"))
}

fn path_value<'a>(attrs: &'a Attrs, path: &[String]) -> Option<&'a Value> {
    let mut value = attrs.get(path.first()?)?;
    for key in &path[1..] {
        value = value.get(key)?;
    }
    Some(value)
}

fn attr_deltas(before: &Attrs, after: &Attrs) -> Vec<AttrDeltaV1> {
    let keys: BTreeSet<_> = before.keys().chain(after.keys()).collect();
    let slot = |value: Option<&Value>| match value {
        Some(value) => SlotV1::Present {
            value: value.clone(),
        },
        None => SlotV1::Missing,
    };
    keys.into_iter()
        .filter(|key| before.get(*key) != after.get(*key))
        .map(|key| AttrDeltaV1 {
            path: vec![key.clone()],
            before: slot(before.get(key)),
            after: slot(after.get(key)),
        })
        .collect()
}

fn apply_bytes(source: &str, mut edits: Vec<(Range<usize>, String)>) -> Result<String, Failure> {
    edits.sort_by_key(|(span, _)| (span.start, span.end));
    let mut output = String::new();
    let mut cursor = 0;
    for (span, replacement) in edits {
        #[cfg(test)]
        work::add(1);
        if span.start < cursor
            || span.start > span.end
            || span.end > source.len()
            || !source.is_char_boundary(span.start)
            || !source.is_char_boundary(span.end)
        {
            return Err(Failure::new(
                PatchErrorCode::PatchConflict,
                "overlapping byte edits",
            ));
        }
        #[cfg(test)]
        work::add(span.start - cursor + replacement.len());
        output.push_str(&source[cursor..span.start]);
        output.push_str(&replacement);
        cursor = span.end;
    }
    #[cfg(test)]
    work::add(source.len() - cursor);
    output.push_str(&source[cursor..]);
    Ok(output)
}

fn insert_into_root(
    source: &str,
    root: &lower::Node,
    fragment: &str,
    first: bool,
) -> Result<String, Failure> {
    if fragment.is_empty() {
        return Ok(source.to_owned());
    }
    if root.empty {
        let opener = source[root.open.clone()]
            .trim_end_matches('>')
            .trim_end_matches('/');
        Ok(format!("{opener}>{fragment}</{}>", root.name))
    } else {
        let at = if first {
            root.open.end
        } else {
            root.close.start
        };
        apply_bytes(source, vec![(at..at, fragment.to_owned())])
    }
}

fn valid_xml_text(text: &str) -> bool {
    text.chars().all(|scalar| {
        matches!(scalar, '\t' | '\n' | '\r' | '\u{20}'..='\u{d7ff}'
            | '\u{e000}'..='\u{fffd}' | '\u{10000}'..='\u{10ffff}')
    })
}
