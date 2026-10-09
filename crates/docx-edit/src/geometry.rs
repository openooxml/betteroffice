use std::collections::{BTreeMap, HashMap, HashSet};

use serde::Serialize;
use yrs::types::text::YChange;
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, Transact};

use crate::{
    EditingDoc, KIND_KEY, PARA_ID, STORIES, is_identity_key, is_pilcrow, map_string, story_ref,
};

#[derive(Debug, PartialEq, Serialize)]
#[serde(untagged)]
pub enum GeometryRead<T> {
    Value(T),
    Sentinel(GeometrySentinel),
}

#[derive(Debug, PartialEq, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum GeometrySentinel {
    Legacy,
    Fallback,
}

pub type GeometryPositionOutline = BTreeMap<String, GeometryStoryOutline>;

#[derive(Default, Debug, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GeometryStoryOutline {
    pub content_start: u64,
    pub size: u64,
    pub paragraphs: Vec<GeometryParagraphOutline>,
}

#[derive(Debug, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GeometryParagraphOutline {
    pub para_id: String,
    pub display_start: u64,
    pub length: u64,
    pub leading: u64,
}

#[derive(Debug, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct OwnedRevisionRange {
    pub revision_id: String,
    pub kind: String,
    pub story: String,
    pub range: GeometryRange,
}

#[derive(Debug, PartialEq, Serialize)]
pub struct GeometryRange {
    pub story: String,
    pub start: GeometryEndpoint,
    pub end: GeometryEndpoint,
}

#[derive(Debug, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct GeometryEndpoint {
    pub para_id: String,
    pub offset: usize,
}

impl EditingDoc {
    pub fn geometry_position_outline(&self, root: &str) -> GeometryRead<GeometryPositionOutline> {
        let txn = self.doc.transact();
        let mut outline = BTreeMap::new();
        match build_story(&txn, &mut outline, root, 0, 0) {
            Ok(_) => GeometryRead::Value(outline),
            Err(()) => GeometryRead::Sentinel(GeometrySentinel::Legacy),
        }
    }

    pub fn owned_revision_ranges(&self, ids: &[String]) -> GeometryRead<Vec<OwnedRevisionRange>> {
        let txn = self.doc.transact();
        match revision_ranges(&txn, ids) {
            Ok(ranges) => GeometryRead::Value(ranges),
            Err(sentinel) => GeometryRead::Sentinel(sentinel),
        }
    }
}

fn plain_value<T: ReadTxn>(map: &MapRef, txn: &T, key: &str) -> Option<Any> {
    match map.get(txn, key) {
        Some(Out::Any(value)) => Some(value),
        _ => None,
    }
}

fn child_story(value: Option<&Any>, default: String) -> Result<String, ()> {
    match value {
        None | Some(Any::Null | Any::Undefined) => Ok(default),
        Some(Any::String(story)) => Ok(story.to_string()),
        _ => Err(()),
    }
}

const MAX_POSITION: u64 = (1 << 53) - 1;

fn build_story<T: ReadTxn>(
    txn: &T,
    outline: &mut GeometryPositionOutline,
    story_id: &str,
    content_start: u64,
    depth: usize,
) -> Result<u64, ()> {
    if let Some(existing) = outline.get(story_id) {
        return Ok(existing.size);
    }
    if depth > 256 || content_start > MAX_POSITION {
        return Err(());
    }
    let story = story_ref(txn, story_id).map_err(|_| ())?;
    outline.insert(
        story_id.to_owned(),
        GeometryStoryOutline {
            content_start,
            ..Default::default()
        },
    );
    let mut cursor = 0;
    let mut inline_length = 0;
    let mut leading = 0;
    let mut paragraph_start = 0;
    let mut table_index = 0;
    let mut paragraphs = Vec::new();
    let mut seen = HashSet::new();
    for diff in story.diff(txn, YChange::identity) {
        match diff.insert {
            Out::Any(Any::String(text)) => inline_length += text.encode_utf16().count() as u64,
            Out::YMap(map) if is_pilcrow(&map, txn) => {
                let para_id = map_string(&map, txn, PARA_ID).unwrap_or_default();
                if seen.insert(para_id.clone()) {
                    paragraphs.push(GeometryParagraphOutline {
                        para_id,
                        display_start: paragraph_start,
                        length: inline_length,
                        leading,
                    });
                }
                cursor = paragraph_start + inline_length + 2;
                paragraph_start = cursor;
                inline_length = 0;
                leading = 0;
            }
            Out::YMap(map) => match map_string(&map, txn, KIND_KEY).as_deref().unwrap_or("") {
                "table" => {
                    cursor += build_table(
                        txn,
                        outline,
                        &map,
                        story_id,
                        table_index,
                        content_start + cursor,
                        depth,
                    )?;
                    table_index += 1;
                    paragraph_start = cursor;
                    leading += 1;
                }
                "blockSdt" => {
                    let child = child_story(
                        plain_value(&map, txn, "story").as_ref(),
                        format!("{story_id}:sdt0"),
                    )?;
                    cursor +=
                        build_story(txn, outline, &child, content_start + cursor + 1, depth + 1)?
                            + 2;
                    paragraph_start = cursor;
                    leading += 1;
                }
                "pageBreak" | "columnBreak" if inline_length == 0 => {
                    cursor += 1;
                    paragraph_start = cursor;
                    leading += 1;
                }
                _ => inline_length += 1,
            },
            _ => inline_length += 1,
        }
        if cursor > MAX_POSITION || inline_length > MAX_POSITION {
            return Err(());
        }
    }
    let size = cursor.max(
        paragraph_start
            + if inline_length > 0 {
                inline_length + 2
            } else {
                0
            },
    );
    if size > MAX_POSITION {
        return Err(());
    }
    let entry = outline.get_mut(story_id).unwrap();
    entry.size = size;
    entry.paragraphs = paragraphs;
    Ok(size)
}

fn build_table<T: ReadTxn>(
    txn: &T,
    outline: &mut GeometryPositionOutline,
    map: &MapRef,
    story_id: &str,
    table_index: usize,
    start: u64,
    depth: usize,
) -> Result<u64, ()> {
    if start > MAX_POSITION {
        return Err(());
    }
    let rows = plain_value(map, txn, "rows");
    let rows = match rows.as_ref() {
        Some(Any::Array(rows)) => rows.as_ref(),
        Some(Any::Buffer(_)) => return Err(()),
        _ => return Ok(2),
    };
    let mut table_size = 0;
    for (row_index, row) in rows.iter().enumerate() {
        let Any::Map(row) = row else { return Err(()) };
        let cells = match row.get("cells") {
            Some(Any::Array(cells)) => cells.as_ref(),
            Some(Any::Buffer(_)) => return Err(()),
            _ => {
                table_size += 2;
                if table_size > MAX_POSITION {
                    return Err(());
                }
                continue;
            }
        };
        let mut row_size = 0;
        for (cell_index, cell) in cells.iter().enumerate() {
            let Any::Map(cell) = cell else { return Err(()) };
            let child = child_story(
                cell.get("story"),
                format!("{story_id}:t{table_index}:r{row_index}c{cell_index}"),
            )?;
            let child_start = start + 3 + table_size + row_size;
            row_size += build_story(txn, outline, &child, child_start, depth + 1)? + 2;
            if row_size > MAX_POSITION {
                return Err(());
            }
        }
        table_size += row_size + 2;
        if table_size > MAX_POSITION {
            return Err(());
        }
    }
    if table_size + 2 > MAX_POSITION {
        return Err(());
    }
    Ok(table_size + 2)
}

fn revision_key(key: &str) -> bool {
    matches!(
        key,
        "id" | "revisionId"
            | "pPrIns"
            | "pPrDel"
            | "pPrChange"
            | "trIns"
            | "trDel"
            | "tableIns"
            | "tableDel"
    )
}

fn revision_properties(value: &Any, depth: usize) -> Result<bool, GeometrySentinel> {
    if depth > 256 {
        return Err(GeometrySentinel::Legacy);
    }
    match value {
        Any::Map(map) => {
            for (key, child) in map.iter() {
                if revision_key(key) || revision_properties(child, depth + 1)? {
                    return Ok(true);
                }
            }
        }
        Any::Array(array) => {
            for child in array.iter() {
                if revision_properties(child, depth + 1)? {
                    return Ok(true);
                }
            }
        }
        _ => {}
    }
    Ok(false)
}

fn map_revision_properties<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    pilcrow: bool,
) -> Result<bool, GeometrySentinel> {
    for (key, value) in map.iter(txn) {
        if (pilcrow && is_identity_key(key)) || (!pilcrow && key == KIND_KEY) {
            continue;
        }
        let Out::Any(value) = value else { continue };
        if revision_key(key) || revision_properties(&value, 0)? {
            return Ok(true);
        }
    }
    Ok(false)
}

fn stamp_id(value: &Any) -> Result<Option<&str>, GeometrySentinel> {
    let Any::Map(attributes) = value else {
        return Ok(None);
    };
    let stamp = match attributes.get("info") {
        Some(Any::Map(info)) => info,
        _ => attributes,
    };
    match stamp.get("id").or_else(|| stamp.get("revisionId")) {
        Some(Any::String(id)) => Ok(Some(id.as_ref())),
        Some(Any::Number(number)) if number.is_finite() => Err(GeometrySentinel::Legacy),
        Some(Any::BigInt(_)) => Err(GeometrySentinel::Legacy),
        _ => Ok(None),
    }
}

struct Change {
    revision_id: String,
    kind: &'static str,
    start: usize,
    end: usize,
}

fn revision_ranges<T: ReadTxn>(
    txn: &T,
    ids: &[String],
) -> Result<Vec<OwnedRevisionRange>, GeometrySentinel> {
    let owned: HashSet<&str> = ids.iter().map(String::as_str).collect();
    let Some(stories) = txn.get_map(STORIES) else {
        return Ok(Vec::new());
    };
    let mut story_ids: Vec<_> = stories.iter(txn).map(|(id, _)| id.to_owned()).collect();
    if story_ids.iter().any(|id| !id.is_ascii()) {
        return Err(GeometrySentinel::Legacy);
    }
    story_ids.sort();
    let mut revisions = Vec::new();
    for story_id in story_ids {
        let story = story_ref(txn, &story_id).map_err(|_| GeometrySentinel::Legacy)?;
        let mut offset = 0;
        let mut paragraphs = HashSet::new();
        let mut changes: Vec<Change> = Vec::new();
        let mut previous: HashMap<&str, usize> = HashMap::new();
        for diff in story.diff(txn, YChange::identity) {
            if let Out::YMap(map) = &diff.insert {
                if is_pilcrow(map, txn) {
                    let para_id = map_string(map, txn, PARA_ID).unwrap_or_default();
                    if map_revision_properties(map, txn, true)?
                        || !paragraphs.insert(para_id.clone())
                    {
                        return Err(GeometrySentinel::Fallback);
                    }
                    changes.sort_by_key(|change| change.start);
                    for change in changes.drain(..) {
                        revisions.push(OwnedRevisionRange {
                            revision_id: change.revision_id,
                            kind: change.kind.to_owned(),
                            story: story_id.clone(),
                            range: GeometryRange {
                                story: story_id.clone(),
                                start: GeometryEndpoint {
                                    para_id: para_id.clone(),
                                    offset: change.start,
                                },
                                end: GeometryEndpoint {
                                    para_id: para_id.clone(),
                                    offset: change.end,
                                },
                            },
                        });
                    }
                    offset = 0;
                    previous.clear();
                    continue;
                }
                if map_revision_properties(map, txn, false)? {
                    return Err(GeometrySentinel::Fallback);
                }
            }
            let length = match &diff.insert {
                Out::Any(Any::String(text)) => text.encode_utf16().count(),
                _ => 1,
            };
            for (key, kind) in [("ins", "insertion"), ("del", "deletion")] {
                let Some(value) = diff.attributes.as_deref().and_then(|attrs| attrs.get(key))
                else {
                    continue;
                };
                let Some(id) = stamp_id(value)? else { continue };
                if !owned.contains(id) {
                    previous.remove(kind);
                    continue;
                }
                if let Some(&index) = previous.get(kind) {
                    let last = &mut changes[index];
                    if last.revision_id == id && last.end == offset {
                        last.end += length;
                        continue;
                    }
                }
                previous.insert(kind, changes.len());
                changes.push(Change {
                    revision_id: id.to_owned(),
                    kind,
                    start: offset,
                    end: offset + length,
                });
            }
            offset += length;
        }
        if !changes.is_empty() {
            return Err(GeometrySentinel::Fallback);
        }
    }
    Ok(revisions)
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::{Value, json};
    use yrs::types::Attrs;
    use yrs::{MapPrelim, Text};

    fn story(doc: &EditingDoc, id: &str, segments: Value) {
        doc.create_story(id, "", "Normal", "left").unwrap();
        let mut txn = doc.doc.transact_mut();
        let story = story_ref(&txn, id).unwrap();
        story.remove_range(&mut txn, 0, 1);
        let mut offset = 0;
        for segment in segments.as_array().unwrap() {
            let attrs: Attrs = segment
                .get("attrs")
                .and_then(Value::as_object)
                .into_iter()
                .flat_map(|attrs| attrs.iter())
                .map(|(key, value)| {
                    (
                        key.as_str().into(),
                        Any::from_json(&value.to_string()).unwrap(),
                    )
                })
                .collect();
            if let Some(text) = segment.get("text").and_then(Value::as_str) {
                story.insert_with_attributes(&mut txn, offset, text, attrs);
                offset += text.encode_utf16().count() as u32;
            } else {
                let payload = segment
                    .as_object()
                    .unwrap()
                    .iter()
                    .filter(|(key, _)| key.as_str() != "attrs")
                    .map(|(key, value)| {
                        (
                            if key == "kind" {
                                KIND_KEY.to_owned()
                            } else {
                                key.clone()
                            },
                            Any::from_json(&value.to_string()).unwrap(),
                        )
                    });
                story.insert_embed_with_attributes(
                    &mut txn,
                    offset,
                    MapPrelim::from_iter(payload),
                    attrs,
                );
                offset += 1;
            }
        }
    }

    fn outline(doc: &EditingDoc) -> Value {
        serde_json::to_value(doc.geometry_position_outline("body")).unwrap()
    }

    fn ranges(doc: &EditingDoc) -> Value {
        serde_json::to_value(doc.owned_revision_ranges(&["owned".into(), "other".into()])).unwrap()
    }

    fn range(story: &str, para: &str, id: &str, kind: &str, start: usize, end: usize) -> Value {
        json!({"revisionId": id, "kind": kind, "story": story, "range": {
            "story": story, "start": {"paraId": para, "offset": start},
            "end": {"paraId": para, "offset": end}
        }})
    }

    #[test]
    fn represented_merged_cells_nested_tables_and_empty_rows() {
        let doc = EditingDoc::new(84001);
        story(
            &doc,
            "body",
            json!([
                {"kind": "table", "rows": [
                    {"cells": [{"story": "a", "tcPr": {"rowspan": 2, "colspan": 2}}, {"story": "b"}]},
                    {"cells": [{}]}, {"cells": []}
                ]}, {"text": "Z"}, {"kind": "pilcrow", "paraId": "p"}
            ]),
        );
        story(
            &doc,
            "a",
            json!([
                {"kind": "table", "rows": []}, {"text": "😀"}, {"kind": "pilcrow", "paraId": "ap"}
            ]),
        );
        story(&doc, "b", json!([{ "kind": "pilcrow", "paraId": "bp" }]));
        story(&doc, "body:t0:r1c0", json!([]));
        assert_eq!(
            outline(&doc),
            json!({
                "body": {"contentStart": 0, "size": 25, "paragraphs": [{"paraId": "p", "displayStart": 22, "length": 1, "leading": 1}]},
                "a": {"contentStart": 3, "size": 6, "paragraphs": [{"paraId": "ap", "displayStart": 2, "length": 2, "leading": 1}]},
                "b": {"contentStart": 11, "size": 2, "paragraphs": [{"paraId": "bp", "displayStart": 0, "length": 0, "leading": 0}]},
                "body:t0:r1c0": {"contentStart": 17, "size": 0, "paragraphs": []}
            })
        );
    }

    #[test]
    fn nested_cells_keep_story_local_table_indices() {
        let doc = EditingDoc::new(84011);
        story(
            &doc,
            "body",
            json!([
                {"text": "x"}, {"kind": "pilcrow", "paraId": "p0"},
                {"kind": "table", "rows": [{"cells": [{}]}]},
                {"kind": "table", "rows": [{"cells": [{}]}]},
                {"text": "z"}, {"kind": "pilcrow", "paraId": "p1"}
            ]),
        );
        story(
            &doc,
            "body:t0:r0c0",
            json!([
                {"kind": "table", "rows": [{"cells": [{"story": null}]}]},
                {"text": "A"}, {"kind": "pilcrow", "paraId": "cp"}
            ]),
        );
        story(
            &doc,
            "body:t0:r0c0:t0:r0c0",
            json!([
                {"text": "😀"}, {"kind": "pilcrow", "paraId": "np"}
            ]),
        );
        story(
            &doc,
            "body:t1:r0c0",
            json!([{ "kind": "pilcrow", "paraId": "ep" }]),
        );
        assert_eq!(
            outline(&doc),
            json!({
                "body": {"contentStart": 0, "size": 33, "paragraphs": [
                    {"paraId": "p0", "displayStart": 0, "length": 1, "leading": 0},
                    {"paraId": "p1", "displayStart": 30, "length": 1, "leading": 2}
                ]},
                "body:t0:r0c0": {"contentStart": 6, "size": 13, "paragraphs": [
                    {"paraId": "cp", "displayStart": 10, "length": 1, "leading": 1}
                ]},
                "body:t0:r0c0:t0:r0c0": {"contentStart": 9, "size": 4, "paragraphs": [
                    {"paraId": "np", "displayStart": 0, "length": 2, "leading": 0}
                ]},
                "body:t1:r0c0": {"contentStart": 25, "size": 2, "paragraphs": [
                    {"paraId": "ep", "displayStart": 0, "length": 0, "leading": 0}
                ]}
            })
        );
    }

    #[test]
    fn block_sdts_breaks_inline_atoms_and_unterminated_text() {
        let doc = EditingDoc::new(84002);
        story(
            &doc,
            "body",
            json!([
                {"kind": "pageBreak"}, {"kind": "columnBreak"},
                {"kind": "blockSdt"}, {"text": "😀"},
                {"kind": "pageBreak"}, {"kind": "columnBreak"},
                {"kind": "field", "text": null}, {"kind": "image"},
                {"kind": "noteRef"}, {"kind": "inlineSdt"},
                {"kind": "table", "rows": []},
                {"kind": "pilcrow", "paraId": "p"}, {"text": "tail"}
            ]),
        );
        story(
            &doc,
            "body:sdt0",
            json!([{ "text": "x" }, {"kind": "pilcrow", "paraId": "child"}]),
        );
        let actual = outline(&doc);
        assert_eq!(
            actual["body"],
            json!({"contentStart": 0, "size": 25, "paragraphs": [
                {"paraId": "p", "displayStart": 9, "length": 8, "leading": 4}
            ]})
        );
        assert_eq!(actual["body:sdt0"]["contentStart"], 3);
    }

    #[test]
    fn first_seen_shared_and_cyclic_stories_and_duplicate_paragraphs() {
        let doc = EditingDoc::new(84003);
        story(
            &doc,
            "body",
            json!([
                {"kind": "blockSdt", "story": "child"},
                {"kind": "blockSdt", "story": "child"},
                {"text": "a"}, {"kind": "pilcrow", "paraId": "p"},
                {"text": "longer"}, {"kind": "pilcrow", "paraId": "p"}
            ]),
        );
        story(
            &doc,
            "child",
            json!([
                {"kind": "blockSdt", "story": "body"}, {"text": "b"}, {"kind": "pilcrow", "paraId": "cp"}
            ]),
        );
        let actual = outline(&doc);
        assert_eq!(
            actual["child"],
            json!({"contentStart": 1, "size": 5, "paragraphs": [
                {"paraId": "cp", "displayStart": 2, "length": 1, "leading": 1}
            ]})
        );
        assert_eq!(
            actual["body"],
            json!({"contentStart": 0, "size": 25, "paragraphs": [
                {"paraId": "p", "displayStart": 14, "length": 1, "leading": 2}
            ]})
        );
        assert_eq!(ranges(&doc), "fallback");
    }

    #[test]
    fn revision_order_coalescing_utf16_stamps_and_paragraph_boundaries() {
        let doc = EditingDoc::new(84004);
        story(
            &doc,
            "z",
            json!([
                {"text": "x", "attrs": {"ins": {"id": "owned"}}}, {"kind": "pilcrow", "paraId": "zp"}
            ]),
        );
        story(
            &doc,
            "body",
            json!([
                {"text": "😀", "attrs": {"ins": {"id": "owned"}, "del": {"info": {"revisionId": "other"}}}},
                {"kind": "image", "attrs": {"ins": {"revisionId": "owned"}, "del": {"id": "other"}}},
                {"text": "x", "attrs": {"ins": {"id": "unowned"}, "del": {"id": "other"}}},
                {"text": "y", "attrs": {"ins": {"info": {"id": "owned"}}, "del": {"id": "other"}}},
                {"kind": "pilcrow", "paraId": "p1", "attrs": {"ins": {"id": 12}}},
                {"text": "ab", "attrs": {"ins": {"id": "owned"}, "del": {"id": null, "revisionId": "other"}}},
                {"text": "c", "attrs": {"ins": {"id": "wrong", "info": {"id": "owned"}}}},
                {"kind": "pilcrow", "paraId": "p2"}
            ]),
        );
        assert_eq!(
            ranges(&doc),
            json!([
                range("body", "p1", "owned", "insertion", 0, 3),
                range("body", "p1", "other", "deletion", 0, 5),
                range("body", "p1", "owned", "insertion", 4, 5),
                range("body", "p2", "owned", "insertion", 0, 3),
                range("z", "zp", "owned", "insertion", 0, 1)
            ])
        );
    }

    #[test]
    fn revision_properties_and_owned_tail_fall_back() {
        for key in [
            "id",
            "revisionId",
            "pPrIns",
            "pPrDel",
            "pPrChange",
            "trIns",
            "trDel",
            "tableIns",
            "tableDel",
        ] {
            for kind in ["pilcrow", "image"] {
                let doc = EditingDoc::new(84005);
                let mut nested = serde_json::Map::new();
                nested.insert(key.to_owned(), Value::Null);
                story(
                    &doc,
                    "body",
                    json!([
                        {"kind": kind, "paraId": "p", "extra": [{"nested": nested}]}
                    ]),
                );
                assert_eq!(ranges(&doc), "fallback", "{key} in {kind}");
            }
        }
        let doc = EditingDoc::new(84006);
        story(
            &doc,
            "body",
            json!([{ "text": "x", "attrs": {"ins": {"id": "owned"}} }]),
        );
        assert_eq!(ranges(&doc), "fallback");
    }

    #[test]
    fn omitted_map_children_and_identity_values_do_not_trigger_fallback() {
        let doc = EditingDoc::new(84007);
        story(
            &doc,
            "body",
            json!([
                {"text": "x", "attrs": {"ins": {"id": "owned"}}},
                {"kind": "pilcrow", "paraId": "p", "sourceParaId": {"id": "ignored"}}
            ]),
        );
        {
            let mut txn = doc.doc.transact_mut();
            let text = story_ref(&txn, "body").unwrap();
            for diff in text.diff(&txn, YChange::identity) {
                if let Out::YMap(map) = diff.insert {
                    map.insert(&mut txn, "pPrIns", MapPrelim::default());
                }
            }
        }
        assert_eq!(
            ranges(&doc),
            json!([range("body", "p", "owned", "insertion", 0, 1)])
        );
    }

    #[test]
    fn numeric_ids_non_ascii_stories_and_odd_shapes_use_legacy() {
        let doc = EditingDoc::new(84008);
        story(
            &doc,
            "body",
            json!([
                {"text": "x", "attrs": {"ins": {"id": 42}}}, {"kind": "pilcrow", "paraId": "p"}
            ]),
        );
        assert_eq!(ranges(&doc), "legacy");
        let doc = EditingDoc::new(84009);
        story(&doc, "😀", json!([]));
        assert_eq!(ranges(&doc), "legacy");
        for segment in [
            json!({"kind": "blockSdt", "story": 42}),
            json!({"kind": "blockSdt", "story": "missing"}),
            json!({"kind": "table", "rows": [null]}),
            json!({"kind": "table", "rows": [{"cells": [false]}]}),
            json!({"kind": "table", "rows": [{"cells": [{"story": []}]}]}),
        ] {
            let doc = EditingDoc::new(84010);
            story(&doc, "body", json!([segment]));
            assert_eq!(outline(&doc), "legacy");
        }
    }

    #[test]
    fn shared_stories_keep_large_positions_or_use_legacy() {
        for depth in [35, 53] {
            let doc = EditingDoc::new(84012);
            story(
                &doc,
                "leaf",
                json!([
                    {"text": "x"}, {"kind": "pilcrow", "paraId": "leaf"}
                ]),
            );
            let mut child = "leaf".to_owned();
            for index in 0..depth {
                let id = if index + 1 == depth {
                    "body".to_owned()
                } else {
                    format!("s{index}")
                };
                story(
                    &doc,
                    &id,
                    json!([
                        {"kind": "blockSdt", "story": child},
                        {"kind": "blockSdt", "story": child},
                        {"kind": "pilcrow", "paraId": id}
                    ]),
                );
                child = id;
            }
            if depth == 35 {
                let GeometryRead::Value(outline) = doc.geometry_position_outline("body") else {
                    panic!("legacy outline for safe positions");
                };
                let size = 9_u64 * (1_u64 << depth) - 6;
                assert_eq!(outline["body"].size, size);
                assert_eq!(outline["body"].paragraphs[0].display_start, size - 2);
            } else {
                assert_eq!(outline(&doc), "legacy");
            }
        }
    }
}
