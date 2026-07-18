//! Deterministic `canonical-stream-v1` projection for the coexistence invariant watchdog.
//!
//! The byte format is deliberately small enough to reproduce exactly in TypeScript:
//!
//! 1. UTF-8 bytes for `canonical-stream-v1`, followed by LF (`0x0a`).
//! 2. One compact JSON object per item, followed by LF (including the final item). Items use
//!    serde's externally-tagged spelling: `{"CharItem":{"ch":"x","marks":{...}}}`,
//!    `{"ParaMark":{"ppr":{...}}}`, `{"Table":{"tblPr":{...},"grid":[],"rows":[]}}`,
//!    `{"Embed":{"kind":"break","payload":{...}}}`, or `{"OpaqueBlk":{"blob":...}}`.
//! 3. Object keys at every depth are ordered by Rust `str`/`BTreeMap` lexicographic order. Arrays
//!    retain their input order. JSON is compact and uses `serde_json`'s escaping and number format.
//! 4. JSON `null` (including yrs `Any::Undefined`) is removed from objects at every depth. A null
//!    array slot remains `null`, because removing it would change array positions. A top-level
//!    opaque blob may also be `null` because its field itself cannot be made absent.
//!
//! Text is projected as one [`CanonicalItem::CharItem`] per Unicode scalar value (Rust `char`),
//! not per UTF-16 code unit and never as coalesced runs. A pilcrow's `_kind` discriminator and
//! volatile `paraId` are excluded from [`CanonicalItem::ParaMark`], so equivalent documents with
//! freshly minted paragraph IDs compare equal. Tabs are ordinary `"\t"` character items. Current
//! non-pilcrow atoms (including hard breaks and native `noteRef` footnote anchors) are map embeds
//! whose `_kind` becomes `kind` and whose remaining map becomes `payload`.
//!
//! Comments (S4a) are projected as PRESENCE/GROUPING, not identity: every char item fully covered
//! by a side-map comment's resolved anchors carries a sorted `commentIds` array of per-story
//! comment ORDINALS. Ordinals are assigned by sorting each covering comment's covered story units
//! (UTF-16 units, every embed = 1) lexicographically, so volatile comment ids (`{client}:{counter}`
//! keys, PM numeric ids) never reach the stream — exactly like the excluded `paraId`. The field is
//! omitted when empty, so documents without comments project byte-identical streams to before.
//! Comments whose anchors no longer resolve, resolve empty, or live in another story are skipped.
//!
//! [`CanonicalItem::OpaqueBlk`] is reserved for the coexistence two-tier representation. The
//! current editing schema has no opaque-block discriminator, so [`project_story`] does not invent
//! one; Wave 1b can add the schema-aware projection while retaining this byte vocabulary.

use std::collections::BTreeMap;

use serde_json::{Map as JsonMap, Value};
use yrs::types::{ToJson, text::YChange};
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, Transact};

use crate::op::OpError;
use crate::{EditingDoc, KIND_KEY, PARA_ID, is_pilcrow, map_string, story_ref};

const VERSION: &str = "canonical-stream-v1";

/// One semantic unit in the cross-language coexistence stream.
#[derive(Clone, Debug, PartialEq)]
pub enum CanonicalItem {
    /// One Unicode scalar of story text and its active formatting attributes.
    /// `comment_ids` carries the sorted per-story ordinals of the comments covering
    /// this scalar (empty and omitted from the wire bytes when uncommented).
    CharItem {
        ch: String,
        marks: BTreeMap<String, Value>,
        comment_ids: Vec<u64>,
    },
    /// One paragraph boundary and its authored properties, excluding `_kind` and `paraId`.
    ParaMark { ppr: BTreeMap<String, Value> },
    /// One native table structure. Cell story references are volatile identity
    /// and are removed; each cell's authored content projects in its own story.
    Table {
        tbl_pr: Value,
        grid: Value,
        rows: Value,
    },
    /// A non-pilcrow inline atom. The `_kind` discriminator is not duplicated in `payload`.
    Embed {
        kind: String,
        payload: BTreeMap<String, Value>,
    },
    /// A sealed PM block JSON blob. Reserved until the Wave 1b opaque tier is added to the model.
    OpaqueBlk { blob: Value },
}

/// Projects one yrs story into canonical semantic units.
pub fn project_story(doc: &EditingDoc, story_id: &str) -> Result<Vec<CanonicalItem>, OpError> {
    let txn = doc.yrs_doc().transact();
    let story = story_ref(&txn, story_id)?;
    let comment_groups = story_comment_groups(&txn, story_id);
    let mut items = Vec::new();
    // The running UTF-16 story-unit index (every embed, pilcrows included, = 1),
    // used to test each scalar against the resolved comment intervals.
    let mut unit = 0_u32;

    for diff in story.diff(&txn, YChange::identity) {
        match diff.insert {
            Out::Any(Any::String(text)) => {
                let marks = canonical_attrs(diff.attributes.as_deref());
                for ch in text.chars() {
                    let width = ch.len_utf16() as u32;
                    items.push(CanonicalItem::CharItem {
                        ch: ch.to_string(),
                        marks: marks.clone(),
                        comment_ids: covering_ordinals(&comment_groups, unit, width),
                    });
                    unit += width;
                }
            }
            Out::YMap(map) if is_pilcrow(&map, &txn) => {
                items.push(CanonicalItem::ParaMark {
                    ppr: canonical_shared_map(&map, &txn, &[KIND_KEY, PARA_ID]),
                });
                unit += 1;
            }
            Out::YMap(map) if map_string(&map, &txn, KIND_KEY).as_deref() == Some("table") => {
                items.push(canonical_table(&map, &txn));
                unit += 1;
            }
            Out::YMap(map) if map_string(&map, &txn, KIND_KEY).as_deref() == Some("blockSdt") => {
                // The child-story key is stable plumbing identity, not authored
                // content (same rule as table-cell story references).
                items.push(CanonicalItem::Embed {
                    kind: "blockSdt".to_owned(),
                    payload: canonical_shared_map(&map, &txn, &[KIND_KEY, "story"]),
                });
                unit += 1;
            }
            Out::YMap(map) => {
                items.push(CanonicalItem::Embed {
                    kind: map_string(&map, &txn, KIND_KEY).unwrap_or_default(),
                    payload: canonical_shared_map(&map, &txn, &[KIND_KEY]),
                });
                unit += 1;
            }
            Out::Any(Any::Map(map)) => {
                let kind = match map.get(KIND_KEY) {
                    Some(Any::String(value)) => value.to_string(),
                    _ => String::new(),
                };
                items.push(CanonicalItem::Embed {
                    kind,
                    payload: canonical_any_map(&map, &[KIND_KEY]),
                });
                unit += 1;
            }
            other => {
                // The current public schema creates map-backed embeds. Preserve a deterministic
                // identity for legacy/non-schema atom values instead of silently dropping them.
                let mut payload = BTreeMap::new();
                if let Some(value) = canonical_out(other, &txn) {
                    payload.insert("value".to_owned(), value);
                }
                items.push(CanonicalItem::Embed {
                    kind: String::new(),
                    payload,
                });
                unit += 1;
            }
        }
    }

    Ok(items)
}

fn canonical_table<T: ReadTxn>(map: &MapRef, txn: &T) -> CanonicalItem {
    let object = || Value::Object(JsonMap::new());
    let array = || Value::Array(Vec::new());
    let tbl_pr = map
        .get(txn, "tblPr")
        .and_then(|value| canonical_out(value, txn))
        .unwrap_or_else(object);
    let grid = map
        .get(txn, "grid")
        .and_then(|value| canonical_out(value, txn))
        .unwrap_or_else(array);
    let mut rows = map
        .get(txn, "rows")
        .and_then(|value| canonical_out(value, txn))
        .unwrap_or_else(array);
    exclude_cell_story_ids(&mut rows);
    CanonicalItem::Table { tbl_pr, grid, rows }
}

fn exclude_cell_story_ids(rows: &mut Value) {
    let Value::Array(rows) = rows else {
        return;
    };
    for row in rows {
        let Some(cells) = row.get_mut("cells").and_then(Value::as_array_mut) else {
            continue;
        };
        for cell in cells {
            if let Value::Object(cell) = cell {
                cell.remove("story");
            }
        }
    }
}

/// Resolves the side-map comments anchored in `story_id` to ordinal-ordered interval
/// lists: index = the comment's per-story ordinal, value = its merged, sorted
/// `[start, end)` UTF-16 story-unit intervals. Ordinals order groups by their covered
/// story units, compared lexicographically — a volatile-id-free grouping that the
/// TypeScript projector (`canonicalStream.ts`) reproduces from PM comment marks.
fn story_comment_groups<T: ReadTxn>(txn: &T, story_id: &str) -> Vec<Vec<(u32, u32)>> {
    let Some(comments) = txn.get_map(crate::COMMENTS) else {
        return Vec::new();
    };
    let mut groups: Vec<(Vec<u32>, Vec<(u32, u32)>)> = Vec::new();
    for (_, value) in comments.iter(txn) {
        let Out::YMap(comment) = value else {
            continue;
        };
        let Some(Out::Any(Any::Array(anchors))) = comment.get(txn, "anchors") else {
            continue;
        };
        let mut intervals: Vec<(u32, u32)> = Vec::new();
        for encoded in anchors.iter() {
            let Ok(anchor) = crate::decode_anchor(encoded) else {
                continue;
            };
            if anchor.story != story_id {
                continue;
            }
            let (Some(start), Some(end)) =
                (anchor.start.get_offset(txn), anchor.end.get_offset(txn))
            else {
                continue;
            };
            if start.index < end.index {
                intervals.push((start.index, end.index));
            }
        }
        if intervals.is_empty() {
            continue;
        }
        let intervals = merge_intervals(intervals);
        let covered: Vec<u32> = intervals
            .iter()
            .flat_map(|&(start, end)| start..end)
            .collect();
        groups.push((covered, intervals));
    }
    // Lexicographic covered-unit order; ties (identical coverage) are interchangeable
    // in the projection, so any stable outcome is deterministic.
    groups.sort_by(|left, right| left.0.cmp(&right.0));
    groups.into_iter().map(|(_, intervals)| intervals).collect()
}

/// Merges sorted-or-unsorted intervals, coalescing overlapping AND adjacent spans.
fn merge_intervals(mut intervals: Vec<(u32, u32)>) -> Vec<(u32, u32)> {
    intervals.sort_unstable();
    let mut merged: Vec<(u32, u32)> = Vec::with_capacity(intervals.len());
    for (start, end) in intervals {
        match merged.last_mut() {
            Some(last) if start <= last.1 => last.1 = last.1.max(end),
            _ => merged.push((start, end)),
        }
    }
    merged
}

/// The sorted ordinals of every comment group fully covering `[unit, unit + width)`.
fn covering_ordinals(groups: &[Vec<(u32, u32)>], unit: u32, width: u32) -> Vec<u64> {
    groups
        .iter()
        .enumerate()
        .filter(|(_, intervals)| {
            intervals
                .iter()
                .any(|&(start, end)| start <= unit && unit + width <= end)
        })
        .map(|(ordinal, _)| ordinal as u64)
        .collect()
}

/// Serializes canonical items to the exact `canonical-stream-v1` bytes documented above.
pub fn to_canonical_bytes(items: &[CanonicalItem]) -> Vec<u8> {
    let mut bytes = Vec::new();
    bytes.extend_from_slice(VERSION.as_bytes());
    bytes.push(b'\n');
    for item in items {
        let encoded = serde_json::to_vec(&canonical_item_value(item))
            .expect("canonical items contain only JSON values");
        bytes.extend_from_slice(&encoded);
        bytes.push(b'\n');
    }
    bytes
}

/// Stable 64-bit FNV-1a over [`to_canonical_bytes`].
pub fn checksum(items: &[CanonicalItem]) -> u64 {
    const OFFSET_BASIS: u64 = 0xcbf2_9ce4_8422_2325;
    const PRIME: u64 = 0x0000_0100_0000_01b3;

    to_canonical_bytes(items)
        .into_iter()
        .fold(OFFSET_BASIS, |hash, byte| {
            (hash ^ u64::from(byte)).wrapping_mul(PRIME)
        })
}

/// Projects and checksums one story in a single convenience call.
pub fn story_checksum(doc: &EditingDoc, story_id: &str) -> Result<u64, OpError> {
    Ok(checksum(&project_story(doc, story_id)?))
}

fn canonical_attrs(attrs: Option<&yrs::types::Attrs>) -> BTreeMap<String, Value> {
    attrs
        .into_iter()
        .flat_map(|attrs| attrs.iter())
        .filter_map(|(key, value)| canonical_any(value).map(|value| (key.to_string(), value)))
        .collect()
}

fn canonical_shared_map<T: ReadTxn>(
    map: &MapRef,
    txn: &T,
    excluded: &[&str],
) -> BTreeMap<String, Value> {
    map.iter(txn)
        .filter_map(|(key, value)| {
            if excluded.contains(&key) {
                return None;
            }
            canonical_out(value, txn).map(|value| (key.to_string(), value))
        })
        .collect()
}

fn canonical_any_map(
    map: &std::collections::HashMap<String, Any>,
    excluded: &[&str],
) -> BTreeMap<String, Value> {
    map.iter()
        .filter_map(|(key, value)| {
            if excluded.contains(&key.as_str()) {
                None
            } else {
                canonical_any(value).map(|value| (key.clone(), value))
            }
        })
        .collect()
}

fn canonical_any(value: &Any) -> Option<Value> {
    let value = serde_json::to_value(value).expect("yrs Any serializes to JSON");
    normalize_value(&value)
}

fn canonical_out<T: ReadTxn>(value: Out, txn: &T) -> Option<Value> {
    canonical_any(&value.to_json(txn))
}

fn normalize_value(value: &Value) -> Option<Value> {
    match value {
        Value::Null => None,
        Value::Array(values) => Some(Value::Array(
            values
                .iter()
                .map(|value| normalize_value(value).unwrap_or(Value::Null))
                .collect(),
        )),
        Value::Object(values) => {
            let sorted: BTreeMap<String, Value> = values
                .iter()
                .filter_map(|(key, value)| normalize_value(value).map(|value| (key.clone(), value)))
                .collect();
            Some(Value::Object(sorted.into_iter().collect()))
        }
        value => Some(value.clone()),
    }
}

fn canonical_map(values: &BTreeMap<String, Value>) -> Value {
    let values: JsonMap<String, Value> = values
        .iter()
        .filter_map(|(key, value)| normalize_value(value).map(|value| (key.clone(), value)))
        .collect();
    Value::Object(values)
}

fn object(entries: impl IntoIterator<Item = (String, Value)>) -> Value {
    let sorted: BTreeMap<String, Value> = entries.into_iter().collect();
    Value::Object(sorted.into_iter().collect())
}

fn canonical_item_value(item: &CanonicalItem) -> Value {
    match item {
        CanonicalItem::CharItem {
            ch,
            marks,
            comment_ids,
        } => {
            let mut entries = vec![
                ("ch".to_owned(), Value::String(ch.clone())),
                ("marks".to_owned(), canonical_map(marks)),
            ];
            if !comment_ids.is_empty() {
                entries.push((
                    "commentIds".to_owned(),
                    Value::Array(comment_ids.iter().map(|id| Value::from(*id)).collect()),
                ));
            }
            object([("CharItem".to_owned(), object(entries))])
        }
        CanonicalItem::ParaMark { ppr } => object([(
            "ParaMark".to_owned(),
            object([("ppr".to_owned(), canonical_map(ppr))]),
        )]),
        CanonicalItem::Table { tbl_pr, grid, rows } => object([(
            "Table".to_owned(),
            object([
                ("tblPr".to_owned(), tbl_pr.clone()),
                ("grid".to_owned(), grid.clone()),
                ("rows".to_owned(), rows.clone()),
            ]),
        )]),
        CanonicalItem::Embed { kind, payload } => object([(
            "Embed".to_owned(),
            object([
                ("kind".to_owned(), Value::String(kind.clone())),
                ("payload".to_owned(), canonical_map(payload)),
            ]),
        )]),
        CanonicalItem::OpaqueBlk { blob } => object([(
            "OpaqueBlk".to_owned(),
            object([(
                "blob".to_owned(),
                normalize_value(blob).unwrap_or(Value::Null),
            )]),
        )]),
    }
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;
    use std::sync::Arc;

    use yrs::Any;
    use yrs::types::Attrs;

    use super::*;
    use crate::{
        ColorPatch, EditCtx, FormatPolicy, InlineFormatDelta, MergeDirection, Patch, Position,
        RawOp, SimpleFormat, StoryRange,
    };

    const DATE: &str = "2026-07-13T12:00:00Z";

    fn direct() -> EditCtx {
        EditCtx::local("Ada", DATE)
    }

    #[test]
    fn table_structure_matches_canonical_golden_and_excludes_cell_story_ids() {
        let doc = EditingDoc::new(40);
        doc.create_story("body", "", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Delete { index: 0, len: 1 },
                RawOp::InsertEmbed {
                    index: 0,
                    kind: "table".to_owned(),
                    payload: vec![
                        (
                            "tblPr".to_owned(),
                            Any::from_json(
                                r#"{"width":5000,"widthType":"pct","bidi":true}"#,
                            )
                            .unwrap(),
                        ),
                        (
                            "grid".to_owned(),
                            Any::from_json("[1200,2400]").unwrap(),
                        ),
                        (
                            "rows".to_owned(),
                            Any::from_json(
                                r#"[{"trPr":{"isHeader":false},"cells":[{"tcPr":{"colspan":1,"rowspan":1,"noWrap":false},"story":"volatile:cell"}]}]"#,
                            )
                            .unwrap(),
                        ),
                    ],
                    attrs: Attrs::new(),
                },
            ],
            &direct(),
        )
        .unwrap();

        let items = project_story(&doc, "body").unwrap();
        const EXPECTED: &str = concat!(
            "canonical-stream-v1\n",
            "{\"Table\":{\"grid\":[1200,2400],\"rows\":[{\"cells\":[{\"tcPr\":{\"colspan\":1,\"noWrap\":false,\"rowspan\":1}}],\"trPr\":{\"isHeader\":false}}],\"tblPr\":{\"bidi\":true,\"width\":5000,\"widthType\":\"pct\"}}}\n",
        );
        assert_eq!(to_canonical_bytes(&items), EXPECTED.as_bytes());
        assert!(!EXPECTED.contains("volatile:cell"));
    }

    #[test]
    fn block_sdt_projects_authored_properties_but_not_child_story_identity() {
        let doc = EditingDoc::new(43);
        doc.create_story("body", "", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Delete { index: 0, len: 1 },
                RawOp::InsertEmbed {
                    index: 0,
                    kind: "blockSdt".to_owned(),
                    payload: vec![
                        ("story".to_owned(), Any::from("body:sdt0")),
                        ("sdtType".to_owned(), Any::from("richText")),
                        ("tag".to_owned(), Any::from("intro")),
                    ],
                    attrs: Attrs::new(),
                },
            ],
            &direct(),
        )
        .unwrap();

        const EXPECTED: &str = concat!(
            "canonical-stream-v1\n",
            "{\"Embed\":{\"kind\":\"blockSdt\",\"payload\":{\"sdtType\":\"richText\",\"tag\":\"intro\"}}}\n",
        );
        assert_eq!(
            to_canonical_bytes(&project_story(&doc, "body").unwrap()),
            EXPECTED.as_bytes()
        );
        assert!(!EXPECTED.contains("body:sdt0"));
    }

    #[test]
    fn note_ref_embed_matches_canonical_golden() {
        let doc = EditingDoc::new(42);
        doc.create_story("body", "", "Normal", "left").unwrap();
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Delete { index: 0, len: 1 },
                RawOp::InsertEmbed {
                    index: 0,
                    kind: "noteRef".to_owned(),
                    payload: vec![("footnoteRefId".to_owned(), Any::Number(5.0))],
                    attrs: Attrs::new(),
                },
            ],
            &direct(),
        )
        .unwrap();

        const EXPECTED: &str = concat!(
            "canonical-stream-v1\n",
            "{\"Embed\":{\"kind\":\"noteRef\",\"payload\":{\"footnoteRefId\":5}}}\n",
        );
        assert_eq!(
            to_canonical_bytes(&project_story(&doc, "body").unwrap()),
            EXPECTED.as_bytes()
        );
    }

    #[test]
    fn representative_story_matches_canonical_stream_v1_golden() {
        let doc = EditingDoc::new(41);
        doc.create_story("body", "A😀 beta", "Normal", "left")
            .unwrap();
        doc.insert_text(
            &direct(),
            Position::new("body", 4),
            "\té",
            FormatPolicy::Plain,
        )
        .unwrap();
        let split = doc
            .split_paragraph(&direct(), Position::new("body", 7), None)
            .unwrap();

        doc.toggle_format(&direct(), StoryRange::new("body", 0, 3), SimpleFormat::Bold)
            .unwrap();
        doc.format_range(
            &direct(),
            StoryRange::new("body", 5, 7),
            &InlineFormatDelta {
                italic: Patch::Set(true),
                color: Patch::Set(ColorPatch::Rgb("112233".to_owned())),
                ..InlineFormatDelta::default()
            },
        )
        .unwrap();
        doc.set_paragraph_attr(&split.first_para_id, "alignment", Any::from("center"))
            .unwrap();
        doc.set_paragraph_attr(&split.first_para_id, "keepNext", Any::Bool(true))
            .unwrap();
        doc.set_paragraph_attr(
            &split.first_para_id,
            "spacingExplicit",
            Any::Map(Arc::new(HashMap::from([
                ("before".to_owned(), Any::Bool(true)),
                ("after".to_owned(), Any::Null),
            ]))),
        )
        .unwrap();
        doc.set_paragraph_attr(&split.second_para_id, "pStyle", Any::from("Heading1"))
            .unwrap();
        doc.insert_hard_break(&direct(), Position::new("body", 9))
            .unwrap();

        let items = project_story(&doc, "body").unwrap();
        const EXPECTED: &str = concat!(
            "canonical-stream-v1\n",
            "{\"CharItem\":{\"ch\":\"A\",\"marks\":{\"bold\":true}}}\n",
            "{\"CharItem\":{\"ch\":\"😀\",\"marks\":{\"bold\":true}}}\n",
            "{\"CharItem\":{\"ch\":\" \",\"marks\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"\\t\",\"marks\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"é\",\"marks\":{\"italic\":true,\"textColor\":{\"rgb\":\"112233\"}}}}\n",
            "{\"CharItem\":{\"ch\":\"b\",\"marks\":{\"italic\":true,\"textColor\":{\"rgb\":\"112233\"}}}}\n",
            "{\"ParaMark\":{\"ppr\":{\"alignment\":\"center\",\"keepNext\":true,\"pStyle\":\"Normal\",\"spacingExplicit\":{\"before\":true}}}}\n",
            "{\"CharItem\":{\"ch\":\"e\",\"marks\":{}}}\n",
            "{\"Embed\":{\"kind\":\"break\",\"payload\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"t\",\"marks\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"a\",\"marks\":{}}}\n",
            "{\"ParaMark\":{\"ppr\":{\"alignment\":\"left\",\"pStyle\":\"Heading1\"}}}\n",
        );
        assert_eq!(to_canonical_bytes(&items), EXPECTED.as_bytes());
    }

    #[test]
    fn section_break_paragraph_projects_sect_pr_in_para_mark() {
        let doc = EditingDoc::new(51);
        let para = doc.create_story("body", "end", "Normal", "left").unwrap();
        doc.set_paragraph_attr(&para, "sectionBreakType", Any::from("nextPage"))
            .unwrap();
        // Integers use Any::BigInt: the coexistence pipeline delivers whole
        // JSON numbers as integers, which serialize without a fraction — the
        // exact bytes the TypeScript projector emits. The null entry checks
        // the shared null-strip inside the sectPr sub-map.
        doc.set_paragraph_attr(
            &para,
            "sectPr",
            Any::Map(Arc::new(HashMap::from([
                ("sectionStart".to_owned(), Any::from("nextPage")),
                ("pageWidth".to_owned(), Any::BigInt(15840)),
                ("pageHeight".to_owned(), Any::BigInt(12240)),
                ("orientation".to_owned(), Any::from("landscape")),
                ("marginTop".to_owned(), Any::BigInt(720)),
                ("columnCount".to_owned(), Any::BigInt(2)),
                ("separator".to_owned(), Any::Bool(true)),
                ("headerDistance".to_owned(), Any::Null),
            ]))),
        )
        .unwrap();

        let items = project_story(&doc, "body").unwrap();
        const EXPECTED: &str = concat!(
            "canonical-stream-v1\n",
            "{\"CharItem\":{\"ch\":\"e\",\"marks\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"n\",\"marks\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"d\",\"marks\":{}}}\n",
            "{\"ParaMark\":{\"ppr\":{\"alignment\":\"left\",\"pStyle\":\"Normal\",\"sectPr\":{\"columnCount\":2,\"marginTop\":720,\"orientation\":\"landscape\",\"pageHeight\":12240,\"pageWidth\":15840,\"sectionStart\":\"nextPage\",\"separator\":true},\"sectionBreakType\":\"nextPage\"}}}\n",
        );
        assert_eq!(to_canonical_bytes(&items), EXPECTED.as_bytes());
    }

    #[test]
    fn paragraph_id_churn_does_not_change_stream_or_checksum() {
        let direct_doc = EditingDoc::new(101);
        direct_doc
            .create_story("body", "stable", "Normal", "left")
            .unwrap();

        let churned_doc = EditingDoc::new(202);
        churned_doc
            .create_story("body", "stable", "Normal", "left")
            .unwrap();
        let split = churned_doc
            .split_paragraph(&direct(), Position::new("body", 3), None)
            .unwrap();
        churned_doc
            .merge_paragraphs(&direct(), &split.first_para_id, MergeDirection::Forward)
            .unwrap();

        let direct_items = project_story(&direct_doc, "body").unwrap();
        let churned_items = project_story(&churned_doc, "body").unwrap();
        assert_ne!(
            direct_doc.paragraphs("body").unwrap()[0].para_id,
            churned_doc.paragraphs("body").unwrap()[0].para_id
        );
        assert_eq!(direct_items, churned_items);
        assert_eq!(checksum(&direct_items), checksum(&churned_items));
        assert_eq!(
            story_checksum(&direct_doc, "body").unwrap(),
            story_checksum(&churned_doc, "body").unwrap()
        );
    }

    #[test]
    fn commented_story_matches_the_commentids_golden() {
        let doc = EditingDoc::new(51);
        doc.create_story("body", "abcde", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        // Keys are deliberately out of positional order: ordinals must come from
        // covered-unit order, not from the volatile comment ids.
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::SetComment {
                    id: "7".to_owned(),
                    ranges: vec![(1, 4)],
                    author: String::new(),
                    date: String::new(),
                    body: Any::Null,
                },
                RawOp::SetComment {
                    id: "3".to_owned(),
                    ranges: vec![(3, 5)],
                    author: String::new(),
                    date: String::new(),
                    body: Any::Null,
                },
            ],
            &ctx,
        )
        .unwrap();

        let items = project_story(&doc, "body").unwrap();
        const EXPECTED: &str = concat!(
            "canonical-stream-v1\n",
            "{\"CharItem\":{\"ch\":\"a\",\"marks\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"b\",\"commentIds\":[0],\"marks\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"c\",\"commentIds\":[0],\"marks\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"d\",\"commentIds\":[0,1],\"marks\":{}}}\n",
            "{\"CharItem\":{\"ch\":\"e\",\"commentIds\":[1],\"marks\":{}}}\n",
            "{\"ParaMark\":{\"ppr\":{\"alignment\":\"left\",\"pStyle\":\"Normal\"}}}\n",
        );
        assert_eq!(to_canonical_bytes(&items), EXPECTED.as_bytes());
    }

    #[test]
    fn comment_id_churn_does_not_change_stream_or_checksum() {
        let ctx = EditCtx::local(String::new(), String::new());
        let raw_keyed = EditingDoc::new(61);
        raw_keyed
            .create_story("body", "hello", "Normal", "left")
            .unwrap();
        raw_keyed
            .apply_raw_ops(
                "body",
                vec![RawOp::SetComment {
                    id: "42".to_owned(),
                    ranges: vec![(0, 5)],
                    author: "Ada".to_owned(),
                    date: DATE.to_owned(),
                    body: Any::from("side-map body"),
                }],
                &ctx,
            )
            .unwrap();

        // Same coverage authored through the typed op with a generated
        // `{client}:{counter}` key and different metadata.
        let typed = EditingDoc::new(62);
        typed
            .create_story("body", "hello", "Normal", "left")
            .unwrap();
        typed
            .add_comment(
                &[StoryRange::new("body", 0, 5)],
                "Grace",
                DATE,
                Any::from("other body"),
            )
            .unwrap();

        assert_eq!(
            project_story(&raw_keyed, "body").unwrap(),
            project_story(&typed, "body").unwrap()
        );
        assert_eq!(
            story_checksum(&raw_keyed, "body").unwrap(),
            story_checksum(&typed, "body").unwrap()
        );
        // And commenting changed the stream relative to the uncommented story.
        let plain = EditingDoc::new(63);
        plain
            .create_story("body", "hello", "Normal", "left")
            .unwrap();
        assert_ne!(
            story_checksum(&plain, "body").unwrap(),
            story_checksum(&typed, "body").unwrap()
        );
    }

    #[test]
    fn unresolvable_and_empty_comment_anchors_are_skipped() {
        let doc = EditingDoc::new(71);
        doc.create_story("body", "xy", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "5".to_owned(),
                ranges: vec![(0, 2)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        )
        .unwrap();
        let commented = story_checksum(&doc, "body").unwrap();

        // Deleting every covered unit collapses the anchors to an empty interval;
        // the projection must then match a never-commented equivalent document.
        doc.apply_raw_ops("body", vec![RawOp::Delete { index: 0, len: 2 }], &ctx)
            .unwrap();
        let plain = EditingDoc::new(72);
        plain.create_story("body", "", "Normal", "left").unwrap();
        assert_ne!(commented, story_checksum(&doc, "body").unwrap());
        assert_eq!(
            story_checksum(&doc, "body").unwrap(),
            story_checksum(&plain, "body").unwrap()
        );
    }

    #[test]
    fn checksum_changes_with_character_marks_and_paragraph_properties() {
        let base = EditingDoc::new(301);
        base.create_story("body", "abc", "Normal", "left").unwrap();
        let marked = EditingDoc::new(302);
        marked
            .create_story("body", "abc", "Normal", "left")
            .unwrap();
        marked
            .toggle_format(&direct(), StoryRange::new("body", 0, 1), SimpleFormat::Bold)
            .unwrap();
        let changed_ppr = EditingDoc::new(303);
        let changed_para = changed_ppr
            .create_story("body", "abc", "Normal", "left")
            .unwrap();
        changed_ppr
            .set_paragraph_attr(&changed_para, "alignment", Any::from("right"))
            .unwrap();

        let base_checksum = story_checksum(&base, "body").unwrap();
        assert_ne!(base_checksum, story_checksum(&marked, "body").unwrap());
        assert_ne!(base_checksum, story_checksum(&changed_ppr, "body").unwrap());
    }
}
