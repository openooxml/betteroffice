//! Raw story mutations using UTF-16 story indices.

use std::collections::BTreeMap;
use std::sync::Arc;

use yrs::types::text::YChange;
use yrs::types::{Attrs, Delta};
use yrs::{
    Any, Assoc, In, IndexedSequence, Map, MapPrelim, MapRef, Out, ReadTxn, Text, TextRef,
    TransactionMut,
};

use crate::op::{OpError, OpResult};
use crate::{COMMENTS, EditCtx, EditingDoc, KIND_KEY, anchor_value, out_len, story_ref};

/// One low-level story mutation. Indices are UTF-16 story units (every embed = 1).
#[derive(Clone, Debug, PartialEq)]
pub enum RawOp {
    /// Inserts text with explicit attributes.
    Insert {
        index: u32,
        text: String,
        attrs: Attrs,
    },
    /// Remove `len` units starting at `index` (text, embeds, or pilcrows alike).
    Delete { index: u32, len: u32 },
    /// Re-format `len` units at `index`; `Any::Null` values clear an attribute.
    Format { index: u32, len: u32, attrs: Attrs },
    /// Insert a map-backed embed at `index` with discriminator `kind` (`pilcrow`,
    /// `break`, `opaque`, …) and `payload` map entries. `attrs` are the embed's
    /// text-level formatting (usually just tracked-change stamps, if any).
    InsertEmbed {
        index: u32,
        kind: String,
        payload: Vec<(String, Any)>,
        attrs: Attrs,
    },
    /// Sets one key on the map-backed embed at the index.
    SetEmbedAttr { index: u32, key: String, value: Any },
    /// Upserts a side-map comment with sticky UTF-16 ranges.
    SetComment {
        id: String,
        ranges: Vec<(u32, u32)>,
        author: String,
        date: String,
        body: Any,
    },
    /// Remove the side-map comment keyed by `id`. Errors when it does not exist.
    RemoveComment { id: String },
}

/// Finds the map-backed embed sitting exactly at story `index`.
fn embed_at<T: ReadTxn>(story: &yrs::TextRef, txn: &T, index: u32) -> OpResult<MapRef> {
    let mut offset = 0u32;
    for diff in story.diff(txn, YChange::identity) {
        if offset == index {
            if let Out::YMap(map) = diff.insert {
                return Ok(map);
            }
            break;
        }
        offset += out_len(&diff.insert);
        if offset > index {
            break;
        }
    }
    Err(OpError::OutOfBounds {
        index,
        len: story.len(txn),
    })
}

impl EditingDoc {
    /// Applies raw story operations in one transaction.
    pub fn apply_raw_ops(&self, story_id: &str, ops: Vec<RawOp>, ctx: &EditCtx) -> OpResult<()> {
        let mut txn = self.transact_for(ctx);
        apply_raw_ops_to_story(&mut txn, story_id, ops, false)
    }

    pub(crate) fn apply_raw_story_batches(
        &self,
        batches: Vec<(String, Vec<RawOp>)>,
        ctx: &EditCtx,
    ) -> OpResult<()> {
        let mut txn = self.transact_for(ctx);
        for (story_id, ops) in batches {
            apply_raw_ops_to_story(&mut txn, &story_id, ops, true)?;
        }
        Ok(())
    }
}

struct InsertRun {
    deltas: Vec<Delta<In>>,
    deterministic: bool,
    embeds: Vec<InsertedEmbed>,
    formats: BTreeMap<Arc<str>, Vec<FormatSpan>>,
    cursor: Option<u32>,
}

struct InsertedEmbed {
    index: u32,
    kind: String,
    payload: Vec<(String, Any)>,
}

struct FormatSpan {
    index: u32,
    len: u32,
    value: Any,
}

impl InsertRun {
    fn new(deterministic: bool) -> Self {
        Self {
            deltas: Vec::new(),
            deterministic,
            embeds: Vec::new(),
            formats: BTreeMap::new(),
            cursor: None,
        }
    }

    fn is_active(&self) -> bool {
        self.cursor.is_some()
    }

    fn is_contiguous(&self, index: u32) -> bool {
        self.cursor.is_none_or(|cursor| cursor == index)
    }

    fn attrs_match(trailing: &Option<Box<Attrs>>, attrs: &Attrs) -> bool {
        if attrs.is_empty() {
            trailing.is_none()
        } else {
            trailing.as_deref() == Some(attrs)
        }
    }

    fn trailing_text_attrs_match(&self, attrs: &Attrs) -> bool {
        match self.deltas.last() {
            Some(Delta::Inserted(In::Any(Any::String(_)), trailing)) => {
                Self::attrs_match(trailing, attrs)
            }
            _ => true,
        }
    }

    fn flush(&mut self, story: &TextRef, txn: &mut TransactionMut<'_>) -> OpResult<()> {
        if !self.deltas.is_empty() {
            story.apply_delta(txn, std::mem::take(&mut self.deltas));
        }
        if self.deterministic {
            self.hydrate_embeds(story, txn)?;
            self.apply_formats(story, txn);
        }
        self.cursor = None;
        Ok(())
    }

    fn start_delta(&mut self, index: u32) {
        if self.deltas.is_empty() && index > 0 {
            self.deltas.push(Delta::Retain(index, None));
        }
    }

    fn insert(&mut self, index: u32, text: String, attrs: Attrs) {
        let len = text.encode_utf16().count() as u32;
        self.cursor = Some(index + len);
        if len == 0 {
            return;
        }
        self.start_delta(index);
        if !self.deterministic {
            let attrs = (!attrs.is_empty()).then_some(Box::new(attrs));
            if let Some(Delta::Inserted(In::Any(Any::String(trailing)), trailing_attrs)) =
                self.deltas.last_mut()
            {
                if trailing_attrs == &attrs {
                    let mut combined = String::with_capacity(trailing.len() + text.len());
                    combined.push_str(trailing);
                    combined.push_str(&text);
                    *trailing = Arc::from(combined);
                    return;
                }
            }
            self.deltas.push(Delta::Inserted(
                In::Any(Any::String(Arc::from(text))),
                attrs,
            ));
            return;
        }
        self.record_formats(index, len, attrs);
        if let Some(Delta::Inserted(In::Any(Any::String(trailing)), None)) = self.deltas.last_mut()
        {
            let mut combined = String::with_capacity(trailing.len() + text.len());
            combined.push_str(trailing);
            combined.push_str(&text);
            *trailing = Arc::from(combined);
            return;
        }
        self.deltas
            .push(Delta::Inserted(In::Any(Any::String(Arc::from(text))), None));
    }

    fn insert_embed(
        &mut self,
        index: u32,
        kind: String,
        payload: Vec<(String, Any)>,
        attrs: Attrs,
    ) {
        self.cursor = Some(index + 1);
        self.start_delta(index);
        if !self.deterministic {
            let entries = std::iter::once((KIND_KEY.to_owned(), Any::from(kind))).chain(payload);
            self.deltas.push(Delta::Inserted(
                In::Map(MapPrelim::from_iter(entries)),
                (!attrs.is_empty()).then_some(Box::new(attrs)),
            ));
            return;
        }
        self.record_formats(index, 1, attrs);
        self.embeds.push(InsertedEmbed {
            index,
            kind,
            payload,
        });
        self.deltas
            .push(Delta::Inserted(In::Map(MapPrelim::default()), None));
    }

    fn record_formats(&mut self, index: u32, len: u32, attrs: Attrs) {
        for (key, value) in attrs {
            let spans = self.formats.entry(key).or_default();
            if let Some(previous) = spans.last_mut()
                && previous.index + previous.len == index
                && previous.value == value
            {
                previous.len += len;
            } else {
                spans.push(FormatSpan { index, len, value });
            }
        }
    }

    fn hydrate_embeds(&mut self, story: &TextRef, txn: &mut TransactionMut<'_>) -> OpResult<()> {
        if self.embeds.is_empty() {
            return Ok(());
        }
        let embeds = std::mem::take(&mut self.embeds);
        let mut targets = Vec::with_capacity(embeds.len());
        let mut next = 0;
        let mut offset = 0;
        for diff in story.diff(txn, YChange::identity) {
            let len = out_len(&diff.insert);
            if next < embeds.len() && offset == embeds[next].index {
                let Out::YMap(map) = diff.insert else {
                    return Err(OpError::OutOfBounds {
                        index: offset,
                        len: story.len(txn),
                    });
                };
                targets.push(map);
                next += 1;
            }
            offset += len;
        }
        if targets.len() != embeds.len() {
            return Err(OpError::OutOfBounds {
                index: embeds[targets.len()].index,
                len: story.len(txn),
            });
        }
        for (map, embed) in targets.into_iter().zip(embeds) {
            map.insert(txn, KIND_KEY, embed.kind);
            for (key, value) in embed.payload {
                map.insert(txn, key, value);
            }
        }
        Ok(())
    }

    fn apply_formats(&mut self, story: &TextRef, txn: &mut TransactionMut<'_>) {
        for (key, spans) in std::mem::take(&mut self.formats) {
            let mut cursor = 0;
            let mut deltas: Vec<Delta<In>> = Vec::with_capacity(spans.len() * 2);
            for span in spans {
                if span.index > cursor {
                    deltas.push(Delta::Retain(span.index - cursor, None));
                }
                deltas.push(Delta::Retain(
                    span.len,
                    Some(Box::new(Attrs::from([(key.clone(), span.value)]))),
                ));
                cursor = span.index + span.len;
            }
            story.apply_delta(txn, deltas);
        }
    }
}

fn apply_raw_ops_to_story(
    txn: &mut TransactionMut<'_>,
    story_id: &str,
    ops: Vec<RawOp>,
    deterministic: bool,
) -> OpResult<()> {
    let story = story_ref(txn, story_id).map_err(OpError::from)?;
    let mut run = InsertRun::new(deterministic);
    for op in ops {
        match op {
            RawOp::Insert { index, text, attrs }
                if run.is_contiguous(index)
                    && (deterministic
                        || text.is_empty()
                        || run.trailing_text_attrs_match(&attrs))
                    && (run.is_active() || is_utf16_boundary(&story, txn, index)) =>
            {
                if !run.is_active() {
                    guard_index(&story, txn, index)?;
                }
                run.insert(index, text, attrs);
            }
            RawOp::InsertEmbed {
                index,
                kind,
                payload,
                attrs,
            } if run.is_contiguous(index)
                && (deterministic || attrs.is_empty())
                && (run.is_active() || is_utf16_boundary(&story, txn, index)) =>
            {
                if !run.is_active() {
                    guard_index(&story, txn, index)?;
                }
                run.insert_embed(index, kind, payload, attrs);
            }
            op => {
                run.flush(&story, txn)?;
                apply_raw_op_absolute(txn, &story, story_id, op)?;
            }
        }
    }
    run.flush(&story, txn)?;
    Ok(())
}

fn is_utf16_boundary<T: ReadTxn>(story: &TextRef, txn: &T, index: u32) -> bool {
    let mut offset = 0;
    for diff in story.diff(txn, YChange::identity) {
        if index == offset {
            return true;
        }
        if let Out::Any(Any::String(text)) = &diff.insert {
            for ch in text.chars() {
                offset += ch.len_utf16() as u32;
                if index == offset {
                    return true;
                }
                if index < offset {
                    return false;
                }
            }
        } else {
            offset += out_len(&diff.insert);
            if index < offset {
                return false;
            }
        }
    }
    index == offset
}

fn apply_raw_op_absolute(
    txn: &mut TransactionMut<'_>,
    story: &TextRef,
    story_id: &str,
    op: RawOp,
) -> OpResult<()> {
    match op {
        RawOp::Insert { index, text, attrs } => {
            guard_index(story, txn, index)?;
            story.insert_with_attributes(txn, index, &text, attrs);
        }
        RawOp::Delete { index, len } => {
            guard_range(story, txn, index, len)?;
            story.remove_range(txn, index, len);
        }
        RawOp::Format { index, len, attrs } => {
            guard_range(story, txn, index, len)?;
            if len > 0 {
                story.format(txn, index, len, attrs);
            }
        }
        RawOp::InsertEmbed {
            index,
            kind,
            payload,
            attrs,
        } => {
            guard_index(story, txn, index)?;
            let embed = story.insert_embed_with_attributes(txn, index, MapPrelim::default(), attrs);
            embed.insert(txn, KIND_KEY, kind.as_str());
            for (key, value) in payload {
                embed.insert(txn, key, value);
            }
        }
        RawOp::SetEmbedAttr { index, key, value } => {
            let embed = embed_at(story, txn, index)?;
            embed.insert(txn, key, value);
        }
        RawOp::SetComment {
            id,
            ranges,
            author,
            date,
            body,
        } => {
            if ranges.is_empty() {
                return Err(OpError::InvalidComment(
                    "at least one anchored range is required".into(),
                ));
            }
            let mut anchors = Vec::with_capacity(ranges.len());
            for (start, end) in ranges {
                let len = end
                    .checked_sub(start)
                    .filter(|len| *len > 0)
                    .ok_or(OpError::InvalidRange { start, end })?;
                guard_range(story, txn, start, len)?;
                let start_anchor =
                    story
                        .sticky_index(txn, start, Assoc::After)
                        .ok_or_else(|| {
                            OpError::InvalidComment("start anchor could not be made".into())
                        })?;
                let end_anchor = story.sticky_index(txn, end, Assoc::Before).ok_or_else(|| {
                    OpError::InvalidComment("end anchor could not be made".into())
                })?;
                anchors.push(anchor_value(story_id, &start_anchor, &end_anchor));
            }
            let comments = txn
                .get_map(COMMENTS)
                .expect("comments root is declared by EditingDoc::new");
            let comment = comments.insert(txn, id.as_str(), MapPrelim::default());
            comment.insert(txn, "author", author.as_str());
            comment.insert(txn, "date", date.as_str());
            comment.insert(txn, "parentId", Any::Null);
            comment.insert(txn, "done", false);
            comment.insert(txn, "body", body);
            comment.insert(txn, "anchors", Any::Array(Arc::from(anchors)));
        }
        RawOp::RemoveComment { id } => {
            let comments = txn
                .get_map(COMMENTS)
                .expect("comments root is declared by EditingDoc::new");
            if comments.remove(txn, &id).is_none() {
                return Err(OpError::UnknownComment(id));
            }
        }
    }
    Ok(())
}

fn guard_index<T: ReadTxn>(story: &yrs::TextRef, txn: &T, index: u32) -> OpResult<()> {
    if index <= story.len(txn) {
        Ok(())
    } else {
        Err(OpError::OutOfBounds {
            index,
            len: story.len(txn),
        })
    }
}

fn guard_range<T: ReadTxn>(story: &yrs::TextRef, txn: &T, index: u32, len: u32) -> OpResult<()> {
    let story_len = story.len(txn);
    if index.checked_add(len).is_some_and(|end| end <= story_len) {
        Ok(())
    } else {
        Err(OpError::OutOfBounds {
            index: index.saturating_add(len),
            len: story_len,
        })
    }
}

#[cfg(test)]
mod tests {
    use std::sync::Arc;

    use super::*;
    use crate::canonical::project_story;
    use crate::{EditCtx, EditingDoc, PILCROW_KIND, seed_from_docx};

    fn attrs(pairs: &[(&str, Any)]) -> Attrs {
        pairs
            .iter()
            .map(|(key, value)| (Arc::from(*key), value.clone()))
            .collect()
    }

    fn apply_raw_ops_legacy_to_doc(
        doc: &EditingDoc,
        story_id: &str,
        ops: Vec<RawOp>,
        ctx: &EditCtx,
    ) -> OpResult<()> {
        let mut txn = doc.transact_for(ctx);
        apply_raw_ops_legacy(&mut txn, story_id, ops)
    }

    /// Legacy absolute-index oracle.
    fn apply_raw_ops_legacy(
        txn: &mut TransactionMut<'_>,
        story_id: &str,
        ops: Vec<RawOp>,
    ) -> OpResult<()> {
        let story = story_ref(txn, story_id).map_err(OpError::from)?;
        let guard_index = |txn: &TransactionMut<'_>, index: u32| -> OpResult<()> {
            if index <= story.len(txn) {
                Ok(())
            } else {
                Err(OpError::OutOfBounds {
                    index,
                    len: story.len(txn),
                })
            }
        };
        let guard_range = |txn: &TransactionMut<'_>, index: u32, len: u32| -> OpResult<()> {
            let story_len = story.len(txn);
            if index.checked_add(len).is_some_and(|end| end <= story_len) {
                Ok(())
            } else {
                Err(OpError::OutOfBounds {
                    index: index.saturating_add(len),
                    len: story_len,
                })
            }
        };
        for op in ops {
            match op {
                RawOp::Insert { index, text, attrs } => {
                    guard_index(txn, index)?;
                    story.insert_with_attributes(txn, index, &text, attrs);
                }
                RawOp::Delete { index, len } => {
                    guard_range(txn, index, len)?;
                    story.remove_range(txn, index, len);
                }
                RawOp::Format { index, len, attrs } => {
                    guard_range(txn, index, len)?;
                    if len > 0 {
                        story.format(txn, index, len, attrs);
                    }
                }
                RawOp::InsertEmbed {
                    index,
                    kind,
                    payload,
                    attrs,
                } => {
                    guard_index(txn, index)?;
                    let embed =
                        story.insert_embed_with_attributes(txn, index, MapPrelim::default(), attrs);
                    embed.insert(txn, KIND_KEY, kind.as_str());
                    for (key, value) in payload {
                        embed.insert(txn, key, value);
                    }
                }
                RawOp::SetEmbedAttr { index, key, value } => {
                    let embed = embed_at(&story, txn, index)?;
                    embed.insert(txn, key, value);
                }
                RawOp::SetComment {
                    id,
                    ranges,
                    author,
                    date,
                    body,
                } => {
                    if ranges.is_empty() {
                        return Err(OpError::InvalidComment(
                            "at least one anchored range is required".into(),
                        ));
                    }
                    let mut anchors = Vec::with_capacity(ranges.len());
                    for (start, end) in ranges {
                        let len = end
                            .checked_sub(start)
                            .filter(|len| *len > 0)
                            .ok_or(OpError::InvalidRange { start, end })?;
                        guard_range(txn, start, len)?;
                        let start_anchor = story
                            .sticky_index(txn, start, Assoc::After)
                            .ok_or_else(|| {
                                OpError::InvalidComment("start anchor could not be made".into())
                            })?;
                        let end_anchor =
                            story.sticky_index(txn, end, Assoc::Before).ok_or_else(|| {
                                OpError::InvalidComment("end anchor could not be made".into())
                            })?;
                        anchors.push(anchor_value(story_id, &start_anchor, &end_anchor));
                    }
                    let comments = txn
                        .get_map(COMMENTS)
                        .expect("comments root is declared by EditingDoc::new");
                    let comment = comments.insert(txn, id.as_str(), MapPrelim::default());
                    comment.insert(txn, "author", author.as_str());
                    comment.insert(txn, "date", date.as_str());
                    comment.insert(txn, "parentId", Any::Null);
                    comment.insert(txn, "done", false);
                    comment.insert(txn, "body", body);
                    comment.insert(txn, "anchors", Any::Array(Arc::from(anchors)));
                }
                RawOp::RemoveComment { id } => {
                    let comments = txn
                        .get_map(COMMENTS)
                        .expect("comments root is declared by EditingDoc::new");
                    if comments.remove(txn, &id).is_none() {
                        return Err(OpError::UnknownComment(id));
                    }
                }
            }
        }
        Ok(())
    }

    fn assert_stories_equivalent(left: &EditingDoc, right: &EditingDoc, story_id: &str) {
        assert_eq!(
            project_story(left, story_id).unwrap(),
            project_story(right, story_id).unwrap()
        );
        assert_eq!(
            left.story_segments(story_id).unwrap(),
            right.story_segments(story_id).unwrap()
        );
    }

    fn pilcrow_payload(para_id: &str) -> Vec<(String, Any)> {
        vec![
            ("paraId".into(), Any::from(para_id)),
            ("pStyle".into(), Any::from("Normal")),
            ("alignment".into(), Any::from("left")),
        ]
    }

    fn assert_raw_ops_equivalent(
        initial_text: &str,
        ops: Vec<RawOp>,
    ) -> (EditingDoc, EditingDoc, OpResult<()>) {
        let ctx = EditCtx::local(String::new(), String::new());
        let legacy = EditingDoc::new(7);
        legacy
            .create_story("body", initial_text, "Normal", "left")
            .unwrap();
        let legacy_result = apply_raw_ops_legacy_to_doc(&legacy, "body", ops.clone(), &ctx);

        let batched = EditingDoc::new(7);
        batched
            .create_story("body", initial_text, "Normal", "left")
            .unwrap();
        let batched_result = batched.apply_raw_ops("body", ops.clone(), &ctx);

        assert_eq!(legacy_result, batched_result, "ops: {ops:?}");
        assert_stories_equivalent(&legacy, &batched, "body");
        (legacy, batched, batched_result)
    }

    fn exhaustive_op_variants() -> Vec<RawOp> {
        let mut ops = Vec::new();
        for index in [0, 1, 4] {
            ops.push(RawOp::Insert {
                index,
                text: "x".into(),
                attrs: Attrs::new(),
            });
            ops.push(RawOp::Insert {
                index,
                text: "𝔘".into(),
                attrs: attrs(&[("bold", Any::Bool(true))]),
            });
            ops.push(RawOp::Delete { index, len: 1 });
            ops.push(RawOp::Format {
                index,
                len: 1,
                attrs: Attrs::new(),
            });
            ops.push(RawOp::Format {
                index,
                len: 1,
                attrs: attrs(&[("italic", Any::Bool(true))]),
            });
            ops.push(RawOp::InsertEmbed {
                index,
                kind: "break".into(),
                payload: Vec::new(),
                attrs: Attrs::new(),
            });
        }
        ops
    }

    /// What `seed_parsed_docx` emits: placeholder delete, appending inserts, trailing comment.
    fn seed_shaped_ops() -> Vec<RawOp> {
        let font = attrs(&[(
            "fontFamily",
            Any::from_json(r#"{"ascii":"Calibri","hAnsi":"Calibri"}"#).unwrap(),
        )]);
        let mut bold = font.clone();
        bold.insert("bold".into(), Any::Bool(true));
        struct Builder {
            ops: Vec<RawOp>,
            index: u32,
        }
        impl Builder {
            fn text(&mut self, value: &str, attrs: &Attrs) {
                self.ops.push(RawOp::Insert {
                    index: self.index,
                    text: value.into(),
                    attrs: attrs.clone(),
                });
                self.index += value.encode_utf16().count() as u32;
            }
            fn embed(&mut self, kind: &str, payload: Vec<(String, Any)>) {
                self.ops.push(RawOp::InsertEmbed {
                    index: self.index,
                    kind: kind.into(),
                    payload,
                    attrs: Attrs::new(),
                });
                self.index += 1;
            }
        }
        let mut builder = Builder {
            ops: vec![RawOp::Delete { index: 0, len: 1 }],
            index: 0,
        };
        builder.text("Heading with stars ★★", &bold);
        builder.embed(PILCROW_KIND, pilcrow_payload("7:p0"));
        builder.text("Body ", &font);
        builder.text("bold span", &bold);
        builder.text(" and surrogate pairs 𝔘𝔫𝔦", &font);
        builder.embed("break", Vec::new());
        builder.text("after the break", &font);
        builder.embed(PILCROW_KIND, pilcrow_payload("7:p1"));
        builder.embed(PILCROW_KIND, pilcrow_payload("7:p2"));
        builder.text("tail paragraph", &font);
        builder.embed(PILCROW_KIND, pilcrow_payload("7:p3"));
        let mut ops = builder.ops;
        ops.push(RawOp::SetComment {
            id: "9".into(),
            ranges: vec![(22, 30), (40, 45)],
            author: String::new(),
            date: String::new(),
            body: Any::Null,
        });
        ops
    }

    #[test]
    fn exhaustive_three_op_vectors_match_legacy() {
        let variants = exhaustive_op_variants();
        let ctx = EditCtx::local(String::new(), String::new());
        let mut vector_count = 0;
        for first in &variants {
            for second in &variants {
                for third in &variants {
                    vector_count += 1;
                    let ops = vec![first.clone(), second.clone(), third.clone()];
                    let comment = vec![RawOp::SetComment {
                        id: "anchor".into(),
                        ranges: vec![(1, 3)],
                        author: String::new(),
                        date: String::new(),
                        body: Any::Null,
                    }];

                    let legacy = EditingDoc::new(7);
                    legacy
                        .create_story("body", "A𝔘B", "Normal", "left")
                        .unwrap();
                    apply_raw_ops_legacy_to_doc(&legacy, "body", comment.clone(), &ctx).unwrap();
                    let legacy_result =
                        apply_raw_ops_legacy_to_doc(&legacy, "body", ops.clone(), &ctx);

                    let batched = EditingDoc::new(7);
                    batched
                        .create_story("body", "A𝔘B", "Normal", "left")
                        .unwrap();
                    apply_raw_ops_legacy_to_doc(&batched, "body", comment.clone(), &ctx).unwrap();
                    let batched_result = batched.apply_raw_ops("body", ops.clone(), &ctx);

                    assert_eq!(legacy_result, batched_result, "ops: {ops:?}");
                    assert_eq!(
                        project_story(&legacy, "body").unwrap(),
                        project_story(&batched, "body").unwrap(),
                        "ops: {ops:?}"
                    );
                    assert_eq!(
                        legacy.story_segments("body").unwrap(),
                        batched.story_segments("body").unwrap(),
                        "ops: {ops:?}"
                    );
                    assert_eq!(
                        legacy.resolve_comment("anchor"),
                        batched.resolve_comment("anchor"),
                        "ops: {ops:?}"
                    );
                }
            }
        }
        assert_eq!(vector_count, 5_832);
    }

    #[test]
    fn formats_and_deletes_do_not_leak_into_inserts() {
        let (_, doc, result) = assert_raw_ops_equivalent(
            "AB",
            vec![
                RawOp::Format {
                    index: 0,
                    len: 2,
                    attrs: attrs(&[("bold", Any::Bool(true))]),
                },
                RawOp::Delete { index: 1, len: 1 },
                RawOp::Insert {
                    index: 1,
                    text: "Y".into(),
                    attrs: attrs(&[("italic", Any::Bool(true))]),
                },
            ],
        );
        result.unwrap();
        let segments = doc.story_segments("body").unwrap();
        let y = segments
            .iter()
            .find(|segment| matches!(&segment.content, crate::SegmentContent::Text(text) if text == "Y"))
            .unwrap();
        assert_eq!(y.attributes.get("italic"), Some(&Any::Bool(true)));
        assert!(!y.attributes.contains_key("bold"));

        let (_, doc, result) = assert_raw_ops_equivalent(
            "ABC",
            vec![
                RawOp::Format {
                    index: 1,
                    len: 2,
                    attrs: attrs(&[("italic", Any::Bool(true))]),
                },
                RawOp::Delete { index: 0, len: 1 },
                RawOp::Delete { index: 0, len: 1 },
            ],
        );
        result.unwrap();
        let segments = doc.story_segments("body").unwrap();
        let c = segments
            .iter()
            .find(|segment| matches!(&segment.content, crate::SegmentContent::Text(text) if text == "C"))
            .unwrap();
        assert_eq!(c.attributes.get("italic"), Some(&Any::Bool(true)));
    }

    #[test]
    fn contiguous_indices_count_utf16_units() {
        let (_, doc, result) = assert_raw_ops_equivalent(
            "ABC",
            vec![
                RawOp::Insert {
                    index: 2,
                    text: "𝔘".into(),
                    attrs: Attrs::new(),
                },
                RawOp::Format {
                    index: 1,
                    len: 2,
                    attrs: attrs(&[("bold", Any::Bool(true))]),
                },
                RawOp::Insert {
                    index: 4,
                    text: "Z".into(),
                    attrs: Attrs::new(),
                },
            ],
        );
        result.unwrap();
        assert_eq!(doc.paragraphs("body").unwrap()[0].text, "AB𝔘ZC");
    }

    #[test]
    fn insert_run_starting_inside_a_surrogate_matches_legacy() {
        let (_, _, result) = assert_raw_ops_equivalent(
            "A𝔘B",
            vec![
                RawOp::Format {
                    index: 0,
                    len: 1,
                    attrs: Attrs::new(),
                },
                RawOp::Insert {
                    index: 2,
                    text: "𝔘".into(),
                    attrs: attrs(&[("bold", Any::Bool(true))]),
                },
                RawOp::Insert {
                    index: 4,
                    text: "x".into(),
                    attrs: Attrs::new(),
                },
            ],
        );
        result.unwrap();
    }

    #[test]
    fn surrogate_delete_flushes_before_live_bounds_check() {
        let (_, doc, result) = assert_raw_ops_equivalent(
            "A𝔘B",
            vec![
                RawOp::Delete { index: 1, len: 1 },
                RawOp::Insert {
                    index: 4,
                    text: "X".into(),
                    attrs: Attrs::new(),
                },
            ],
        );
        assert_eq!(result, Err(OpError::OutOfBounds { index: 4, len: 3 }));
        assert_eq!(doc.paragraphs("body").unwrap()[0].text, "AB");
        assert_eq!(doc.story_len("body").unwrap(), 3);
    }

    #[test]
    fn insert_runs_preserve_maximal_text_segments() {
        let bold = attrs(&[("bold", Any::Bool(true))]);
        let (_, doc, result) = assert_raw_ops_equivalent(
            "",
            vec![
                RawOp::Insert {
                    index: 0,
                    text: "X".into(),
                    attrs: bold.clone(),
                },
                RawOp::Insert {
                    index: 1,
                    text: "Y".into(),
                    attrs: bold.clone(),
                },
                RawOp::Insert {
                    index: 2,
                    text: "Z".into(),
                    attrs: bold,
                },
            ],
        );
        result.unwrap();
        let text = doc
            .story_segments("body")
            .unwrap()
            .into_iter()
            .filter_map(|segment| match segment.content {
                crate::SegmentContent::Text(text) => Some(text),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(text, ["XYZ"]);

        let (_, doc, result) = assert_raw_ops_equivalent(
            "",
            vec![
                RawOp::Insert {
                    index: 0,
                    text: "AB".into(),
                    attrs: attrs(&[("bold", Any::Bool(true))]),
                },
                RawOp::Insert {
                    index: 2,
                    text: "CDE".into(),
                    attrs: attrs(&[("italic", Any::Bool(true))]),
                },
                RawOp::Insert {
                    index: 2,
                    text: "x".into(),
                    attrs: Attrs::new(),
                },
                RawOp::Delete { index: 2, len: 1 },
            ],
        );
        result.unwrap();
        let text = doc
            .story_segments("body")
            .unwrap()
            .into_iter()
            .filter_map(|segment| match segment.content {
                crate::SegmentContent::Text(text) => Some(text),
                _ => None,
            })
            .collect::<Vec<_>>();
        assert_eq!(text, ["AB", "CDE"]);
    }

    #[test]
    fn insert_batch_matches_legacy_for_seed_shaped_ops() {
        let ctx = EditCtx::local(String::new(), String::new());
        let legacy = EditingDoc::new(7);
        legacy.create_empty_stories(&["body".into()]).unwrap();
        apply_raw_ops_legacy_to_doc(&legacy, "body", seed_shaped_ops(), &ctx).unwrap();

        let batched = EditingDoc::new(7);
        batched.create_empty_stories(&["body".into()]).unwrap();
        batched
            .apply_raw_ops("body", seed_shaped_ops(), &ctx)
            .unwrap();

        assert_stories_equivalent(&legacy, &batched, "body");
        assert_eq!(
            legacy.paragraphs("body").unwrap().len(),
            batched.paragraphs("body").unwrap().len()
        );
        let legacy_anchor = legacy.resolve_comment("9").unwrap();
        let batched_anchor = batched.resolve_comment("9").unwrap();
        assert_eq!(
            legacy_anchor
                .iter()
                .map(|anchor| (anchor.start, anchor.end))
                .collect::<Vec<_>>(),
            batched_anchor
                .iter()
                .map(|anchor| (anchor.start, anchor.end))
                .collect::<Vec<_>>()
        );
    }

    #[test]
    fn fallback_path_matches_legacy_for_backward_jumps() {
        let ops = vec![
            RawOp::Insert {
                index: 4,
                text: " tail".into(),
                attrs: Attrs::new(),
            },
            RawOp::Insert {
                index: 2,
                text: "XY".into(),
                attrs: attrs(&[("bold", Any::Bool(true))]),
            },
            RawOp::Format {
                index: 1,
                len: 5,
                attrs: attrs(&[("italic", Any::Bool(true))]),
            },
            RawOp::Delete { index: 3, len: 2 },
            RawOp::InsertEmbed {
                index: 0,
                kind: "break".into(),
                payload: Vec::new(),
                attrs: Attrs::new(),
            },
            RawOp::Insert {
                index: 10,
                text: "end".into(),
                attrs: Attrs::new(),
            },
            RawOp::SetEmbedAttr {
                index: 0,
                key: "cleared".into(),
                value: Any::Bool(true),
            },
        ];
        let ctx = EditCtx::local(String::new(), String::new());
        let legacy = EditingDoc::new(7);
        legacy
            .create_story("body", "seed", "Normal", "left")
            .unwrap();
        apply_raw_ops_legacy_to_doc(&legacy, "body", ops.clone(), &ctx).unwrap();

        let batched = EditingDoc::new(7);
        batched
            .create_story("body", "seed", "Normal", "left")
            .unwrap();
        batched.apply_raw_ops("body", ops, &ctx).unwrap();

        assert_stories_equivalent(&legacy, &batched, "body");
    }

    #[test]
    fn a_failing_op_leaves_the_same_applied_prefix_as_legacy() {
        let ops = vec![
            RawOp::Insert {
                index: 2,
                text: "one".into(),
                attrs: Attrs::new(),
            },
            RawOp::Insert {
                index: 0,
                text: "two".into(),
                attrs: attrs(&[("bold", Any::Bool(true))]),
            },
            RawOp::Delete { index: 90, len: 5 },
            RawOp::Insert {
                index: 0,
                text: "never applied".into(),
                attrs: Attrs::new(),
            },
        ];
        let ctx = EditCtx::local(String::new(), String::new());
        let legacy = EditingDoc::new(7);
        legacy.create_story("body", "AB", "Normal", "left").unwrap();
        let legacy_err =
            apply_raw_ops_legacy_to_doc(&legacy, "body", ops.clone(), &ctx).unwrap_err();

        let batched = EditingDoc::new(7);
        batched
            .create_story("body", "AB", "Normal", "left")
            .unwrap();
        let batched_err = batched.apply_raw_ops("body", ops, &ctx).unwrap_err();

        assert_eq!(legacy_err, batched_err);
        assert_stories_equivalent(&legacy, &batched, "body");
    }

    #[test]
    fn every_fallback_error_keeps_the_legacy_prefix() {
        let failures = vec![
            RawOp::Insert {
                index: 99,
                text: "x".into(),
                attrs: Attrs::new(),
            },
            RawOp::Delete { index: 99, len: 1 },
            RawOp::Format {
                index: 99,
                len: 1,
                attrs: Attrs::new(),
            },
            RawOp::InsertEmbed {
                index: 99,
                kind: "break".into(),
                payload: Vec::new(),
                attrs: Attrs::new(),
            },
            RawOp::SetEmbedAttr {
                index: 0,
                key: "value".into(),
                value: Any::Bool(true),
            },
            RawOp::SetComment {
                id: "comment".into(),
                ranges: vec![(0, 99)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            },
            RawOp::RemoveComment {
                id: "missing".into(),
            },
        ];
        for failure in failures {
            let (_, doc, result) = assert_raw_ops_equivalent(
                "AB",
                vec![
                    RawOp::Insert {
                        index: 0,
                        text: "prefix".into(),
                        attrs: Attrs::new(),
                    },
                    failure,
                    RawOp::Insert {
                        index: 0,
                        text: "unreachable".into(),
                        attrs: Attrs::new(),
                    },
                ],
            );
            assert!(result.is_err());
            assert_eq!(doc.paragraphs("body").unwrap()[0].text, "prefixAB");
        }
    }

    #[test]
    fn comment_and_embed_fallbacks_match_legacy() {
        let ops = vec![
            RawOp::InsertEmbed {
                index: 0,
                kind: "break".into(),
                payload: Vec::new(),
                attrs: Attrs::new(),
            },
            RawOp::Insert {
                index: 1,
                text: "x".into(),
                attrs: Attrs::new(),
            },
            RawOp::SetEmbedAttr {
                index: 0,
                key: "value".into(),
                value: Any::Bool(true),
            },
            RawOp::SetComment {
                id: "comment".into(),
                ranges: vec![(1, 2)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            },
            RawOp::RemoveComment {
                id: "comment".into(),
            },
        ];
        let (_, doc, result) = assert_raw_ops_equivalent("AB", ops);
        result.unwrap();
        assert!(doc.resolve_comment("comment").is_err());
    }

    #[test]
    fn deleting_an_attributed_embed_preserves_its_neighbor_attributes() {
        let ops = vec![
            RawOp::InsertEmbed {
                index: 0,
                kind: "break".into(),
                payload: vec![("audit".into(), Any::from(0_i64))],
                attrs: attrs(&[("bold", Any::Bool(true))]),
            },
            RawOp::InsertEmbed {
                index: 1,
                kind: "break".into(),
                payload: vec![("audit".into(), Any::from(1_i64))],
                attrs: attrs(&[("bold", Any::Bool(true)), ("italic", Any::Bool(true))]),
            },
            RawOp::Insert {
                index: 0,
                text: "yz".into(),
                attrs: attrs(&[("bold", Any::Bool(true)), ("italic", Any::Bool(true))]),
            },
            RawOp::Delete { index: 2, len: 1 },
        ];
        let (_, doc, result) = assert_raw_ops_equivalent("A𝔘BC", ops);
        result.unwrap();
        let segments = doc.story_segments("body").unwrap();
        assert_eq!(segments[1].attributes.get("bold"), Some(&Any::Bool(true)));
    }

    #[test]
    fn formatting_adjacent_insertions_to_equal_attributes_preserves_layout() {
        let ops = vec![
            RawOp::Insert {
                index: 0,
                text: "yz".into(),
                attrs: attrs(&[("bold", Any::Bool(true)), ("italic", Any::Bool(true))]),
            },
            RawOp::Insert {
                index: 2,
                text: "x".into(),
                attrs: attrs(&[("italic", Any::Bool(true))]),
            },
            RawOp::Format {
                index: 2,
                len: 1,
                attrs: attrs(&[("bold", Any::Bool(true))]),
            },
        ];
        let (_, doc, result) = assert_raw_ops_equivalent("", ops);
        result.unwrap();
        assert!(matches!(
            &doc.story_segments("body").unwrap()[0].content,
            crate::SegmentContent::Text(text) if text == "yzx"
        ));
    }

    #[test]
    fn deleting_an_embed_between_differently_formatted_text_preserves_layout() {
        let ops = vec![
            RawOp::Insert {
                index: 0,
                text: "A".into(),
                attrs: attrs(&[("bold", Any::Bool(true)), ("italic", Any::Bool(true))]),
            },
            RawOp::InsertEmbed {
                index: 1,
                kind: "break".into(),
                payload: Vec::new(),
                attrs: Attrs::new(),
            },
            RawOp::Insert {
                index: 2,
                text: "B".into(),
                attrs: attrs(&[("italic", Any::Bool(true))]),
            },
            RawOp::Delete { index: 1, len: 1 },
            RawOp::Format {
                index: 1,
                len: 1,
                attrs: attrs(&[("bold", Any::Bool(true))]),
            },
        ];
        let (_, doc, result) = assert_raw_ops_equivalent("", ops);
        result.unwrap();
        assert!(matches!(
            &doc.story_segments("body").unwrap()[0].content,
            crate::SegmentContent::Text(text) if text == "AB"
        ));
    }

    #[test]
    fn attributed_embed_after_text_matches_the_absolute_layout() {
        let ops = vec![
            RawOp::Insert {
                index: 6,
                text: "x".into(),
                attrs: attrs(&[("italic", Any::Bool(true))]),
            },
            RawOp::InsertEmbed {
                index: 7,
                kind: "break".into(),
                payload: vec![("audit".into(), Any::from(0_i64))],
                attrs: attrs(&[("bold", Any::Bool(true)), ("italic", Any::Bool(true))]),
            },
            RawOp::Delete { index: 7, len: 0 },
        ];
        let (_, _, result) = assert_raw_ops_equivalent("A𝔘BC", ops);
        result.unwrap();
    }

    #[test]
    fn legacy_seeded_state_loads_and_merges_under_the_batched_applier() {
        let ctx = EditCtx::local(String::new(), String::new());
        let original = EditingDoc::new(1);
        original.create_empty_stories(&["body".into()]).unwrap();
        apply_raw_ops_legacy_to_doc(&original, "body", seed_shaped_ops(), &ctx).unwrap();

        let restored = EditingDoc::new(2);
        restored
            .apply_update_v1(&original.encode_state_as_update_v1())
            .unwrap();
        assert_stories_equivalent(&original, &restored, "body");

        restored
            .apply_raw_ops(
                "body",
                vec![RawOp::Insert {
                    index: 5,
                    text: "restored edit ".into(),
                    attrs: attrs(&[("italic", Any::Bool(true))]),
                }],
                &ctx,
            )
            .unwrap();
        original
            .apply_update_v1(&restored.encode_state_as_update_v1())
            .unwrap();
        restored
            .apply_update_v1(&original.encode_state_as_update_v1())
            .unwrap();
        assert_stories_equivalent(&original, &restored, "body");
        assert!(
            original
                .story_segments("body")
                .unwrap()
                .iter()
                .any(|segment| matches!(
                    &segment.content,
                    crate::SegmentContent::Text(text) if text.contains("restored edit")
                ))
        );
    }

    #[test]
    fn seeded_docx_state_round_trips_through_encode_and_load() {
        let bytes = include_bytes!("../tests/fixtures/footnote-anchor.docx");
        let seeded = EditingDoc::new(1);
        seed_from_docx(&seeded, bytes).unwrap();

        let loaded = EditingDoc::new(2);
        loaded
            .apply_update_v1(&seeded.encode_state_as_update_v1())
            .unwrap();
        let txn = yrs::Transact::transact(seeded.yrs_doc());
        let stories = txn.get_map(crate::STORIES).unwrap();
        let story_ids: Vec<String> = stories.keys(&txn).map(|key| key.to_string()).collect();
        drop(txn);
        assert!(!story_ids.is_empty());
        for story_id in story_ids {
            assert_stories_equivalent(&seeded, &loaded, &story_id);
        }
    }

    #[test]
    fn insert_format_and_delete_apply_in_one_batch() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        // story: A(0) B(1) pilcrow(2)
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![
                RawOp::Insert {
                    index: 1,
                    text: "X".into(),
                    attrs: attrs(&[("bold", Any::Bool(true))]),
                },
                // story: A(0) X(1) B(2) pilcrow(3)
                RawOp::Format {
                    index: 0,
                    len: 1,
                    attrs: attrs(&[("italic", Any::Bool(true))]),
                },
                RawOp::Delete { index: 2, len: 1 }, // remove B
                                                    // story: A(0) X(1) pilcrow(2)
            ],
            &ctx,
        )
        .unwrap();

        let paras = doc.paragraphs("body").unwrap();
        assert_eq!(paras.len(), 1);
        assert_eq!(paras[0].text, "AX");
    }

    #[test]
    fn insert_embed_pilcrow_splits_a_paragraph() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![RawOp::InsertEmbed {
                index: 1,
                kind: PILCROW_KIND.to_owned(),
                payload: vec![
                    ("paraId".into(), Any::from("7:seed")),
                    ("pStyle".into(), Any::from("Normal")),
                    ("alignment".into(), Any::from("left")),
                ],
                attrs: Attrs::new(),
            }],
            &ctx,
        )
        .unwrap();

        let paras = doc.paragraphs("body").unwrap();
        assert_eq!(paras.len(), 2);
        assert_eq!(paras[0].text, "A");
        assert_eq!(paras[1].text, "B");
    }

    #[test]
    fn set_embed_attr_updates_a_pilcrow_property() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        // The pilcrow sits at index 2 (after "AB").
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetEmbedAttr {
                index: 2,
                key: "alignment".into(),
                value: Any::from("center"),
            }],
            &ctx,
        )
        .unwrap();

        let paras = doc.paragraphs("body").unwrap();
        assert_eq!(
            paras[0].properties.get("alignment"),
            Some(&Any::from("center"))
        );
    }

    #[test]
    fn set_comment_upserts_and_remove_comment_deletes() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "ABCDE", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "9".into(),
                ranges: vec![(1, 3)],
                author: "Ada".into(),
                date: "2026-07-13T12:00:00Z".into(),
                body: Any::from("hi"),
            }],
            &ctx,
        )
        .unwrap();
        let resolved = doc.resolve_comment("9").unwrap();
        assert_eq!((resolved[0].start, resolved[0].end), (1, 3));

        // Sticky anchors ride an insertion before the range.
        doc.apply_raw_ops(
            "body",
            vec![RawOp::Insert {
                index: 0,
                text: "xx".into(),
                attrs: Attrs::new(),
            }],
            &ctx,
        )
        .unwrap();
        let resolved = doc.resolve_comment("9").unwrap();
        assert_eq!((resolved[0].start, resolved[0].end), (3, 5));

        // Upsert on the same key re-anchors.
        doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "9".into(),
                ranges: vec![(0, 2), (4, 6)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        )
        .unwrap();
        let resolved = doc.resolve_comment("9").unwrap();
        assert_eq!(resolved.len(), 2);
        assert_eq!((resolved[0].start, resolved[0].end), (0, 2));
        assert_eq!((resolved[1].start, resolved[1].end), (4, 6));

        doc.apply_raw_ops("body", vec![RawOp::RemoveComment { id: "9".into() }], &ctx)
            .unwrap();
        assert!(doc.resolve_comment("9").is_err());
        // Removing an unknown comment is a typed error, not silence.
        let missing =
            doc.apply_raw_ops("body", vec![RawOp::RemoveComment { id: "9".into() }], &ctx);
        assert!(matches!(missing, Err(OpError::UnknownComment(_))));
    }

    #[test]
    fn set_comment_rejects_empty_and_out_of_bounds_ranges() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        let empty = doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "1".into(),
                ranges: vec![(1, 1)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        );
        assert!(matches!(empty, Err(OpError::InvalidRange { .. })));
        let outside = doc.apply_raw_ops(
            "body",
            vec![RawOp::SetComment {
                id: "1".into(),
                ranges: vec![(0, 99)],
                author: String::new(),
                date: String::new(),
                body: Any::Null,
            }],
            &ctx,
        );
        assert!(matches!(outside, Err(OpError::OutOfBounds { .. })));
    }

    #[test]
    fn out_of_bounds_index_is_a_typed_error() {
        let doc = EditingDoc::new(7);
        doc.create_story("body", "AB", "Normal", "left").unwrap();
        let ctx = EditCtx::local(String::new(), String::new());
        let result = doc.apply_raw_ops("body", vec![RawOp::Delete { index: 99, len: 1 }], &ctx);
        assert!(matches!(result, Err(OpError::OutOfBounds { .. })));
    }
}
