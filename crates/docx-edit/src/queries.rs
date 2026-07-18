//! Read-query surface (op-contract §3): pure reads over a transaction snapshot.
//!
//! [`TextView::Vanilla`] is the agent's view (the `paraText.ts` vanilla contract): it EXCLUDES
//! `ins`-attributed text (pending insertions are "not in the document yet"), INCLUDES
//! `del`-attributed text (still there until accepted), and makes comment anchors, tabs, breaks,
//! and every other atom invisible — no character substitution. [`TextView::Raw`] maps every
//! story unit 1:1 (tab = `\t`, hard break = `\n`, other embeds = U+FFFC) so raw offsets equal
//! story indices.

use std::collections::{HashMap, HashSet};

use unicode_segmentation::UnicodeSegmentation;
use yrs::{Any, Map, Out, ReadTxn, TextRef, Transact};

use crate::op::{Loc, LocRange, OpError, OpResult, global_of_loc, loc_of_global};
use crate::ops::table::{TableRowChangeKind, table_row_changes};
use crate::ops::{ChunkKind, snapshot};
use crate::{
    BREAK_KIND, COMMENTS, DEL, EditingDoc, INS, KIND_KEY, PARA_ID, ParagraphId, RevisionId,
    map_string, story_ref,
};

/// Which projection of the story text a read query uses.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq)]
pub enum TextView {
    /// Every story unit, 1:1 with story indices.
    Raw,
    /// The agent's view: pending insertions excluded, pending deletions included, atoms and
    /// comment anchors invisible.
    #[default]
    Vanilla,
}

/// One contiguous run of view text mapped back to raw story indices. Within a span, view and
/// raw offsets advance in lockstep (UTF-16 units).
struct ViewSpan {
    view_start: u32,
    raw_start: u32,
    len: u32,
}

/// One paragraph's view text plus the mapping back to story indices.
pub(crate) struct ParaView {
    pub para_id: ParagraphId,
    pub style_id: Option<String>,
    pub text: String,
    spans: Vec<ViewSpan>,
    /// Story-global index of the paragraph's first unit.
    pub raw_start: u32,
    /// Story-global index of the paragraph's pilcrow.
    pub pilcrow: u32,
}

impl ParaView {
    /// Maps a view offset (UTF-16) to a raw story index. `end_bias` resolves offsets that sit
    /// exactly on a span end to the end of that span rather than the start of the next.
    fn raw_of_view(&self, offset: u32, end_bias: bool) -> u32 {
        for span in &self.spans {
            let span_end = span.view_start + span.len;
            if offset < span_end || (end_bias && offset == span_end) {
                if offset < span.view_start {
                    return span.raw_start;
                }
                return span.raw_start + (offset - span.view_start);
            }
        }
        self.pilcrow
    }

    /// Appends the view text overlapping the raw interval `[from, to)` to `out`.
    fn view_slice_of_raw(&self, from: u32, to: u32, out: &mut String) {
        for span in &self.spans {
            let overlap_start = span.raw_start.max(from);
            let overlap_end = (span.raw_start + span.len).min(to);
            if overlap_end <= overlap_start {
                continue;
            }
            let view_from = span.view_start + (overlap_start - span.raw_start);
            let view_to = span.view_start + (overlap_end - span.raw_start);
            push_utf16_slice(&self.text, view_from, view_to, out);
        }
    }
}

fn push_utf16_slice(text: &str, from: u32, to: u32, out: &mut String) {
    let mut offset = 0u32;
    for ch in text.chars() {
        let width = ch.len_utf16() as u32;
        if offset >= to {
            break;
        }
        if offset >= from {
            out.push(ch);
        }
        offset += width;
    }
}

fn utf16_of_byte(text: &str, byte: usize) -> u32 {
    text[..byte].encode_utf16().count() as u32
}

pub(crate) fn para_views<T: ReadTxn>(story: &TextRef, txn: &T, view: TextView) -> Vec<ParaView> {
    let mut views = Vec::new();
    let mut text = String::new();
    let mut spans: Vec<ViewSpan> = Vec::new();
    let mut view_len = 0u32;
    let mut para_start = 0u32;

    let push_unit =
        |text: &mut String, spans: &mut Vec<ViewSpan>, view_len: &mut u32, ch: char, raw: u32| {
            let width = ch.len_utf16() as u32;
            match spans.last_mut() {
                Some(span)
                    if span.view_start + span.len == *view_len
                        && span.raw_start + span.len == raw =>
                {
                    span.len += width;
                }
                _ => spans.push(ViewSpan {
                    view_start: *view_len,
                    raw_start: raw,
                    len: width,
                }),
            }
            text.push(ch);
            *view_len += width;
        };

    for chunk in snapshot(story, txn) {
        match &chunk.kind {
            ChunkKind::Pilcrow(map) => {
                views.push(ParaView {
                    para_id: map_string(map, txn, PARA_ID).unwrap_or_default(),
                    style_id: map_string(map, txn, "pStyle"),
                    text: std::mem::take(&mut text),
                    spans: std::mem::take(&mut spans),
                    raw_start: para_start,
                    pilcrow: chunk.start,
                });
                view_len = 0;
                para_start = chunk.start + 1;
            }
            ChunkKind::Text(value) => {
                if view == TextView::Vanilla && chunk.attr_active(INS) {
                    continue;
                }
                let mut raw = chunk.start;
                for ch in value.chars() {
                    let width = ch.len_utf16() as u32;
                    if !(view == TextView::Vanilla && ch == '\t') {
                        push_unit(&mut text, &mut spans, &mut view_len, ch, raw);
                    }
                    raw += width;
                }
            }
            ChunkKind::Embed(map) => {
                if view == TextView::Vanilla {
                    continue;
                }
                let is_break = map.as_ref().is_some_and(|map| {
                    map_string(map, txn, KIND_KEY).as_deref() == Some(BREAK_KIND)
                });
                let ch = if is_break { '\n' } else { '\u{FFFC}' };
                push_unit(&mut text, &mut spans, &mut view_len, ch, chunk.start);
            }
        }
    }
    views
}

/// One find-in-document hit (port of `FindInDocumentMatch`).
#[derive(Clone, Debug, PartialEq)]
pub struct FindMatch {
    pub para_id: ParagraphId,
    pub match_text: String,
    pub before: String,
    pub after: String,
}

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct FindOptions {
    pub case_sensitive: bool,
    pub limit: usize,
}

impl Default for FindOptions {
    fn default() -> Self {
        Self {
            case_sensitive: false,
            limit: 20,
        }
    }
}

/// Port of `SelectionInfo` in `queries.ts` — all fields use the Vanilla view.
#[derive(Clone, Debug, PartialEq)]
pub struct SelectionInfo {
    pub para_id: ParagraphId,
    pub selected_text: String,
    pub paragraph_text: String,
    pub before: String,
    pub after: String,
}

#[derive(Clone, Debug, PartialEq)]
pub struct CommentInfo {
    pub id: String,
    pub author: String,
    pub date: String,
    pub done: bool,
    pub parent_id: Option<String>,
    pub body: Any,
    pub ranges: Vec<LocRange>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ChangeKind {
    Insertion,
    Deletion,
    ParagraphMarkInsertion,
    ParagraphMarkDeletion,
    ParagraphPropertiesChanged,
    TableRowInsertion,
    TableRowDeletion,
    TableInsertion,
    TableDeletion,
}

#[derive(Clone, Debug, PartialEq)]
pub struct ChangeInfo {
    pub revision_id: RevisionId,
    pub kind: ChangeKind,
    pub author: String,
    pub date: String,
    pub range: LocRange,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum NavUnit {
    Grapheme,
    Word,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum NavDirection {
    Prev,
    Next,
}

/// Pagination lives OUTSIDE the CRDT; the host injects the page → paragraph mapping.
pub trait LayoutBridge {
    fn page_count(&self) -> u32;
    /// Ordered paraIds of the paragraph fragments laid out on the 1-based page. A paragraph
    /// split across pages may repeat; `page_content` dedupes by paraId (first fragment wins).
    fn paragraphs_on_page(&self, page_number: u32) -> Vec<ParagraphId>;
}

#[derive(Clone, Debug, PartialEq)]
pub struct PageParagraph {
    pub para_id: ParagraphId,
    pub text: String,
    pub style_id: Option<String>,
}

#[derive(Clone, Debug, PartialEq)]
pub struct PageContent {
    pub page_number: u32,
    /// `[paraId] text` lines joined with `\n` (port of `getPageContent`).
    pub text: String,
    pub paragraphs: Vec<PageParagraph>,
}

fn lower_char(ch: char) -> char {
    ch.to_lowercase().next().unwrap_or(ch)
}

/// Parses a `{id, author, date}` revision value (shared with the resolve ops).
pub(crate) fn revision_parts(value: &Any) -> Option<(String, String, String)> {
    let Any::Map(map) = value else {
        return None;
    };
    let identity = match map.get("info") {
        Some(Any::Map(info)) => info.as_ref(),
        _ => map.as_ref(),
    };
    let get_string = |key: &str| match identity.get(key) {
        Some(Any::String(value)) => Some(value.to_string()),
        _ => None,
    };
    // Native revisions use cooperative string ids (`client:counter`), while
    // coexistence seeds preserve imported OOXML `w:id` values as numbers.
    // Both are valid revision identities for read/resolve queries.
    let id = match identity.get("id").or_else(|| identity.get("revisionId")) {
        Some(Any::String(value)) => value.to_string(),
        Some(Any::Number(value)) if value.is_finite() => value.to_string(),
        Some(Any::BigInt(value)) => value.to_string(),
        _ => return None,
    };
    Some((
        id,
        get_string("author").unwrap_or_default(),
        get_string("date").unwrap_or_default(),
    ))
}

impl EditingDoc {
    fn views_for_story(&self, story_id: &str, view: TextView) -> OpResult<Vec<ParaView>> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, story_id)?;
        Ok(para_views(&story, &txn, view))
    }

    fn views_everywhere(&self, view: TextView) -> Vec<(String, Vec<ParaView>)> {
        let txn = self.yrs_doc().transact();
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return Vec::new();
        };
        let mut story_ids: Vec<String> = {
            use yrs::Map;
            stories.keys(&txn).map(|key| key.to_string()).collect()
        };
        story_ids.sort();
        story_ids
            .into_iter()
            .filter_map(|story_id| {
                use yrs::Map;
                match stories.get(&txn, &story_id) {
                    Some(Out::YText(story)) => {
                        Some((story_id.clone(), para_views(&story, &txn, view)))
                    }
                    _ => None,
                }
            })
            .collect()
    }

    /// The text of one paragraph in the requested view.
    pub fn para_text(&self, para_id: &str, view: TextView) -> OpResult<String> {
        for (_, views) in self.views_everywhere(view) {
            if let Some(para) = views.into_iter().find(|para| para.para_id == para_id) {
                return Ok(para.text);
            }
        }
        Err(OpError::UnknownPara(para_id.to_owned()))
    }

    /// The view text between two Locs (paragraph boundaries contribute nothing, like
    /// `getVanillaTextBetween`).
    pub fn text_between(&self, range: &LocRange, view: TextView) -> OpResult<String> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, &range.start.story)?;
        let from = global_of_loc(&story, &txn, &range.start)?;
        let to = global_of_loc(&story, &txn, &range.end)?;
        if to < from {
            return Err(OpError::InvalidRange {
                start: from,
                end: to,
            });
        }
        let views = para_views(&story, &txn, view);
        let mut out = String::new();
        for para in &views {
            para.view_slice_of_raw(from, to, &mut out);
        }
        Ok(out)
    }

    /// Resolves a search phrase to the ONE place it occurs (op-contract §3): zero occurrences →
    /// [`OpError::SearchNotFound`]; more than one → [`OpError::AmbiguousSearch`] so an LLM
    /// narrows instead of mistargeting. `within` scopes the search to one paragraph.
    /// Case-sensitive (mutation precision), searching the given view but returning RAW offsets.
    pub fn resolve_search(
        &self,
        story_id: &str,
        within: Option<&str>,
        needle: &str,
        view: TextView,
    ) -> OpResult<LocRange> {
        if needle.is_empty() {
            return Err(OpError::SearchNotFound(String::new()));
        }
        let views = self.views_for_story(story_id, view)?;
        if let Some(para_id) = within
            && !views.iter().any(|para| para.para_id == para_id)
        {
            return Err(OpError::UnknownPara(para_id.to_owned()));
        }
        let mut hits: Vec<(usize, usize)> = Vec::new(); // (para index, byte offset)
        for (index, para) in views.iter().enumerate() {
            if within.is_some_and(|para_id| para_id != para.para_id) {
                continue;
            }
            let mut from = 0usize;
            while let Some(found) = para.text[from..].find(needle) {
                hits.push((index, from + found));
                from += found + needle.len().max(1);
            }
        }
        match hits.len() {
            0 => Err(OpError::SearchNotFound(needle.to_owned())),
            1 => {
                let (index, byte) = hits[0];
                let para = &views[index];
                let view_start = utf16_of_byte(&para.text, byte);
                let view_end = view_start + needle.encode_utf16().count() as u32;
                let raw_start = para.raw_of_view(view_start, false);
                let raw_end = para.raw_of_view(view_end, true);
                Ok(LocRange {
                    start: Loc::new(story_id, para.para_id.clone(), raw_start - para.raw_start),
                    end: Loc::new(story_id, para.para_id.clone(), raw_end - para.raw_start),
                })
            }
            occurrences => Err(OpError::AmbiguousSearch {
                needle: needle.to_owned(),
                occurrences,
            }),
        }
    }

    /// Port of `findInDocument`: Vanilla view, case-insensitive by default, at most one match
    /// per paragraph (paragraphs where the needle occurs more than once are SKIPPED — the agent
    /// narrows), 40-character context windows, limit 20.
    pub fn find_in_document(
        &self,
        story_id: &str,
        needle: &str,
        options: FindOptions,
    ) -> OpResult<Vec<FindMatch>> {
        if needle.is_empty() {
            return Ok(Vec::new());
        }
        const CONTEXT: usize = 40;
        let views = self.views_for_story(story_id, TextView::Vanilla)?;
        let needle_chars: Vec<char> = if options.case_sensitive {
            needle.chars().collect()
        } else {
            needle.chars().map(lower_char).collect()
        };
        let mut results = Vec::new();
        for para in &views {
            if results.len() >= options.limit {
                break;
            }
            let original: Vec<char> = para.text.chars().collect();
            let haystack: Vec<char> = if options.case_sensitive {
                original.clone()
            } else {
                original.iter().copied().map(lower_char).collect()
            };
            let find_from = |start: usize| -> Option<usize> {
                if needle_chars.is_empty() || haystack.len() < needle_chars.len() {
                    return None;
                }
                (start..=haystack.len() - needle_chars.len())
                    .find(|&at| haystack[at..at + needle_chars.len()] == needle_chars[..])
            };
            let Some(at) = find_from(0) else {
                continue;
            };
            if find_from(at + 1).is_some() {
                continue; // Ambiguous within the paragraph — reject, agent narrows.
            }
            let match_end = at + needle_chars.len();
            results.push(FindMatch {
                para_id: para.para_id.clone(),
                match_text: original[at..match_end].iter().collect(),
                before: original[at.saturating_sub(CONTEXT)..at].iter().collect(),
                after: original[match_end..(match_end + CONTEXT).min(original.len())]
                    .iter()
                    .collect(),
            });
        }
        Ok(results)
    }

    /// Port of `getSelectionInfo`: the paragraph containing the selection start, with vanilla
    /// before/selected/after slices.
    pub fn selection_info(&self, anchor: &Loc, head: &Loc) -> OpResult<SelectionInfo> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, &anchor.story)?;
        let a = global_of_loc(&story, &txn, anchor)?;
        let h = global_of_loc(&story, &txn, head)?;
        let (from, to) = (a.min(h), a.max(h));
        let views = para_views(&story, &txn, TextView::Vanilla);
        let para = views
            .iter()
            .find(|para| from <= para.pilcrow)
            .ok_or_else(|| OpError::UnknownStory(anchor.story.clone()))?;
        let mut before = String::new();
        para.view_slice_of_raw(para.raw_start, from, &mut before);
        let mut selected = String::new();
        for view in &views {
            view.view_slice_of_raw(from, to, &mut selected);
        }
        let mut after = String::new();
        if to <= para.pilcrow {
            para.view_slice_of_raw(to, para.pilcrow, &mut after);
        }
        Ok(SelectionInfo {
            para_id: para.para_id.clone(),
            paragraph_text: format!("{before}{selected}{after}"),
            selected_text: selected,
            before,
            after,
        })
    }

    /// The current Loc range of a comment's first anchored range.
    pub fn find_comment_range(&self, comment_id: &str) -> OpResult<LocRange> {
        let anchors = self.resolve_comment(comment_id)?;
        let anchor = anchors
            .first()
            .ok_or_else(|| OpError::UnknownComment(comment_id.to_owned()))?;
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, &anchor.story)?;
        Ok(LocRange {
            start: loc_of_global(&anchor.story, &story, &txn, anchor.start)?,
            end: loc_of_global(&anchor.story, &story, &txn, anchor.end)?,
        })
    }

    /// Every comment in the side map, with best-effort resolved anchor ranges.
    pub fn list_comments(&self) -> OpResult<Vec<CommentInfo>> {
        let ids: Vec<String> = {
            let txn = self.yrs_doc().transact();
            let Some(comments) = txn.get_map(COMMENTS) else {
                return Ok(Vec::new());
            };
            use yrs::Map;
            let mut ids: Vec<String> = comments.keys(&txn).map(|key| key.to_string()).collect();
            ids.sort();
            ids
        };
        let mut result = Vec::new();
        for id in ids {
            let ranges = self
                .resolve_comment(&id)
                .ok()
                .map(|anchors| {
                    let txn = self.yrs_doc().transact();
                    anchors
                        .into_iter()
                        .filter_map(|anchor| {
                            let story = story_ref(&txn, &anchor.story).ok()?;
                            Some(LocRange {
                                start: loc_of_global(&anchor.story, &story, &txn, anchor.start)
                                    .ok()?,
                                end: loc_of_global(&anchor.story, &story, &txn, anchor.end).ok()?,
                            })
                        })
                        .collect()
                })
                .unwrap_or_default();
            let txn = self.yrs_doc().transact();
            let Some(comments) = txn.get_map(COMMENTS) else {
                continue;
            };
            use yrs::Map;
            let Some(comment) = comments
                .get(&txn, &id)
                .and_then(|value| value.cast::<yrs::MapRef>().ok())
            else {
                continue;
            };
            let string_of = |key: &str| match comment.get(&txn, key) {
                Some(Out::Any(Any::String(value))) => Some(value.to_string()),
                _ => None,
            };
            result.push(CommentInfo {
                author: string_of("author").unwrap_or_default(),
                date: string_of("date").unwrap_or_default(),
                done: matches!(comment.get(&txn, "done"), Some(Out::Any(Any::Bool(true)))),
                parent_id: string_of("parentId"),
                body: match comment.get(&txn, "body") {
                    Some(Out::Any(value)) => value,
                    _ => Any::Null,
                },
                ranges,
                id,
            });
        }
        Ok(result)
    }

    /// Every tracked change in a story: `ins`/`del` runs (adjacent chunks with the same revision
    /// ID merged), paragraph-mark revisions (`pPrIns`/`pPrDel`), and table-row
    /// revisions (`trIns`/`trDel`), ordered by position.
    pub fn list_changes(&self, story_id: &str) -> OpResult<Vec<ChangeInfo>> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, story_id)?;
        let chunks = snapshot(&story, &txn);
        struct RawChange {
            id: String,
            kind: ChangeKind,
            author: String,
            date: String,
            start: u32,
            end: u32,
        }
        let mut raw: Vec<RawChange> = Vec::new();
        for chunk in &chunks {
            if let ChunkKind::Pilcrow(map) = &chunk.kind {
                for (key, kind) in [
                    (crate::PPR_INS, ChangeKind::ParagraphMarkInsertion),
                    (crate::PPR_DEL, ChangeKind::ParagraphMarkDeletion),
                ] {
                    if let Some(Out::Any(value)) = map.get(&txn, key)
                        && let Some((id, author, date)) = revision_parts(&value)
                    {
                        raw.push(RawChange {
                            id,
                            kind,
                            author,
                            date,
                            start: chunk.start,
                            end: chunk.start + 1,
                        });
                    }
                }
                if let Some(Out::Any(Any::Array(changes))) = map.get(&txn, crate::PPR_CHANGE) {
                    for change in changes.iter() {
                        if let Some((id, author, date)) = revision_parts(change) {
                            raw.push(RawChange {
                                id,
                                kind: ChangeKind::ParagraphPropertiesChanged,
                                author,
                                date,
                                start: chunk.start,
                                end: chunk.start + 1,
                            });
                        }
                    }
                }
                continue;
            }
            for (key, kind) in [(INS, ChangeKind::Insertion), (DEL, ChangeKind::Deletion)] {
                let Some(value) = chunk.attrs.get(key) else {
                    continue;
                };
                let Some((id, author, date)) = revision_parts(value) else {
                    continue;
                };
                if let Some(last) = raw
                    .iter_mut()
                    .rev()
                    .find(|change| change.kind == kind)
                    .filter(|change| change.id == id && change.end == chunk.start)
                {
                    last.end = chunk.end();
                } else {
                    raw.push(RawChange {
                        id,
                        kind,
                        author,
                        date,
                        start: chunk.start,
                        end: chunk.end(),
                    });
                }
            }
        }
        raw.extend(
            table_row_changes(&story, &txn)
                .into_iter()
                .map(|change| RawChange {
                    id: change.revision_id,
                    kind: match change.kind {
                        TableRowChangeKind::Insertion => ChangeKind::TableRowInsertion,
                        TableRowChangeKind::Deletion => ChangeKind::TableRowDeletion,
                        TableRowChangeKind::TableInsertion => ChangeKind::TableInsertion,
                        TableRowChangeKind::TableDeletion => ChangeKind::TableDeletion,
                    },
                    author: change.author,
                    date: change.date,
                    start: change.start,
                    end: change.start + 1,
                }),
        );
        raw.sort_by_key(|change| change.start);
        raw.into_iter()
            .map(|change| {
                Ok(ChangeInfo {
                    revision_id: change.id,
                    kind: change.kind,
                    author: change.author,
                    date: change.date,
                    range: crate::op::loc_range_in_txn(
                        story_id,
                        &story,
                        &txn,
                        change.start,
                        change.end,
                    )?,
                })
            })
            .collect()
    }

    /// The Loc range covering every unit stamped with the revision ID (any story).
    pub fn find_change_range(&self, revision_id: &str) -> OpResult<LocRange> {
        let txn = self.yrs_doc().transact();
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return Err(OpError::UnknownChange(revision_id.to_owned()));
        };
        let mut story_ids: Vec<String> = {
            use yrs::Map;
            stories.keys(&txn).map(|key| key.to_string()).collect()
        };
        story_ids.sort();
        for story_id in story_ids {
            use yrs::Map;
            let Some(Out::YText(story)) = stories.get(&txn, &story_id) else {
                continue;
            };
            let mut min: Option<u32> = None;
            let mut max: Option<u32> = None;
            for chunk in snapshot(&story, &txn) {
                let mut matched = [INS, DEL].iter().any(|key| {
                    chunk
                        .attrs
                        .get(*key)
                        .and_then(revision_parts)
                        .map(|(id, ..)| id)
                        == Some(revision_id.to_owned())
                });
                if let ChunkKind::Pilcrow(map) = &chunk.kind {
                    matched = matched
                        || [crate::PPR_INS, crate::PPR_DEL].iter().any(|key| {
                            matches!(
                                map.get(&txn, key),
                                Some(Out::Any(value))
                                    if revision_parts(&value).map(|(id, ..)| id)
                                        == Some(revision_id.to_owned())
                            )
                        })
                        || matches!(
                            map.get(&txn, crate::PPR_CHANGE),
                            Some(Out::Any(Any::Array(changes)))
                                if changes.iter().any(|change| {
                                    revision_parts(change).map(|(id, ..)| id)
                                        == Some(revision_id.to_owned())
                                })
                        );
                }
                if matched {
                    min = Some(min.map_or(chunk.start, |value| value.min(chunk.start)));
                    max = Some(max.map_or(chunk.end(), |value| value.max(chunk.end())));
                }
            }
            for change in table_row_changes(&story, &txn) {
                if change.revision_id == revision_id {
                    min = Some(min.map_or(change.start, |value| value.min(change.start)));
                    max = Some(max.map_or(change.start + 1, |value| value.max(change.start + 1)));
                }
            }
            if let (Some(start), Some(end)) = (min, max) {
                return crate::op::loc_range_in_txn(&story_id, &story, &txn, start, end);
            }
        }
        Err(OpError::UnknownChange(revision_id.to_owned()))
    }

    /// The nearest grapheme or word boundary from a Loc, crossing paragraph edges within the
    /// story. Boundaries are computed on the Raw view (embeds are their own boundary).
    pub fn nav_boundary(&self, loc: &Loc, unit: NavUnit, direction: NavDirection) -> OpResult<Loc> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, &loc.story)?;
        let views = para_views(&story, &txn, TextView::Raw);
        let index = views
            .iter()
            .position(|para| para.para_id == loc.para)
            .ok_or_else(|| OpError::UnknownPara(loc.para.clone()))?;
        let para = &views[index];
        let para_len = para.pilcrow - para.raw_start;
        if loc.offset > para_len {
            return Err(OpError::OutOfBounds {
                index: loc.offset,
                len: para_len,
            });
        }
        let mut boundaries: Vec<u32> = match unit {
            NavUnit::Grapheme => para
                .text
                .grapheme_indices(true)
                .map(|(byte, _)| utf16_of_byte(&para.text, byte))
                .collect(),
            NavUnit::Word => para
                .text
                .split_word_bound_indices()
                .map(|(byte, _)| utf16_of_byte(&para.text, byte))
                .collect(),
        };
        boundaries.push(para_len);
        boundaries.dedup();
        match direction {
            NavDirection::Next => match boundaries.iter().find(|&&b| b > loc.offset) {
                Some(&next) => Ok(Loc::new(loc.story.clone(), loc.para.clone(), next)),
                None => match views.get(index + 1) {
                    Some(next_para) => {
                        Ok(Loc::new(loc.story.clone(), next_para.para_id.clone(), 0))
                    }
                    None => Ok(Loc::new(loc.story.clone(), loc.para.clone(), para_len)),
                },
            },
            NavDirection::Prev => match boundaries.iter().rev().find(|&&b| b < loc.offset) {
                Some(&prev) => Ok(Loc::new(loc.story.clone(), loc.para.clone(), prev)),
                None => match index.checked_sub(1).and_then(|i| views.get(i)) {
                    Some(prev_para) => Ok(Loc::new(
                        loc.story.clone(),
                        prev_para.para_id.clone(),
                        prev_para.pilcrow - prev_para.raw_start,
                    )),
                    None => Ok(Loc::new(loc.story.clone(), loc.para.clone(), 0)),
                },
            },
        }
    }

    /// Port of `getPageContent`: the paragraphs on one 1-based page, deduped by paraId (a
    /// paragraph split across a page boundary is reported once). Returns `None` for an
    /// out-of-range page.
    pub fn page_content(
        &self,
        page_number: u32,
        bridge: &dyn LayoutBridge,
        view: TextView,
    ) -> OpResult<Option<PageContent>> {
        if page_number == 0 || page_number > bridge.page_count() {
            return Ok(None);
        }
        let mut lookup: HashMap<String, (String, Option<String>)> = HashMap::new();
        for (_, views) in self.views_everywhere(view) {
            for para in views {
                lookup
                    .entry(para.para_id.clone())
                    .or_insert((para.text, para.style_id));
            }
        }
        let mut seen: HashSet<String> = HashSet::new();
        let mut paragraphs = Vec::new();
        for para_id in bridge.paragraphs_on_page(page_number) {
            if !seen.insert(para_id.clone()) {
                continue;
            }
            let Some((text, style_id)) = lookup.get(&para_id) else {
                continue;
            };
            paragraphs.push(PageParagraph {
                para_id,
                text: text.clone(),
                style_id: style_id.clone(),
            });
        }
        let text = paragraphs
            .iter()
            .map(|para| format!("[{}] {}", para.para_id, para.text))
            .collect::<Vec<_>>()
            .join("\n");
        Ok(Some(PageContent {
            page_number,
            text,
            paragraphs,
        }))
    }
}
