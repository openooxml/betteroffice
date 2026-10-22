//! Read queries over transaction snapshots.

use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::sync::Arc;

use unicode_segmentation::UnicodeSegmentation;
use yrs::{Any, Map, Out, ReadTxn, Transact};

use crate::op::{Loc, LocRange, OpError, OpResult, global_of_loc, loc_of_global};
use crate::ops::table::{
    TableRowChangeKind, push_table_row_changes, table_revision_stamps, table_row_changes,
};
use crate::ops::{Chunk, ChunkKind};
use crate::{
    BREAK_KIND, COMMENTS, DEL, EditingDoc, INS, KIND_KEY, PARA_ID, ParagraphId, RevisionId,
    map_string, story_ref,
};

/// A tracked change with its story-global `[start, end)`, before its paragraph
/// locations are resolved.
pub(crate) struct RawChange {
    pub(crate) id: String,
    kind: ChangeKind,
    author: String,
    date: String,
    pub(crate) start: u32,
    pub(crate) end: u32,
}

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
    pub(crate) fn view_slice_of_raw(&self, from: u32, to: u32, out: &mut String) {
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

pub(crate) fn para_views<T: ReadTxn>(txn: &T, view: TextView, chunks: &[Chunk]) -> Vec<ParaView> {
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

    for chunk in chunks {
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

/// One document-search match.
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

/// Selection details in the vanilla text view.
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
    /// Paragraph lines joined with newlines.
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
    // Imported OOXML revision IDs remain numeric.
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

fn visit_chunk_revisions<T: ReadTxn>(
    chunk: &Chunk,
    txn: &T,
    mut visit: impl FnMut(ChangeKind, (String, String, String)),
) {
    if let ChunkKind::Pilcrow(map) = &chunk.kind {
        for (key, kind) in [
            (crate::PPR_INS, ChangeKind::ParagraphMarkInsertion),
            (crate::PPR_DEL, ChangeKind::ParagraphMarkDeletion),
        ] {
            if let Some(Out::Any(value)) = map.get(txn, key)
                && let Some(stamp) = revision_parts(&value)
            {
                visit(kind, stamp);
            }
        }
        if let Some(Out::Any(Any::Array(changes))) = map.get(txn, crate::PPR_CHANGE) {
            for change in changes.iter() {
                if let Some(stamp) = revision_parts(change) {
                    visit(ChangeKind::ParagraphPropertiesChanged, stamp);
                }
            }
        }
    } else {
        for (key, kind) in [(INS, ChangeKind::Insertion), (DEL, ChangeKind::Deletion)] {
            if let Some(stamp) = chunk.attrs.get(key).and_then(revision_parts) {
                visit(kind, stamp);
            }
        }
        if let ChunkKind::Embed(Some(map)) = &chunk.kind
            && map_string(map, txn, KIND_KEY).as_deref() == Some("sdt")
            && let Some(Out::Any(content)) = map.get(txn, "content")
        {
            let mut seen: Vec<_> = [(INS, ChangeKind::Insertion), (DEL, ChangeKind::Deletion)]
                .into_iter()
                .filter_map(|(key, kind)| {
                    chunk
                        .attrs
                        .get(key)
                        .and_then(revision_parts)
                        .map(|stamp| (kind, stamp))
                })
                .collect();
            crate::inline_content::visit(&content, &mut |child| {
                let Some(Any::Map(attrs)) = child.get("attrs") else {
                    return;
                };
                for (key, kind) in [(INS, ChangeKind::Insertion), (DEL, ChangeKind::Deletion)] {
                    if let Some(stamp) = attrs.get(key).and_then(revision_parts) {
                        let entry = (kind, stamp);
                        if !seen.contains(&entry) {
                            seen.push(entry.clone());
                            visit(entry.0, entry.1);
                        }
                    }
                }
            });
        }
    }
}

/// The tracked changes of `chunks` and `tables` (one story's, in story order) as
/// [`EditingDoc::story_changes`] finds them, ordered by position. A run merges only
/// with the run of its kind ending where it starts, and paragraph marks start their own.
fn raw_changes<'a, T: ReadTxn>(
    chunks: impl IntoIterator<Item = &'a Chunk>,
    tables: Vec<crate::ops::table::TableRowChange>,
    txn: &T,
) -> Vec<RawChange> {
    let mut raw: Vec<RawChange> = Vec::new();
    for chunk in chunks {
        visit_chunk_revisions(chunk, txn, |kind, (id, author, date)| {
            let pilcrow = matches!(&chunk.kind, ChunkKind::Pilcrow(_));
            if !pilcrow
                && let Some(last) = raw
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
                    end: if pilcrow {
                        chunk.start + 1
                    } else {
                        chunk.end()
                    },
                });
            }
        });
    }
    raw.extend(tables.into_iter().map(|change| RawChange {
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
    }));
    raw.sort_by_key(|change| change.start);
    raw
}

impl EditingDoc {
    fn views_for_story(&self, story_id: &str, view: TextView) -> OpResult<Vec<ParaView>> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, story_id)?;
        let chunks = self.chunk_snapshot(story_id, &story, &txn);
        Ok(para_views(&txn, view, &chunks))
    }

    fn views_everywhere(&self, view: TextView) -> Vec<(String, Vec<ParaView>)> {
        let txn = self.yrs_doc().transact();
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return Vec::new();
        };
        let mut story_ids: Vec<String> = crate::live_stories(&stories, &txn)
            .map(|(id, _)| id.to_owned())
            .collect();
        story_ids.sort();
        story_ids
            .into_iter()
            .filter_map(|story_id| {
                use yrs::Map;
                match stories.get(&txn, &story_id) {
                    Some(Out::YText(story)) => {
                        let chunks = self.chunk_snapshot(&story_id, &story, &txn);
                        Some((story_id.clone(), para_views(&txn, view, &chunks)))
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

    /// Returns view text without paragraph-boundary text.
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
        let views = para_views(
            &txn,
            view,
            &self.chunk_snapshot(&range.start.story, &story, &txn),
        );
        let mut out = String::new();
        for para in &views {
            para.view_slice_of_raw(from, to, &mut out);
        }
        Ok(out)
    }

    /// Resolves a unique case-sensitive search match to raw offsets.
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

    /// Finds unambiguous paragraph-local matches in the vanilla view.
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

    /// Returns vanilla-view slices around a selection.
    pub fn selection_info(&self, anchor: &Loc, head: &Loc) -> OpResult<SelectionInfo> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, &anchor.story)?;
        let a = global_of_loc(&story, &txn, anchor)?;
        let h = global_of_loc(&story, &txn, head)?;
        let (from, to) = (a.min(h), a.max(h));
        let views = para_views(
            &txn,
            TextView::Vanilla,
            &self.chunk_snapshot(&anchor.story, &story, &txn),
        );
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
        Ok(self
            .story_changes(story_id)?
            .into_iter()
            .map(|(change, _)| change)
            .collect())
    }

    /// Distinct author and date pairs for requested revisions across all stories.
    pub fn revision_stamps(
        &self,
        ids: &[String],
    ) -> OpResult<BTreeMap<String, BTreeSet<(String, String)>>> {
        let requested: HashSet<&str> = ids.iter().map(String::as_str).collect();
        let mut result: BTreeMap<String, BTreeSet<(String, String)>> = BTreeMap::new();
        if requested.is_empty() {
            return Ok(result);
        }
        let txn = self.yrs_doc().transact();
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return Ok(result);
        };
        let mut collect = |(id, author, date): (String, String, String)| {
            if requested.contains(id.as_str()) {
                result.entry(id).or_default().insert((author, date));
            }
        };
        for (story_id, value) in crate::live_stories(&stories, &txn) {
            let Out::YText(story) = value else {
                continue;
            };
            for chunk in self.chunk_snapshot(story_id, &story, &txn).iter() {
                visit_chunk_revisions(chunk, &txn, |_, stamp| collect(stamp));
                if let ChunkKind::Embed(Some(map)) = &chunk.kind {
                    for stamp in table_revision_stamps(map, &txn) {
                        collect(stamp);
                    }
                }
            }
        }
        Ok(result)
    }

    /// [`Self::list_changes`] with each change's story-global `[start, end)`.
    pub(crate) fn story_changes(&self, story_id: &str) -> OpResult<Vec<(ChangeInfo, (u32, u32))>> {
        let txn = self.yrs_doc().transact();
        let story = story_ref(&txn, story_id)?;
        let raw = self.story_raw_changes(story_id, &story, &txn);
        if raw.is_empty() {
            return Ok(Vec::new());
        }
        let bounds = crate::op::para_bounds(&story, &txn);
        raw.into_iter()
            .map(|change| {
                let range = LocRange {
                    start: crate::op::loc_in_bounds(story_id, &bounds, change.start)?,
                    end: crate::op::loc_in_bounds(story_id, &bounds, change.end)?,
                };
                Ok((
                    ChangeInfo {
                        revision_id: change.id,
                        kind: change.kind,
                        author: change.author,
                        date: change.date,
                        range,
                    },
                    (change.start, change.end),
                ))
            })
            .collect()
    }

    /// A story's tracked changes as [`Self::story_changes`] finds them, ordered by
    /// position, without resolving paragraph locations (a story may hold none).
    pub(crate) fn story_raw_changes<T: ReadTxn>(
        &self,
        story_id: &str,
        story: &yrs::TextRef,
        txn: &T,
    ) -> Vec<RawChange> {
        let chunks = self.chunk_snapshot(story_id, story, txn);
        raw_changes(chunks.iter(), table_row_changes(story, txn), txn)
    }

    /// Ids of the tracked changes of `story_id` whose range holds the caret at
    /// `para_id`/`offset`, in [`Self::list_changes`] order, both range ends
    /// resolved through the first paragraph carrying their paragraph's id. Reads
    /// the caret paragraph and the mark before it, or the whole story when an
    /// offset past the paragraph or a repeated paragraph id needs it.
    #[cfg_attr(not(feature = "wasm"), allow(dead_code))]
    pub(crate) fn revision_ids_at(
        &self,
        story_id: &str,
        para_id: &str,
        offset: u32,
    ) -> OpResult<Vec<String>> {
        let segments = self.segment_index(story_id)?;
        let index = self.paragraph_index(story_id)?;
        let slot = index
            .para_slot(para_id)
            .ok_or_else(|| OpError::UnknownPara(para_id.to_owned()))? as usize;
        let (start, pilcrow) = index.para_span(para_id).expect("slot has a span");
        let position = u64::from(start) + u64::from(offset);
        let at = |index_at: u32| {
            let para = index.para_at_clamped(index_at)?;
            let offset = index_at
                .min(para.pilcrow)
                .saturating_sub(para.start)
                .min(para.pilcrow - para.start);
            index
                .para_span(&para.para_id)
                .map(|(start, _)| u64::from(start) + u64::from(offset))
        };
        // No change crosses a pilcrow: with unique paragraph ids only the previous mark,
        // the caret paragraph and any units past the last mark can hold the caret.
        let raw = if index.has_repeated() || offset > pilcrow - start {
            let txn = self.yrs_doc().transact();
            let story = story_ref(&txn, story_id)?;
            self.story_raw_changes(story_id, &story, &txn)
        } else {
            let lo = slot
                .checked_sub(1)
                .and_then(|previous| index.pilcrow_at(previous))
                .unwrap_or(0);
            let hi = if slot + 1 == index.para_count() {
                u32::MAX
            } else {
                pilcrow + 1
            };
            let txn = self.yrs_doc().transact();
            let story = story_ref(&txn, story_id)?;
            let indexed = crate::ops::IndexedStory::new(segments, Arc::clone(&index));
            let snapshot;
            let read;
            let chunks: &[Chunk] = match indexed.snapshot_window(&txn, lo, hi, false) {
                Some(window) => {
                    read = window;
                    &read
                }
                None => {
                    snapshot = self.chunk_snapshot(story_id, &story, &txn);
                    let first = snapshot.partition_point(|chunk| chunk.end() <= lo);
                    let last = snapshot.partition_point(|chunk| chunk.start < hi);
                    &snapshot[first..last]
                }
            };
            let mut tables = Vec::new();
            for chunk in chunks {
                if let ChunkKind::Embed(Some(map)) = &chunk.kind {
                    push_table_row_changes(map, &txn, chunk.start, &mut tables);
                }
            }
            raw_changes(chunks.iter(), tables, &txn)
        };
        Ok(raw
            .into_iter()
            .filter(|change| {
                at(change.start).is_some_and(|start| start <= position)
                    && at(change.end).is_some_and(|end| position <= end)
            })
            .map(|change| change.id)
            .collect())
    }

    /// The Loc range covering every unit stamped with the revision ID (any story).
    pub fn find_change_range(&self, revision_id: &str) -> OpResult<LocRange> {
        let txn = self.yrs_doc().transact();
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return Err(OpError::UnknownChange(revision_id.to_owned()));
        };
        let mut story_ids: Vec<String> = crate::live_stories(&stories, &txn)
            .map(|(id, _)| id.to_owned())
            .collect();
        story_ids.sort();
        for story_id in story_ids {
            use yrs::Map;
            let Some(Out::YText(story)) = stories.get(&txn, &story_id) else {
                continue;
            };
            let chunks = self.chunk_snapshot(&story_id, &story, &txn);
            let mut min: Option<u32> = None;
            let mut max: Option<u32> = None;
            for chunk in chunks.iter() {
                let mut matched = [INS, DEL].iter().any(|key| {
                    chunk
                        .attrs
                        .get(*key)
                        .and_then(revision_parts)
                        .map(|(id, ..)| id)
                        == Some(revision_id.to_owned())
                });
                if let ChunkKind::Embed(Some(map)) = &chunk.kind
                    && map_string(map, &txn, KIND_KEY).as_deref() == Some("sdt")
                {
                    visit_chunk_revisions(chunk, &txn, |_, (id, ..)| {
                        matched |= id == revision_id;
                    });
                }
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
        let views = para_views(
            &txn,
            TextView::Raw,
            &self.chunk_snapshot(&loc.story, &story, &txn),
        );
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

    /// Returns deduplicated paragraphs for a 1-based page.
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

#[cfg(test)]
mod caret_revision_tests {
    use super::*;
    use crate::format::FormatPolicy;
    use crate::{EditCtx, Position};

    const DATE: &str = "2026-01-01T00:00:00Z";

    fn stamp(id: u32) -> String {
        format!(r#"w:id="{id}" w:author="A" w:date="{DATE}""#)
    }

    fn body() -> String {
        let mark = |kind: &str, id: u32| {
            format!("<w:pPr><w:rPr><w:{kind} {}/></w:rPr></w:pPr>", stamp(id))
        };
        let run = |text: &str| format!(r#"<w:r><w:t xml:space="preserve">{text}</w:t></w:r>"#);
        let bold = |text: &str| {
            format!(r#"<w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">{text}</w:t></w:r>"#)
        };
        let ins = |id: u32, runs: String| format!("<w:ins {}>{runs}</w:ins>", stamp(id));
        let del = format!(
            r#"<w:del {}><w:r><w:delText>gone</w:delText></w:r></w:del>"#,
            stamp(23)
        );
        format!(
            concat!(
                r#"<w:p w14:paraId="0000000A">{}{}{}</w:p>"#,
                r#"<w:p w14:paraId="0000000B">{}{}{}</w:p>"#,
                r#"<w:p w14:paraId="0000000C">{}{}</w:p>"#,
                r#"<w:p w14:paraId="0000000D">{}</w:p>"#,
                r#"<w:tbl><w:tblGrid><w:gridCol w:w="2000"/></w:tblGrid>"#,
                r#"<w:tr><w:trPr><w:ins {}/></w:trPr>"#,
                r#"<w:tc><w:p w14:paraId="0000000E">{}</w:p></w:tc></w:tr></w:tbl>"#,
                r#"<w:p w14:paraId="0000000F">{}</w:p>"#,
            ),
            mark("ins", 21),
            run("Lead "),
            ins(21, run("inserted")),
            ins(21, format!("{}{}", bold("still "), run("inserted"))),
            run(" plain"),
            del,
            mark("del", 24),
            run("Mark deleted"),
            run("After"),
            stamp(25),
            run("Cell"),
            run("Tail"),
        )
    }

    const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#;
    const PACKAGE_RELS: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#;
    const DOCUMENT_OPEN: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>"#;

    fn seeded(body: &str) -> EditingDoc {
        let document = format!("{DOCUMENT_OPEN}{body}</w:body></w:document>");
        let parts = [
            ("[Content_Types].xml", CONTENT_TYPES.to_owned()),
            ("_rels/.rels", PACKAGE_RELS.to_owned()),
            ("word/document.xml", document),
        ];
        let bytes = ooxml_opc::rezip_parts(
            &parts
                .into_iter()
                .map(|(name, value)| (name.to_owned(), value.into_bytes()))
                .collect::<Vec<_>>(),
        )
        .unwrap();
        let doc = EditingDoc::new(41);
        crate::seed::seed_from_docx(&doc, &bytes).unwrap();
        doc
    }

    fn scanned(doc: &EditingDoc, para_id: &str, offset: u32) -> Vec<String> {
        let index = doc.paragraph_index("body").unwrap();
        let at = |loc: &Loc| {
            index
                .para_span(&loc.para)
                .map(|(start, _)| u64::from(start) + u64::from(loc.offset))
        };
        let position = u64::from(index.para_span(para_id).unwrap().0) + u64::from(offset);
        doc.story_changes("body")
            .unwrap()
            .into_iter()
            .filter(|(change, _)| {
                at(&change.range.start).is_some_and(|start| start <= position)
                    && at(&change.range.end).is_some_and(|end| position <= end)
            })
            .map(|(change, _)| change.revision_id)
            .collect()
    }

    fn expect_every_caret_matches(doc: &EditingDoc) -> usize {
        let index = doc.paragraph_index("body").unwrap();
        let mut held = 0;
        for slot in 0..index.para_count() {
            let pilcrow = index.pilcrow_at(slot).unwrap();
            let start = index.unit_start(slot).unwrap();
            let para_id = index.para_at(pilcrow).unwrap().para_id.to_string();
            if index.para_slot(&para_id) != Some(slot as u32) {
                continue;
            }
            for offset in 0..=pilcrow - start + 2 {
                let ids = doc.revision_ids_at("body", &para_id, offset).unwrap();
                assert_eq!(ids, scanned(doc, &para_id, offset), "{para_id}@{offset}");
                held += usize::from(!ids.is_empty());
            }
        }
        held
    }

    #[test]
    fn caret_revision_reads_match_whole_story_scans() {
        for repeat_last in [false, true] {
            let doc = seeded(&body());
            // Opening a package renames repeated ids, so share one after seeding.
            if repeat_last {
                crate::seed::fixture::share_key(&doc, "body", "0000000F", "0000000B");
            }
            assert_eq!(
                doc.paragraph_index("body").unwrap().has_repeated(),
                repeat_last
            );
            assert!(expect_every_caret_matches(&doc) > 20);
            let held = |para_id: &str, offset: u32| {
                doc.revision_ids_at("body", para_id, offset).unwrap().len()
            };
            let caret_paragraph = [held("0000000B", 0), held("0000000B", 3)];
            let after_deleted_mark = [held("0000000D", 0), held("0000000D", 1)];
            // A shared id resolves the row insertion before the tail through the first paragraph.
            assert_eq!(caret_paragraph, [2 + usize::from(repeat_last), 1]);
            assert_eq!(after_deleted_mark, [1, 0]);
            let suggesting = EditCtx::local("B", DATE).suggesting();
            for index in [9, 20, 40] {
                let at = Position::new("body", index);
                doc.insert_text(&suggesting, at, "zz", FormatPolicy::Plain)
                    .unwrap();
                expect_every_caret_matches(&doc);
            }
        }
    }
}
