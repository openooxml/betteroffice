//! Versioned paragraph projections and the target resolver shared by reads, edit batches and the
//! legacy agent helpers.
//!
//! A projection gives each paragraph of a story its text in one revision view. Offsets are
//! half-open UTF-16 positions into that text: every inline atom (hard break, image, control, note
//! reference, field, other embed) occupies one U+FFFC, tabs stay `\t`, and neither the paragraph
//! mark nor the block embeds leading a paragraph are part of it. The accepted view shows pending
//! insertions and hides pending deletions; the original view does the reverse.

use std::collections::{BTreeMap, HashMap};
use std::rc::Rc;

use serde::{Deserialize, Serialize};
use yrs::{Any, Map, MapRef, Out, ReadTxn, Transact};

use crate::batch::{
    DocumentVersion, EditFailure, EditFailureCode, EditRefusal, EditTarget, failure, refusal,
};
use crate::content_controls::Inventory;
use crate::format::InlineFormatDelta;
use crate::ops::{ChunkKind, capture_pilcrow, utf16_len};
use crate::policy::Ownership;
use crate::queries::SelectionInfo;
use crate::segments::is_block_embed;
use crate::{
    DEL, EditCtx, EditingDoc, INS, KIND_KEY, Loc, LocRange, OpError, OpResult, PPR_CHANGE, PPR_DEL,
    PPR_INS, RawOp, StoryRange, map_string, story_ref,
};

/// Paragraphs one read may return.
const READ_PARAGRAPH_LIMIT: usize = 65_536;
/// UTF-16 units of text one read may return.
const READ_TEXT_LIMIT: usize = 8_388_608;
/// Matches a search returns when the request names no limit.
const FIND_DEFAULT_LIMIT: u32 = 100;
/// Matches one search may return.
const FIND_MAX_LIMIT: u32 = 10_000;

/// Which revision projection a read or target uses.
#[derive(Clone, Copy, Debug, Default, Deserialize, Eq, Hash, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum EditTextView {
    /// Pending insertions included, pending deletions excluded.
    #[default]
    Accepted,
    /// Pending deletions included, pending insertions excluded.
    Original,
}

/// One paragraph addressed by story and paragraph id.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ParagraphTarget {
    pub story: String,
    pub para_id: String,
}

/// A UTF-16 offset into one paragraph's projected text.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TextPosition {
    pub para_id: String,
    pub offset: u32,
}

/// A half-open range of projected text in one story and view.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TextRange {
    pub story: String,
    pub start: TextPosition,
    pub end: TextPosition,
    pub view: EditTextView,
}

/// Where a search looks.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase", deny_unknown_fields)]
pub enum SearchScope {
    Story { story: String },
    Paragraph(ParagraphTarget),
}

/// The text a step addresses.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(tag = "kind", rename_all = "camelCase", deny_unknown_fields)]
pub enum TextTarget {
    /// The paragraph's whole accepted-view text.
    Paragraph(ParagraphTarget),
    /// An explicit range; v1 requires both ends in one paragraph.
    Range(TextRange),
    /// The one exact, case-sensitive, paragraph-local match of `text` in `within`.
    Search {
        text: String,
        within: SearchScope,
        view: EditTextView,
    },
}

/// What an inline atom in projected text stands for.
#[derive(Clone, Copy, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum AtomKind {
    LineBreak,
    Image,
    ContentControl,
    NoteReference,
    Field,
    Other,
}

/// One U+FFFC in projected text that stands for an inline atom.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextAtom {
    pub offset: u32,
    pub kind: AtomKind,
}

/// One paragraph's projected text.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphText {
    pub story: String,
    pub para_id: String,
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_id: Option<String>,
    pub atoms: Vec<TextAtom>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ReadParagraphsRequest {
    /// Defaults to `body`; child stories are never included implicitly.
    #[serde(default)]
    pub story: Option<String>,
    /// Restricts the read to these paragraphs, returned in story order.
    #[serde(default)]
    pub para_ids: Option<Vec<String>>,
    pub view: EditTextView,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ReadParagraphsResponse {
    pub version: DocumentVersion,
    pub view: EditTextView,
    pub paragraphs: Vec<ParagraphText>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FindTextRequest {
    pub text: String,
    pub within: SearchScope,
    pub view: EditTextView,
    #[serde(default)]
    pub limit: Option<u32>,
}

/// One search hit; `range` is reusable as a [`TextTarget::Range`].
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextMatch {
    pub text: String,
    pub range: TextRange,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FindTextResponse {
    pub version: DocumentVersion,
    pub matches: Vec<TextMatch>,
    /// More matches exist than were returned.
    pub truncated: bool,
}

/// Visible units whose view and raw offsets advance together.
#[derive(Clone, Copy, Debug)]
struct Span {
    view: u32,
    raw: u32,
    len: u32,
    stamped: bool,
    atom: bool,
}

/// A paragraph mark's properties and whether it carries a revision.
pub(crate) struct Mark {
    pub properties: BTreeMap<String, Any>,
    pub revision: bool,
}

impl Mark {
    fn has(&self, key: &str) -> bool {
        self.properties
            .get(key)
            .is_some_and(|value| !matches!(value, Any::Null | Any::Undefined))
    }

    pub fn property_change(&self) -> bool {
        matches!(self.properties.get(PPR_CHANGE), Some(Any::Array(changes)) if !changes.is_empty())
    }

    pub fn section(&self) -> bool {
        self.has("sectPr") || self.has("sectionBreakType")
    }

    pub fn numbering(&self) -> bool {
        self.has("numPr")
    }

    /// The source runs retained for save carry tracked formatting changes.
    fn run_revisions(&self) -> bool {
        let Some(Any::Array(runs)) = self.properties.get("_originalRunBoundaries") else {
            return false;
        };
        runs.iter().any(|run| {
            matches!(run, Any::Map(run) if matches!(
                run.get("propertyChanges"),
                Some(Any::Array(changes)) if !changes.is_empty()
            ))
        })
    }

    pub fn style_id(&self) -> Option<String> {
        match self.properties.get("pStyle") {
            Some(Any::String(value)) => Some(value.to_string()),
            _ => None,
        }
    }

    fn original_style_id(&self) -> Option<String> {
        let Some(Any::Array(changes)) = self.properties.get(PPR_CHANGE) else {
            return self.style_id();
        };
        let Some(Any::Map(first)) = changes.first() else {
            return self.style_id();
        };
        match first.get("previousFormatting") {
            Some(Any::Map(previous)) => {
                match previous.get("pStyle").or_else(|| previous.get("styleId")) {
                    Some(Any::String(value)) => Some(value.to_string()),
                    _ => None,
                }
            }
            _ => self.style_id(),
        }
    }
}

/// One paragraph of a story projection.
pub(crate) struct ParagraphView {
    pub para_id: String,
    pub style_id: Option<String>,
    pub text: String,
    pub atoms: Vec<TextAtom>,
    spans: Vec<Span>,
    /// Raw intervals of inline units carrying `ins` or `del`, visible or hidden.
    stamped: Vec<(u32, u32)>,
    /// Story index after the previous paragraph mark.
    pub start: u32,
    /// Story index of the first inline unit, after any leading block embeds.
    pub node_start: u32,
    /// Story index of this paragraph's mark.
    pub pilcrow: u32,
    pub mark: Mark,
    /// Its runs carry tracked formatting changes that editing its text would drop.
    pub run_revisions: bool,
    /// Embeds in the inline region, visible or hidden, with their kind.
    pub embeds: Vec<(String, Option<MapRef>)>,
}

impl ParagraphView {
    pub fn len(&self) -> u32 {
        self.spans.last().map_or(0, |span| span.view + span.len)
    }

    /// Raw index of the visible unit at `offset`, or the paragraph mark at the end.
    pub fn raw_at(&self, offset: u32) -> u32 {
        self.spans
            .iter()
            .find(|span| offset < span.view + span.len)
            .map_or(self.pilcrow, |span| span.raw + (offset - span.view))
    }

    /// Raw index after the visible unit before `offset`, or the inline start at zero.
    pub fn raw_after(&self, offset: u32) -> u32 {
        if offset == 0 {
            return self.node_start;
        }
        self.spans
            .iter()
            .find(|span| offset - 1 < span.view + span.len)
            .map_or(self.pilcrow, |span| span.raw + (offset - span.view))
    }

    /// Projected offset of raw story index `raw`, clamped into this paragraph.
    pub fn offset_of_raw(&self, raw: u32) -> u32 {
        match self.spans.partition_point(|span| span.raw <= raw) {
            0 => 0,
            after => {
                let span = &self.spans[after - 1];
                span.view + (raw - span.raw).min(span.len)
            }
        }
    }

    /// Whether `offset` falls between Unicode scalar values of the projected text.
    pub fn is_scalar_boundary(&self, offset: u32) -> bool {
        let mut position = 0u32;
        for ch in self.text.chars() {
            if position >= offset {
                return position == offset;
            }
            position += ch.len_utf16() as u32;
        }
        position == offset
    }

    pub fn has_atom_in(&self, start: u32, end: u32) -> bool {
        self.spans
            .iter()
            .any(|span| span.atom && span.view < end && start < span.view + span.len)
    }

    pub fn has_stamped_in(&self, start: u32, end: u32) -> bool {
        self.spans
            .iter()
            .any(|span| span.stamped && span.view < end && start < span.view + span.len)
    }

    pub fn has_revisions(&self) -> bool {
        !self.stamped.is_empty()
            || self.mark.revision
            || self.mark.property_change()
            || self.run_revisions
    }

    /// Whether the raw unit at `raw` carries a revision.
    pub fn stamped_at(&self, raw: u32) -> bool {
        self.stamped
            .iter()
            .any(|(start, end)| *start <= raw && raw < *end)
    }

    pub fn slice(&self, start: u32, end: u32) -> String {
        let mut out = String::new();
        let mut position = 0u32;
        for ch in self.text.chars() {
            if position >= end {
                break;
            }
            if position >= start {
                out.push(ch);
            }
            position += ch.len_utf16() as u32;
        }
        out
    }

    fn record(&self, story: &str) -> ParagraphText {
        ParagraphText {
            story: story.to_owned(),
            para_id: self.para_id.clone(),
            text: self.text.clone(),
            style_id: self.style_id.clone(),
            atoms: self.atoms.clone(),
        }
    }

    fn occurrences<'a>(&'a self, needle: &'a str) -> Occurrences<'a> {
        Occurrences {
            text: &self.text,
            needle,
            needle_units: utf16_len(needle),
            byte: 0,
            units: 0,
        }
    }
}

/// Overlapping matches of a needle, found lazily with UTF-16 offsets kept incrementally.
struct Occurrences<'a> {
    text: &'a str,
    needle: &'a str,
    needle_units: u32,
    byte: usize,
    units: u32,
}

impl Iterator for Occurrences<'_> {
    type Item = (u32, u32);

    fn next(&mut self) -> Option<Self::Item> {
        let found = self.byte + self.text.get(self.byte..)?.find(self.needle)?;
        self.units += utf16_len(&self.text[self.byte..found]);
        let start = self.units;
        let first = self.text[found..].chars().next()?;
        self.byte = found + first.len_utf8();
        self.units += first.len_utf16() as u32;
        Some((start, start + self.needle_units))
    }
}

fn atom_kind(kind: &str) -> AtomKind {
    match kind {
        "break" => AtomKind::LineBreak,
        "image" => AtomKind::Image,
        "sdt" => AtomKind::ContentControl,
        "noteRef" => AtomKind::NoteReference,
        "field" => AtomKind::Field,
        _ => AtomKind::Other,
    }
}

fn visible(view: EditTextView, ins: bool, del: bool) -> bool {
    match view {
        EditTextView::Accepted => !del,
        EditTextView::Original => !ins,
    }
}

/// One story's paragraphs in one view.
pub(crate) struct StoryView {
    pub story: String,
    pub view: EditTextView,
    pub paragraphs: Vec<ParagraphView>,
}

pub(crate) enum Lookup {
    Missing,
    One(usize),
    Many,
}

struct ParagraphBuilder {
    text: String,
    atoms: Vec<TextAtom>,
    spans: Vec<Span>,
    stamped: Vec<(u32, u32)>,
    embeds: Vec<(String, Option<MapRef>)>,
    start: u32,
    node_start: u32,
    view_len: u32,
}

impl ParagraphBuilder {
    fn new(start: u32) -> Self {
        Self {
            text: String::new(),
            atoms: Vec::new(),
            spans: Vec::new(),
            stamped: Vec::new(),
            embeds: Vec::new(),
            start,
            node_start: start,
            view_len: 0,
        }
    }

    fn push(&mut self, raw: u32, len: u32, stamped: bool, atom: bool) {
        match self.spans.last_mut() {
            Some(span)
                if !atom
                    && !span.atom
                    && span.stamped == stamped
                    && span.view + span.len == self.view_len
                    && span.raw + span.len == raw =>
            {
                span.len += len;
            }
            _ => self.spans.push(Span {
                view: self.view_len,
                raw,
                len,
                stamped,
                atom,
            }),
        }
        self.view_len += len;
    }

    fn stamp(&mut self, start: u32, end: u32) {
        match self.stamped.last_mut() {
            Some(last) if last.1 == start => last.1 = end,
            _ => self.stamped.push((start, end)),
        }
    }
}

impl StoryView {
    pub fn build<T: ReadTxn>(
        doc: &EditingDoc,
        txn: &T,
        story_id: &str,
        view: EditTextView,
    ) -> Option<Self> {
        Self::build_within(doc, txn, story_id, view, u32::MAX).map(|(view, _)| view)
    }

    /// The paragraphs of `story_id` that end within its first `limit` units, and whether they
    /// are all of it. Nothing past the last of them is copied.
    pub fn build_within<T: ReadTxn>(
        doc: &EditingDoc,
        txn: &T,
        story_id: &str,
        view: EditTextView,
        limit: u32,
    ) -> Option<(Self, bool)> {
        let story = story_ref(txn, story_id).ok()?;
        let chunks = doc.chunk_snapshot(story_id, &story, txn);
        let source = doc.source_metadata();
        let mut paragraphs = Vec::new();
        let mut current = ParagraphBuilder::new(0);
        let mut complete = true;
        for chunk in chunks.iter() {
            if chunk.end() > limit {
                complete = false;
                break;
            }
            let ins = chunk.attr_active(INS);
            let del = chunk.attr_active(DEL);
            match &chunk.kind {
                ChunkKind::Pilcrow(map) => {
                    let (para_id, properties) = capture_pilcrow(map, txn);
                    let properties: BTreeMap<String, Any> = properties.into_iter().collect();
                    let revision = ins
                        || del
                        || [PPR_INS, PPR_DEL].into_iter().any(|key| {
                            properties
                                .get(key)
                                .is_some_and(|value| !matches!(value, Any::Null | Any::Undefined))
                        });
                    let mark = Mark {
                        properties,
                        revision,
                    };
                    let style_id = match view {
                        EditTextView::Accepted => mark.style_id(),
                        EditTextView::Original => mark.original_style_id(),
                    };
                    let built = std::mem::replace(&mut current, ParagraphBuilder::new(chunk.end()));
                    let run_revisions = mark.run_revisions()
                        || source
                            .as_ref()
                            .is_some_and(|source| source.run_revision(story_id, &para_id));
                    paragraphs.push(ParagraphView {
                        para_id,
                        style_id,
                        text: built.text,
                        atoms: built.atoms,
                        spans: built.spans,
                        stamped: built.stamped,
                        start: built.start,
                        node_start: built.node_start,
                        pilcrow: chunk.start,
                        mark,
                        run_revisions,
                        embeds: built.embeds,
                    });
                }
                ChunkKind::Text(text) => {
                    if ins || del {
                        current.stamp(chunk.start, chunk.end());
                    }
                    if visible(view, ins, del) {
                        current.text.push_str(text);
                        current.push(chunk.start, chunk.len, ins || del, false);
                    }
                }
                ChunkKind::Embed(map) => {
                    let kind = map
                        .as_ref()
                        .and_then(|map| map_string(map, txn, KIND_KEY))
                        .unwrap_or_default();
                    if chunk.start == current.node_start && is_block_embed(&kind) {
                        current.node_start = chunk.end();
                        continue;
                    }
                    current.embeds.push((kind.clone(), map.clone()));
                    if ins || del {
                        current.stamp(chunk.start, chunk.end());
                    }
                    if visible(view, ins, del) {
                        current.atoms.push(TextAtom {
                            offset: current.view_len,
                            kind: atom_kind(&kind),
                        });
                        current.text.push('\u{FFFC}');
                        current.push(chunk.start, 1, ins || del, true);
                    }
                }
            }
        }
        Some((
            Self {
                story: story_id.to_owned(),
                view,
                paragraphs,
            },
            complete,
        ))
    }

    pub fn lookup(&self, para_id: &str) -> Lookup {
        let mut found = self
            .paragraphs
            .iter()
            .enumerate()
            .filter(|(_, paragraph)| paragraph.para_id == para_id)
            .map(|(index, _)| index);
        match (found.next(), found.next()) {
            (None, _) => Lookup::Missing,
            (Some(index), None) => Lookup::One(index),
            _ => Lookup::Many,
        }
    }

    /// A paragraph whose boundaries a pending paragraph-mark revision moves.
    pub fn structurally_revised(&self, index: usize) -> bool {
        self.paragraphs[index].mark.revision
            || index
                .checked_sub(1)
                .is_some_and(|previous| self.paragraphs[previous].mark.revision)
    }

    /// The final-state position of story index `raw`, clamped into the nearest paragraph.
    pub fn position_of_raw(&self, raw: u32) -> Option<TextPosition> {
        let paragraph = self
            .paragraphs
            .get(
                self.paragraphs
                    .partition_point(|paragraph| paragraph.pilcrow < raw),
            )
            .or_else(|| self.paragraphs.last())?;
        Some(TextPosition {
            para_id: paragraph.para_id.clone(),
            offset: paragraph.offset_of_raw(raw.min(paragraph.pilcrow)),
        })
    }

    pub fn range_of_raw(&self, start: u32, end: u32) -> Option<TextRange> {
        Some(TextRange {
            story: self.story.clone(),
            start: self.position_of_raw(start)?,
            end: self.position_of_raw(end)?,
            view: self.view,
        })
    }
}

/// Story projections for one transaction, built on first use.
pub(crate) struct Views<'a, T: ReadTxn> {
    doc: &'a EditingDoc,
    txn: &'a T,
    cache: HashMap<(String, EditTextView), Option<Rc<StoryView>>>,
    ownership: Option<Rc<Ownership>>,
    controls: Option<Rc<Inventory>>,
}

/// A resolved text selection in one paragraph of one view.
#[derive(Clone)]
pub(crate) struct Selection {
    pub view: Rc<StoryView>,
    pub paragraph: usize,
    pub start: u32,
    pub end: u32,
}

impl Selection {
    pub fn paragraph(&self) -> &ParagraphView {
        &self.view.paragraphs[self.paragraph]
    }

    pub fn text(&self) -> String {
        self.paragraph().slice(self.start, self.end)
    }

    pub fn range(&self) -> TextRange {
        let para_id = self.paragraph().para_id.clone();
        TextRange {
            story: self.view.story.clone(),
            start: TextPosition {
                para_id: para_id.clone(),
                offset: self.start,
            },
            end: TextPosition {
                para_id,
                offset: self.end,
            },
            view: self.view.view,
        }
    }
}

fn paragraph_target(target: &ParagraphTarget) -> EditTarget {
    EditTarget::Paragraph(target.clone())
}

impl<'a, T: ReadTxn> Views<'a, T> {
    pub fn new(doc: &'a EditingDoc, txn: &'a T) -> Self {
        Self {
            doc,
            txn,
            cache: HashMap::new(),
            ownership: None,
            controls: None,
        }
    }

    pub fn doc(&self) -> &'a EditingDoc {
        self.doc
    }

    pub fn txn(&self) -> &'a T {
        self.txn
    }

    pub fn story(&mut self, story: &str, view: EditTextView) -> Option<Rc<StoryView>> {
        let (doc, txn) = (self.doc, self.txn);
        self.cache
            .entry((story.to_owned(), view))
            .or_insert_with(|| StoryView::build(doc, txn, story, view).map(Rc::new))
            .clone()
    }

    #[cfg(test)]
    pub fn built(&self, story: &str, view: EditTextView) -> bool {
        self.cache.contains_key(&(story.to_owned(), view))
    }

    /// The story's projection when it is built or has at most `limit` units, else the
    /// uncached prefix [`StoryView::build_within`] reads, with whether it is complete.
    pub fn story_within(
        &mut self,
        story: &str,
        view: EditTextView,
        limit: u32,
    ) -> Option<(Rc<StoryView>, bool)> {
        let cached = self.cache.contains_key(&(story.to_owned(), view));
        let units = yrs::Text::len(&story_ref(self.txn, story).ok()?, self.txn);
        if cached || units <= limit {
            return self.story(story, view).map(|built| (built, true));
        }
        StoryView::build_within(self.doc, self.txn, story, view, limit)
            .map(|(built, complete)| (Rc::new(built), complete))
    }

    pub fn ownership(&mut self) -> Rc<Ownership> {
        let (doc, txn) = (self.doc, self.txn);
        Rc::clone(
            self.ownership
                .get_or_insert_with(|| Rc::new(Ownership::build(doc, txn))),
        )
    }

    /// Every content control of this state, read once.
    pub fn controls(&mut self) -> Result<Rc<Inventory>, EditFailure> {
        if let Some(controls) = &self.controls {
            return Ok(Rc::clone(controls));
        }
        let inventory = Rc::new(
            Inventory::build(self.doc, self.txn)
                .map_err(|export| failure(EditFailureCode::LimitExceeded, export.message, None))?,
        );
        self.controls = Some(Rc::clone(&inventory));
        Ok(inventory)
    }

    pub fn paragraph(
        &mut self,
        target: &ParagraphTarget,
        view: EditTextView,
    ) -> Result<(Rc<StoryView>, usize), EditFailure> {
        let Some(story) = self.story(&target.story, view) else {
            return Err(failure(
                EditFailureCode::MissingTarget,
                format!("story {:?} was not found", target.story),
                Some(paragraph_target(target)),
            ));
        };
        match story.lookup(&target.para_id) {
            Lookup::One(index) => Ok((story, index)),
            Lookup::Missing => Err(failure(
                EditFailureCode::MissingTarget,
                format!(
                    "paragraph {:?} was not found in story {:?}",
                    target.para_id, target.story
                ),
                Some(paragraph_target(target)),
            )),
            Lookup::Many => Err(failure(
                EditFailureCode::AmbiguousTarget,
                format!(
                    "paragraph id {:?} occurs more than once in story {:?}",
                    target.para_id, target.story
                ),
                Some(paragraph_target(target)),
            )),
        }
    }

    /// Resolves `target` against this state; the selection is always inside one paragraph.
    pub fn text(&mut self, target: &TextTarget) -> Result<Selection, EditFailure> {
        let described = || Some(EditTarget::from(target.clone()));
        match target {
            TextTarget::Paragraph(paragraph) => {
                let (view, index) = self.paragraph(paragraph, EditTextView::Accepted)?;
                let end = view.paragraphs[index].len();
                Ok(Selection {
                    view,
                    paragraph: index,
                    start: 0,
                    end,
                })
            }
            TextTarget::Range(range) => {
                if range.start.para_id != range.end.para_id {
                    return Err(failure(
                        EditFailureCode::Unsupported,
                        "a text range must start and end in one paragraph".to_owned(),
                        described(),
                    ));
                }
                let (view, index) = self.paragraph(
                    &ParagraphTarget {
                        story: range.story.clone(),
                        para_id: range.start.para_id.clone(),
                    },
                    range.view,
                )?;
                let paragraph = &view.paragraphs[index];
                let (start, end) = (range.start.offset, range.end.offset);
                if end < start {
                    return Err(failure(
                        EditFailureCode::InvalidStep,
                        format!("range end {end} precedes its start {start}"),
                        described(),
                    ));
                }
                if end > paragraph.len() {
                    return Err(failure(
                        EditFailureCode::InvalidStep,
                        format!(
                            "offset {end} exceeds the paragraph length {}",
                            paragraph.len()
                        ),
                        described(),
                    ));
                }
                if !paragraph.is_scalar_boundary(start) || !paragraph.is_scalar_boundary(end) {
                    return Err(failure(
                        EditFailureCode::InvalidStep,
                        "range offsets must not split a surrogate pair".to_owned(),
                        described(),
                    ));
                }
                Ok(Selection {
                    view,
                    paragraph: index,
                    start,
                    end,
                })
            }
            TextTarget::Search { text, within, view } => {
                if text.is_empty() {
                    return Err(failure(
                        EditFailureCode::InvalidStep,
                        "search text must not be empty".to_owned(),
                        described(),
                    ));
                }
                let (story, candidates) = self.scope(within, *view)?;
                if let Some(index) = candidates
                    .iter()
                    .copied()
                    .find(|index| story.structurally_revised(*index))
                {
                    return Err(failure(
                        EditFailureCode::TrackedRevisionConflict,
                        format!(
                            "paragraph {:?} has a pending paragraph-mark revision",
                            story.paragraphs[index].para_id
                        ),
                        described(),
                    ));
                }
                let mut hits = candidates.into_iter().flat_map(|index| {
                    story.paragraphs[index]
                        .occurrences(text)
                        .map(move |(start, end)| (index, start, end))
                });
                let Some((paragraph, start, end)) = hits.next() else {
                    return Err(failure(
                        EditFailureCode::MissingTarget,
                        format!("search text {text:?} was not found"),
                        described(),
                    ));
                };
                if hits.next().is_some() {
                    return Err(failure(
                        EditFailureCode::AmbiguousTarget,
                        format!("search text {text:?} occurs more than once"),
                        described(),
                    ));
                }
                drop(hits);
                if !matches!(
                    story.lookup(&story.paragraphs[paragraph].para_id),
                    Lookup::One(_)
                ) {
                    return Err(failure(
                        EditFailureCode::AmbiguousTarget,
                        format!(
                            "paragraph id {:?} occurs more than once in story {:?}",
                            story.paragraphs[paragraph].para_id, story.story
                        ),
                        described(),
                    ));
                }
                Ok(Selection {
                    view: story,
                    paragraph,
                    start,
                    end,
                })
            }
        }
    }

    fn scope(
        &mut self,
        within: &SearchScope,
        view: EditTextView,
    ) -> Result<(Rc<StoryView>, Vec<usize>), EditFailure> {
        match within {
            SearchScope::Story { story } => {
                let Some(projection) = self.story(story, view) else {
                    return Err(failure(
                        EditFailureCode::MissingTarget,
                        format!("story {story:?} was not found"),
                        None,
                    ));
                };
                let indices = (0..projection.paragraphs.len()).collect();
                Ok((projection, indices))
            }
            SearchScope::Paragraph(paragraph) => {
                let (projection, index) = self.paragraph(paragraph, view)?;
                Ok((projection, vec![index]))
            }
        }
    }

    /// Fails when a control owning `story` locks its content or a tracked structure owns it.
    pub fn check_story_writable(
        &mut self,
        story: &str,
        target: &EditTarget,
    ) -> Result<(), EditFailure> {
        let ownership = self.ownership();
        let chain = ownership.chain(story).map_err(|message| {
            failure(
                EditFailureCode::LimitExceeded,
                message,
                Some(target.clone()),
            )
        })?;
        if chain.iter().any(|owner| owner.content_locked()) {
            return Err(failure(
                EditFailureCode::LockedTarget,
                format!("story {story:?} is inside a content-locked control"),
                Some(target.clone()),
            ));
        }
        if chain.iter().any(|owner| owner.revision) {
            return Err(failure(
                EditFailureCode::TrackedRevisionConflict,
                format!("story {story:?} belongs to a tracked table or control revision"),
                Some(target.clone()),
            ));
        }
        Ok(())
    }

    fn check_story_readable(&mut self, story: &str) -> Result<(), EditFailure> {
        let ownership = self.ownership();
        let chain = ownership
            .chain(story)
            .map_err(|message| failure(EditFailureCode::LimitExceeded, message, None))?;
        if chain.iter().any(|owner| owner.revision) {
            return Err(failure(
                EditFailureCode::Unsupported,
                format!("story {story:?} belongs to a tracked table or control revision"),
                None,
            ));
        }
        Ok(())
    }
}

fn structural_read_failure(story: &StoryView, index: usize) -> EditFailure {
    let paragraph = &story.paragraphs[index];
    failure(
        EditFailureCode::Unsupported,
        format!(
            "paragraph {:?} has a pending paragraph-mark revision",
            paragraph.para_id
        ),
        Some(EditTarget::Paragraph(ParagraphTarget {
            story: story.story.clone(),
            para_id: paragraph.para_id.clone(),
        })),
    )
}

impl EditingDoc {
    fn read_scope<R>(
        &self,
        read: impl FnOnce(&mut Views<'_, yrs::Transaction<'_>>) -> Result<R, EditFailure>,
    ) -> Result<(DocumentVersion, R), EditRefusal> {
        let version = self.version();
        let txn = self.yrs_doc().transact();
        let mut views = Views::new(self, &txn);
        match read(&mut views) {
            Ok(result) => Ok((version, result)),
            Err(failure) => Err(refusal(version, failure)),
        }
    }

    /// Paragraph texts in one view, with the version they were read at.
    pub fn read_paragraphs(
        &self,
        request: &ReadParagraphsRequest,
    ) -> Result<ReadParagraphsResponse, EditRefusal> {
        let story_id = request.story.as_deref().unwrap_or("body");
        let (version, paragraphs) = self.read_scope(|views| {
            let Some(story) = views.story(story_id, request.view) else {
                return Err(failure(
                    EditFailureCode::MissingTarget,
                    format!("story {story_id:?} was not found"),
                    None,
                ));
            };
            views.check_story_readable(story_id)?;
            let indices: Vec<usize> = match &request.para_ids {
                None => (0..story.paragraphs.len()).collect(),
                Some(ids) => {
                    let mut indices = Vec::with_capacity(ids.len());
                    for para_id in ids {
                        let target = ParagraphTarget {
                            story: story_id.to_owned(),
                            para_id: para_id.clone(),
                        };
                        let (_, index) = views.paragraph(&target, request.view)?;
                        indices.push(index);
                    }
                    indices.sort_unstable();
                    indices.dedup();
                    indices
                }
            };
            if indices.len() > READ_PARAGRAPH_LIMIT {
                return Err(failure(
                    EditFailureCode::LimitExceeded,
                    format!("a read returns at most {READ_PARAGRAPH_LIMIT} paragraphs"),
                    None,
                ));
            }
            let mut text_units = 0usize;
            let mut paragraphs = Vec::with_capacity(indices.len());
            for index in indices {
                if story.structurally_revised(index) {
                    return Err(structural_read_failure(&story, index));
                }
                let paragraph = &story.paragraphs[index];
                text_units += paragraph.len() as usize;
                if text_units > READ_TEXT_LIMIT {
                    return Err(failure(
                        EditFailureCode::LimitExceeded,
                        format!("a read returns at most {READ_TEXT_LIMIT} UTF-16 units of text"),
                        None,
                    ));
                }
                paragraphs.push(paragraph.record(story_id));
            }
            Ok(paragraphs)
        })?;
        Ok(ReadParagraphsResponse {
            version,
            view: request.view,
            paragraphs,
        })
    }

    /// Exact, case-sensitive, paragraph-local matches; overlapping occurrences count separately.
    pub fn find_text(&self, request: &FindTextRequest) -> Result<FindTextResponse, EditRefusal> {
        let limit = request
            .limit
            .unwrap_or(FIND_DEFAULT_LIMIT)
            .min(FIND_MAX_LIMIT) as usize;
        let (version, (matches, truncated)) = self.read_scope(|views| {
            if request.text.is_empty() {
                return Err(failure(
                    EditFailureCode::InvalidStep,
                    "search text must not be empty".to_owned(),
                    None,
                ));
            }
            let (story, candidates) = views.scope(&request.within, request.view)?;
            views.check_story_readable(&story.story)?;
            let mut matches = Vec::new();
            for index in candidates {
                if story.structurally_revised(index) {
                    return Err(structural_read_failure(&story, index));
                }
                let paragraph = &story.paragraphs[index];
                for (start, end) in paragraph.occurrences(&request.text) {
                    if matches.len() == limit {
                        return Ok((matches, true));
                    }
                    matches.push(TextMatch {
                        text: request.text.clone(),
                        range: Selection {
                            view: Rc::clone(&story),
                            paragraph: index,
                            start,
                            end,
                        }
                        .range(),
                    });
                }
            }
            Ok((matches, false))
        })?;
        Ok(FindTextResponse {
            version,
            matches,
            truncated,
        })
    }

    /// The story span an annotation over `target` covers, including hidden revision units inside
    /// it; `None` when the target selects no text.
    pub fn resolve_text_span(
        &self,
        target: &TextTarget,
    ) -> Result<Option<StoryRange>, EditRefusal> {
        self.read_scope(|views| {
            let selection = views.text(target)?;
            if selection.view.structurally_revised(selection.paragraph) {
                return Err(failure(
                    EditFailureCode::TrackedRevisionConflict,
                    format!(
                        "paragraph {:?} has a pending paragraph-mark revision",
                        selection.paragraph().para_id
                    ),
                    Some(EditTarget::from(target.clone())),
                ));
            }
            if selection.start == selection.end {
                return Ok(None);
            }
            let paragraph = selection.paragraph();
            Ok(Some(StoryRange::new(
                selection.view.story.clone(),
                paragraph.raw_at(selection.start),
                paragraph.raw_after(selection.end),
            )))
        })
        .map(|(_, span)| span)
    }

    /// Resolves `target` and formats it in one call, as the legacy agent helpers do. Refuses a
    /// target inside a content-locked control; an empty selection changes nothing.
    pub fn format_text_target(
        &self,
        target: &TextTarget,
        delta: &InlineFormatDelta,
    ) -> OpResult<Result<(), EditRefusal>> {
        let span = match self.resolve_text_span(target) {
            Ok(span) => span,
            Err(refusal) => return Ok(Err(refusal)),
        };
        let Some(span) = span else {
            return Ok(Ok(()));
        };
        if let Err(refusal) = self.writable(&span.story, target) {
            return Ok(Err(refusal));
        }
        self.format_range(&EditCtx::local(String::new(), String::new()), span, delta)?;
        Ok(Ok(()))
    }

    /// Resolves `target` and anchors side-map comment `id` over it in one call.
    pub fn comment_text_target(
        &self,
        target: &TextTarget,
        id: &str,
        author: &str,
        date: &str,
        body: Any,
    ) -> OpResult<Result<(), EditRefusal>> {
        let span = match self.resolve_text_span(target) {
            Ok(span) => span,
            Err(refusal) => return Ok(Err(refusal)),
        };
        let Some(span) = span else {
            return Ok(Err(refusal(
                self.version(),
                failure(
                    EditFailureCode::InvalidStep,
                    "a comment target must select text".to_owned(),
                    Some(EditTarget::from(target.clone())),
                ),
            )));
        };
        self.apply_raw_ops(
            &span.story,
            vec![RawOp::SetComment {
                id: id.to_owned(),
                ranges: vec![(span.start, span.end)],
                author: author.to_owned(),
                date: date.to_owned(),
                body,
            }],
            &EditCtx::local(String::new(), String::new()),
        )?;
        Ok(Ok(()))
    }

    fn writable(&self, story: &str, target: &TextTarget) -> Result<(), EditRefusal> {
        self.read_scope(|views| {
            views.check_story_writable(story, &EditTarget::from(target.clone()))
        })
        .map(|_| ())
    }

    /// Projected texts around a paragraph-keyed selection: the start paragraph's text, the text
    /// before and after the selection, and the selection itself with `\n` between paragraphs.
    pub fn selection_text(&self, range: &LocRange, view: EditTextView) -> OpResult<SelectionInfo> {
        let txn = self.yrs_doc().transact();
        let story_id = &range.start.story;
        story_ref(&txn, story_id)?;
        let story = StoryView::build(self, &txn, story_id, view)
            .ok_or_else(|| OpError::UnknownStory(story_id.clone()))?;
        let locate = |loc: &Loc| -> OpResult<(usize, u32)> {
            let index = story
                .paragraphs
                .iter()
                .position(|paragraph| paragraph.para_id == loc.para)
                .ok_or_else(|| OpError::UnknownPara(loc.para.clone()))?;
            let paragraph = &story.paragraphs[index];
            let len = paragraph.pilcrow - paragraph.start;
            if loc.offset > len {
                return Err(OpError::OutOfBounds {
                    index: loc.offset,
                    len,
                });
            }
            Ok((index, paragraph.offset_of_raw(paragraph.start + loc.offset)))
        };
        let (mut first, mut first_offset) = locate(&range.start)?;
        let (mut last, mut last_offset) = locate(&range.end)?;
        if (last, last_offset) < (first, first_offset) {
            std::mem::swap(&mut first, &mut last);
            std::mem::swap(&mut first_offset, &mut last_offset);
        }
        let start = &story.paragraphs[first];
        let end = &story.paragraphs[last];
        let selected = if first == last {
            start.slice(first_offset, last_offset)
        } else {
            let mut parts = vec![start.slice(first_offset, start.len())];
            parts.extend(
                story.paragraphs[first + 1..last]
                    .iter()
                    .map(|paragraph| paragraph.text.clone()),
            );
            parts.push(end.slice(0, last_offset));
            parts.join("\n")
        };
        Ok(SelectionInfo {
            para_id: start.para_id.clone(),
            selected_text: selected,
            paragraph_text: start.text.clone(),
            before: start.slice(0, first_offset),
            after: end.slice(last_offset, end.len()),
        })
    }
}

/// Resolves the comment anchors of `story` to story indices, skipping ones that no longer resolve.
pub(crate) fn comment_spans<T: ReadTxn>(txn: &T, story: &str) -> Vec<(u32, u32)> {
    let Some(comments) = txn.get_map(crate::COMMENTS) else {
        return Vec::new();
    };
    let mut spans = Vec::new();
    for (_, value) in comments.iter(txn) {
        let Ok(comment) = value.cast::<MapRef>() else {
            continue;
        };
        let Some(Out::Any(Any::Array(anchors))) = comment.get(txn, "anchors") else {
            continue;
        };
        for anchor in anchors.iter() {
            let Ok(anchor) = crate::decode_anchor(anchor) else {
                continue;
            };
            if anchor.story != story {
                continue;
            }
            if let (Some(start), Some(end)) =
                (anchor.start.get_offset(txn), anchor.end.get_offset(txn))
            {
                spans.push((start.index, end.index));
            }
        }
    }
    spans
}

/// Paragraph ids a field embed binds as its cached result blocks.
pub(crate) fn field_result_blocks<T: ReadTxn>(story: &StoryView, txn: &T) -> Vec<String> {
    let mut ids = Vec::new();
    for paragraph in &story.paragraphs {
        for (kind, map) in &paragraph.embeds {
            if kind != "field" {
                continue;
            }
            let Some(map) = map else { continue };
            if let Some(Out::Any(Any::Array(blocks))) = map.get(txn, "fieldResultBlocks") {
                ids.extend(blocks.iter().filter_map(|block| match block {
                    Any::String(id) => Some(id.to_string()),
                    _ => None,
                }));
            }
        }
    }
    ids
}

impl From<TextTarget> for EditTarget {
    fn from(target: TextTarget) -> Self {
        match target {
            TextTarget::Paragraph(paragraph) => EditTarget::Paragraph(paragraph),
            TextTarget::Range(range) => EditTarget::Range(range),
            TextTarget::Search { text, within, view } => EditTarget::Search { text, within, view },
        }
    }
}
