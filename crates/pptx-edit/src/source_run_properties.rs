use std::collections::{BTreeMap, HashMap, HashSet};
use std::ops::Range;

use serde_json::Value;
use yrs::branch::{Branch, BranchID, BranchPtr};
use yrs::types::Attrs;
use yrs::types::text::YChange;
use yrs::{
    Any, Assoc, ClientID, Doc, ID, IdSet, IndexedSequence, Map, MapRef, Out, ReadTxn, Snapshot,
    StateVector, StickyIndex, Text, TextRef, Transact,
};

use crate::deck::{SourceImport, map_string};
use crate::story::style_from_attrs;
use crate::{
    BOOTSTRAP_CLIENT_ID, DeckSession, EditError, EditResult, META, MIGRATE_ORIGIN, PARA_ID,
    ShapeSnapshot, StorySnapshot, TextStyle,
};

#[derive(Clone, Copy)]
pub(crate) enum SourceProperty {
    Baseline,
    Spacing,
    Caps,
    Color,
}

impl SourceProperty {
    /// The pending flag, the package keys the source restores, and the run attribute.
    fn keys(self) -> (&'static str, &'static [&'static str], &'static str) {
        match self {
            Self::Baseline => ("baselinesPendingSource", &["baselinePct"], "baseline"),
            Self::Spacing => ("spacingPendingSource", &["spacingPt"], "spacing"),
            Self::Caps => ("capsPendingSource", &["caps"], "caps"),
            Self::Color => (
                "colorsPendingSource",
                &["colorMap", "colorMapOverride"],
                "color",
            ),
        }
    }

    /// Whether surviving text from any client counts, as it did before caps and
    /// colours were recovered; caps and colours take only text the seed wrote.
    fn matches_any_client(self) -> bool {
        matches!(self, Self::Baseline | Self::Spacing)
    }

    fn value(self, style: &TextStyle) -> Option<Any> {
        match self {
            Self::Baseline => style.baseline_pct.map(Any::Number),
            Self::Spacing => style.spacing_pt.map(Any::Number),
            Self::Caps => style.caps.map(|caps| Any::from(caps.as_attribute())),
            Self::Color => style.color.as_deref().map(Any::from),
        }
    }
}

/// Characters a pass could not recover, by story, as `(client, clock)`.
type Unresolved = BTreeMap<String, Vec<(u64, u32)>>;

/// Which stories a pass still has to recover: all of them after migration, or
/// only the characters an earlier pass left unresolved.
enum Pending {
    All,
    Only(Unresolved),
}

/// The run attribute, baseline or spacing, whose recovery from the source left
/// characters unresolved, still in the deck, that a save would strip it from.
/// Caps and colours were never kept before, so their unresolved characters save
/// as they are.
pub(crate) fn unrecovered_attribute(doc: &Doc) -> Option<&'static str> {
    let txn = doc.transact();
    let meta = txn.get_map(META)?;
    [SourceProperty::Baseline, SourceProperty::Spacing]
        .into_iter()
        .map(SourceProperty::keys)
        .find(|(pending_key, ..)| match meta.get(&txn, pending_key) {
            Some(Out::Any(Any::String(json))) => {
                serde_json::from_str::<Unresolved>(&json).map_or(true, |stories| {
                    stories
                        .values()
                        .flatten()
                        .any(|&(client, clock)| alive(&txn, ID::new(ClientID::new(client), clock)))
                })
            }
            _ => false,
        })
        .map(|(_, _, attribute)| attribute)
}

/// Whether the character `id` is still in its story rather than deleted.
fn alive<T: ReadTxn>(txn: &T, id: ID) -> bool {
    StickyIndex::from_id(id, Assoc::After)
        .get_offset(txn)
        .is_some_and(|offset| {
            TextRef::from(offset.branch)
                .sticky_index(txn, offset.index, Assoc::After)
                .and_then(|index| index.id().copied())
                == Some(id)
        })
}

pub(crate) fn import_source(
    session: &DeckSession,
    import: &mut SourceImport<'_>,
    property: SourceProperty,
) -> EditResult<()> {
    let (pending_key, json_keys, attribute) = property.keys();
    let pending = {
        let txn = session.doc.transact();
        match txn
            .get_map(META)
            .and_then(|meta| meta.get(&txn, pending_key))
        {
            Some(Out::Any(Any::Bool(true))) => Pending::All,
            Some(Out::Any(Any::String(json))) => Pending::Only(
                serde_json::from_str(&json).map_err(|error| EditError::Json(error.to_string()))?,
            ),
            _ => return Ok(()),
        }
    };
    let source_json =
        serde_json::to_value(import.source).map_err(|error| EditError::Json(error.to_string()))?;
    if !json_keys.iter().any(|key| has_property(&source_json, key)) {
        return Ok(());
    }
    let legacy_snapshot = match property {
        SourceProperty::Color => Some(crate::deck::legacy_baseline_snapshot(import.source)?),
        _ => None,
    };
    let mut legacy = HashMap::new();
    for slide in legacy_snapshot.iter().flat_map(|snapshot| &snapshot.slides) {
        collect_stories(&slide.shapes, &mut legacy);
    }
    let mut sources = Vec::new();
    for slide in &import.source_snapshot()?.slides {
        collect_source_tokens(&slide.shapes, &legacy, property, &mut sources);
    }
    let mut package = serde_json::to_value(&import.package)
        .map_err(|error| EditError::Json(error.to_string()))?;
    for key in json_keys {
        merge_property(&mut package, &source_json, key);
    }
    import.package =
        serde_json::from_value(package).map_err(|error| EditError::Json(error.to_string()))?;
    let mut patches = Vec::new();
    let mut unresolved = Unresolved::new();
    for (id, source) in &sources {
        let only = match &pending {
            Pending::All => None,
            Pending::Only(stories) => match stories.get(id) {
                Some(ids) => Some(
                    ids.iter()
                        .map(|(client, clock)| ID::new(ClientID::new(*client), *clock))
                        .collect::<HashSet<_>>(),
                ),
                None => continue,
            },
        };
        let Some((story, replayed)) = import.recovery.story(&session.doc, session.package(), id)?
        else {
            continue;
        };
        let (ranges, left) = recover_story(story, replayed, source, property, only.as_ref());
        for (start, end, value) in ranges {
            patches.push((story.text.clone(), start, end, value));
        }
        if !left.is_empty() {
            unresolved.insert(
                id.clone(),
                left.into_iter()
                    .map(|id| (id.client.get(), id.clock))
                    .collect(),
            );
        }
    }
    let mut txn = session.doc.transact_mut_with(MIGRATE_ORIGIN);
    for (story, start, end, value) in patches {
        story.format(
            &mut txn,
            start,
            end - start,
            Attrs::from([(attribute.into(), value)]),
        );
    }
    let meta = txn
        .get_map(META)
        .ok_or_else(|| EditError::InvalidState("missing metadata".into()))?;
    if unresolved.is_empty() {
        meta.remove(&mut txn, pending_key);
    } else {
        let json = serde_json::to_string(&unresolved)
            .map_err(|error| EditError::Json(error.to_string()))?;
        meta.insert(&mut txn, pending_key, json);
    }
    Ok(())
}

#[derive(Clone, Copy, PartialEq)]
enum Token {
    Char(char),
    Paragraph(usize),
    Break,
}

struct SourceToken {
    token: Token,
    value: Option<Any>,
    legacy: Option<Any>,
}

/// A story's live content, read once per attachment and shared by every pass.
struct Story {
    text: TextRef,
    /// The clock of the item holding the story, when the seed created it.
    seed: Option<u32>,
    units: Vec<Unit>,
    styles: Vec<TextStyle>,
}

/// A character with the item that holds it, or a paragraph end.
struct Unit {
    token: Token,
    start: u32,
    end: u32,
    id: Option<ID>,
    style: usize,
}

/// A story as the seed wrote it: its tokens, by clock relative to the story.
struct Replayed {
    tokens: Vec<Token>,
    positions: HashMap<u32, usize>,
}

/// What the recovery passes of one attachment share.
#[derive(Default)]
pub(crate) struct RecoveryCache {
    visible: Option<Visible>,
    stories: HashMap<String, Option<Story>>,
    replay: Option<(Doc, Snapshot)>,
    replayed: HashMap<String, Option<Replayed>>,
}

/// The doc's state as a snapshot yrs 0.27.3 can split, which leaves out clients
/// with at most two clock ticks, and those clients' characters, located by id.
/// It carries no deletions: collected items hold no text, and anything a
/// snapshot shows that the plain diff does not fails the merge in `read_story`.
struct Visible {
    snapshot: Snapshot,
    tiny: Vec<(BranchPtr, u32, ID)>,
}

impl RecoveryCache {
    /// The story's live content and its replayed seed, the stored package
    /// seeded again to reproduce the legacy seed's ids.
    fn story(
        &mut self,
        doc: &Doc,
        package: &pptx_parse::PptxPackage,
        story_id: &str,
    ) -> EditResult<Option<(&Story, Option<&Replayed>)>> {
        if self.replay.is_none() {
            let replay = crate::doc_with_client_id(BOOTSTRAP_CLIENT_ID);
            crate::deck::seed_doc(&replay, package, "")?;
            let state = replay.transact().state_vector();
            self.replay = Some((replay, Snapshot::new(state, IdSet::new())));
        }
        if !self.stories.contains_key(story_id) {
            let visible = self.visible.get_or_insert_with(|| visible(doc));
            let story = read_story(doc, story_id, visible);
            self.stories.insert(story_id.to_owned(), story);
        }
        if let Some((replay, snapshot)) = &self.replay {
            self.replayed
                .entry(story_id.to_owned())
                .or_insert_with(|| read_replayed(replay, snapshot, story_id));
        }
        Ok(self.stories[story_id]
            .as_ref()
            .map(|story| (story, self.replayed[story_id].as_ref())))
    }
}

fn visible(doc: &Doc) -> Visible {
    let txn = doc.transact();
    let mut state = StateVector::default();
    let mut tiny = Vec::new();
    for (&client, &clock) in txn.state_vector().iter() {
        if clock > 2 {
            state.set_max(client, clock);
            continue;
        }
        for clock in 0..clock {
            let id = ID::new(client, clock);
            if let Some(offset) = StickyIndex::from_id(id, Assoc::After).get_offset(&txn)
                && TextRef::from(offset.branch)
                    .sticky_index(&txn, offset.index, Assoc::After)
                    .and_then(|index| index.id().copied())
                    == Some(id)
            {
                tiny.push((offset.branch, offset.index, id));
            }
        }
    }
    Visible {
        snapshot: Snapshot::new(state, IdSet::new()),
        tiny,
    }
}

fn read_story(doc: &Doc, story_id: &str, visible: &Visible) -> Option<Story> {
    let mut txn = doc.transact_mut_with(MIGRATE_ORIGIN);
    let text = txn
        .get_map(crate::STORIES)?
        .get(&txn, story_id)?
        .cast::<TextRef>()
        .ok()?;
    let seed = match <TextRef as AsRef<Branch>>::as_ref(&text).id() {
        BranchID::Nested(id) if id.client == ClientID::new(BOOTSTRAP_CLIENT_ID) => Some(id.clock),
        _ => None,
    };
    let mut identified = Vec::new();
    for diff in text.diff_range(
        &mut txn,
        Some(&visible.snapshot),
        Some(&Snapshot::default()),
        YChange::identity,
    ) {
        match (&diff.insert, diff.ychange) {
            (Out::Any(Any::String(chunk)), Some(change)) => {
                let mut clock = change.id.clock;
                for unit in chunk.chars() {
                    identified.push(Some((unit, ID::new(change.id.client, clock))));
                    clock += unit.len_utf16() as u32;
                }
            }
            _ => identified.push(None),
        }
    }
    let branch = BranchPtr::from(<TextRef as AsRef<Branch>>::as_ref(&text));
    let tiny: HashMap<_, _> = visible
        .tiny
        .iter()
        .filter(|(owner, ..)| *owner == branch)
        .map(|(_, index, id)| (*index, *id))
        .collect();
    let prefix = format!("para:{story_id}:");
    let mut identified = identified.into_iter();
    let mut matched = true;
    let mut units = Vec::new();
    let mut styles = Vec::new();
    let mut offset = 0;
    for diff in text.diff(&txn, YChange::identity) {
        match &diff.insert {
            Out::Any(Any::String(chunk)) => {
                styles.push(style_from_attrs(diff.attributes.as_deref()));
                for unit in chunk.chars() {
                    let end = offset + unit.len_utf16() as u32;
                    let id = match tiny.get(&offset) {
                        Some(id) => Some(*id),
                        None => match identified.next() {
                            Some(Some((found, id))) if found == unit => Some(id),
                            _ => {
                                matched = false;
                                None
                            }
                        },
                    };
                    units.push(Unit {
                        token: Token::Char(unit),
                        start: offset,
                        end,
                        id,
                        style: styles.len() - 1,
                    });
                    offset = end;
                }
            }
            other => {
                matched &= matches!(identified.next(), Some(None));
                let (token, id) = match other {
                    Out::YMap(pilcrow) => (pilcrow_token(pilcrow, &txn, &prefix), item_id(pilcrow)),
                    _ => (Token::Break, None),
                };
                units.push(Unit {
                    token,
                    start: offset,
                    end: offset + 1,
                    id,
                    style: 0,
                });
                offset += 1;
            }
        }
    }
    if !matched || identified.next().is_some() {
        for unit in units
            .iter_mut()
            .filter(|unit| matches!(unit.token, Token::Char(_)))
        {
            unit.id = text
                .sticky_index(&txn, unit.start, Assoc::After)
                .and_then(|index| index.id().copied());
        }
    }
    Some(Story {
        text,
        seed,
        units,
        styles,
    })
}

fn read_replayed(replay: &Doc, snapshot: &Snapshot, story_id: &str) -> Option<Replayed> {
    let mut txn = replay.transact_mut();
    let text = txn
        .get_map(crate::STORIES)?
        .get(&txn, story_id)?
        .cast::<TextRef>()
        .ok()?;
    let BranchID::Nested(base) = <TextRef as AsRef<Branch>>::as_ref(&text).id() else {
        return None;
    };
    let prefix = format!("para:{story_id}:");
    let mut tokens = Vec::new();
    let mut positions = HashMap::new();
    for diff in text.diff_range(
        &mut txn,
        Some(snapshot),
        Some(&Snapshot::default()),
        YChange::identity,
    ) {
        match (&diff.insert, diff.ychange) {
            (Out::Any(Any::String(chunk)), Some(change)) => {
                let mut clock = change.id.clock;
                for unit in chunk.chars() {
                    positions.insert(clock - base.clock, tokens.len());
                    tokens.push(Token::Char(unit));
                    clock += unit.len_utf16() as u32;
                }
            }
            (Out::YMap(pilcrow), _) => {
                if let Some(id) = item_id(pilcrow) {
                    positions.insert(id.clock - base.clock, tokens.len());
                }
                tokens.push(pilcrow_token(pilcrow, &txn, &prefix));
            }
            _ => tokens.push(Token::Break),
        }
    }
    Some(Replayed { tokens, positions })
}

fn item_id(pilcrow: &MapRef) -> Option<ID> {
    match <MapRef as AsRef<Branch>>::as_ref(pilcrow).id() {
        BranchID::Nested(id) => Some(id),
        BranchID::Root(_) => None,
    }
}

fn pilcrow_token<T: ReadTxn>(pilcrow: &MapRef, txn: &T, prefix: &str) -> Token {
    map_string(pilcrow, txn, PARA_ID)
        .and_then(|id| id.strip_prefix(prefix)?.parse().ok())
        .map_or(Token::Break, Token::Paragraph)
}

/// The ranges to write and the characters left unresolved. A seeded character
/// takes its source position from its clock when the replayed seed agrees with
/// the source; otherwise from its in-order match. Baseline and spacing also
/// pair the text other clients left between those positions with the source
/// it stands in for.
fn recover_story(
    story: &Story,
    replayed: Option<&Replayed>,
    source: &[SourceToken],
    property: SourceProperty,
    only: Option<&HashSet<ID>>,
) -> (Vec<(u32, u32, Any)>, Vec<ID>) {
    let bootstrap = ClientID::new(BOOTSTRAP_CLIENT_ID);
    let considered = |unit: &Unit| {
        matches!(unit.token, Token::Char(_))
            && unit
                .id
                .is_some_and(|id| only.is_none_or(|only| only.contains(&id)))
    };
    let seeded = |unit: &Unit| considered(unit) && unit.id.is_some_and(|id| id.client == bootstrap);
    let foreign = |unit: &Unit| property.matches_any_client() && considered(unit) && !seeded(unit);
    let value = |unit: &Unit| property.value(&story.styles[unit.style]);
    let mut ranges = Vec::new();
    let mut unresolved = Vec::new();
    if let Some(positions) = exact_positions(story, replayed, source) {
        let mut from = 0;
        let mut gap = Vec::new();
        for (unit, position) in story.units.iter().zip(positions) {
            match position {
                Some(position) if position >= from => {
                    let window = from..position;
                    recover_gap(
                        source,
                        window,
                        &gap,
                        property,
                        story,
                        &mut ranges,
                        &mut unresolved,
                    );
                    gap.clear();
                    from = position + 1;
                    let candidate = &source[position];
                    if seeded(unit)
                        && candidate.value != candidate.legacy
                        && value(unit) == candidate.legacy
                    {
                        push_range(&mut ranges, unit, candidate);
                    }
                }
                None if foreign(unit) => gap.push(unit),
                _ => {}
            }
        }
        let window = from..source.len();
        recover_gap(
            source,
            window,
            &gap,
            property,
            story,
            &mut ranges,
            &mut unresolved,
        );
        return (ranges, unresolved);
    }
    let mut paragraphs = HashMap::new();
    let mut changes = Vec::with_capacity(source.len());
    let mut previous = None;
    let mut count = 0;
    for (position, token) in source.iter().enumerate() {
        match token.token {
            Token::Paragraph(index) => {
                paragraphs.insert(index, position);
            }
            _ => {
                let values = (&token.value, &token.legacy);
                count += usize::from(previous.is_some_and(|previous| previous != values));
                previous = Some(values);
            }
        }
        changes.push(count);
    }
    let mut recover_window = |window: Range<usize>, stretch: &[&Unit]| {
        if property.matches_any_client() {
            recover_gap(
                source,
                window,
                stretch,
                property,
                story,
                &mut ranges,
                &mut unresolved,
            );
        } else {
            let stretch: Vec<_> = stretch.iter().map(|unit| (*unit, value(unit))).collect();
            recover_stretch(
                source,
                &changes,
                window,
                &stretch,
                &mut ranges,
                &mut unresolved,
            );
        }
    };
    let mut from = 0;
    let mut stretch = Vec::new();
    for unit in &story.units {
        match unit.token {
            Token::Char(_) if seeded(unit) || foreign(unit) => stretch.push(unit),
            Token::Paragraph(index) => {
                if let Some(&until) = paragraphs.get(&index).filter(|until| **until >= from) {
                    recover_window(from..until, &stretch);
                    from = until + 1;
                    stretch.clear();
                }
            }
            _ => {}
        }
    }
    recover_window(from..source.len(), &stretch);
    (ranges, unresolved)
}

/// Pairs `stretch` with the source it stands in for in `window` by the longest
/// common subsequence, as whole stories were paired before; past the bound its
/// characters stay unresolved.
fn recover_gap(
    source: &[SourceToken],
    window: Range<usize>,
    stretch: &[&Unit],
    property: SourceProperty,
    story: &Story,
    ranges: &mut Vec<(u32, u32, Any)>,
    unresolved: &mut Vec<ID>,
) {
    if stretch.is_empty() || window.is_empty() {
        return;
    }
    let window = &source[window];
    let Some(pairs) = common_pairs(window, stretch) else {
        unresolved.extend(stretch.iter().filter_map(|unit| unit.id));
        return;
    };
    for (position, index) in pairs {
        let (candidate, unit) = (&window[position], stretch[index]);
        if candidate.value != candidate.legacy
            && property.value(&story.styles[unit.style]) == candidate.legacy
        {
            push_range(ranges, unit, candidate);
        }
    }
}

/// The longest common subsequence of tokens: equal ends pair directly, and
/// the rest is diffed, `None` past four million cells.
fn common_pairs(window: &[SourceToken], stretch: &[&Unit]) -> Option<Vec<(usize, usize)>> {
    let same = |i: usize, j: usize| window[i].token == stretch[j].token;
    let mut prefix = 0;
    while prefix < window.len().min(stretch.len()) && same(prefix, prefix) {
        prefix += 1;
    }
    let mut suffix = 0;
    while suffix < window.len().min(stretch.len()) - prefix
        && same(window.len() - suffix - 1, stretch.len() - suffix - 1)
    {
        suffix += 1;
    }
    let rows = window.len() - prefix - suffix + 1;
    let cols = stretch.len() - prefix - suffix + 1;
    let cells = rows.checked_mul(cols).filter(|cells| *cells <= 4_000_000)?;
    let mut lengths = vec![0u32; cells];
    for i in (0..rows - 1).rev() {
        for j in (0..cols - 1).rev() {
            lengths[i * cols + j] = if same(prefix + i, prefix + j) {
                lengths[(i + 1) * cols + j + 1] + 1
            } else {
                lengths[(i + 1) * cols + j].max(lengths[i * cols + j + 1])
            };
        }
    }
    let mut pairs: Vec<_> = (0..prefix).map(|i| (i, i)).collect();
    let (mut i, mut j) = (0, 0);
    while i + 1 < rows && j + 1 < cols {
        if same(prefix + i, prefix + j) {
            pairs.push((prefix + i, prefix + j));
            i += 1;
            j += 1;
        } else if lengths[(i + 1) * cols + j] >= lengths[i * cols + j + 1] {
            i += 1;
        } else {
            j += 1;
        }
    }
    pairs.extend((0..suffix).map(|i| (window.len() - suffix + i, stretch.len() - suffix + i)));
    Some(pairs)
}

/// Each seeded character's source position by clock, or `None` when the
/// replayed seed does not reproduce this story, its paragraph ends included,
/// or the source text.
fn exact_positions(
    story: &Story,
    replayed: Option<&Replayed>,
    source: &[SourceToken],
) -> Option<Vec<Option<usize>>> {
    let replayed = replayed?;
    if replayed.tokens.len() != source.len()
        || replayed
            .tokens
            .iter()
            .zip(source)
            .any(|(token, source)| *token != source.token)
    {
        return None;
    }
    let bootstrap = ClientID::new(BOOTSTRAP_CLIENT_ID);
    story
        .units
        .iter()
        .map(|unit| match unit.id {
            Some(id) if id.client == bootstrap => {
                let position = *replayed
                    .positions
                    .get(&id.clock.checked_sub(story.seed?)?)?;
                (replayed.tokens[position] == unit.token).then_some(Some(position))
            }
            _ => Some(None),
        })
        .collect()
}

fn push_range(ranges: &mut Vec<(u32, u32, Any)>, unit: &Unit, candidate: &SourceToken) {
    let value = candidate.value.clone().unwrap_or(Any::Null);
    match ranges.last_mut() {
        Some((_, end, last)) if *end == unit.start && *last == value => *end = unit.end,
        _ => ranges.push((unit.start, unit.end, value)),
    }
}

/// A character's candidate source positions run from its earliest to its latest
/// in-order match within `window`; it is written when they all carry the same
/// values and left unresolved when they do not.
fn recover_stretch(
    source: &[SourceToken],
    changes: &[usize],
    window: Range<usize>,
    stretch: &[(&Unit, Option<Any>)],
    ranges: &mut Vec<(u32, u32, Any)>,
    unresolved: &mut Vec<ID>,
) {
    let from = window.start;
    let window = &source[window];
    let (Some(earliest), Some(latest)) = (earliest(window, stretch), latest(window, stretch))
    else {
        unresolved.extend(stretch.iter().filter_map(|(unit, _)| unit.id));
        return;
    };
    for (((unit, value), earliest), latest) in stretch.iter().zip(earliest).zip(latest) {
        let (earliest, latest) = (from + earliest, from + latest);
        let candidate = &source[earliest];
        if changes[earliest] != changes[latest] {
            unresolved.extend(unit.id);
        } else if candidate.value != candidate.legacy && *value == candidate.legacy {
            push_range(ranges, unit, candidate);
        }
    }
}

fn earliest(window: &[SourceToken], stretch: &[(&Unit, Option<Any>)]) -> Option<Vec<usize>> {
    let mut positions = Vec::with_capacity(stretch.len());
    let mut next = 0;
    for (unit, _) in stretch {
        next += window[next..]
            .iter()
            .position(|source| source.token == unit.token)?;
        positions.push(next);
        next += 1;
    }
    Some(positions)
}

fn latest(window: &[SourceToken], stretch: &[(&Unit, Option<Any>)]) -> Option<Vec<usize>> {
    let mut positions = vec![0; stretch.len()];
    let mut end = window.len();
    for (position, (unit, _)) in positions.iter_mut().zip(stretch).rev() {
        end = window[..end]
            .iter()
            .rposition(|source| source.token == unit.token)?;
        *position = end;
    }
    Some(positions)
}

fn collect_stories<'a>(
    shapes: &'a [ShapeSnapshot],
    stories: &mut HashMap<&'a str, &'a StorySnapshot>,
) {
    for shape in shapes {
        for story in &shape.text_stories {
            stories.insert(&story.id, story);
        }
        collect_stories(&shape.children, stories);
    }
}

/// Tokenizes each source story whose values differ anywhere from what a legacy seed stored.
fn collect_source_tokens(
    shapes: &[ShapeSnapshot],
    legacy: &HashMap<&str, &StorySnapshot>,
    property: SourceProperty,
    sources: &mut Vec<(String, Vec<SourceToken>)>,
) {
    for shape in shapes {
        for story in &shape.text_stories {
            let values = token_values(story, property);
            let legacy = match legacy.get(story.id.as_str()) {
                Some(legacy) => token_values(legacy, property),
                None => vec![None; values.len()],
            };
            if values == legacy {
                continue;
            }
            let tokens = story
                .paragraphs
                .iter()
                .enumerate()
                .flat_map(|(index, paragraph)| {
                    paragraph
                        .runs
                        .iter()
                        .flat_map(|run| run.text.chars().map(Token::Char))
                        .chain([Token::Paragraph(index)])
                })
                .zip(values.into_iter().zip(legacy))
                .map(|(token, (value, legacy))| SourceToken {
                    token,
                    value,
                    legacy,
                })
                .collect();
            sources.push((story.id.clone(), tokens));
        }
        collect_source_tokens(&shape.children, legacy, property, sources);
    }
}

fn token_values(story: &StorySnapshot, property: SourceProperty) -> Vec<Option<Any>> {
    let mut values = Vec::new();
    for paragraph in &story.paragraphs {
        for run in &paragraph.runs {
            let value = property.value(&run.style);
            values.extend(run.text.chars().map(|_| value.clone()));
        }
        values.push(None);
    }
    values
}

fn merge_property(target: &mut Value, source: &Value, key: &str) {
    match (target, source) {
        (Value::Object(target), Value::Object(source)) => {
            if let Some(baseline) = source.get(key) {
                target.insert(key.into(), baseline.clone());
            }
            for (child_key, target) in target {
                if let Some(source) = source.get(child_key) {
                    merge_property(target, source, key);
                }
            }
        }
        (Value::Array(target), Value::Array(source)) => {
            for (target, source) in target.iter_mut().zip(source) {
                merge_property(target, source, key);
            }
        }
        _ => {}
    }
}

fn has_property(value: &Value, key: &str) -> bool {
    match value {
        Value::Object(object) => {
            object.contains_key(key) || object.values().any(|value| has_property(value, key))
        }
        Value::Array(array) => array.iter().any(|value| has_property(value, key)),
        _ => false,
    }
}
