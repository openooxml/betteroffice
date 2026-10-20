//! View-only layout reservations a host passes in the render environment.

use std::collections::{BTreeMap, BTreeSet, HashMap};

use docx_layout::types::{BlockId, LayoutBlock, ReservationRef, ReservedBlock, ReservedRun};
use serde::{Deserialize, Deserializer, Serialize, Serializer};
use serde_json::Value;

use super::{DrawingMarker, RawRun, RawRunKind, RunFormatting, utf16_len, utf16_slice};

/// What a [`Reservation`] does at its anchor.
#[derive(Clone, Copy, Debug, Eq, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum ReservationKind {
    Inline,
    Block,
    Hide,
}

/// The side of its paragraph a block reservation takes.
#[derive(Clone, Copy, Debug, Default, Eq, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "lowercase")]
pub enum ReservationSide {
    #[default]
    Before,
    After,
}

/// One reservation resolved to a paragraph. `start..end` count the paragraph's
/// UTF-16 units, one per embed, with its mark at the text's length.
#[derive(Clone, Debug, PartialEq, Deserialize, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct Reservation {
    pub owner: String,
    pub id: String,
    pub story: String,
    pub para_id: String,
    pub start: u32,
    pub end: u32,
    pub kind: ReservationKind,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub width: Option<f64>,
    #[serde(default)]
    pub height: f64,
    #[serde(default)]
    pub side: ReservationSide,
    /// An inline box that also hides `start..end`.
    #[serde(default, skip_serializing_if = "std::ops::Not::not")]
    pub hide: bool,
}

impl Reservation {
    fn valid(&self) -> bool {
        let size = |value: f64| value.is_finite() && value >= 0.0;
        self.start <= self.end
            && size(self.height)
            && self.width.is_none_or(size)
            && (self.kind != ReservationKind::Inline || self.width.is_some())
    }

    fn reference(&self) -> ReservationRef {
        ReservationRef {
            owner: self.owner.clone(),
            id: self.id.clone(),
        }
    }

    fn hides(&self) -> bool {
        self.kind == ReservationKind::Hide || (self.kind == ReservationKind::Inline && self.hide)
    }
}

/// Reservations by host key, indexed by the paragraph each names. Invalid
/// entries are dropped when read.
#[derive(Clone, Debug, Default)]
pub struct Reservations {
    entries: BTreeMap<String, Reservation>,
    paragraphs: HashMap<String, HashMap<String, Vec<Reservation>>>,
}

impl Reservations {
    pub fn is_empty(&self) -> bool {
        self.entries.is_empty()
    }

    pub fn insert(&mut self, key: impl Into<String>, reservation: Reservation) {
        let mut entries = std::mem::take(&mut self.entries);
        entries.insert(key.into(), reservation);
        *self = Self::from_entries(entries);
    }

    pub(super) fn paragraph(&self, story: &str, para_id: &str) -> &[Reservation] {
        self.paragraphs
            .get(story)
            .and_then(|paragraphs| paragraphs.get(para_id))
            .map_or(&[], Vec::as_slice)
    }

    fn from_entries(entries: BTreeMap<String, Reservation>) -> Self {
        let mut paragraphs: HashMap<String, HashMap<String, Vec<Reservation>>> = HashMap::new();
        for reservation in entries.values().filter(|reservation| reservation.valid()) {
            paragraphs
                .entry(reservation.story.clone())
                .or_default()
                .entry(reservation.para_id.clone())
                .or_default()
                .push(reservation.clone());
        }
        Self {
            entries,
            paragraphs,
        }
    }
}

impl PartialEq for Reservations {
    fn eq(&self, other: &Self) -> bool {
        self.entries == other.entries
    }
}

impl Serialize for Reservations {
    fn serialize<S: Serializer>(&self, serializer: S) -> Result<S::Ok, S::Error> {
        self.entries.serialize(serializer)
    }
}

impl<'de> Deserialize<'de> for Reservations {
    fn deserialize<D: Deserializer<'de>>(deserializer: D) -> Result<Self, D::Error> {
        let Value::Object(entries) = Value::deserialize(deserializer)? else {
            return Ok(Self::default());
        };
        Ok(Self::from_entries(
            entries
                .into_iter()
                .filter_map(|(key, value)| {
                    serde_json::from_value::<Reservation>(value)
                        .ok()
                        .filter(Reservation::valid)
                        .map(|reservation| (key, reservation))
                })
                .collect(),
        ))
    }
}

/// What one paragraph's reservations add around it.
#[derive(Default)]
pub(super) struct Placed {
    pub(super) before: Vec<LayoutBlock>,
    pub(super) after: Vec<LayoutBlock>,
    pub(super) hides_mark: bool,
}

/// Hides the reserved ranges of one paragraph's runs and drawings and inserts
/// its inline boxes. `start` and `mark` are the story indices of the
/// paragraph's first unit and of its mark; an anchor past the mark is stale
/// and skipped.
pub(super) fn place(
    reservations: &[Reservation],
    runs: &mut Vec<RawRun>,
    drawings: &mut [DrawingMarker],
    start: u32,
    mark: u32,
) -> Placed {
    let mut placed = Placed::default();
    let length = mark.saturating_sub(start);
    let current: Vec<&Reservation> = reservations
        .iter()
        .filter(|reservation| reservation.end <= length + 1 && reservation.start <= length)
        .collect();
    if current.is_empty() {
        return placed;
    }
    let hidden: Vec<(u32, u32)> = current
        .iter()
        .filter(|reservation| reservation.hides() && reservation.start < reservation.end)
        .map(|reservation| (start + reservation.start, start + reservation.end))
        .collect();
    let boxes: Vec<&&Reservation> = current
        .iter()
        .filter(|reservation| reservation.kind == ReservationKind::Inline)
        .collect();
    let cuts: BTreeSet<u32> = hidden
        .iter()
        .flat_map(|&(from, to)| [from, to])
        .chain(boxes.iter().map(|reservation| start + reservation.start))
        .collect();
    split_runs(runs, &cuts);
    let is_hidden = |from: u32, to: u32| hidden.iter().any(|&(a, b)| a <= from && to <= b);
    for run in runs.iter_mut() {
        if run.story_start < run.story_end && is_hidden(run.story_start, run.story_end) {
            run.revision_hidden = true;
        }
    }
    for drawing in drawings.iter_mut() {
        if is_hidden(drawing.story_index, drawing.story_index + 1) {
            drawing.revision_hidden = true;
        }
    }
    placed.hides_mark = is_hidden(mark, mark + 1);
    for reservation in boxes {
        let at = start + reservation.start;
        let index = runs
            .iter()
            .position(|run| {
                run.story_start > at
                    || (run.story_start == at && !matches!(run.kind, RawRunKind::Reserved(_)))
            })
            .unwrap_or(runs.len());
        runs.insert(
            index,
            RawRun {
                kind: RawRunKind::Reserved(ReservedRun {
                    reservation: reservation.reference(),
                    width: reservation.width.unwrap_or_default(),
                    height: reservation.height,
                    pm_start: None,
                    pm_end: None,
                }),
                formatting: RunFormatting::default(),
                inherited_hyperlink: (false, false),
                story_start: at,
                story_end: at,
                pm_start: reservation.start,
                pm_end: reservation.start,
                inline_sdt_widget: None,
                atom: true,
                revision_hidden: false,
            },
        );
    }
    for reservation in current
        .iter()
        .filter(|reservation| reservation.kind == ReservationKind::Block)
    {
        let block = LayoutBlock::Reserved(ReservedBlock {
            id: BlockId::Str(format!(
                "reserved:{}",
                serde_json::json!([reservation.owner, reservation.id])
            )),
            reservation: reservation.reference(),
            width: reservation.width,
            height: reservation.height,
        });
        match reservation.side {
            ReservationSide::Before => placed.before.push(block),
            ReservationSide::After => placed.after.push(block),
        }
    }
    placed
}

/// Cuts story text runs at `cuts`, leaving embeds and cuts inside a
/// surrogate pair whole.
fn split_runs(runs: &mut Vec<RawRun>, cuts: &BTreeSet<u32>) {
    if cuts.is_empty() {
        return;
    }
    let mut split = Vec::with_capacity(runs.len() + cuts.len());
    for run in runs.drain(..) {
        let RawRunKind::Text(text) = &run.kind else {
            split.push(run);
            continue;
        };
        if run.atom || run.story_end - run.story_start != utf16_len(text) {
            split.push(run);
            continue;
        }
        let mut boundaries = BTreeSet::new();
        let mut offset = 0;
        for ch in text.chars() {
            offset += ch.len_utf16() as u32;
            boundaries.insert(run.story_start + offset);
        }
        let inner: Vec<u32> = cuts
            .range(run.story_start + 1..run.story_end)
            .filter(|cut| boundaries.contains(cut))
            .copied()
            .collect();
        if inner.is_empty() {
            split.push(run);
            continue;
        }
        let mut from = run.story_start;
        for to in inner.into_iter().chain(std::iter::once(run.story_end)) {
            let (left, right) = (from - run.story_start, to - run.story_start);
            let mut piece = run.clone();
            piece.kind = RawRunKind::Text(utf16_slice(text, left, right));
            piece.story_start = from;
            piece.story_end = to;
            piece.pm_start = run.pm_start + left;
            piece.pm_end = run.pm_start + right;
            split.push(piece);
            from = to;
        }
    }
    *runs = split;
}
