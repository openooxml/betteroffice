//! Versioned story projections and the target resolver shared by reads, searches and edit
//! batches.
//!
//! A story projects to its paragraphs' text joined by one `\n` with no trailing separator, so
//! offsets are the story's own UTF-16 positions: each paragraph mark is a separator's unit and
//! the final mark is never addressable. Soft line breaks read as `\n` inside a paragraph; the
//! paragraph records tell them apart from separators.

use pptx_parse::{GraphicFrameData, PptxPackage, ShapeNode, TextBody, TextParagraph};
use serde::{Deserialize, Serialize};

use crate::batch::{
    ByteBudget, DocumentVersion, EditFailure, EditFailureCode, EditRefusal, EditTarget,
    MAX_REQUEST_BYTES, SlideTarget, failure, quoted, refusal, request_limit,
};
use crate::search::Needle;
use crate::story::selects_paragraph;
use crate::{DeckSession, DeckSnapshot, EditResult, ShapeSnapshot, SlideSnapshot, StorySnapshot};

/// Deepest shape nesting a target may address.
const MAX_OWNERSHIP_DEPTH: usize = 32;
/// JSON bytes one read or search response may hold, slide snapshots included.
const MAX_RESPONSE_BYTES: usize = 64 * 1024 * 1024;
/// Matches a search returns when the request names no limit.
const FIND_DEFAULT_LIMIT: u32 = 100;
/// Matches one search may return.
const FIND_MAX_LIMIT: u32 = 10_000;

/// A shape addressed through the slide that owns it.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ShapeTarget {
    pub slide_id: String,
    pub shape_id: String,
}

/// A text story addressed through its slide and owning shape.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct StoryTarget {
    pub slide_id: String,
    pub shape_id: String,
    pub story_id: String,
}

/// A half-open range of story-local UTF-16 offsets.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct TextRange {
    pub slide_id: String,
    pub shape_id: String,
    pub story_id: String,
    pub start: u32,
    pub end: u32,
}

impl TextRange {
    fn story(&self) -> StoryTarget {
        StoryTarget {
            slide_id: self.slide_id.clone(),
            shape_id: self.shape_id.clone(),
            story_id: self.story_id.clone(),
        }
    }
}

/// The text a step addresses.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(
    tag = "kind",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum TextTarget {
    Range(TextRange),
    /// The one exact, case-sensitive, paragraph-local occurrence of `text` in the story.
    Search {
        within: StoryTarget,
        text: String,
    },
}

#[derive(Clone, Debug, Default, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct ReadRequest {
    /// Restricts the read to these slides, returned in deck order; every slide when omitted.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub slide_ids: Option<Vec<String>>,
}

#[derive(Clone, Debug, Deserialize, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ReadResponse {
    pub version: DocumentVersion,
    pub slides: Vec<SlideSnapshot>,
    /// Every story of the read slides, in document order.
    pub stories: Vec<StoryText>,
}

/// One story's projected text.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryText {
    pub slide_id: String,
    pub shape_id: String,
    pub story_id: String,
    /// Paragraph texts joined by `\n`.
    pub text: String,
    pub paragraphs: Vec<ParagraphText>,
}

/// One paragraph's place in its story's projected text.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ParagraphText {
    pub paragraph_id: String,
    /// Story offsets of the paragraph's text; a following separator sits at `end`.
    pub start: u32,
    pub end: u32,
    /// Story offsets of soft line breaks, which read as `\n` inside the paragraph.
    pub line_breaks: Vec<u32>,
    /// Field results, such as slide numbers, that text steps leave whole.
    pub fields: Vec<TextField>,
    /// False when saving could not keep a field once the paragraph changes: the field is empty
    /// or left the unchanged leading or trailing text. Steps changing the paragraph refuse it.
    pub editable: bool,
}

/// The story offsets of one field result.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct TextField {
    pub start: u32,
    pub end: u32,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub field_type: Option<String>,
}

/// Where a search looks: a slide, one shape and its descendants, or one story.
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FindScope {
    pub slide_id: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shape_id: Option<String>,
    /// Requires `shape_id`.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub story_id: Option<String>,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct FindRequest {
    pub text: String,
    /// The whole deck when omitted.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub within: Option<FindScope>,
    /// Defaults to 100; at most 10,000.
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub limit: Option<u32>,
}

/// One search hit; `range` is reusable as a [`TextTarget::Range`].
#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FindMatch {
    pub text: String,
    pub range: TextRange,
}

#[derive(Clone, Debug, Deserialize, Eq, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct FindResponse {
    pub version: DocumentVersion,
    pub matches: Vec<FindMatch>,
    /// More matches exist than were returned.
    pub truncated: bool,
}

pub type ReadOutcome = Result<ReadResponse, EditRefusal>;
pub type FindOutcome = Result<FindResponse, EditRefusal>;

fn missing(message: String, target: &EditTarget) -> EditFailure {
    failure(EditFailureCode::MissingTarget, message, Some(target))
}

fn invalid(message: impl Into<String>, target: &EditTarget) -> EditFailure {
    failure(EditFailureCode::InvalidStep, message, Some(target))
}

/// One paragraph of a projected story.
pub(crate) struct ParagraphView {
    pub id: String,
    pub start: u32,
    pub end: u32,
    pub alignment: Option<String>,
    pub line_breaks: Vec<u32>,
    pub fields: Vec<TextField>,
    pub editable: bool,
}

/// A story's projected text in one captured deck state.
pub(crate) struct StoryView<'a> {
    pub slide: &'a SlideSnapshot,
    pub shape: &'a ShapeSnapshot,
    pub story: &'a StorySnapshot,
    units: Vec<u16>,
    pub paragraphs: Vec<ParagraphView>,
}

impl<'a> StoryView<'a> {
    pub(crate) fn build(
        package: &PptxPackage,
        slide: &'a SlideSnapshot,
        shape: &'a ShapeSnapshot,
        story: &'a StorySnapshot,
    ) -> Result<Self, String> {
        let body = source_body(package, slide, shape, &story.id);
        let mut units: Vec<u16> = Vec::new();
        let mut paragraphs = Vec::with_capacity(story.paragraphs.len());
        for (index, paragraph) in story.paragraphs.iter().enumerate() {
            if index > 0 {
                units.push(u16::from(b'\n'));
            }
            let start = units.len() as u32;
            let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
            units.extend(text.encode_utf16());
            let end = units.len() as u32;
            let line_breaks = (start..end)
                .filter(|offset| units[*offset as usize] == u16::from(b'\n'))
                .collect();
            let (fields, editable) = body
                .and_then(|body| source_paragraph(body, &story.id, &paragraph.id))
                .map_or((Vec::new(), true), |source| {
                    located_fields(source, &text, start)
                });
            paragraphs.push(ParagraphView {
                id: paragraph.id.clone(),
                start,
                end,
                alignment: paragraph.alignment.clone(),
                line_breaks,
                fields,
                editable,
            });
        }
        if units.len() as u64 + 1 != u64::from(story.length) {
            return Err(format!(
                "story {} holds content its text projection cannot address",
                quoted(&story.id)
            ));
        }
        Ok(Self {
            slide,
            shape,
            story,
            units,
            paragraphs,
        })
    }

    pub fn len(&self) -> u32 {
        self.units.len() as u32
    }

    pub fn slice(&self, start: u32, end: u32) -> String {
        String::from_utf16_lossy(&self.units[start as usize..end as usize])
    }

    fn is_scalar_boundary(&self, offset: u32) -> bool {
        self.units
            .get(offset as usize)
            .is_none_or(|unit| !(0xDC00..=0xDFFF).contains(unit))
    }

    pub fn range(&self, start: u32, end: u32) -> TextRange {
        TextRange {
            slide_id: self.slide.id.clone(),
            shape_id: self.shape.id.clone(),
            story_id: self.story.id.clone(),
            start,
            end,
        }
    }

    fn target(&self) -> EditTarget {
        EditTarget::Story(StoryTarget {
            slide_id: self.slide.id.clone(),
            shape_id: self.shape.id.clone(),
            story_id: self.story.id.clone(),
        })
    }

    /// Visits exact, case-sensitive, paragraph-local occurrences in order, overlapping ones
    /// included, until `visit` returns false.
    fn each_occurrence(&self, text: &str, mut visit: impl FnMut(u32, u32) -> bool) {
        let needle = Needle::new(text, true);
        let span = text.encode_utf16().count() as u32;
        for paragraph in &self.paragraphs {
            let content = self.slice(paragraph.start, paragraph.end);
            let (mut byte, mut position) = (0, paragraph.start);
            for (from, _) in needle.find_overlapping(&content) {
                position += content[byte..from].encode_utf16().count() as u32;
                byte = from;
                if !visit(position, position + span) {
                    return;
                }
            }
        }
    }

    fn record(&self) -> StoryText {
        StoryText {
            slide_id: self.slide.id.clone(),
            shape_id: self.shape.id.clone(),
            story_id: self.story.id.clone(),
            text: String::from_utf16_lossy(&self.units),
            paragraphs: self
                .paragraphs
                .iter()
                .map(|paragraph| ParagraphText {
                    paragraph_id: paragraph.id.clone(),
                    start: paragraph.start,
                    end: paragraph.end,
                    line_breaks: paragraph.line_breaks.clone(),
                    fields: paragraph.fields.clone(),
                    editable: paragraph.editable,
                })
                .collect(),
        }
    }
}

/// A resolved text target: always inside one story, possibly across paragraphs.
pub(crate) struct Selection<'a> {
    pub view: StoryView<'a>,
    pub start: u32,
    pub end: u32,
}

impl Selection<'_> {
    pub fn text(&self) -> String {
        self.view.slice(self.start, self.end)
    }

    pub fn range(&self) -> TextRange {
        self.view.range(self.start, self.end)
    }

    /// The paragraph holding the whole selection, when one does.
    pub fn paragraph(&self) -> Option<&ParagraphView> {
        self.view
            .paragraphs
            .iter()
            .find(|paragraph| paragraph.start <= self.start && self.end <= paragraph.end)
    }

    /// The paragraphs a paragraph-level step at this selection changes.
    pub fn selected(&self) -> Vec<&ParagraphView> {
        self.view
            .paragraphs
            .iter()
            .filter(|paragraph| {
                selects_paragraph(self.start, self.end, paragraph.start, paragraph.end)
            })
            .collect()
    }
}

/// A shape's position in a slide's tree, one-based from the top level.
fn find_shape<'s>(
    shapes: &'s [ShapeSnapshot],
    id: &str,
    depth: usize,
) -> Option<(&'s ShapeSnapshot, usize)> {
    shapes.iter().find_map(|shape| {
        if shape.id == id {
            Some((shape, depth))
        } else {
            find_shape(&shape.children, id, depth + 1)
        }
    })
}

/// Shapes and their descendants in document order.
fn shapes_in_order(shapes: &[ShapeSnapshot]) -> Vec<&ShapeSnapshot> {
    let mut ordered = Vec::new();
    let mut pending: Vec<&ShapeSnapshot> = shapes.iter().rev().collect();
    while let Some(shape) = pending.pop() {
        ordered.push(shape);
        pending.extend(shape.children.iter().rev());
    }
    ordered
}

/// The parsed text body a seeded story came from, found by the shape's seeded tree path.
fn source_body<'p>(
    package: &'p PptxPackage,
    slide: &SlideSnapshot,
    shape: &ShapeSnapshot,
    story_id: &str,
) -> Option<&'p TextBody> {
    let source = package
        .slides
        .iter()
        .find(|source| Some(&source.part_path) == slide.source_part_path.as_ref())?;
    let path = shape
        .id
        .strip_prefix(slide.id.as_str())?
        .strip_prefix(":shape:")?;
    let mut nodes = source.shapes.as_slice();
    let mut node = None;
    for segment in path.split('.') {
        let found = nodes.get(segment.parse::<usize>().ok()?)?;
        nodes = match found {
            ShapeNode::Group(group) => &group.children,
            _ => &[],
        };
        node = Some(found);
    }
    let node = node.filter(|node| node.id() == shape.source_id)?;
    let suffix = story_id
        .strip_prefix("story:")?
        .strip_prefix(shape.id.as_str())?;
    match (node, suffix) {
        (ShapeNode::Shape(shape), ":0") => shape.text.as_ref(),
        (ShapeNode::GraphicFrame(frame), suffix) => {
            let (row, cell) = suffix.strip_prefix(":table:")?.split_once(':')?;
            match &frame.data {
                GraphicFrameData::Table(table) => Some(
                    &table
                        .rows
                        .get(row.parse::<usize>().ok()?)?
                        .cells
                        .get(cell.parse::<usize>().ok()?)?
                        .text,
                ),
                _ => None,
            }
        }
        _ => None,
    }
}

/// The source paragraph a seeded paragraph id names.
pub(crate) fn source_paragraph<'b>(
    body: &'b TextBody,
    story_id: &str,
    paragraph_id: &str,
) -> Option<&'b TextParagraph> {
    let index = paragraph_id
        .strip_prefix("para:")?
        .strip_prefix(story_id)?
        .strip_prefix(':')?
        .parse::<usize>()
        .ok()?;
    body.paragraphs.get(index)
}

/// The leading and trailing bytes `source` and `target` share, compared by character, the
/// trailing ones within what the leading ones leave. Saving keeps source runs verbatim there.
pub(crate) fn common_ends(source: &str, target: &str) -> (usize, usize) {
    let prefix: usize = source
        .chars()
        .zip(target.chars())
        .take_while(|(left, right)| left == right)
        .map(|(value, _)| value.len_utf8())
        .sum();
    let suffix = source[prefix..]
        .chars()
        .rev()
        .zip(target[prefix..].chars().rev())
        .take_while(|(left, right)| left == right)
        .map(|(value, _)| value.len_utf8())
        .sum();
    (prefix, suffix)
}

/// A source paragraph's field results as byte spans of `text`, its current text, empty results
/// included. `None` when a field left the unchanged leading or trailing text: only there does
/// saving provably keep a field.
fn field_spans(
    source: &TextParagraph,
    text: &str,
) -> Option<Vec<(std::ops::Range<usize>, Option<String>)>> {
    let mut seeded = String::new();
    let mut fields = Vec::new();
    for run in &source.runs {
        let start = seeded.len();
        seeded.push_str(&run.text);
        if run.field_id.is_some() || run.field_type.is_some() {
            fields.push((start..seeded.len(), run.field_type.clone()));
        }
    }
    let (prefix, suffix) = common_ends(&seeded, text);
    let (tail, moved_tail) = (seeded.len() - suffix, text.len() - suffix);
    fields
        .into_iter()
        .map(|(span, field_type)| {
            if span.end <= prefix {
                Some((span, field_type))
            } else if span.start >= tail {
                let start = moved_tail + (span.start - tail);
                Some((start..start + span.len(), field_type))
            } else {
                None
            }
        })
        .collect()
}

/// Field results of a seeded paragraph at their story offsets, and whether saving provably keeps
/// them through a change to the paragraph: all were located and none is empty, since saving a
/// changed paragraph drops empty results.
fn located_fields(source: &TextParagraph, live: &str, start: u32) -> (Vec<TextField>, bool) {
    let Some(spans) = field_spans(source, live) else {
        return (Vec::new(), false);
    };
    let editable = spans.iter().all(|(span, _)| !span.is_empty());
    let offset = |byte: usize| start + live[..byte].encode_utf16().count() as u32;
    let fields = spans
        .into_iter()
        .map(|(span, field_type)| TextField {
            start: offset(span.start),
            end: offset(span.end),
            field_type,
        })
        .collect();
    (fields, editable)
}

/// Whether saving `paragraph_id` of `story_id` as `snapshot` holds it, changed, keeps every
/// field result of its source paragraph: each non-empty, in the unchanged leading or trailing
/// text, inside one run.
pub(crate) fn keeps_fields(
    package: &PptxPackage,
    snapshot: &DeckSnapshot,
    story_id: &str,
    paragraph_id: &str,
) -> bool {
    let located = snapshot.slides.iter().find_map(|slide| {
        shapes_in_order(&slide.shapes)
            .into_iter()
            .find_map(|shape| {
                let story = shape
                    .text_stories
                    .iter()
                    .find(|story| story.id == story_id)?;
                Some((slide, shape, story))
            })
    });
    let Some((slide, shape, story)) = located else {
        return true;
    };
    let source = source_body(package, slide, shape, story_id)
        .and_then(|body| source_paragraph(body, story_id, paragraph_id));
    let paragraph = story
        .paragraphs
        .iter()
        .find(|paragraph| paragraph.id == paragraph_id);
    let (Some(source), Some(paragraph)) = (source, paragraph) else {
        return true;
    };
    let text: String = paragraph.runs.iter().map(|run| run.text.as_str()).collect();
    let Some(spans) = field_spans(source, &text) else {
        return false;
    };
    let mut boundaries = Vec::with_capacity(paragraph.runs.len());
    let mut end = 0;
    for run in &paragraph.runs {
        end += run.text.len();
        boundaries.push(end);
    }
    spans.iter().all(|(span, _)| {
        !span.is_empty()
            && !boundaries
                .iter()
                .any(|boundary| span.start < *boundary && *boundary < span.end)
    })
}

/// The captured deck state targets resolve against.
pub(crate) struct Deck<'a> {
    snapshot: &'a DeckSnapshot,
    package: &'a PptxPackage,
}

impl<'a> Deck<'a> {
    pub fn new(snapshot: &'a DeckSnapshot, package: &'a PptxPackage) -> Self {
        Self { snapshot, package }
    }

    pub fn slide(
        &self,
        slide_id: &str,
        requested: &EditTarget,
    ) -> Result<&'a SlideSnapshot, EditFailure> {
        self.snapshot
            .slides
            .iter()
            .find(|slide| slide.id == slide_id)
            .ok_or_else(|| {
                missing(
                    format!("slide {} was not found", quoted(slide_id)),
                    requested,
                )
            })
    }

    /// Resolves a shape through its slide; the flag marks a shape at the top of the slide.
    pub fn shape(
        &self,
        target: &ShapeTarget,
        requested: &EditTarget,
    ) -> Result<(&'a SlideSnapshot, &'a ShapeSnapshot, bool), EditFailure> {
        let slide = self.slide(&target.slide_id, requested)?;
        match find_shape(&slide.shapes, &target.shape_id, 1) {
            Some((_, depth)) if depth > MAX_OWNERSHIP_DEPTH => Err(failure(
                EditFailureCode::LimitExceeded,
                format!(
                    "shapes nested deeper than {MAX_OWNERSHIP_DEPTH} levels are not addressable"
                ),
                Some(requested),
            )),
            Some((shape, depth)) => Ok((slide, shape, depth == 1)),
            None => Err(
                match self
                    .snapshot
                    .slides
                    .iter()
                    .find(|other| find_shape(&other.shapes, &target.shape_id, 1).is_some())
                {
                    Some(owner) => invalid(
                        format!(
                            "shape {} belongs to slide {}, not {}",
                            quoted(&target.shape_id),
                            quoted(&owner.id),
                            quoted(&target.slide_id)
                        ),
                        requested,
                    ),
                    None => missing(
                        format!("shape {} was not found", quoted(&target.shape_id)),
                        requested,
                    ),
                },
            ),
        }
    }

    pub fn story(
        &self,
        target: &StoryTarget,
        requested: &EditTarget,
    ) -> Result<StoryView<'a>, EditFailure> {
        let shape_target = ShapeTarget {
            slide_id: target.slide_id.clone(),
            shape_id: target.shape_id.clone(),
        };
        let (slide, shape, _) = self.shape(&shape_target, requested)?;
        let Some(story) = shape
            .text_stories
            .iter()
            .find(|story| story.id == target.story_id)
        else {
            let owner = self
                .snapshot
                .slides
                .iter()
                .flat_map(|slide| shapes_in_order(&slide.shapes))
                .find(|other| {
                    other
                        .text_stories
                        .iter()
                        .any(|story| story.id == target.story_id)
                });
            return Err(match owner {
                Some(owner) => invalid(
                    format!(
                        "story {} belongs to shape {}, not {}",
                        quoted(&target.story_id),
                        quoted(&owner.id),
                        quoted(&target.shape_id)
                    ),
                    requested,
                ),
                None => missing(
                    format!("story {} was not found", quoted(&target.story_id)),
                    requested,
                ),
            });
        };
        StoryView::build(self.package, slide, shape, story)
            .map_err(|message| failure(EditFailureCode::Unsupported, message, Some(requested)))
    }

    /// Resolves `target` against this state.
    pub fn text(&self, target: &TextTarget) -> Result<Selection<'a>, EditFailure> {
        let requested = EditTarget::from(target.clone());
        match target {
            TextTarget::Range(range) => {
                let view = self.story(&range.story(), &requested)?;
                let (start, end) = (range.start, range.end);
                if end < start {
                    return Err(invalid(
                        format!("range end {end} precedes its start {start}"),
                        &requested,
                    ));
                }
                if end > view.len() {
                    return Err(invalid(
                        format!(
                            "offset {end} exceeds the story's text length {}",
                            view.len()
                        ),
                        &requested,
                    ));
                }
                if !view.is_scalar_boundary(start) || !view.is_scalar_boundary(end) {
                    return Err(invalid(
                        "range offsets must not split a surrogate pair",
                        &requested,
                    ));
                }
                Ok(Selection { view, start, end })
            }
            TextTarget::Search { within, text } => {
                if text.is_empty() {
                    return Err(invalid("search text must not be empty", &requested));
                }
                let view = self.story(within, &requested)?;
                let mut hits = Vec::with_capacity(2);
                view.each_occurrence(text, |start, end| {
                    hits.push((start, end));
                    hits.len() < 2
                });
                match hits.as_slice() {
                    [] => Err(missing(
                        format!("search text {} was not found", quoted(text)),
                        &requested,
                    )),
                    [(start, end)] => Ok(Selection {
                        start: *start,
                        end: *end,
                        view,
                    }),
                    _ => Err(failure(
                        EditFailureCode::AmbiguousTarget,
                        format!("search text {} occurs more than once", quoted(text)),
                        Some(&requested),
                    )),
                }
            }
        }
    }

    fn stories_under(
        &self,
        slide: &'a SlideSnapshot,
        shapes: &'a [ShapeSnapshot],
    ) -> Result<Vec<StoryView<'a>>, EditFailure> {
        let mut views = Vec::new();
        for shape in shapes_in_order(shapes) {
            for story in &shape.text_stories {
                let view =
                    StoryView::build(self.package, slide, shape, story).map_err(|message| {
                        let target = EditTarget::Story(StoryTarget {
                            slide_id: slide.id.clone(),
                            shape_id: shape.id.clone(),
                            story_id: story.id.clone(),
                        });
                        failure(EditFailureCode::Unsupported, message, Some(&target))
                    })?;
                views.push(view);
            }
        }
        Ok(views)
    }

    /// The stories a search looks through, in document order.
    fn scope(&self, within: Option<&FindScope>) -> Result<Vec<StoryView<'a>>, EditFailure> {
        let Some(within) = within else {
            let mut views = Vec::new();
            for slide in &self.snapshot.slides {
                views.extend(self.stories_under(slide, &slide.shapes)?);
            }
            return Ok(views);
        };
        match (&within.shape_id, &within.story_id) {
            (None, None) => {
                let requested = EditTarget::Slide(SlideTarget {
                    slide_id: within.slide_id.clone(),
                });
                let slide = self.slide(&within.slide_id, &requested)?;
                self.stories_under(slide, &slide.shapes)
            }
            (None, Some(_)) => Err(invalid(
                "a story scope needs its shape",
                &EditTarget::Slide(SlideTarget {
                    slide_id: within.slide_id.clone(),
                }),
            )),
            (Some(shape_id), None) => {
                let target = ShapeTarget {
                    slide_id: within.slide_id.clone(),
                    shape_id: shape_id.clone(),
                };
                let (slide, shape, _) = self.shape(&target, &EditTarget::Shape(target.clone()))?;
                self.stories_under(slide, std::slice::from_ref(shape))
            }
            (Some(shape_id), Some(story_id)) => {
                let target = StoryTarget {
                    slide_id: within.slide_id.clone(),
                    shape_id: shape_id.clone(),
                    story_id: story_id.clone(),
                };
                Ok(vec![
                    self.story(&target, &EditTarget::Story(target.clone()))?,
                ])
            }
        }
    }
}

impl DeckSession {
    /// Slides and the projected text of their stories, with the version they were read at.
    pub fn read_content(&self, request: &ReadRequest) -> EditResult<ReadOutcome> {
        let version = self.version();
        if !ByteBudget::new(MAX_REQUEST_BYTES).charge(request) {
            return Ok(Err(refusal(version, request_limit())));
        }
        let snapshot = self.snapshot()?;
        let deck = Deck::new(&snapshot, &self.package);
        let read = || -> Result<(Vec<SlideSnapshot>, Vec<StoryText>), EditFailure> {
            let slides: Vec<&SlideSnapshot> = match &request.slide_ids {
                None => snapshot.slides.iter().collect(),
                Some(ids) => {
                    for id in ids {
                        deck.slide(
                            id,
                            &EditTarget::Slide(SlideTarget {
                                slide_id: id.clone(),
                            }),
                        )?;
                    }
                    snapshot
                        .slides
                        .iter()
                        .filter(|slide| ids.contains(&slide.id))
                        .collect()
                }
            };
            let mut budget = ByteBudget::new(MAX_RESPONSE_BYTES);
            let over = |target: EditTarget| {
                failure(
                    EditFailureCode::LimitExceeded,
                    format!("a read returns at most {MAX_RESPONSE_BYTES} bytes; read fewer slides"),
                    Some(&target),
                )
            };
            let mut stories = Vec::new();
            for slide in &slides {
                if !budget.charge(slide) {
                    return Err(over(EditTarget::Slide(SlideTarget {
                        slide_id: slide.id.clone(),
                    })));
                }
                for view in deck.stories_under(slide, &slide.shapes)? {
                    let record = view.record();
                    if !budget.charge(&record) {
                        return Err(over(view.target()));
                    }
                    stories.push(record);
                }
            }
            Ok((slides.into_iter().cloned().collect(), stories))
        };
        Ok(match read() {
            Ok((slides, stories)) => Ok(ReadResponse {
                version,
                slides,
                stories,
            }),
            Err(failure) => Err(refusal(version, failure)),
        })
    }

    /// Exact, case-sensitive, paragraph-local matches; overlapping occurrences count separately.
    pub fn find_text(&self, request: &FindRequest) -> EditResult<FindOutcome> {
        let version = self.version();
        if !ByteBudget::new(MAX_REQUEST_BYTES).charge(request) {
            return Ok(Err(refusal(version, request_limit())));
        }
        let limit = request
            .limit
            .unwrap_or(FIND_DEFAULT_LIMIT)
            .min(FIND_MAX_LIMIT) as usize;
        let snapshot = self.snapshot()?;
        let deck = Deck::new(&snapshot, &self.package);
        let find = || -> Result<(Vec<FindMatch>, bool), EditFailure> {
            if request.text.is_empty() {
                return Err(failure(
                    EditFailureCode::InvalidStep,
                    "search text must not be empty",
                    None,
                ));
            }
            let mut budget = ByteBudget::new(MAX_RESPONSE_BYTES);
            let mut matches = Vec::new();
            let mut truncated = false;
            for view in deck.scope(request.within.as_ref())? {
                view.each_occurrence(&request.text, |start, end| {
                    let found = FindMatch {
                        text: request.text.clone(),
                        range: view.range(start, end),
                    };
                    truncated = matches.len() == limit || !budget.charge(&found);
                    if !truncated {
                        matches.push(found);
                    }
                    !truncated
                });
                if truncated {
                    break;
                }
            }
            Ok((matches, truncated))
        };
        Ok(match find() {
            Ok((matches, truncated)) => Ok(FindResponse {
                version,
                matches,
                truncated,
            }),
            Err(failure) => Err(refusal(version, failure)),
        })
    }
}
