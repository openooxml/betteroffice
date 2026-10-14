//! The one walker behind every structured export. It reads the committed deck one slide
//! snapshot at a time under a single read transaction, projects stories through the batch
//! story view, maps seeded shapes to the retained source parts, and admits records one at a time
//! within the budget.

use std::collections::{BTreeSet, HashMap, HashSet};

use ooxml_drawingml::Theme;
use pptx_parse::{
    Bullet, BulletFont, GraphicFrameData, MediaKind, OmittedElement, OmittedInline, PptxPackage,
    Presentation, RunProperties, ShapeNode, SourceShape, TableCell, TableRow, TargetMode, TextBody,
    TextParagraph,
};
use serde::{Deserialize, Serialize};
use sha2::{Digest, Sha256};
use yrs::{Any, Map, MapRef, Out, ReadTxn, Text, Transact};

use super::*;
use crate::comments::{comment_keys, snapshot_comment};
use crate::deck::{
    ShapeParts, SlideParts, graphic_json, map_bool, map_string, required_map, required_order,
    shape_base, shape_parts, slide_notes, slide_parts, slide_ref, string_array_ref,
};
use crate::paragraph::{
    ListCounters, ListMarker, ParagraphCascade, SlideParents, find_placeholder,
};
use crate::story::{snapshot_story, story_ref};
use crate::target::{ParagraphView, StoryView, common_ends, source_paragraph};
use crate::{
    COMMENTS, CommentSnapshot, DeckSession, EditError, META, SHAPES, SLIDES, ShapeKind,
    ShapeSnapshot, SlideSnapshot, StorySnapshot, TextCaps, TextStyle,
};

/// Shape nesting the export descends.
const MAX_DEPTH: usize = 64;
/// Slides, shapes, stories, paragraphs, cells and comments one export visits before it stops.
const MAX_VISITED: usize = 4_000_000;
/// Diagnostics one export returns.
const MAX_DIAGNOSTICS: usize = 1_000;
/// Diagnostics of one code an export returns before it summarizes the rest.
const MAX_DIAGNOSTICS_PER_CODE: usize = 100;
/// Lookups one table spends finding the origins of its merged cells.
const MAX_MERGE_WORK: usize = 1_000_000;
const OLE_URI: &str = "http://schemas.openxmlformats.org/presentationml/2006/ole";

const TRUNCATED_MESSAGE: &str =
    "The export stopped at its limits; content after this point is not included.";
const SUMMARY_MESSAGE: &str = "Further diagnostics with this code were omitted.";

/// Serialized JSON length, counted without allocating the text.
fn json_len<T: Serialize>(value: &T) -> usize {
    struct Counter(usize);
    impl std::io::Write for Counter {
        fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
            self.0 += bytes.len();
            Ok(bytes.len())
        }

        fn flush(&mut self) -> std::io::Result<()> {
            Ok(())
        }
    }
    let mut counter = Counter(0);
    let _ = serde_json::to_writer(&mut counter, value);
    counter.0
}

/// The JSON-escaped length of `text`, a lower bound of any record that carries it.
fn escaped_len(text: &str) -> usize {
    text.chars()
        .map(|character| match character {
            '"' | '\\' | '\n' | '\r' | '\t' | '\u{8}' | '\u{c}' => 2,
            character if (character as u32) < 0x20 => 6,
            character => character.len_utf8(),
        })
        .sum()
}

fn utf16_len(text: &str) -> u32 {
    text.encode_utf16().count() as u32
}

/// The cost of pushing a record of `bytes` onto a list already holding `len` records.
fn push_cost(len: usize, bytes: usize) -> usize {
    bytes + usize::from(len > 0)
}

/// The cost of replacing a `null` field with a value of `bytes`.
fn fill_cost(bytes: usize) -> usize {
    bytes.saturating_sub("null".len())
}

fn warning(
    code: ExportDiagnosticCode,
    anchor: &PptxAnchor,
    message: impl Into<String>,
) -> ExportDiagnostic {
    ExportDiagnostic {
        code,
        severity: ExportSeverity::Warning,
        anchor: Some(anchor.clone()),
        message: message.into(),
    }
}

fn truncation_diagnostic() -> ExportDiagnostic {
    ExportDiagnostic {
        code: ExportDiagnosticCode::Truncated,
        severity: ExportSeverity::Warning,
        anchor: None,
        message: TRUNCATED_MESSAGE.to_owned(),
    }
}

struct Budget {
    max_bytes: usize,
    bytes: usize,
    max_blocks: usize,
    blocks: usize,
    visited: usize,
    stopped: bool,
}

impl Budget {
    fn admit(&mut self, bytes: usize, blocks: usize) -> bool {
        if self.stopped
            || self.bytes + bytes > self.max_bytes
            || self.blocks + blocks > self.max_blocks
        {
            self.stopped = true;
            return false;
        }
        self.bytes += bytes;
        self.blocks += blocks;
        true
    }

    fn visit(&mut self) -> bool {
        if self.stopped {
            return false;
        }
        self.visited += 1;
        if self.visited > MAX_VISITED {
            self.stopped = true;
        }
        !self.stopped
    }
}

/// One slide as the walk sees it before reading its shapes.
struct SlideEntry {
    id: String,
    source_part_path: Option<String>,
    hidden: Option<bool>,
}

/// What a slide's shapes are read against.
struct SlideContext<'a> {
    slide: &'a SlideSnapshot,
    theme: &'a Theme,
    presentation: &'a Presentation,
    parents: SlideParents<'a>,
    part: Option<(String, String)>,
    sld_id: Option<u32>,
}

impl SlideContext<'_> {
    fn provenance(&self, path: &[u32], source_id: Option<u32>) -> Option<SourceProvenance> {
        let (part, part_sha256) = self.part.clone()?;
        Some(SourceProvenance {
            part,
            part_sha256,
            path: path.to_vec(),
            sld_id: self.sld_id,
            source_id,
        })
    }

    fn slide_anchor(&self) -> PptxAnchor {
        PptxAnchor::Slide {
            slide_id: self.slide.id.clone(),
        }
    }
}

/// One level of a shape tree: the current shape ids and what their seeded source holds.
struct Level<'s> {
    nodes: Option<&'s [ShapeNode]>,
    inventory: Option<&'s [SourceShape]>,
    omitted: &'s [OmittedElement],
    /// The id prefix of shapes seeded at this level.
    seeded_prefix: Option<String>,
    export_prefix: String,
    hidden: bool,
}

/// A story and what its paragraphs are read against.
struct StoryInput<'s> {
    shape: &'s ShapeSnapshot,
    story: &'s StorySnapshot,
    id: String,
    cascade: ParagraphCascade<'s>,
    source: Option<&'s TextBody>,
    inlines: Vec<&'s OmittedInline>,
}

struct Walker<'a, T: ReadTxn> {
    txn: &'a T,
    shapes: MapRef,
    comments: MapRef,
    package: &'a PptxPackage,
    options: &'a Resolved,
    budget: Budget,
    diagnostics: Vec<ExportDiagnostic>,
    per_code: HashMap<ExportDiagnosticCode, usize>,
    hashes: HashMap<String, Option<String>>,
    inherited_parts: HashSet<String>,
    noted: HashSet<ExportDiagnosticCode>,
    hidden_shapes: usize,
    /// Stored table rows read into memory.
    rows_read: usize,
}

pub(super) fn export(
    session: &DeckSession,
    options: &Resolved,
    scope: AnchorScope,
) -> EditResult<PptxStructuredContent> {
    walk(session, options, scope).map(|(content, _)| content)
}

/// The export and the number of stored table rows it read into memory.
fn walk(
    session: &DeckSession,
    options: &Resolved,
    scope: AnchorScope,
) -> EditResult<(PptxStructuredContent, usize)> {
    let package = session.package();
    let txn = session.doc.transact();
    let order = required_order(&txn)?;
    let slides = required_map(&txn, SLIDES)?;
    let meta = required_map(&txn, META)?;
    let mut seen = HashSet::new();
    let mut entries = Vec::new();
    let mut has_notes = false;
    for id in string_array_ref(&order, &txn) {
        if !seen.insert(id.clone()) {
            continue;
        }
        let Some(slide) = slides
            .get(&txn, &id)
            .and_then(|value| value.cast::<MapRef>().ok())
        else {
            continue;
        };
        let source_part_path = map_string(&slide, &txn, "sourcePartPath");
        let hidden = match &source_part_path {
            None => Some(false),
            Some(path) => package
                .slides
                .iter()
                .find(|source| &source.part_path == path)
                .and_then(|source| source.hidden),
        };
        let stored_notes = match slide.get(&txn, "notes") {
            Some(Out::Any(Any::String(notes))) => Some(!notes.is_empty()),
            _ => None,
        };
        has_notes |= !options.included.notes
            && (stored_notes.unwrap_or_else(|| {
                source_part_path.as_deref().is_some_and(|path| {
                    package
                        .slides
                        .iter()
                        .any(|source| source.part_path == path && !source.notes.is_empty())
                })
            }) || source_part_path.as_deref().is_some_and(|path| {
                package
                    .relationships
                    .get(path)
                    .is_some_and(|relationships| {
                        relationships.iter().any(|relationship| {
                            relationship.is_type(pptx_parse::relationship_types::NOTES_SLIDE)
                        })
                    })
            }));
        entries.push(SlideEntry {
            id,
            source_part_path,
            hidden,
        });
    }
    let comment_keys = comment_keys(&txn)?;

    let mut content = PptxStructuredContent {
        schema_version: SCHEMA_VERSION,
        anchor_scope: scope,
        reading_order: ReadingOrder::ShapeTree,
        included: options.included,
        slides: Vec::new(),
        diagnostics: Vec::new(),
        truncated: false,
    };
    let reserve = json_len(&truncation_diagnostic()) + 1;
    let mut walker = Walker {
        txn: &txn,
        shapes: required_map(&txn, SHAPES)?,
        comments: required_map(&txn, COMMENTS)?,
        package,
        options,
        budget: Budget {
            max_bytes: options.max_bytes.saturating_sub(reserve),
            bytes: json_len(&content),
            max_blocks: options.max_blocks,
            blocks: 0,
            visited: 0,
            stopped: false,
        },
        diagnostics: Vec::new(),
        per_code: HashMap::new(),
        hashes: HashMap::new(),
        inherited_parts: HashSet::new(),
        noted: HashSet::new(),
        hidden_shapes: 0,
        rows_read: 0,
    };
    walker.preamble(
        &entries,
        has_notes,
        !comment_keys.is_empty(),
        map_bool(&meta, &txn, "commentsPendingSource") == Some(true),
    );
    for (index, entry) in entries.iter().enumerate() {
        if walker.budget.stopped {
            break;
        }
        if entry.hidden == Some(true) && !options.included.hidden_slides {
            continue;
        }
        let comments: Vec<&str> = comment_keys
            .iter()
            .filter(|key| options.included.comments && key.slide_id == entry.id)
            .map(|key| key.id.as_str())
            .collect();
        walker.slide(&mut content.slides, &slides, index, entry, &comments)?;
    }
    content.truncated = walker.budget.stopped;
    content.diagnostics = walker.diagnostics;
    if content.truncated {
        content.diagnostics.push(truncation_diagnostic());
    }
    Ok((content, walker.rows_read))
}

impl<T: ReadTxn> Walker<'_, T> {
    fn diagnose(
        &mut self,
        code: ExportDiagnosticCode,
        severity: ExportSeverity,
        anchor: Option<PptxAnchor>,
        message: impl Into<String>,
    ) {
        if self.budget.stopped || self.diagnostics.len() >= MAX_DIAGNOSTICS {
            return;
        }
        let count = self.per_code.entry(code).or_default();
        *count += 1;
        let diagnostic = match (*count).cmp(&(MAX_DIAGNOSTICS_PER_CODE + 1)) {
            std::cmp::Ordering::Less => ExportDiagnostic {
                code,
                severity,
                anchor,
                message: message.into(),
            },
            std::cmp::Ordering::Equal => ExportDiagnostic {
                code,
                severity: ExportSeverity::Info,
                anchor: None,
                message: SUMMARY_MESSAGE.to_owned(),
            },
            std::cmp::Ordering::Greater => return,
        };
        let cost = push_cost(self.diagnostics.len(), json_len(&diagnostic));
        if self.budget.admit(cost, 0) {
            self.diagnostics.push(diagnostic);
        }
    }

    /// A diagnostic reported once per export.
    fn note_once(&mut self, code: ExportDiagnosticCode, anchor: Option<PptxAnchor>, message: &str) {
        if self.noted.insert(code) {
            self.diagnose(code, ExportSeverity::Info, anchor, message);
        }
    }

    fn preamble(
        &mut self,
        entries: &[SlideEntry],
        has_notes: bool,
        has_comments: bool,
        comments_pending: bool,
    ) {
        use ExportDiagnosticCode as Code;
        let hidden = entries
            .iter()
            .filter(|entry| entry.hidden == Some(true))
            .count();
        let unknown = entries
            .iter()
            .filter(|entry| entry.hidden.is_none())
            .count();
        let included = self.options.included;
        if hidden > 0 && !included.hidden_slides {
            self.diagnose(
                Code::HiddenContentExcluded,
                ExportSeverity::Info,
                None,
                format!(
                    "{hidden} hidden slide(s) were not exported; set includeHiddenSlides to export them."
                ),
            );
        }
        if unknown > 0 {
            self.diagnose(
                Code::VisibilityUnknown,
                ExportSeverity::Warning,
                None,
                format!(
                    "{unknown} slide(s) do not record whether they are hidden, so they are \
                     exported with hidden: null; open the session with its source file to read \
                     their visibility."
                ),
            );
        }
        if has_notes {
            self.diagnose(
                Code::StoriesOmitted,
                ExportSeverity::Info,
                None,
                "Speaker notes are not included; set includeNotes to export them.",
            );
        }
        if has_comments && !included.comments {
            self.diagnose(
                Code::StoriesOmitted,
                ExportSeverity::Info,
                None,
                "Comments are not included; set includeComments to export them.",
            );
        }
        if comments_pending && included.comments {
            self.diagnose(
                Code::ProvenanceUnavailable,
                ExportSeverity::Warning,
                None,
                "This session holds no comments until it is opened with its source file.",
            );
        }
        if !included.formatting {
            self.diagnose(
                Code::FormattingOmitted,
                ExportSeverity::Info,
                None,
                "Formatting marks are not included.",
            );
        }
        if !self.package.has_parts() && entries.iter().any(|entry| entry.source_part_path.is_some())
        {
            self.diagnose(
                Code::ProvenanceUnavailable,
                ExportSeverity::Warning,
                None,
                "This session carries no source file bytes: source provenance, alternative-text \
                 titles and content the deck model does not represent are unavailable.",
            );
        }
    }

    fn part_hash(&mut self, part: &str) -> Option<String> {
        if let Some(hash) = self.hashes.get(part) {
            return hash.clone();
        }
        let hash = self
            .package
            .part_bytes(part)
            .map(|bytes| format!("{:x}", Sha256::digest(bytes)));
        self.hashes.insert(part.to_owned(), hash.clone());
        hash
    }

    /// A story's current text, or `None` when even its text would not fit the budget.
    fn load_story(&mut self, story_id: &str) -> EditResult<Option<StorySnapshot>> {
        let story = story_ref(self.txn, story_id)?;
        if self.budget.bytes + story.len(self.txn) as usize > self.budget.max_bytes {
            self.budget.stopped = true;
            return Ok(None);
        }
        snapshot_story(&story, self.txn, story_id).map(Some)
    }

    fn slide(
        &mut self,
        out: &mut Vec<ExportSlide>,
        slides: &MapRef,
        index: usize,
        entry: &SlideEntry,
        comments: &[&str],
    ) -> EditResult<()> {
        if !self.budget.visit() {
            return Ok(());
        }
        let SlideParts {
            snapshot,
            shape_ids,
            theme,
        } = slide_parts(slides, self.package, self.txn, &entry.id)?;
        let anchor = PptxAnchor::Slide {
            slide_id: snapshot.id.clone(),
        };
        let part_path = snapshot
            .source_part_path
            .as_deref()
            .filter(|path| self.package.part_bytes(path).is_some());
        let unhashed = part_path.map(|path| (path.to_owned(), "0".repeat(64)));
        let sld_id = snapshot.source_part_path.as_deref().and_then(|path| {
            self.package
                .presentation
                .slides
                .iter()
                .find(|reference| reference.part_path == path)
                .map(|reference| reference.id)
        });
        let record = ExportSlide {
            id: format!("s{index}"),
            index: index as u32,
            anchor: anchor.clone(),
            name: snapshot.name.clone(),
            hidden: entry.hidden,
            provenance: unhashed.map(|(part, part_sha256)| SourceProvenance {
                part,
                part_sha256,
                path: Vec::new(),
                sld_id,
                source_id: None,
            }),
            shapes: Vec::new(),
            notes: None,
            comments: Vec::new(),
        };
        if !self
            .budget
            .admit(push_cost(out.len(), json_len(&record)), 1)
        {
            return Ok(());
        }
        out.push(record);
        let record = out.last_mut().expect("just pushed");
        let part =
            part_path.and_then(|path| self.part_hash(path).map(|hash| (path.to_owned(), hash)));
        if let (Some(provenance), Some((_, hash))) = (&mut record.provenance, &part) {
            provenance.part_sha256.clone_from(hash);
        }
        let source = part
            .as_ref()
            .and_then(|(path, _)| self.package.slide_source(path));
        let context = SlideContext {
            slide: &snapshot,
            theme: &theme,
            presentation: &self.package.presentation,
            parents: SlideParents::resolve(
                self.package,
                snapshot.source_part_path.as_deref(),
                snapshot.layout_part_path.as_deref(),
            ),
            part,
            sld_id,
        };
        self.inherited_content(&context);
        let level = Level {
            nodes: context.parents.slide.map(|slide| slide.shapes.as_slice()),
            inventory: source.map(|source| source.shapes.as_slice()),
            omitted: source
                .map(|source| source.omitted.as_slice())
                .unwrap_or_default(),
            seeded_prefix: snapshot
                .source_part_path
                .as_ref()
                .map(|_| format!("{}:shape:", snapshot.id)),
            export_prefix: record.id.clone(),
            hidden: false,
        };
        self.hidden_shapes = 0;
        self.shapes(&mut record.shapes, &shape_ids, &context, &level, 0)?;
        if self.hidden_shapes > 0 {
            self.diagnose(
                ExportDiagnosticCode::HiddenContentExcluded,
                ExportSeverity::Info,
                Some(context.slide_anchor()),
                format!(
                    "{} hidden shape(s) on this slide were not exported; set \
                     includeHiddenShapes to export them.",
                    self.hidden_shapes
                ),
            );
        }
        self.notes(record, &context);
        for comment in comments {
            if self.budget.stopped {
                break;
            }
            let comment = snapshot_comment(&self.comments, self.txn, comment)?;
            self.comment(record, &comment);
        }
        Ok(())
    }

    /// Reports once per part the layout and master shapes drawn on exported slides.
    fn inherited_content(&mut self, context: &SlideContext<'_>) {
        let parents = context.parents;
        let show_layout = parents.slide.is_none_or(|slide| slide.show_master_shapes);
        let show_master = show_layout
            && parents
                .layout
                .is_none_or(|layout| layout.show_master_shapes);
        let drawn = |shapes: &[ShapeNode]| {
            shapes
                .iter()
                .filter(|shape| shape_base(shape).placeholder.is_none())
                .count()
        };
        let mut parts = Vec::new();
        if let Some(master) = parents.master.filter(|_| show_master) {
            parts.push((master.part_path.clone(), drawn(&master.shapes), "master"));
        }
        if let Some(layout) = parents.layout.filter(|_| show_layout) {
            parts.push((layout.part_path.clone(), drawn(&layout.shapes), "layout"));
        }
        for (part, count, role) in parts {
            if count == 0 || !self.inherited_parts.insert(part.clone()) {
                continue;
            }
            let anchor = self
                .part_hash(&part)
                .map(|part_sha256| PptxAnchor::SourcePart {
                    part: part.clone(),
                    part_sha256,
                    path: Vec::new(),
                });
            self.diagnose(
                ExportDiagnosticCode::InheritedContentOmitted,
                ExportSeverity::Info,
                anchor,
                format!(
                    "The {role} {part} draws {count} shape(s) on the slides that use it; layout \
                     and master content is not exported."
                ),
            );
        }
    }

    fn shapes(
        &mut self,
        out: &mut Vec<ExportShape>,
        shape_ids: &[String],
        context: &SlideContext<'_>,
        level: &Level<'_>,
        depth: usize,
    ) -> EditResult<()> {
        let mut omitted = level.omitted.iter().enumerate().peekable();
        for (index, shape_id) in shape_ids.iter().enumerate() {
            if self.budget.stopped {
                return Ok(());
            }
            let seeded = level
                .seeded_prefix
                .as_deref()
                .and_then(|prefix| shape_id.strip_prefix(prefix))
                .and_then(|rest| rest.parse::<usize>().ok());
            if let Some(position) = seeded {
                while let Some((ordinal, element)) =
                    omitted.next_if(|(_, element)| element.position <= position)
                {
                    self.omitted(out, ordinal, element, context, level);
                }
            }
            if !self.budget.visit() {
                return Ok(());
            }
            let ShapeParts {
                snapshot: shape,
                story_ids,
                child_ids,
            } = shape_parts(&self.shapes, self.txn, shape_id, Some(context.theme))?;
            let node = seeded
                .and_then(|position| level.nodes?.get(position))
                .filter(|node| node.id() == shape.source_id);
            let inventory = seeded
                .and_then(|position| level.inventory?.get(position))
                .filter(|source| node.is_some() && source.id == shape.source_id);
            let anchor = PptxAnchor::Shape {
                slide_id: context.slide.id.clone(),
                shape_id: shape.id.clone(),
            };
            if seeded.is_some() && self.package.has_parts() && inventory.is_none() {
                self.diagnose(
                    ExportDiagnosticCode::ProvenanceUnavailable,
                    ExportSeverity::Warning,
                    Some(anchor.clone()),
                    "This shape no longer matches its source element; its source provenance is \
                     unavailable.",
                );
            }
            let hidden = level.hidden || shape.hidden;
            if hidden && !self.options.included.hidden_shapes {
                self.hidden_shapes += 1;
                continue;
            }
            if depth >= MAX_DEPTH {
                self.diagnose(
                    ExportDiagnosticCode::UnsupportedContent,
                    ExportSeverity::Warning,
                    Some(anchor),
                    format!("Shapes nested deeper than {MAX_DEPTH} levels are not exported."),
                );
                continue;
            }
            let record = ExportShape {
                id: format!("{}.h{index}", level.export_prefix),
                anchor: anchor.clone(),
                kind: match shape.kind {
                    ShapeKind::Shape => ExportShapeKind::Shape,
                    ShapeKind::Picture => ExportShapeKind::Picture,
                    ShapeKind::GraphicFrame => ExportShapeKind::GraphicFrame,
                    ShapeKind::Group => ExportShapeKind::Group,
                },
                name: shape.name.clone(),
                title: inventory.and_then(|source| source.title.clone()),
                description: inventory
                    .map(|source| source.description.clone())
                    .unwrap_or_else(|| node.and_then(|node| shape_base(node).description.clone())),
                hidden,
                placeholder: shape
                    .placeholder
                    .as_ref()
                    .map(|placeholder| ExportPlaceholder {
                        placeholder_type: placeholder.placeholder_type.clone(),
                        index: placeholder.index,
                    }),
                provenance: inventory
                    .and_then(|source| context.provenance(&source.path, Some(shape.source_id))),
                stories: Vec::new(),
                table: None,
                object: None,
                children: Vec::new(),
            };
            if !self
                .budget
                .admit(push_cost(out.len(), json_len(&record)), 1)
            {
                return Ok(());
            }
            out.push(record);
            let record = out.last_mut().expect("just pushed");
            match shape.kind {
                ShapeKind::Group => {
                    let children = Level {
                        nodes: match node {
                            Some(ShapeNode::Group(group)) => Some(group.children.as_slice()),
                            _ => None,
                        },
                        inventory: inventory.map(|source| source.children.as_slice()),
                        omitted: inventory
                            .map(|source| source.omitted.as_slice())
                            .unwrap_or_default(),
                        seeded_prefix: seeded.map(|_| format!("{}.", shape.id)),
                        export_prefix: record.id.clone(),
                        hidden,
                    };
                    self.shapes(
                        &mut record.children,
                        &child_ids,
                        context,
                        &children,
                        depth + 1,
                    )?;
                }
                ShapeKind::Shape => {
                    let cascade = shape_cascade(context, &shape, node);
                    let source = match node {
                        Some(ShapeNode::Shape(source)) => source.text.as_ref(),
                        _ => None,
                    };
                    for (story_index, story_id) in story_ids.iter().enumerate() {
                        let Some(story) = self.load_story(story_id)? else {
                            return Ok(());
                        };
                        let input = StoryInput {
                            shape: &shape,
                            story: &story,
                            id: format!("{}.t{story_index}", record.id),
                            cascade,
                            source,
                            inlines: inventory
                                .map(|source| {
                                    source
                                        .omitted_inlines
                                        .iter()
                                        .filter(|inline| inline.cell.is_none())
                                        .collect()
                                })
                                .unwrap_or_default(),
                        };
                        if let Some(view) = self.open_story(&input, context) {
                            let story_record = ExportStory {
                                id: input.id.clone(),
                                anchor: text_anchor(&view, 0, view.len()),
                                paragraphs: Vec::new(),
                            };
                            let cost = push_cost(record.stories.len(), json_len(&story_record));
                            if !self.budget.admit(cost, 0) {
                                return Ok(());
                            }
                            record.stories.push(story_record);
                            let story_record = record.stories.last_mut().expect("just pushed");
                            self.paragraphs(story_record, &view, &input, context);
                        }
                    }
                }
                ShapeKind::Picture => {
                    let (kind, code, severity, message) = match inventory
                        .and_then(|source| source.media)
                    {
                        Some(MediaKind::Video) => (
                            ExportObjectKind::Video,
                            ExportDiagnosticCode::UnsupportedContent,
                            ExportSeverity::Warning,
                            "Video is exported as a placeholder, without its media or poster image.",
                        ),
                        Some(MediaKind::Audio) => (
                            ExportObjectKind::Audio,
                            ExportDiagnosticCode::UnsupportedContent,
                            ExportSeverity::Warning,
                            "Audio is exported as a placeholder, without its media.",
                        ),
                        None => (
                            ExportObjectKind::Picture,
                            ExportDiagnosticCode::ImageDataOmitted,
                            ExportSeverity::Info,
                            "Image data is not exported; the picture keeps its alternative text.",
                        ),
                    };
                    let relationship_ids = match (inventory, node) {
                        (Some(source), _) => source.relationship_ids.clone(),
                        (None, Some(ShapeNode::Picture(picture))) => {
                            picture.relationship_id.iter().cloned().collect()
                        }
                        _ => Vec::new(),
                    };
                    let mut object = self.object_record(
                        record,
                        kind,
                        element_name(inventory, "p:pic"),
                        None,
                        relationship_ids,
                        context,
                    );
                    if let Some(part) = &shape.media_part_path
                        && !object.parts.contains(part)
                    {
                        object.parts.push(part.clone());
                    }
                    self.object(record, object, code, severity, message);
                }
                ShapeKind::GraphicFrame => match graphic_json(self.txn, &shape.id)? {
                    Some(json) if is_table(&json)? => {
                        self.table(record, &shape, &story_ids, &json, context, inventory)?;
                    }
                    json => {
                        let graphic = json
                            .map(|json| serde_json::from_str::<GraphicFrameData>(&json))
                            .transpose()
                            .map_err(|error| EditError::InvalidState(error.to_string()))?;
                        let object =
                            self.graphic_object(record, graphic.as_ref(), inventory, context);
                        let message = match object.kind {
                            ExportObjectKind::Chart => {
                                "Charts are exported as placeholders, without their data."
                            }
                            ExportObjectKind::SmartArt => {
                                "SmartArt is exported as a placeholder, without its content."
                            }
                            ExportObjectKind::EmbeddedObject => {
                                "Embedded objects are exported as placeholders, without their data."
                            }
                            _ => "This graphic frame's content is not represented in the export.",
                        };
                        self.object(
                            record,
                            object,
                            ExportDiagnosticCode::UnsupportedContent,
                            ExportSeverity::Warning,
                            message,
                        );
                    }
                },
            }
        }
        for (ordinal, element) in omitted {
            if self.budget.stopped {
                return Ok(());
            }
            self.omitted(out, ordinal, element, context, level);
        }
        Ok(())
    }

    /// A shape-tree element only the source holds, placed after the current shapes seeded
    /// before it, and hidden or excluded like a shape.
    fn omitted(
        &mut self,
        out: &mut Vec<ExportShape>,
        ordinal: usize,
        element: &OmittedElement,
        context: &SlideContext<'_>,
        level: &Level<'_>,
    ) {
        let Some((part, part_sha256)) = context.part.clone() else {
            return;
        };
        let hidden = level.hidden || element.hidden;
        if hidden && !self.options.included.hidden_shapes {
            self.hidden_shapes += 1;
            return;
        }
        if !self.budget.visit() {
            return;
        }
        let anchor = PptxAnchor::SourcePart {
            part,
            part_sha256,
            path: element.path.clone(),
        };
        let id = format!("{}.x{ordinal}", level.export_prefix);
        let mut record = ExportShape {
            id: id.clone(),
            anchor: anchor.clone(),
            kind: ExportShapeKind::Unknown,
            name: element.name.clone().unwrap_or_default(),
            title: None,
            description: element.description.clone(),
            hidden,
            placeholder: None,
            provenance: context.provenance(&element.path, None),
            stories: Vec::new(),
            table: None,
            object: None,
            children: Vec::new(),
        };
        record.object = Some(self.object_record(
            &record,
            ExportObjectKind::Unknown,
            element.element.clone(),
            None,
            element.relationship_ids.clone(),
            context,
        ));
        if !self
            .budget
            .admit(push_cost(out.len(), json_len(&record)), 1)
        {
            return;
        }
        out.push(record);
        self.diagnose(
            ExportDiagnosticCode::UnsupportedContent,
            ExportSeverity::Warning,
            Some(anchor),
            format!(
                "{} in the shape tree is not represented in the export.",
                element.element
            ),
        );
    }

    fn object(
        &mut self,
        record: &mut ExportShape,
        object: ExportObject,
        code: ExportDiagnosticCode,
        severity: ExportSeverity,
        message: &str,
    ) {
        if !self.budget.admit(fill_cost(json_len(&object)), 0) {
            return;
        }
        record.object = Some(object);
        self.diagnose(code, severity, Some(record.anchor.clone()), message);
    }

    /// An object placeholder whose relationships resolve through the slide part.
    fn object_record(
        &self,
        record: &ExportShape,
        kind: ExportObjectKind,
        element: String,
        uri: Option<String>,
        relationship_ids: Vec<String>,
        context: &SlideContext<'_>,
    ) -> ExportObject {
        let relationships = context
            .slide
            .source_part_path
            .as_deref()
            .and_then(|part| self.package.relationships.get(part));
        let mut parts = Vec::new();
        let mut external_targets = Vec::new();
        for id in &relationship_ids {
            let Some(relationship) = relationships.and_then(|relationships| {
                relationships
                    .iter()
                    .find(|relationship| &relationship.id == id)
            }) else {
                continue;
            };
            match (relationship.target_mode, &relationship.resolved_target) {
                (TargetMode::Internal, Some(part)) => parts.push(part.clone()),
                _ => external_targets.push(relationship.target.clone()),
            }
        }
        ExportObject {
            id: format!("{}.o", record.id),
            kind,
            element,
            uri,
            relationship_ids,
            parts,
            external_targets,
        }
    }

    fn graphic_object(
        &self,
        record: &ExportShape,
        graphic: Option<&GraphicFrameData>,
        inventory: Option<&SourceShape>,
        context: &SlideContext<'_>,
    ) -> ExportObject {
        let (kind, uri, modeled) = match graphic {
            Some(GraphicFrameData::Chart {
                relationship_id, ..
            }) => (ExportObjectKind::Chart, None, vec![relationship_id.clone()]),
            Some(GraphicFrameData::Diagram {
                relationship_ids, ..
            }) => (ExportObjectKind::SmartArt, None, relationship_ids.clone()),
            Some(GraphicFrameData::Unknown { uri, .. }) => (
                if uri.as_deref() == Some(OLE_URI) {
                    ExportObjectKind::EmbeddedObject
                } else {
                    ExportObjectKind::Unknown
                },
                uri.clone(),
                Vec::new(),
            ),
            _ => (ExportObjectKind::Unknown, None, Vec::new()),
        };
        let mut object = self.object_record(
            record,
            kind,
            element_name(inventory, "p:graphicFrame"),
            uri.or_else(|| inventory.and_then(|source| source.graphic_uri.clone())),
            inventory.map_or(modeled, |source| source.relationship_ids.clone()),
            context,
        );
        let preview = match graphic {
            Some(GraphicFrameData::Unknown {
                picture: Some(picture),
                ..
            }) => picture.media_part_path.clone(),
            Some(GraphicFrameData::Chart {
                part_path: Some(part),
                ..
            }) => Some(part.clone()),
            _ => None,
        };
        if let Some(part) = preview
            && !object.parts.contains(&part)
        {
            object.parts.push(part);
        }
        object
    }

    /// Streams the stored table into `record` one row at a time, reading no row the budget has
    /// no room for.
    fn table(
        &mut self,
        record: &mut ExportShape,
        shape: &ShapeSnapshot,
        story_ids: &[String],
        json: &str,
        context: &SlideContext<'_>,
        inventory: Option<&SourceShape>,
    ) -> EditResult<()> {
        let reserved = ExportTable {
            columns: u32::MAX,
            rows: Vec::new(),
        };
        if !self.budget.admit(fill_cost(json_len(&reserved)), 0) {
            return Ok(());
        }
        let shape_anchor = record.anchor.clone();
        let shape_record_id = record.id.clone();
        let table = record.table.insert(ExportTable {
            columns: 0,
            rows: Vec::new(),
        });
        let mut merges = Merges::default();
        let mut read = 0;
        let result = read_table_rows(json, &mut read, &mut |grid, row_index, row| {
            let columns = grid.len().max(row.cells.len()).max(table.columns as usize);
            table.columns = columns as u32;
            merges.start_row(row_index);
            let row_record = ExportTableRow { cells: Vec::new() };
            if !self
                .budget
                .admit(push_cost(table.rows.len(), json_len(&row_record)), 0)
            {
                return Ok(false);
            }
            table.rows.push(row_record);
            let row_record = table.rows.last_mut().expect("just pushed");
            for (column, cell) in row.cells.iter().enumerate() {
                let grid_span = (cell.grid_span as usize).clamp(1, columns - column) as u32;
                let merge_origin = if cell.merged {
                    merges.origin(row_index, column)
                } else {
                    merges.open(row_index, column, cell.row_span.max(1), grid_span);
                    None
                };
                if !self.budget.visit() {
                    return Ok(false);
                }
                let story_id = format!("story:{}:table:{row_index}:{column}", shape.id);
                let story = if story_ids.contains(&story_id) {
                    let Some(story) = self.load_story(&story_id)? else {
                        return Ok(false);
                    };
                    Some(story)
                } else {
                    None
                };
                let id = format!("{shape_record_id}.r{row_index}c{column}");
                let input = story.as_ref().map(|story| StoryInput {
                    shape,
                    story,
                    id: format!("{id}.t0"),
                    cascade: ParagraphCascade {
                        primary: Some(&cell.text),
                        ..ParagraphCascade::default()
                    },
                    source: Some(&cell.text),
                    inlines: inventory
                        .map(|source| {
                            source
                                .omitted_inlines
                                .iter()
                                .filter(|inline| inline.cell == Some((row_index, column)))
                                .collect()
                        })
                        .unwrap_or_default(),
                });
                let view = input
                    .as_ref()
                    .and_then(|input| self.open_story(input, context));
                let anchor = view.as_ref().map_or_else(
                    || shape_anchor.clone(),
                    |view| text_anchor(view, 0, view.len()),
                );
                let cell_record = ExportTableCell {
                    id,
                    anchor: anchor.clone(),
                    row: row_index as u32,
                    column: column as u32,
                    grid_span,
                    row_span: cell.row_span.max(1),
                    merged: cell.merged,
                    merge_origin,
                    story: view
                        .as_ref()
                        .zip(input.as_ref())
                        .map(|(_, input)| ExportStory {
                            id: input.id.clone(),
                            anchor,
                            paragraphs: Vec::new(),
                        }),
                };
                if !self
                    .budget
                    .admit(push_cost(row_record.cells.len(), json_len(&cell_record)), 0)
                {
                    return Ok(false);
                }
                row_record.cells.push(cell_record);
                let cell_record = row_record.cells.last_mut().expect("just pushed");
                if let (Some(view), Some(input), Some(story_record)) =
                    (view, &input, cell_record.story.as_mut())
                {
                    self.paragraphs(story_record, &view, input, context);
                }
            }
            Ok(!self.budget.stopped)
        });
        self.rows_read += read;
        clip_row_spans(table);
        if merges.lost {
            self.diagnose(
                ExportDiagnosticCode::UnsupportedContent,
                ExportSeverity::Warning,
                Some(shape_anchor),
                "The table has too many merged cells to resolve; some merge origins are left out.",
            );
        }
        result
    }

    /// The story's projected text, or a diagnostic when it cannot be projected.
    fn open_story<'s>(
        &mut self,
        input: &StoryInput<'s>,
        context: &SlideContext<'s>,
    ) -> Option<StoryView<'s>> {
        if !self.budget.visit() {
            return None;
        }
        match StoryView::build(self.package, context.slide, input.shape, input.story) {
            Ok(view) => Some(view),
            Err(message) => {
                self.diagnose(
                    ExportDiagnosticCode::UnsupportedContent,
                    ExportSeverity::Warning,
                    Some(PptxAnchor::Shape {
                        slide_id: context.slide.id.clone(),
                        shape_id: input.shape.id.clone(),
                    }),
                    message,
                );
                None
            }
        }
    }

    fn paragraphs(
        &mut self,
        story: &mut ExportStory,
        view: &StoryView<'_>,
        input: &StoryInput<'_>,
        context: &SlideContext<'_>,
    ) {
        let mut counters = ListCounters::default();
        for (index, (paragraph, snapshot)) in view
            .paragraphs
            .iter()
            .zip(&input.story.paragraphs)
            .enumerate()
        {
            if !self.budget.visit() {
                return;
            }
            let text_bytes: usize = snapshot.runs.iter().map(|run| escaped_len(&run.text)).sum();
            if self.budget.bytes + text_bytes > self.budget.max_bytes {
                self.budget.stopped = true;
                return;
            }
            let anchor = text_anchor(view, paragraph.start, paragraph.end);
            let mut pending = Vec::new();
            let authored = match &snapshot.bullet_json {
                Some(json) => match serde_json::from_str::<Bullet>(json) {
                    Ok(bullet) => Some(bullet),
                    Err(_) => {
                        pending.push(warning(
                            ExportDiagnosticCode::UnsupportedNumbering,
                            &anchor,
                            "The paragraph's stored bullet could not be read; it is treated as \
                             inherited.",
                        ));
                        None
                    }
                },
                None => None,
            };
            let properties = input
                .cascade
                .properties(index, snapshot.level, authored.as_ref());
            let has_text = snapshot.runs.iter().any(|run| !run.text.is_empty());
            let list = match counters.next(properties.bullet.as_ref(), snapshot.level, has_text) {
                None => None,
                Some(ListMarker::Character(character)) => Some(ExportList::Bullet {
                    character,
                    font: match &properties.bullet_font {
                        Some(BulletFont::Typeface(font)) => Some(font.clone()),
                        _ => None,
                    },
                }),
                Some(ListMarker::Number { value, marker }) => {
                    let (scheme, start_at) = match &properties.bullet {
                        Some(Bullet::AutoNumber {
                            scheme, start_at, ..
                        }) => (scheme.clone(), *start_at),
                        _ => (String::new(), 1),
                    };
                    if marker.is_none() {
                        pending.push(warning(
                            ExportDiagnosticCode::UnsupportedNumbering,
                            &anchor,
                            format!(
                                "The numbering scheme {scheme:?} cannot be formatted; the list \
                                 marker is left unresolved."
                            ),
                        ));
                    }
                    Some(ExportList::Number {
                        scheme,
                        start_at,
                        value,
                        marker,
                    })
                }
            };
            let runs = self.runs(
                view,
                paragraph,
                snapshot,
                input,
                properties.default_run.as_ref(),
                context,
                &mut pending,
            );
            let record = ExportParagraph {
                id: format!("{}.p{index}", input.id),
                anchor,
                paragraph_id: paragraph.id.clone(),
                level: snapshot.level,
                alignment: paragraph
                    .alignment
                    .clone()
                    .or_else(|| properties.alignment.clone()),
                bullet_json: snapshot.bullet_json.clone(),
                list,
                runs,
            };
            let cost = push_cost(story.paragraphs.len(), json_len(&record));
            if !self.budget.admit(cost, 1) {
                return;
            }
            story.paragraphs.push(record);
            for diagnostic in pending {
                if diagnostic.code == ExportDiagnosticCode::FieldCachedResult
                    && !self.noted.insert(diagnostic.code)
                {
                    continue;
                }
                self.diagnose(
                    diagnostic.code,
                    diagnostic.severity,
                    diagnostic.anchor,
                    diagnostic.message,
                );
            }
        }
    }

    /// The paragraph's runs; diagnostics about them go to `pending`, reported once the paragraph
    /// is admitted.
    #[allow(clippy::too_many_arguments)]
    fn runs(
        &self,
        view: &StoryView<'_>,
        paragraph: &ParagraphView,
        snapshot: &crate::ParagraphSnapshot,
        input: &StoryInput<'_>,
        inherited: Option<&RunProperties>,
        context: &SlideContext<'_>,
        pending: &mut Vec<ExportDiagnostic>,
    ) -> Vec<ExportRun> {
        let anchor = text_anchor(view, paragraph.start, paragraph.end);
        let live: String = snapshot.runs.iter().map(|run| run.text.as_str()).collect();
        let source = input
            .source
            .and_then(|body| source_paragraph(body, &input.story.id, &paragraph.id));
        let located = source.map(|source| Located::new(source, &live));
        let offset = |byte: usize| paragraph.start + utf16_len(&live[..byte]);

        if let (Some(source), Some(_)) = (source, &located)
            && paragraph.fields.is_empty()
            && source
                .runs
                .iter()
                .any(|run| run.field_id.is_some() || run.field_type.is_some())
        {
            pending.push(warning(
                ExportDiagnosticCode::ProvenanceUnavailable,
                &anchor,
                "Fields in this paragraph could not be located after edits; their cached text is \
                 exported as plain text.",
            ));
        }
        if let Some(field) = paragraph.fields.first() {
            pending.push(ExportDiagnostic {
                code: ExportDiagnosticCode::FieldCachedResult,
                severity: ExportSeverity::Info,
                anchor: Some(text_anchor(view, field.start, field.end)),
                message: "Fields are exported with their cached results and never evaluated."
                    .to_owned(),
            });
        }

        let mut links = Vec::new();
        if let (Some(source), Some(located)) = (source, &located) {
            let mut unlocated = false;
            let mut start = 0;
            for run in &source.runs {
                let span = start..start + run.text.len();
                start = span.end;
                let Some(id) = &run.properties.hyperlink_relationship_id else {
                    continue;
                };
                let Some(link) = self.link(context, id, &anchor, pending) else {
                    continue;
                };
                match located.span(span) {
                    Some(span) if span.is_empty() => {}
                    Some(span) => links.push((offset(span.start), offset(span.end), link)),
                    None => unlocated = true,
                }
            }
            if unlocated {
                pending.push(warning(
                    ExportDiagnosticCode::ProvenanceUnavailable,
                    &anchor,
                    "Links in this paragraph could not be located after edits; their text is \
                     exported without them.",
                ));
            }
        }

        let mut points: Vec<(u32, String)> = Vec::new();
        if let Some(source) = source {
            let source_index = source_paragraph_index(&input.story.id, &paragraph.id);
            for inline in input
                .inlines
                .iter()
                .filter(|inline| Some(inline.paragraph) == source_index)
            {
                let byte: usize = source.runs[..inline.position.min(source.runs.len())]
                    .iter()
                    .map(|run| run.text.len())
                    .sum();
                let at = located
                    .as_ref()
                    .and_then(|located| located.point(byte))
                    .map_or(paragraph.end, offset);
                points.push((at, inline.element.clone()));
                pending.push(warning(
                    ExportDiagnosticCode::UnsupportedContent,
                    &text_anchor(view, at, at),
                    format!(
                        "{} in this paragraph is not represented in the export.",
                        inline.element
                    ),
                ));
            }
        }

        let formatting = self.options.included.formatting;
        let styles: Vec<(u32, u32, &TextStyle)> = {
            let mut at = paragraph.start;
            snapshot
                .runs
                .iter()
                .map(|run| {
                    let start = at;
                    at += utf16_len(&run.text);
                    (start, at, &run.style)
                })
                .collect()
        };
        let mut cuts: BTreeSet<u32> = BTreeSet::from([paragraph.start, paragraph.end]);
        cuts.extend(styles.iter().flat_map(|(start, end, _)| [*start, *end]));
        cuts.extend(
            paragraph
                .fields
                .iter()
                .flat_map(|field| [field.start, field.end]),
        );
        cuts.extend(links.iter().flat_map(|(start, end, _)| [*start, *end]));
        cuts.extend(paragraph.line_breaks.iter().flat_map(|at| [*at, at + 1]));
        cuts.extend(points.iter().map(|(at, _)| *at));
        let cuts: Vec<u32> = cuts
            .into_iter()
            .filter(|cut| (paragraph.start..=paragraph.end).contains(cut))
            .collect();

        let range = |start: u32, end: u32| text_anchor(view, start, end);
        let marks_at = |at: u32| {
            formatting.then(|| {
                styles
                    .iter()
                    .find(|(start, end, _)| *start <= at && at < *end)
                    .map(|(_, _, style)| marks(style, inherited))
                    .unwrap_or_default()
            })
        };
        let link_at = |at: u32| {
            links
                .iter()
                .find(|(start, end, _)| *start <= at && at < *end)
                .map(|(_, _, link)| link.clone())
        };
        let mut runs: Vec<ExportRun> = Vec::new();
        let mut open_field: Option<usize> = None;
        let emit_points = |runs: &mut Vec<ExportRun>, at: u32| {
            for (point, element) in points.iter().filter(|(point, _)| *point == at) {
                runs.push(ExportRun {
                    anchor: range(*point, *point),
                    marks: formatting.then(Vec::new),
                    link: None,
                    content: ExportRunKind::Unsupported {
                        element: element.clone(),
                    },
                });
            }
            for field in paragraph
                .fields
                .iter()
                .filter(|field| field.start == at && field.end == at)
            {
                runs.push(ExportRun {
                    anchor: range(at, at),
                    marks: marks_at(at),
                    link: link_at(at),
                    content: ExportRunKind::Field {
                        field_type: field.field_type.clone(),
                        text: String::new(),
                    },
                });
            }
        };
        for window in cuts.windows(2) {
            let (start, end) = (window[0], window[1]);
            emit_points(&mut runs, start);
            if start == end {
                continue;
            }
            let text = view.slice(start, end);
            if let Some(field_index) = paragraph
                .fields
                .iter()
                .position(|field| field.start <= start && end <= field.end)
            {
                if open_field == Some(field_index)
                    && let Some(ExportRun {
                        anchor: PptxAnchor::Range(range),
                        content: ExportRunKind::Field { text: cached, .. },
                        ..
                    }) = runs.last_mut()
                {
                    cached.push_str(&text);
                    range.end = end;
                    continue;
                }
                open_field = Some(field_index);
                runs.push(ExportRun {
                    anchor: range(start, end),
                    marks: marks_at(start),
                    link: link_at(start),
                    content: ExportRunKind::Field {
                        field_type: paragraph.fields[field_index].field_type.clone(),
                        text,
                    },
                });
                continue;
            }
            open_field = None;
            if paragraph.line_breaks.contains(&start) && end == start + 1 {
                runs.push(ExportRun {
                    anchor: range(start, end),
                    marks: formatting.then(Vec::new),
                    link: None,
                    content: ExportRunKind::LineBreak,
                });
                continue;
            }
            let marks = marks_at(start);
            let link = link_at(start);
            if let Some(ExportRun {
                anchor: PptxAnchor::Range(range),
                marks: last_marks,
                link: last_link,
                content: ExportRunKind::Text { text: last_text },
            }) = runs.last_mut()
                && range.end == start
                && *last_marks == marks
                && *last_link == link
            {
                last_text.push_str(&text);
                range.end = end;
                continue;
            }
            runs.push(ExportRun {
                anchor: range(start, end),
                marks,
                link,
                content: ExportRunKind::Text { text },
            });
        }
        emit_points(&mut runs, paragraph.end);
        runs
    }

    fn link(
        &self,
        context: &SlideContext<'_>,
        id: &str,
        anchor: &PptxAnchor,
        pending: &mut Vec<ExportDiagnostic>,
    ) -> Option<ExportLink> {
        let relationship = context
            .slide
            .source_part_path
            .as_deref()
            .and_then(|part| self.package.relationships.get(part))
            .and_then(|relationships| {
                relationships
                    .iter()
                    .find(|relationship| relationship.id == id)
            });
        match relationship {
            Some(relationship) => Some(match relationship.target_mode {
                TargetMode::External => ExportLink {
                    href: relationship.target.clone(),
                    external: true,
                },
                TargetMode::Internal => ExportLink {
                    href: relationship
                        .resolved_target
                        .clone()
                        .unwrap_or_else(|| relationship.target.clone()),
                    external: false,
                },
            }),
            None => {
                pending.push(warning(
                    ExportDiagnosticCode::ProvenanceUnavailable,
                    anchor,
                    "A hyperlink's relationship could not be resolved; its text is exported \
                     without it.",
                ));
                None
            }
        }
    }

    /// The slide's notes when requested, and diagnostics for notes-page text they leave out,
    /// even when the notes themselves are empty.
    fn notes(&mut self, record: &mut ExportSlide, context: &SlideContext<'_>) {
        if !self.options.included.notes || self.budget.stopped {
            return;
        }
        let Ok(slide) = slide_ref(self.txn, &context.slide.id) else {
            return;
        };
        let length = match slide.get(self.txn, "notes") {
            Some(Out::Any(Any::String(notes))) => notes.len(),
            _ => context.parents.slide.map_or(0, |source| source.notes.len()),
        };
        if self.budget.bytes + length > self.budget.max_bytes {
            self.budget.stopped = true;
            return;
        }
        let text = &slide_notes(&slide, self.txn, self.package);
        let anchor = PptxAnchor::Notes {
            slide_id: context.slide.id.clone(),
            range: TextSpan {
                start: 0,
                end: utf16_len(text),
            },
        };
        let notes_source = context
            .slide
            .source_part_path
            .as_deref()
            .and_then(|part| self.package.notes_source(part));
        if !text.is_empty() {
            if !self.budget.visit() {
                return;
            }
            let provenance = notes_source.as_ref().and_then(|source| {
                let part_sha256 = self.part_hash(&source.part_path)?;
                Some(SourceProvenance {
                    part: source.part_path.clone(),
                    part_sha256,
                    path: Vec::new(),
                    sld_id: context.sld_id,
                    source_id: None,
                })
            });
            let notes = ExportNotes {
                id: format!("{}.notes", record.id),
                anchor: anchor.clone(),
                text: text.clone(),
                provenance,
            };
            if !self.budget.admit(fill_cost(json_len(&notes)), 1) {
                return;
            }
            record.notes = Some(notes);
            self.note_once(
                ExportDiagnosticCode::NotesStructureOmitted,
                None,
                "Speaker notes are exported as plain text: their formatting, lists and notes-page \
                 layout are not.",
            );
        }
        if let Some(source) = notes_source.filter(|source| source.other_text_shapes > 0) {
            self.diagnose(
                ExportDiagnosticCode::UnsupportedContent,
                ExportSeverity::Warning,
                Some(anchor),
                format!(
                    "The notes page holds {} more text shape(s) that are not exported.",
                    source.other_text_shapes
                ),
            );
        }
    }

    fn comment(&mut self, record: &mut ExportSlide, comment: &CommentSnapshot) {
        if !self.budget.visit() {
            return;
        }
        let exported = ExportComment {
            id: format!("{}.c{}", record.id, record.comments.len()),
            anchor: PptxAnchor::Comment {
                slide_id: comment.slide_id.clone(),
                comment_id: comment.id.clone(),
                range: TextSpan {
                    start: 0,
                    end: utf16_len(&comment.text),
                },
            },
            comment_id: comment.id.clone(),
            author: Some(comment.author.clone()).filter(|author| !author.is_empty()),
            date: comment.created.clone(),
            parent_id: comment.parent_id.clone(),
            resolved: comment.resolved,
            text: comment.text.clone(),
        };
        if self
            .budget
            .admit(push_cost(record.comments.len(), json_len(&exported)), 1)
        {
            record.comments.push(exported);
        }
    }
}

fn shape_cascade<'s>(
    context: &SlideContext<'s>,
    shape: &'s ShapeSnapshot,
    node: Option<&'s ShapeNode>,
) -> ParagraphCascade<'s> {
    let placeholder = shape.placeholder.as_ref();
    let inherited = |shapes: &'s [ShapeNode]| {
        placeholder
            .and_then(|placeholder| find_placeholder(shapes, placeholder))
            .and_then(node_text)
    };
    ParagraphCascade {
        primary: node.and_then(node_text),
        layout: context
            .parents
            .layout
            .and_then(|layout| inherited(&layout.shapes)),
        master: context
            .parents
            .master
            .and_then(|master| inherited(&master.shapes)),
        master_slide: context.parents.master,
        default_style: &context.presentation.default_text_style,
        default_paragraph: context.presentation.default_text_paragraph.as_deref(),
        placeholder,
        style_color: None,
    }
}

/// Where earlier cells of a table span to, for finding the origin of a merged cell.
#[derive(Default)]
struct Merges {
    /// `(row, column, row span, grid span)` of cells that span, in row-major order.
    open: Vec<(usize, usize, u32, u32)>,
    work: usize,
    lost: bool,
}

impl Merges {
    fn start_row(&mut self, row: usize) {
        self.open
            .retain(|&(origin, _, span, _)| origin + span as usize > row);
    }

    fn open(&mut self, row: usize, column: usize, row_span: u32, grid_span: u32) {
        if row_span > 1 || grid_span > 1 {
            self.open.push((row, column, row_span, grid_span));
        }
    }

    /// The first spanning cell covering `(row, column)`, as PowerPoint draws it.
    fn origin(&mut self, row: usize, column: usize) -> Option<CellPosition> {
        self.work += self.open.len();
        if self.work > MAX_MERGE_WORK {
            self.lost = true;
            return None;
        }
        self.open
            .iter()
            .find(|&&(origin_row, origin_column, row_span, grid_span)| {
                origin_row + row_span as usize > row
                    && origin_column <= column
                    && origin_column + grid_span as usize > column
            })
            .map(|&(origin_row, origin_column, _, _)| CellPosition {
                row: origin_row as u32,
                column: origin_column as u32,
            })
    }
}

/// Whether a stored graphic payload is a table, read without materializing it.
fn is_table(json: &str) -> EditResult<bool> {
    #[derive(Deserialize)]
    struct Tag {
        #[serde(rename = "type")]
        kind: String,
    }
    serde_json::from_str::<Tag>(json)
        .map(|tag| tag.kind == "table")
        .map_err(|error| EditError::InvalidState(error.to_string()))
}

/// A stored table row, in the current or the released encoding.
#[derive(Deserialize)]
#[serde(untagged)]
enum StoredRow {
    Modern(TableRow),
    Legacy(Vec<TextBody>),
}

type RowHandler<'h> = dyn FnMut(&[i64], usize, TableRow) -> EditResult<bool> + 'h;

/// Hands a stored table's rows to `on_row` one at a time, with the grid, until it returns
/// false; the rows after that are skipped without being read into memory. `read` counts the
/// rows that were.
fn read_table_rows(json: &str, read: &mut usize, on_row: &mut RowHandler<'_>) -> EditResult<()> {
    struct Table<'a, 'h> {
        on_row: &'a mut RowHandler<'h>,
        read: &'a mut usize,
        failure: Option<EditError>,
    }

    struct Rows<'b, 'a, 'h> {
        table: &'b mut Table<'a, 'h>,
        grid: &'b [i64],
    }

    impl<'de> serde::de::Visitor<'de> for &mut Table<'_, '_> {
        type Value = ();

        fn expecting(&self, formatter: &mut std::fmt::Formatter) -> std::fmt::Result {
            formatter.write_str("a stored table")
        }

        fn visit_map<A: serde::de::MapAccess<'de>>(self, mut map: A) -> Result<(), A::Error> {
            let mut grid = Vec::new();
            while let Some(key) = map.next_key::<String>()? {
                match key.as_str() {
                    "grid" => grid = map.next_value()?,
                    "rows" => map.next_value_seed(Rows {
                        table: &mut *self,
                        grid: &grid,
                    })?,
                    _ => {
                        map.next_value::<serde::de::IgnoredAny>()?;
                    }
                }
            }
            Ok(())
        }
    }

    impl<'de> serde::de::DeserializeSeed<'de> for Rows<'_, '_, '_> {
        type Value = ();

        fn deserialize<D: serde::Deserializer<'de>>(self, deserializer: D) -> Result<(), D::Error> {
            deserializer.deserialize_seq(self)
        }
    }

    impl<'de> serde::de::Visitor<'de> for Rows<'_, '_, '_> {
        type Value = ();

        fn expecting(&self, formatter: &mut std::fmt::Formatter) -> std::fmt::Result {
            formatter.write_str("table rows")
        }

        fn visit_seq<A: serde::de::SeqAccess<'de>>(self, mut rows: A) -> Result<(), A::Error> {
            let mut index = 0;
            while let Some(row) = rows.next_element::<StoredRow>()? {
                *self.table.read += 1;
                let row = match row {
                    StoredRow::Modern(row) => row,
                    StoredRow::Legacy(cells) => TableRow {
                        height: 0,
                        cells: cells.into_iter().map(TableCell::from_text).collect(),
                    },
                };
                match (self.table.on_row)(self.grid, index, row) {
                    Ok(true) => index += 1,
                    Ok(false) => {
                        while rows.next_element::<serde::de::IgnoredAny>()?.is_some() {}
                        return Ok(());
                    }
                    Err(error) => {
                        self.table.failure = Some(error);
                        return Err(serde::de::Error::custom("the row handler failed"));
                    }
                }
            }
            Ok(())
        }
    }

    let mut table = Table {
        on_row,
        read,
        failure: None,
    };
    let mut deserializer = serde_json::Deserializer::from_str(json);
    let read = serde::Deserializer::deserialize_map(&mut deserializer, &mut table);
    match (table.failure, read) {
        (Some(failure), _) => Err(failure),
        (None, Err(error)) => Err(EditError::InvalidState(error.to_string())),
        (None, Ok(())) => Ok(()),
    }
}

/// Clips row spans past the last row a truncated table kept.
fn clip_row_spans(table: &mut ExportTable) {
    let rows = table.rows.len() as u32;
    for cell in table.rows.iter_mut().flat_map(|row| &mut row.cells) {
        cell.row_span = cell.row_span.min(rows - cell.row);
    }
}

fn text_anchor(view: &StoryView<'_>, start: u32, end: u32) -> PptxAnchor {
    PptxAnchor::Range(view.range(start, end))
}

fn node_text(node: &ShapeNode) -> Option<&TextBody> {
    match node {
        ShapeNode::Shape(shape) => shape.text.as_ref(),
        _ => None,
    }
}

fn element_name(inventory: Option<&SourceShape>, default: &str) -> String {
    inventory.map_or_else(|| default.to_owned(), |source| source.element.clone())
}

/// The source paragraph index a seeded paragraph id names.
fn source_paragraph_index(story_id: &str, paragraph_id: &str) -> Option<usize> {
    paragraph_id
        .strip_prefix("para:")?
        .strip_prefix(story_id)?
        .strip_prefix(':')?
        .parse()
        .ok()
}

/// Maps byte offsets of a seeded paragraph's text into its current text, where they fall in the
/// unchanged leading or trailing text.
struct Located {
    prefix: usize,
    tail: usize,
    moved_tail: usize,
}

impl Located {
    fn new(source: &TextParagraph, live: &str) -> Self {
        let seeded: String = source.runs.iter().map(|run| run.text.as_str()).collect();
        let (prefix, suffix) = common_ends(&seeded, live);
        Self {
            prefix,
            tail: seeded.len() - suffix,
            moved_tail: live.len() - suffix,
        }
    }

    fn point(&self, byte: usize) -> Option<usize> {
        if byte <= self.prefix {
            Some(byte)
        } else if byte >= self.tail {
            Some(self.moved_tail + (byte - self.tail))
        } else {
            None
        }
    }

    fn span(&self, span: std::ops::Range<usize>) -> Option<std::ops::Range<usize>> {
        if span.end <= self.prefix {
            Some(span)
        } else if span.start >= self.tail {
            let start = self.moved_tail + (span.start - self.tail);
            Some(start..start + span.len())
        } else {
            None
        }
    }
}

fn marks(style: &TextStyle, inherited: Option<&RunProperties>) -> Vec<ExportMark> {
    let mut marks = Vec::new();
    if style.bold.or_else(|| inherited?.bold) == Some(true) {
        marks.push(ExportMark::Bold);
    }
    if style.italic.or_else(|| inherited?.italic) == Some(true) {
        marks.push(ExportMark::Italic);
    }
    if let Some(underline) = style
        .underline
        .clone()
        .or_else(|| inherited?.underline.clone())
        .filter(|underline| underline != "none")
    {
        marks.push(ExportMark::Underline { style: underline });
    }
    match style.baseline_pct.or_else(|| inherited?.baseline_pct) {
        Some(baseline) if baseline > 0.0 => marks.push(ExportMark::Superscript),
        Some(baseline) if baseline < 0.0 => marks.push(ExportMark::Subscript),
        _ => {}
    }
    match style.caps.or_else(|| inherited?.caps) {
        Some(TextCaps::Small) => marks.push(ExportMark::SmallCaps),
        Some(TextCaps::All) => marks.push(ExportMark::AllCaps),
        _ => {}
    }
    marks
}

#[cfg(test)]
pub(super) mod tests {
    use super::*;

    const DECK: &[u8] = include_bytes!("../../../../apps/demo/public/betteroffice-demo.pptx");

    /// The demo deck with its first slide's shape tree holding only `shapes`.
    pub(in crate::structured) fn deck_with_shapes(shapes: &str) -> Vec<u8> {
        let tree = format!(
            r#"<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>{shapes}</p:spTree>"#
        );
        let parts: Vec<(String, Vec<u8>)> = ooxml_opc::unzip_parts(DECK)
            .unwrap()
            .into_iter()
            .map(|(path, bytes)| {
                if path != "ppt/slides/slide1.xml" {
                    return (path, bytes);
                }
                let xml = String::from_utf8(bytes).unwrap();
                let start = xml.find("<p:spTree>").unwrap();
                let end = xml.find("</p:spTree>").unwrap() + "</p:spTree>".len();
                (
                    path,
                    format!("{}{tree}{}", &xml[..start], &xml[end..]).into_bytes(),
                )
            })
            .collect();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    /// The demo deck with its first slide holding only a two-column table of `rows` rows.
    pub(in crate::structured) fn deck_with_table(rows: usize) -> Vec<u8> {
        let cell = |text: String| {
            format!(
                r#"<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>{text}</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>"#
            )
        };
        let rows: String = (0..rows)
            .map(|row| {
                format!(
                    r#"<a:tr h="300000">{}{}</a:tr>"#,
                    cell(format!("left {row}")),
                    cell(format!("right {row}"))
                )
            })
            .collect();
        deck_with_shapes(&format!(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Long table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="4000000" cy="600000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>{rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>"#
        ))
    }

    fn rows_read(
        session: &DeckSession,
        options: PptxExportOptions,
    ) -> (PptxStructuredContent, usize) {
        let resolved = Resolved::new(&options).unwrap();
        walk(session, &resolved, AnchorScope::Session).unwrap()
    }

    #[test]
    fn table_rows_are_read_only_as_the_budget_admits_them() {
        let session = DeckSession::open(&deck_with_table(2_000), 7).unwrap();
        let blocks = |max_blocks| PptxExportOptions {
            max_blocks: Some(max_blocks),
            ..PptxExportOptions::default()
        };
        let (content, read) = rows_read(&session, blocks(1));
        assert!(content.truncated && content.slides[0].shapes.is_empty());
        assert_eq!(read, 0);
        let shape_id = session.snapshot().unwrap().slides[0].shapes[0].id.clone();
        let txn = session.doc.transact();
        let shapes = required_map(&txn, SHAPES).unwrap();
        assert!(
            shape_parts(&shapes, &txn, &shape_id, None)
                .unwrap()
                .snapshot
                .graphic
                .is_none()
        );
        drop(txn);
        let (content, read) = rows_read(&session, blocks(2));
        assert_eq!(content.slides[0].shapes[0].name, "Long table");
        assert_eq!(read, 1);
        let (_, read) = rows_read(&session, blocks(12));
        assert_eq!(read, 6);
        let (content, read) = rows_read(
            &session,
            PptxExportOptions {
                max_bytes: Some(16_384),
                ..PptxExportOptions::default()
            },
        );
        let table = content.slides[0].shapes[0].table.as_ref().unwrap();
        assert!(
            read <= table.rows.len() + 1 && read < 40,
            "{read} rows read"
        );
        let (content, read) = rows_read(&session, PptxExportOptions::default());
        assert!(read >= 2_000);
        assert_eq!(
            content.slides[0].shapes[0]
                .table
                .as_ref()
                .unwrap()
                .rows
                .len(),
            2_000
        );
    }
}
