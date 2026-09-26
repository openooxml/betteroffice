//! Package context the export keeps from the parsed source: which story is which header, footer
//! or note, comment metadata and bodies, section, numbering and relationship data, the source
//! structure seeding does not keep in the stream, and an inventory of the content it leaves out.

use std::collections::{HashMap, HashSet};
use std::sync::{Arc, Mutex};

use serde_json::Value;
use yrs::types::{DeepObservable, Event, PathSegment};
use yrs::{Assoc, IndexedSequence, ReadTxn, StickyIndex, Text, Transact};

use super::{Anchor, BreakType, Revision, StoryKind};
use crate::{COMMENTS, EditingDoc, story_ref};

/// A header, footer or note story and the part it was read from.
pub(crate) struct SourceStory {
    pub story: String,
    pub kind: StoryKind,
    pub part: Option<String>,
    pub note_id: Option<String>,
}

pub(crate) struct SourceComment {
    pub id: String,
    pub author: Option<String>,
    pub date: Option<String>,
    pub parent_id: Option<String>,
    pub done: bool,
    /// The comment body as parsed blocks.
    pub body: Vec<Value>,
}

/// An internal part or external target a relationship names.
#[derive(Clone)]
pub(crate) enum RelationshipTarget {
    Part(String),
    External(String),
}

/// A story position recorded as a unit index at seeding, pinned to the stream of the replica that
/// seeded it so it follows later edits.
pub(crate) struct Pin {
    pub story: String,
    pub unit: u32,
    pub position: Option<StickyIndex>,
}

impl Pin {
    pub(crate) fn new(story: &str, unit: u32) -> Self {
        Self {
            story: story.to_owned(),
            unit,
            position: None,
        }
    }

    fn sticky<T: ReadTxn>(txn: &T, story: &str, unit: u32) -> Option<StickyIndex> {
        story_ref(txn, story)
            .ok()
            .filter(|text| unit < text.len(txn))
            .and_then(|text| text.sticky_index(txn, unit, Assoc::After))
    }

    /// The pinned unit's current story index, each position resolved once per `resolved`.
    pub(crate) fn resolve<T: ReadTxn>(
        &self,
        txn: &T,
        resolved: &mut HashMap<StickyIndex, Option<u32>>,
    ) -> Option<u32> {
        let position = self.position.as_ref()?;
        *resolved
            .entry(position.clone())
            .or_insert_with(|| position.get_offset(txn).map(|offset| offset.index))
    }
}

/// Inline source content the stream does not carry where the source had it.
#[derive(Clone)]
pub(crate) enum InlineSource {
    /// A raw XML node, an unmodelled drawing or object, or a control child the control's frozen
    /// content drops.
    Omitted { element: String },
    /// A page or column break, which seeding moves out of its paragraph or drops.
    Break {
        kind: BreakType,
        revision: Option<Revision>,
    },
}

/// How the stream still carries a source break.
#[derive(Clone, Copy)]
pub(crate) enum Witness {
    /// Nothing in the stream stands for it.
    Invisible,
    /// The paragraph's `pageBreakBeforeRun` flag.
    Leading,
    /// The relocated break embed at this index of [`Provenance::relocated`].
    Embed(usize),
}

/// Inline source content, pinned to the unit it precedes.
pub(crate) struct InlineRecord {
    pub pin: Pin,
    pub para_id: String,
    /// Inside the content control whose embed is the pinned unit.
    pub in_control: bool,
    /// Where in that control's content, as UTF-16 offsets into it and each nested control,
    /// when known.
    pub control_offset: Option<Vec<u32>>,
    pub content: InlineSource,
    pub witness: Witness,
}

/// How a source cell takes part in a vertical merge.
#[derive(Clone, Copy, Eq, PartialEq)]
pub(crate) enum SourceMerge {
    None,
    Restart,
    Continue,
}

/// A source cell on the table grid.
pub(crate) struct CellLayout {
    pub column: u32,
    pub span: u32,
    pub merge: SourceMerge,
    /// The cell story seeding made for it; continuation cells usually have none.
    pub story: Option<String>,
    /// The source cell holds text or drawings.
    pub content: bool,
}

pub(crate) struct RowLayout {
    pub grid_before: u32,
    pub grid_after: u32,
    pub cells: Vec<CellLayout>,
}

/// A table's grid as the source defines it, which the stream's layout-oriented spans simplify.
pub(crate) struct TableLayout {
    pub grid_columns: u32,
    pub rows: Vec<RowLayout>,
}

/// The identity of a tracked insertion (`inserted`) or deletion, for matching a stream stamp to
/// the move it came from.
pub(crate) fn move_key(inserted: bool, id: &str, author: &str, date: &str) -> String {
    format!(
        "{}\u{1}{id}\u{1}{author}\u{1}{date}",
        if inserted { "ins" } else { "del" }
    )
}

/// A break embed seeding moved out of the paragraph `para_id`.
pub(crate) struct Relocated {
    pub pin: Pin,
    pub para_id: String,
}

/// Source structure and content the editing stream does not carry.
#[derive(Default)]
pub(crate) struct Provenance {
    /// Each story's raw XML block elements, in source order.
    pub raw_blocks: HashMap<String, Vec<String>>,
    /// The source block order of the stories with raw XML blocks: the id seeding gives each other
    /// block (paragraph id, `{story}:t{n}`, `{story}:sdt{n}`), `None` for a raw block.
    pub block_order: HashMap<String, Vec<Option<String>>>,
    pub inline: Vec<InlineRecord>,
    /// Where each raw XML block sits in its part, until resolved against the package.
    pub raw_sources: Vec<RawSource>,
    /// Source grids, keyed by table id (`{story}:t{index}`).
    pub tables: HashMap<String, TableLayout>,
    /// The break embeds seeding moved out of paragraphs.
    pub relocated: Vec<Relocated>,
    /// The revisions of moved content, whose stream stamps read as plain insertions and
    /// deletions; see [`move_key`].
    pub moves: HashSet<String>,
    /// Once pinned, the indices into `inline` and into `relocated` of each story's records.
    pub by_story: HashMap<String, (Vec<usize>, Vec<usize>)>,
    /// Once pinned, the relocated breaks an inline record witnesses.
    pub witnessed: HashSet<usize>,
}

impl Provenance {
    /// Pins every recorded position to `doc`, which seeding just filled, each distinct one once,
    /// and indexes the records by story.
    pub(crate) fn pin(&mut self, doc: &EditingDoc) {
        let txn = doc.yrs_doc().transact();
        let mut pinned: HashMap<(String, u32), Option<StickyIndex>> = HashMap::new();
        let pins = self.inline.iter_mut().map(|record| &mut record.pin).chain(
            self.relocated
                .iter_mut()
                .map(|relocated| &mut relocated.pin),
        );
        for pin in pins {
            pin.position = pinned
                .entry((pin.story.clone(), pin.unit))
                .or_insert_with(|| Pin::sticky(&txn, &pin.story, pin.unit))
                .clone();
        }
        self.by_story.clear();
        for (index, record) in self.inline.iter().enumerate() {
            let entry = self.by_story.entry(record.pin.story.clone()).or_default();
            entry.0.push(index);
        }
        for (index, relocated) in self.relocated.iter().enumerate() {
            let entry = self
                .by_story
                .entry(relocated.pin.story.clone())
                .or_default();
            entry.1.push(index);
        }
        self.witnessed = self
            .inline
            .iter()
            .filter_map(|record| match record.witness {
                Witness::Embed(index) => Some(index),
                _ => None,
            })
            .collect();
    }
}

/// One step from a story part's root element to a block container.
#[derive(Clone)]
pub(crate) enum Step {
    /// The main document's `w:body`.
    Body,
    /// The `w:footnote` or `w:endnote` root child with this `w:id`.
    Note(&'static str, String),
    /// The `w:comment` root child with this `w:id`.
    Comment(String),
    /// The index-th block the story dispatcher reads from the current container.
    Block(usize),
    /// The index-th row of the current table.
    Row(usize),
    /// The index-th cell of the current row.
    Cell(usize),
    /// The current content control's `w:sdtContent`.
    Content,
}

/// A raw XML block and the steps to it from its part's root element.
pub(crate) struct RawSource {
    pub story: String,
    /// Its position among the story's raw blocks.
    pub index: usize,
    pub steps: Vec<Step>,
    pub xml: String,
}

/// The package parts provenance resolves against: the main document part, the note and comment
/// parts, the header and footer parts the main document relates, wherever they sit, and their
/// relationship parts. Part names match ignoring ASCII case, as the parser matches them.
pub(crate) struct SourceParts {
    pub document: String,
    pub parts: Vec<(String, Vec<u8>)>,
}

impl SourceParts {
    pub(crate) fn new(parts: Vec<(String, Vec<u8>)>) -> Self {
        let limits = docx_parse::ParseLimits::default();
        let mut budget = docx_parse::ParseBudget::new(&limits);
        let document = docx_parse::relationships::office_document_path(&parts, &mut budget)
            .unwrap_or_else(|_| DOCUMENT_PART.to_owned());
        let mut stories = vec![
            document.clone(),
            FOOTNOTES_PART.to_owned(),
            ENDNOTES_PART.to_owned(),
            COMMENTS_PART.to_owned(),
        ];
        let rels = relationship_part(&document);
        if let Some((path, bytes)) = parts
            .iter()
            .find(|(path, _)| path.eq_ignore_ascii_case(&rels))
            && let Ok(relationships) =
                docx_parse::relationships::parse_relationships(bytes, path, &mut budget)
        {
            use docx_parse::relationships::relationship_types::{FOOTER, HEADER};
            stories.extend(
                relationships
                    .values()
                    .filter(|relationship| {
                        matches!(relationship.relationship_type.as_str(), HEADER | FOOTER)
                    })
                    .filter_map(|relationship| {
                        match docx_parse::resolve_relationship_target(&document, relationship) {
                            Ok(docx_parse::RelationshipTarget::Internal(part)) => Some(part),
                            _ => None,
                        }
                    }),
            );
        }
        let kept: HashSet<String> = stories
            .iter()
            .flat_map(|part| {
                [
                    part.to_ascii_lowercase(),
                    relationship_part(part).to_ascii_lowercase(),
                ]
            })
            .collect();
        let parts = parts
            .into_iter()
            .filter(|(path, _)| kept.contains(&path.to_ascii_lowercase()))
            .collect();
        Self { document, parts }
    }

    fn part(&self, path: &str) -> Option<&[u8]> {
        self.parts
            .iter()
            .find(|(name, _)| name.eq_ignore_ascii_case(path))
            .map(|(_, bytes)| bytes.as_slice())
    }
}

/// The relationship part of `part`.
fn relationship_part(part: &str) -> String {
    match part.rsplit_once('/') {
        Some((directory, name)) => format!("{directory}/_rels/{name}.rels"),
        None => format!("_rels/{part}.rels"),
    }
}

/// The observer key of the comment-store watch; watching again replaces the previous watch.
const COMMENT_WRITES: &str = "structured-export-comment-writes";

/// Comment store keys written since the export's source was retained, by comment id; `None`
/// marks a whole entry written anew.
#[derive(Default)]
pub(crate) struct CommentWrites {
    written: Arc<Mutex<HashSet<(String, Option<String>)>>>,
}

impl CommentWrites {
    /// Starts recording the comment writes committed to `doc` from now on.
    pub(crate) fn watch(doc: &EditingDoc) -> Self {
        let written: Arc<Mutex<HashSet<(String, Option<String>)>>> = Arc::default();
        let recorded = Arc::clone(&written);
        if let Some(comments) = doc.yrs_doc().transact().get_map(COMMENTS) {
            comments.observe_deep_with(COMMENT_WRITES, move |txn, events| {
                let mut written = recorded.lock().unwrap_or_else(|error| error.into_inner());
                for event in events.iter() {
                    let Event::Map(event) = event else {
                        continue;
                    };
                    let path = event.path();
                    let keys = event.keys(txn);
                    match (path.len(), path.front()) {
                        (0, _) => {
                            for id in keys.keys() {
                                written.insert((id.to_string(), None));
                            }
                        }
                        (1, Some(PathSegment::Key(id))) => {
                            for key in keys.keys() {
                                written.insert((id.to_string(), Some(key.to_string())));
                            }
                        }
                        _ => {}
                    }
                }
            });
        }
        Self { written }
    }

    /// Whether `key` of comment `id` was written since the source was retained.
    pub(crate) fn written(&self, id: &str, key: &str) -> bool {
        let written = self
            .written
            .lock()
            .unwrap_or_else(|error| error.into_inner());
        written.contains(&(id.to_owned(), None))
            || written.contains(&(id.to_owned(), Some(key.to_owned())))
    }
}

/// Everything the export reads from the source package rather than the editing stream.
#[derive(Default)]
pub(crate) struct ReadSource {
    pub document_part: String,
    pub stories: Vec<SourceStory>,
    /// The index in `stories` of each story.
    story_index: HashMap<String, usize>,
    pub footnote_separators: usize,
    pub endnote_separators: usize,
    pub comments: Vec<SourceComment>,
    /// Comments seeding anchored into the comment store; the others never had a range.
    pub seeded_comments: HashSet<String>,
    /// Source provenance of each story's raw XML blocks, parallel to `provenance.raw_blocks`, and
    /// of each source comment body's (keyed `comment:{id}`).
    pub raw_block_anchors: HashMap<String, Vec<Option<Anchor>>>,
    /// Each source comment's `w:comment` element.
    pub comment_anchors: HashMap<String, Anchor>,
    pub final_section: Option<Value>,
    /// Each part's relationships, the main document part's included.
    pub relationships: HashMap<String, HashMap<String, RelationshipTarget>>,
    pub numbering: Arc<docx_parse::NumberingMap>,
    pub provenance: Provenance,
    pub warnings: Vec<String>,
    /// This replica's stories were seeded from the package, so recorded positions are pinned.
    pub pinned: bool,
    pub comment_writes: CommentWrites,
}

pub(crate) const DOCUMENT_PART: &str = "word/document.xml";
pub(crate) const FOOTNOTES_PART: &str = "word/footnotes.xml";
pub(crate) const ENDNOTES_PART: &str = "word/endnotes.xml";
pub(crate) const COMMENTS_PART: &str = "word/comments.xml";

fn field<'a>(value: Option<&'a Value>, key: &str) -> Option<&'a Value> {
    value.and_then(|value| value.get(key))
}

fn array(value: Option<&Value>) -> &[Value] {
    value.and_then(Value::as_array).map_or(&[], Vec::as_slice)
}

fn text(value: Option<&Value>) -> Option<String> {
    match value? {
        Value::String(value) if !value.is_empty() => Some(value.clone()),
        Value::Number(number) => Some(number_text(number)),
        _ => None,
    }
}

/// A JSON number the way JavaScript prints it, as the story ids seeding mints do.
pub(crate) fn number_text(number: &serde_json::Number) -> String {
    match number.as_f64() {
        Some(value) if value.is_finite() && value.fract() == 0.0 && value.abs() < 1e15 => {
            format!("{value:.0}")
        }
        _ => number.to_string(),
    }
}

fn targets(
    part: &str,
    relationships: impl IntoIterator<Item = docx_parse::Relationship>,
) -> HashMap<String, RelationshipTarget> {
    relationships
        .into_iter()
        .filter_map(|relationship| {
            let target = docx_parse::resolve_relationship_target(part, &relationship).ok()?;
            Some((
                relationship.id.clone(),
                match target {
                    docx_parse::RelationshipTarget::Internal(part) => {
                        RelationshipTarget::Part(part)
                    }
                    docx_parse::RelationshipTarget::External(url) => {
                        RelationshipTarget::External(url)
                    }
                },
            ))
        })
        .collect()
}

impl ReadSource {
    /// Reads the parts of `package` (the parsed package as JSON) the export needs.
    pub(crate) fn from_package(
        package: &Value,
        document_part: &str,
        warnings: Vec<String>,
    ) -> Self {
        let package = Some(package);
        let main = targets(
            document_part,
            array(field(package, "relationshipEntries"))
                .iter()
                .filter_map(|entry| serde_json::from_value(entry.get(1)?.clone()).ok()),
        );
        let mut stories: Vec<SourceStory> = Vec::new();
        let mut story_index = HashMap::new();
        for (key, kind) in [
            ("headerEntries", StoryKind::Header),
            ("footerEntries", StoryKind::Footer),
        ] {
            for entry in array(field(package, key)) {
                let Some(id) = entry.get(0).and_then(Value::as_str) else {
                    continue;
                };
                let story = format!("hf:{id}");
                if story_index.contains_key(&story) {
                    continue;
                }
                story_index.insert(story.clone(), stories.len());
                let part = match main.get(id) {
                    Some(RelationshipTarget::Part(part)) => Some(part.clone()),
                    _ => None,
                };
                stories.push(SourceStory {
                    story,
                    kind,
                    part,
                    note_id: None,
                });
            }
        }
        for (key, prefix, kind, part) in [
            ("footnotes", "fn", StoryKind::Footnote, FOOTNOTES_PART),
            ("endnotes", "en", StoryKind::Endnote, ENDNOTES_PART),
        ] {
            for note in array(field(package, key)) {
                let Some(id) = text(note.get("id")) else {
                    continue;
                };
                story_index
                    .entry(format!("{prefix}:{id}"))
                    .or_insert(stories.len());
                stories.push(SourceStory {
                    story: format!("{prefix}:{id}"),
                    kind,
                    part: Some(part.to_owned()),
                    note_id: Some(id),
                });
            }
        }
        let comments = array(field(field(package, "document"), "comments"))
            .iter()
            .filter_map(|comment| {
                let blocks = array(comment.get("blockContent"));
                Some(SourceComment {
                    id: text(comment.get("id"))?,
                    author: text(comment.get("author")),
                    date: text(comment.get("date")),
                    parent_id: text(comment.get("parentId")),
                    done: comment.get("done") == Some(&Value::Bool(true)),
                    body: if blocks.is_empty() {
                        array(comment.get("content")).to_vec()
                    } else {
                        blocks.to_vec()
                    },
                })
            })
            .collect();
        let numbering = field(package, "numbering")
            .and_then(|value| serde_json::from_value(value.clone()).ok())
            .map(|definitions| docx_parse::NumberingMap { definitions })
            .unwrap_or_default();
        Self {
            document_part: document_part.to_owned(),
            stories,
            story_index,
            footnote_separators: array(field(package, "footnoteSeparators")).len(),
            endnote_separators: array(field(package, "endnoteSeparators")).len(),
            comments,
            seeded_comments: HashSet::new(),
            raw_block_anchors: HashMap::new(),
            comment_anchors: HashMap::new(),
            final_section: field(field(package, "document"), "finalSectionProperties")
                .filter(|value| !value.is_null())
                .cloned(),
            relationships: HashMap::from([(document_part.to_owned(), main)]),
            numbering: Arc::new(numbering),
            provenance: Provenance::default(),
            warnings,
            pinned: false,
            comment_writes: CommentWrites::default(),
        }
    }

    /// Pins every recorded position to `doc`, which seeding just filled.
    pub(crate) fn pin(&mut self, doc: &EditingDoc) {
        self.provenance.pin(doc);
        self.pinned = true;
    }

    pub(crate) fn raw_block_anchor(&self, story: &str, index: usize) -> Option<Anchor> {
        self.raw_block_anchors.get(story)?.get(index)?.clone()
    }

    pub(crate) fn story(&self, story: &str) -> Option<&SourceStory> {
        self.stories.get(*self.story_index.get(story)?)
    }

    /// The part a story of the session belongs to.
    pub(crate) fn story_part(&self, story: &str) -> Option<String> {
        let root = story_root(story);
        if root == "body" {
            Some(self.document_part.clone())
        } else if root.starts_with("fn:") {
            Some(FOOTNOTES_PART.to_owned())
        } else if root.starts_with("en:") {
            Some(ENDNOTES_PART.to_owned())
        } else if root.starts_with("comment:") {
            Some(COMMENTS_PART.to_owned())
        } else {
            self.story(root).and_then(|story| story.part.clone())
        }
    }

    /// What relationship `id` of `part` targets; `None` when the part's relationships are unknown.
    pub(crate) fn relationship(&self, part: &str, id: &str) -> Option<Option<&RelationshipTarget>> {
        self.relationships
            .get(part)
            .map(|relationships| relationships.get(id))
    }
}

/// The story a table-cell or block-control story belongs to.
pub(crate) fn story_root(story: &str) -> &str {
    let cut = [":t", ":sdt"]
        .into_iter()
        .filter_map(|marker| story.find(marker))
        .min();
    cut.map_or(story, |index| &story[..index])
}

/// The table a cell story (`{story}:t{index}:r{row}c{cell}`) belongs to.
pub(crate) fn cell_table(cell_story: &str) -> Option<&str> {
    let cut = cell_story.rfind(":r")?;
    let cell = &cell_story[cut + 2..];
    let (row, column) = cell.split_once('c')?;
    (!row.is_empty()
        && !column.is_empty()
        && row.bytes().all(|byte| byte.is_ascii_digit())
        && column.bytes().all(|byte| byte.is_ascii_digit()))
    .then(|| &cell_story[..cut])
}

type Located<'a> = (Vec<u32>, &'a docx_parse::XmlElement);

/// What [`walk`] reads of each container, read once per container however many sources it
/// holds: the blocks, rows or cells in order, and the first `w:` child per name and per name and
/// `w:id`. Containers are keyed by address.
#[derive(Default)]
struct Containers<'a> {
    blocks: HashMap<usize, Vec<Located<'a>>>,
    rows: HashMap<usize, Vec<Located<'a>>>,
    cells: HashMap<usize, Vec<Located<'a>>>,
    children: HashMap<usize, HashMap<(String, Option<String>), (u32, &'a docx_parse::XmlElement)>>,
    /// How many containers have been read.
    reads: usize,
}

impl<'a> Containers<'a> {
    fn listed(
        lists: &mut HashMap<usize, Vec<Located<'a>>>,
        reads: &mut usize,
        container: &'a docx_parse::XmlElement,
        index: usize,
        list: impl FnOnce(&'a docx_parse::XmlElement) -> Vec<Located<'a>>,
    ) -> Option<Located<'a>> {
        let key = std::ptr::from_ref(container) as usize;
        lists
            .entry(key)
            .or_insert_with(|| {
                *reads += 1;
                list(container)
            })
            .get(index)
            .cloned()
    }

    fn child(
        &mut self,
        container: &'a docx_parse::XmlElement,
        local: &str,
        id: Option<&str>,
    ) -> Option<Located<'a>> {
        let key = std::ptr::from_ref(container) as usize;
        let reads = &mut self.reads;
        let children = self.children.entry(key).or_insert_with(|| {
            *reads += 1;
            let mut children = HashMap::new();
            for (index, element) in container.child_elements().enumerate() {
                let local = element.local_name();
                if !element.matches_name("w", local) {
                    continue;
                }
                let found = (index as u32, element);
                children.entry((local.to_owned(), None)).or_insert(found);
                if let Some(id) = element.attribute(Some("w"), "id") {
                    children
                        .entry((local.to_owned(), Some(id.trim().to_owned())))
                        .or_insert(found);
                }
            }
            children
        });
        let (index, element) = *children.get(&(local.to_owned(), id.map(str::to_owned)))?;
        Some((vec![index], element))
    }
}

fn walk<'a>(
    root: &'a docx_parse::XmlElement,
    steps: &[Step],
    containers: &mut Containers<'a>,
) -> Option<Located<'a>> {
    let mut path = Vec::new();
    let mut current = root;
    for step in steps {
        let (relative, next) = match step {
            Step::Body => containers.child(current, "body", None)?,
            Step::Note(local, id) => containers.child(current, local, Some(id))?,
            Step::Comment(id) => containers.child(current, "comment", Some(id))?,
            Step::Block(index) => Containers::listed(
                &mut containers.blocks,
                &mut containers.reads,
                current,
                *index,
                docx_parse::story_block_elements,
            )?,
            Step::Row(index) => Containers::listed(
                &mut containers.rows,
                &mut containers.reads,
                current,
                *index,
                |table| docx_parse::table_part_elements(table, "tr"),
            )?,
            Step::Cell(index) => Containers::listed(
                &mut containers.cells,
                &mut containers.reads,
                current,
                *index,
                |row| docx_parse::table_part_elements(row, "tc"),
            )?,
            Step::Content => containers.child(current, "sdtContent", None)?,
        };
        path.extend(relative);
        current = next;
    }
    Some((path, current))
}

fn sha256(bytes: &[u8]) -> String {
    use sha2::{Digest, Sha256};
    Sha256::digest(bytes)
        .iter()
        .fold(String::with_capacity(64), |mut output, byte| {
            use std::fmt::Write as _;
            let _ = write!(output, "{byte:02x}");
            output
        })
}

/// A parsed part and its digest.
struct ParsedPart {
    document: docx_parse::XmlDocument,
    digest: String,
}

fn parse_part(parts: &SourceParts, part: &str) -> Option<ParsedPart> {
    let bytes = parts.part(part)?;
    let limits = docx_parse::ParseLimits::default();
    let document =
        docx_parse::parse_xml(bytes, part, &mut docx_parse::ParseBudget::new(&limits)).ok()?;
    Some(ParsedPart {
        document,
        digest: sha256(bytes),
    })
}

impl ReadSource {
    /// Resolves recorded source locations against the package: each raw block and comment body
    /// raw block to its element (kept only when it serializes to the recorded XML), each comment
    /// to its element, and the relationships of every part a story belongs to. Reads each
    /// container once, and returns how many it read.
    pub(crate) fn resolve_sources(
        &mut self,
        parts: &SourceParts,
        comment_raw: Vec<RawSource>,
    ) -> usize {
        let mut reads = 0;
        self.resolve_relationships(parts);
        let mut by_part: HashMap<String, Vec<RawSource>> = HashMap::new();
        for source in std::mem::take(&mut self.provenance.raw_sources)
            .into_iter()
            .chain(comment_raw)
        {
            if let Some(part) = self.story_part(&source.story) {
                by_part.entry(part).or_default().push(source);
            }
        }
        if !self.comments.is_empty() {
            by_part.entry(COMMENTS_PART.to_owned()).or_default();
        }
        for (part, items) in by_part {
            let Some(parsed) = parse_part(parts, &part) else {
                continue;
            };
            let Some(root) = parsed.document.root() else {
                continue;
            };
            let anchor = |path: Vec<u32>| Anchor::SourcePart {
                part: part.clone(),
                part_sha256: parsed.digest.clone(),
                path,
            };
            if part == COMMENTS_PART {
                let ids: HashSet<&str> = self
                    .comments
                    .iter()
                    .map(|comment| comment.id.as_str())
                    .collect();
                for (index, element) in root.child_elements().enumerate() {
                    let Some(id) = element
                        .matches_name("w", "comment")
                        .then(|| element.attribute(Some("w"), "id").map(str::trim))
                        .flatten()
                    else {
                        continue;
                    };
                    if ids.contains(id) {
                        self.comment_anchors
                            .entry(id.to_owned())
                            .or_insert_with(|| anchor(vec![index as u32]));
                    }
                }
            }
            let mut containers = Containers::default();
            for source in items {
                let Some((path, element)) = walk(root, &source.steps, &mut containers) else {
                    continue;
                };
                if element.to_raw_inline_xml() != source.xml {
                    continue;
                }
                let anchors = self.raw_block_anchors.entry(source.story).or_default();
                if anchors.len() <= source.index {
                    anchors.resize(source.index + 1, None);
                }
                anchors[source.index] = Some(anchor(path));
            }
            reads += containers.reads;
        }
        reads
    }

    fn resolve_relationships(&mut self, parts: &SourceParts) {
        let mut owners: Vec<String> = self
            .stories
            .iter()
            .filter_map(|story| story.part.clone())
            .collect();
        owners.extend([FOOTNOTES_PART, ENDNOTES_PART, COMMENTS_PART].map(str::to_owned));
        let limits = docx_parse::ParseLimits::default();
        for part in owners {
            if self.relationships.contains_key(&part) || parts.part(&part).is_none() {
                continue;
            }
            let rels = relationship_part(&part);
            let relationships = match parts.part(&rels) {
                Some(bytes) => docx_parse::relationships::parse_relationships(
                    bytes,
                    &rels,
                    &mut docx_parse::ParseBudget::new(&limits),
                )
                .map(|map| targets(&part, map.into_values()))
                .unwrap_or_default(),
                None => HashMap::new(),
            };
            self.relationships.insert(part, relationships);
        }
    }
}

/// The qualified name of the element a raw XML string opens with.
pub(crate) fn element_name(xml: &str) -> String {
    let start = xml.find('<').map_or(0, |index| index + 1);
    let name: String = xml[start..]
        .chars()
        .take_while(|ch| !ch.is_whitespace() && !matches!(ch, '/' | '>'))
        .collect();
    if name.is_empty() {
        "rawXml".to_owned()
    } else {
        name
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolving_source_paths_reads_each_container_once() {
        let blocks = 20_000;
        let foreign = r#"<x:note xmlns:x="urn:example"/>"#;
        let xml = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{}<w:p/></w:body></w:document>"#,
            foreign.repeat(blocks)
        );
        let parts = SourceParts::new(vec![(DOCUMENT_PART.to_owned(), xml.into_bytes())]);
        let parsed = parse_part(&parts, DOCUMENT_PART).unwrap();
        let body = parsed
            .document
            .root()
            .unwrap()
            .child_elements()
            .next()
            .unwrap();
        let recorded = body.child_elements().next().unwrap().to_raw_inline_xml();
        let mut read = ReadSource {
            document_part: DOCUMENT_PART.to_owned(),
            ..ReadSource::default()
        };
        read.provenance.raw_sources = (0..blocks)
            .map(|index| RawSource {
                story: "body".to_owned(),
                index,
                steps: vec![Step::Body, Step::Block(index)],
                xml: recorded.clone(),
            })
            .collect();
        assert_eq!(read.resolve_sources(&parts, Vec::new()), 2);
        for index in [0, blocks - 1] {
            assert!(matches!(
                read.raw_block_anchor("body", index),
                Some(Anchor::SourcePart { path, .. }) if path == [0, index as u32]
            ));
        }
    }
}
