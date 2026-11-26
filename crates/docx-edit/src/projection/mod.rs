use std::collections::{BTreeMap, BTreeSet};
use std::sync::{Arc, Mutex, OnceLock};

use docx_parse::paragraph::{Paragraph, ParagraphContent};
use docx_parse::s9::{S9PackageWire, S9SectionWire, S9WireEnvelope};
use docx_parse::{BlockContent, S9DocumentWire};
use serde::{Deserialize, Deserializer, Serialize};
use serde_json::Value;
use yrs::{Any, ReadTxn, Transact, Transaction};

use crate::identity::{SourceIndex, SourcePackage};
use crate::seed::RunInheritance;
use crate::structured::source::Provenance;
use crate::{EditError, EditingDoc, ParagraphProperties, StorySegment};

mod fields_dispatch;
mod marker_index;
mod markers_dispatch;
mod paragraph;
mod runs;
mod stories_aux_dispatch;
mod story;
mod tables_dispatch;

pub(crate) use marker_index::MarkerIndex;

#[derive(Clone, Debug, Default, PartialEq, Deserialize)]
#[serde(default, rename_all = "camelCase", deny_unknown_fields)]
pub struct ProjectionInput {
    pub root_stories: Option<BTreeSet<String>>,
    pub host: ProjectionHostMetadata,
    pub comment_ids: Vec<(f64, String)>,
    pub crossing_comment_fallback_ids: Vec<f64>,
    pub revision_ids: BTreeMap<String, ProjectionRevisionIds>,
}

#[derive(Clone, Debug, Default, PartialEq, Deserialize)]
#[serde(default, rename_all = "camelCase", deny_unknown_fields)]
pub struct ProjectionRevisionIds {
    pub insertion: Option<f64>,
    pub deletion: Option<f64>,
}

type Override<T> = Option<Option<T>>;

fn deserialize_override<'de, D, T>(deserializer: D) -> Result<Override<T>, D::Error>
where
    D: Deserializer<'de>,
    T: Deserialize<'de>,
{
    Option::<T>::deserialize(deserializer).map(Some)
}

#[derive(Clone, Debug, Default, PartialEq, Deserialize)]
#[serde(default, rename_all = "camelCase", deny_unknown_fields)]
pub struct ProjectionHostMetadata {
    #[serde(deserialize_with = "deserialize_override")]
    pub styles: Override<docx_parse::styles::StyleDefinitions>,
    #[serde(deserialize_with = "deserialize_override")]
    pub theme: Override<docx_parse::theme::Theme>,
    #[serde(deserialize_with = "deserialize_override")]
    pub settings: Override<docx_parse::settings::DocumentSettings>,
    #[serde(deserialize_with = "deserialize_override")]
    pub font_table: Override<docx_parse::fonts::FontTable>,
    #[serde(deserialize_with = "deserialize_override")]
    pub relationships: Override<Vec<(String, docx_parse::relationships::Relationship)>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub warnings: Override<Vec<String>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub comments: Override<Vec<docx_parse::comments::Comment>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub headers: Override<Vec<(String, docx_parse::header_footer::HeaderFooter)>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub footers: Override<Vec<(String, docx_parse::header_footer::HeaderFooter)>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub footnotes: Override<Vec<docx_parse::notes::Note>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub endnotes: Override<Vec<docx_parse::notes::Note>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub footnote_separators: Override<Vec<docx_parse::notes::Note>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub endnote_separators: Override<Vec<docx_parse::notes::Note>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub sections: Override<Vec<S9SectionWire>>,
    #[serde(deserialize_with = "deserialize_override")]
    pub final_section_properties: Override<docx_parse::section::SectionProperties>,
    pub media_urls: Option<BTreeMap<String, String>>,
}

#[derive(Debug)]
pub enum ProjectionError {
    MissingSource,
    SourceParse(String),
    StoryRead { story_id: String, error: EditError },
    InvalidInput(String),
    Json(serde_json::Error),
}

impl std::fmt::Display for ProjectionError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::MissingSource => f.write_str("the DOCX session has no retained source package"),
            Self::SourceParse(error) => {
                write!(f, "could not parse the retained DOCX source: {error}")
            }
            Self::StoryRead { story_id, error } => {
                write!(f, "could not read projection story {story_id}: {error}")
            }
            Self::InvalidInput(error) => write!(f, "invalid projection input: {error}"),
            Self::Json(error) => write!(f, "invalid projection data: {error}"),
        }
    }
}

impl std::error::Error for ProjectionError {}

impl From<serde_json::Error> for ProjectionError {
    fn from(error: serde_json::Error) -> Self {
        Self::Json(error)
    }
}

pub struct ProjectionOutput {
    pub document: docx_parse::S9DocumentWire,
    pub paragraph_session_keys: Vec<String>,
    pub session_revision_placeholders: std::collections::BTreeSet<u64>,
    pub created_list_sites: Vec<CreatedListSite>,
    pub deferred: Vec<ProjectionDeferral>,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct CreatedListSite {
    pub paragraph: usize,
    pub slot: ListSlot,
    pub key: String,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ListSlot {
    Current,
    Previous,
}

#[derive(Clone, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct ProjectionDeferral {
    pub family: ProjectionFamily,
    pub kind: String,
    pub story_id: String,
    pub paragraph: Option<usize>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
#[serde(rename_all = "camelCase")]
pub enum ProjectionFamily {
    Fields,
    Tables,
    Markers,
    StoriesAux,
}

#[doc(hidden)]
#[derive(Clone, Debug, Default, PartialEq, Serialize)]
#[serde(rename_all = "camelCase")]
pub struct StoryProjection {
    pub story_id: String,
    pub blocks: Vec<docx_parse::BlockContent>,
    pub paragraph_session_keys: Vec<String>,
    pub session_revision_placeholders: BTreeSet<u64>,
    pub created_list_sites: Vec<CreatedListSite>,
    pub deferred: Vec<ProjectionDeferral>,
}

impl StoryProjection {
    fn defer(
        &mut self,
        family: ProjectionFamily,
        kind: &str,
        story_id: &str,
        paragraph: Option<usize>,
    ) {
        self.deferred.push(ProjectionDeferral {
            family,
            kind: kind.to_owned(),
            story_id: story_id.to_owned(),
            paragraph,
        });
    }

    fn append(&mut self, mut child: Self) {
        self.append_side_tables(&child);
        self.blocks.append(&mut child.blocks);
    }

    fn append_side_tables(&mut self, child: &Self) {
        let offset = self.paragraph_session_keys.len();
        self.created_list_sites
            .extend(child.created_list_sites.iter().cloned().map(|mut site| {
                site.paragraph += offset;
                site
            }));
        self.paragraph_session_keys
            .extend_from_slice(&child.paragraph_session_keys);
        self.session_revision_placeholders
            .extend(&child.session_revision_placeholders);
        self.deferred.extend_from_slice(&child.deferred);
    }
}

#[derive(Clone, Copy)]
enum SourceRoot {
    Header(usize),
    Footer(usize),
    Footnote(usize),
    Endnote(usize),
}

struct SourceCache {
    envelope: Arc<S9WireEnvelope>,
    roots: OnceLock<BTreeMap<String, SourceRoot>>,
    stories: BTreeMap<String, Arc<SourceStory>>,
    materialize_media: bool,
    media: Option<Arc<docx_parse::media::MediaTable>>,
    materialized_document: Option<Arc<S9DocumentWire>>,
}

impl SourceCache {
    fn new(
        envelope: Arc<S9WireEnvelope>,
        materialize_media: bool,
        media: Option<docx_parse::media::MediaTable>,
    ) -> Self {
        Self {
            envelope,
            roots: OnceLock::new(),
            stories: BTreeMap::new(),
            materialize_media,
            media: media.map(Arc::new),
            materialized_document: None,
        }
    }

    fn roots(&self) -> &BTreeMap<String, SourceRoot> {
        self.roots.get_or_init(|| {
            let mut roots = BTreeMap::new();
            for (id, part, _) in auxiliary_roots(&self.envelope.document.package) {
                roots.entry(id).or_insert(part);
            }
            roots
        })
    }

    fn media(&self, doc: &EditingDoc) -> Option<Arc<docx_parse::media::MediaTable>> {
        if !self.materialize_media {
            return None;
        }
        doc.media_table()
            .or_else(|| self.media.clone())
            .filter(|table| !table.is_empty())
    }

    fn document(&mut self, doc: &EditingDoc) -> Result<S9DocumentWire, ProjectionError> {
        if let Some(document) = &self.materialized_document {
            return Ok((**document).clone());
        }
        let mut document = self.envelope.document.clone();
        if let Some(table) = self.media(doc) {
            materialize_package(&mut document, &table)?;
            self.materialized_document = Some(Arc::new(document.clone()));
        }
        Ok(document)
    }

    fn blocks(&self, id: &str) -> &[BlockContent] {
        let package = &self.envelope.document.package;
        if id == "body" {
            return &package.document.content;
        }
        match self.roots().get(id) {
            Some(SourceRoot::Header(index)) => {
                &package.header_entries.as_ref().unwrap()[*index].1.content
            }
            Some(SourceRoot::Footer(index)) => {
                &package.footer_entries.as_ref().unwrap()[*index].1.content
            }
            Some(SourceRoot::Footnote(index)) => {
                &package.footnotes.as_ref().unwrap()[*index].content
            }
            Some(SourceRoot::Endnote(index)) => &package.endnotes.as_ref().unwrap()[*index].content,
            None => &[],
        }
    }

    fn story(&mut self, doc: &EditingDoc, id: &str) -> Result<Arc<SourceStory>, ProjectionError> {
        if let Some(story) = self.stories.get(id) {
            return Ok(Arc::clone(story));
        }
        if id != "body" && !self.roots().contains_key(id) {
            return Ok(Arc::new(SourceStory::default()));
        }
        let blocks = self.blocks(id);
        let blocks = if let Some(table) = self.media(doc) {
            materialize(&blocks.to_vec(), &table)?
        } else {
            blocks.to_vec()
        };
        let mut source = SourceStory {
            blocks,
            ..SourceStory::default()
        };
        if id == "body" {
            source.index(id);
        }
        let source = Arc::new(source);
        self.stories.insert(id.to_owned(), Arc::clone(&source));
        Ok(source)
    }
}

#[derive(Default)]
struct SourceStory {
    blocks: Vec<BlockContent>,
    paragraphs: BTreeMap<String, Arc<Paragraph>>,
    containers: BTreeMap<String, BlockContent>,
    children: Mutex<BTreeMap<String, Arc<SourceStory>>>,
    table_cells: Mutex<BTreeMap<String, Vec<Vec<usize>>>>,
    has_comment_ranges: bool,
    has_revision_ranges: bool,
}

impl SourceStory {
    fn index(&mut self, story_id: &str) {
        let (mut paragraph, mut table, mut sdt) = (0, 0, 0);
        for block in &self.blocks {
            match block {
                BlockContent::Paragraph(source) => {
                    let key = source
                        .para_id
                        .as_ref()
                        .filter(|id| !id.is_empty() && source.repeated_para_id != Some(true))
                        .cloned()
                        .unwrap_or_else(|| format!("{story_id}:p{paragraph}"));
                    self.paragraphs
                        .entry(key)
                        .or_insert_with(|| Arc::clone(source));
                    self.has_comment_ranges |= story::has_comment_ranges(block);
                    self.has_revision_ranges |= story::has_revision_ranges(block);
                    paragraph += 1;
                }
                BlockContent::Table(_) => {
                    self.containers
                        .insert(format!("{story_id}:t{table}"), block.clone());
                    table += 1;
                }
                BlockContent::BlockSdt(_) => {
                    self.containers
                        .insert(format!("{story_id}:sdt{sdt}"), block.clone());
                    sdt += 1;
                }
                BlockContent::RawXml(_) => {}
            }
        }
    }

    fn child(&self, id: &str) -> Result<Option<Arc<Self>>, ProjectionError> {
        let mut children = self.children.lock().unwrap();
        if let Some(child) = children.get(id) {
            return Ok(Some(Arc::clone(child)));
        }
        let blocks = if let Some(BlockContent::BlockSdt(source)) = self.containers.get(id) {
            &source.content
        } else {
            let Some((table_id, table, row, cell)) = self.cell_location(id) else {
                return Ok(None);
            };
            let mut table_cells = self.table_cells.lock().unwrap();
            let cells = match table_cells.entry(table_id.to_owned()) {
                std::collections::btree_map::Entry::Occupied(entry) => entry.into_mut(),
                std::collections::btree_map::Entry::Vacant(entry) => {
                    entry.insert(crate::seed::source_cells(&json_value(table)?))
                }
            };
            let Some(cell) = cells
                .get(row)
                .and_then(|cells| cells.get(cell))
                .and_then(|&cell| table.rows.get(row)?.cells.get(cell))
            else {
                return Ok(None);
            };
            &cell.content
        };
        let mut child = Self {
            blocks: blocks.clone(),
            ..Self::default()
        };
        child.index(id);
        let child = Arc::new(child);
        children.insert(id.to_owned(), Arc::clone(&child));
        Ok(Some(child))
    }

    fn cell_location<'a>(
        &self,
        id: &'a str,
    ) -> Option<(&'a str, &docx_parse::table::Table, usize, usize)> {
        let table_id = crate::structured::source::cell_table(id)?;
        let BlockContent::Table(table) = self.containers.get(table_id)? else {
            return None;
        };
        let (row, cell) = id.get(table_id.len() + 2..)?.split_once('c')?;
        Some((table_id, table, row.parse().ok()?, cell.parse().ok()?))
    }
}

#[derive(Default)]
pub(crate) struct ProjectionCache {
    source: Option<Result<SourceCache, String>>,
    generation: Option<String>,
    aliases: crate::header_footer::ProjectionAliases,
    styles: Option<(
        Override<docx_parse::styles::StyleDefinitions>,
        RunInheritance,
    )>,
    theme: Option<Override<docx_parse::theme::Theme>>,
    sections: Option<bool>,
    body_source_index: Option<Arc<SourceIndex>>,
    stories: BTreeMap<String, CachedStory>,
}

struct CachedStory {
    revision: u64,
    markers_revision: u64,
    projection: StoryProjection,
}

impl ProjectionCache {
    fn resolve_source(source: &SourcePackage) -> Result<SourceCache, String> {
        let (bytes, digest) = match source {
            SourcePackage::Ready(index) => {
                if let Some(envelope) = &index.projection_envelope {
                    return Ok(SourceCache::new(
                        Arc::clone(envelope),
                        true,
                        index.projection_media.clone(),
                    ));
                }
                (index.bytes(), Some(index.package_digest().to_owned()))
            }
            SourcePackage::Pending(bytes, digest) => (bytes.clone(), digest.clone()),
        };
        let digest = digest.unwrap_or_else(|| crate::seed::package_digest(&bytes));
        crate::seed::parse_docx_package_with_digest(&bytes, digest)
            .map(|(envelope, _)| SourceCache::new(Arc::new(envelope), false, None))
    }

    fn source(&mut self, doc: &EditingDoc) -> Result<&mut SourceCache, ProjectionError> {
        if self.source.is_none() {
            let retained = doc.source.lock().unwrap();
            let source = retained.as_ref().ok_or(ProjectionError::MissingSource)?;
            self.source = Some(Self::resolve_source(source));
        }
        self.source
            .as_mut()
            .ok_or(ProjectionError::MissingSource)?
            .as_mut()
            .map_err(|error| ProjectionError::SourceParse(error.clone()))
    }

    pub(crate) fn reset_opening(&mut self) {
        let source = self.source.take();
        *self = Self {
            source,
            ..Self::default()
        };
        if let Some(Ok(source)) = &mut self.source {
            source.stories.clear();
            source.materialized_document = None;
        }
    }

    pub(crate) fn invalidate_stories(&mut self) {
        self.stories.clear();
        self.styles = None;
        self.theme = None;
        self.sections = None;
    }

    fn begin_generation(&mut self, generation: Option<String>) {
        if self.generation != generation {
            self.generation = generation;
            self.invalidate_stories();
            if let Some(Ok(source)) = &mut self.source {
                source.stories.clear();
                source.materialized_document = None;
            }
        }
    }
}

struct ProjectionContext<'a, 'doc> {
    doc: &'a EditingDoc,
    txn: &'a Transaction<'doc>,
    input: &'a ProjectionInput,
    source: &'a SourceStory,
    package_body_has_sections: bool,
    markers: &'a MarkerIndex,
    source_index: Option<&'a SourceIndex>,
    source_provenance: Option<&'a Provenance>,
    inheritance: &'a mut RunInheritance,
    theme: &'a docx_parse::theme::Theme,
    visiting: BTreeSet<String>,
}

impl ProjectionContext<'_, '_> {
    #[allow(dead_code)]
    fn project_content_story(
        &mut self,
        id: &str,
        state: StoryState,
    ) -> Result<Arc<StoryProjection>, ProjectionError> {
        if !self.visiting.insert(id.to_owned()) {
            return Err(ProjectionError::InvalidInput(format!("cyclic story {id}")));
        }
        let source = self.source.child(id)?;
        let empty = SourceStory::default();
        let mut child = ProjectionContext {
            doc: self.doc,
            txn: self.txn,
            input: self.input,
            source: source.as_deref().unwrap_or(&empty),
            package_body_has_sections: false,
            markers: self.markers,
            source_index: self.source_index,
            source_provenance: self.source_provenance,
            inheritance: self.inheritance,
            theme: self.theme,
            visiting: std::mem::take(&mut self.visiting),
        };
        let result = story::project(&mut child, id, state);
        self.visiting = child.visiting;
        self.visiting.remove(id);
        result.map(Arc::new)
    }
}

type Attrs = BTreeMap<String, Any>;

#[derive(Default)]
#[allow(dead_code)]
struct StoryState {
    paragraph_index: usize,
    offset: usize,
    paragraph_start: usize,
    table_index: usize,
    sdt_index: usize,
    table_paragraph_formatting: Option<Value>,
    table_run_formatting: Option<Value>,
    attributes: Attrs,
    fields: fields_dispatch::StoryData,
    tables: tables_dispatch::StoryData,
    markers: markers_dispatch::StoryData,
    aux: stories_aux_dispatch::StoryData,
}

#[derive(Default)]
struct StoryPosition<'a> {
    id: &'a str,
    state: StoryState,
    result: StoryProjection,
}

enum Lowered {
    Inline(Vec<ParagraphContent>),
    Blocks(StoryProjection),
    Deferred,
}

#[allow(dead_code)]
struct LoweringInput<'a> {
    kind: &'a str,
    payload: &'a Attrs,
    segment: &'a StorySegment,
}

#[allow(dead_code)]
impl LoweringInput<'_> {
    fn control_metadata(&self, control_id: String) -> crate::read_types::ControlMetadata {
        crate::read_types::control_metadata(control_id, |key| self.payload.get(key))
    }

    fn drawing(&self, image: bool) -> Option<crate::DrawingProperties> {
        let payload = self
            .payload
            .iter()
            .map(|(key, value)| (key.clone(), value.clone()))
            .collect();
        crate::target::drawing_from_payload(&payload, image)
    }

    fn payload_json(&self, key: &str) -> Result<Option<Value>, ProjectionError> {
        self.payload
            .get(key)
            .map(|value| {
                let Any::String(json) = value else {
                    return Err(ProjectionError::InvalidInput(format!(
                        "{key} must be a JSON string"
                    )));
                };
                crate::json_tree::parse(json)
                    .map(crate::json_tree::OrderedValue::into_value)
                    .map_err(ProjectionError::InvalidInput)
            })
            .transpose()
    }
}

trait ProjectionHooks: Sync {
    fn family(&self) -> ProjectionFamily;

    fn setup_story(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        _story: &mut StoryPosition<'_>,
        _segments: &[StorySegment],
    ) -> Result<(), ProjectionError> {
        Ok(())
    }

    fn lower(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        _story: &mut StoryPosition<'_>,
        _input: LoweringInput<'_>,
    ) -> Result<Lowered, ProjectionError> {
        Ok(Lowered::Deferred)
    }

    fn wrap(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        _story: &mut StoryPosition<'_>,
        _segment: &StorySegment,
        _lowered: &mut Lowered,
    ) -> Result<(), ProjectionError> {
        Ok(())
    }

    fn finish_paragraph(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        _story: &mut StoryPosition<'_>,
        _properties: Option<&ParagraphProperties>,
        _source: Option<&Paragraph>,
        _paragraph: &mut Paragraph,
    ) -> Result<(), ProjectionError> {
        Ok(())
    }

    fn finish_story(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        _story: &mut StoryPosition<'_>,
    ) -> Result<(), ProjectionError> {
        Ok(())
    }
}

const HOOKS: [&dyn ProjectionHooks; 4] = [
    &fields_dispatch::Hooks,
    &tables_dispatch::Hooks,
    &markers_dispatch::Hooks,
    &stories_aux_dispatch::Hooks,
];

fn embed_hook(kind: &str, payload: &Attrs) -> &'static dyn ProjectionHooks {
    match kind {
        "table" => &tables_dispatch::Hooks,
        "blockSdt" => &stories_aux_dispatch::Hooks,
        "opaque" if story::opaque_type(payload) == Some("blockSdt") => &stories_aux_dispatch::Hooks,
        "rawInline" => &markers_dispatch::Hooks,
        "field" if attr_string(payload, "modelKind") == Some("commentReference") => {
            &markers_dispatch::Hooks
        }
        _ => &fields_dispatch::Hooks,
    }
}

pub(crate) fn projection_root(id: &str) -> Option<&str> {
    let mut parts = id.splitn(3, ':');
    match (parts.next()?, parts.next()) {
        ("body", _) => Some("body"),
        ("hf" | "fn" | "en", Some(item)) if !item.is_empty() => Some(&id[..3 + item.len()]),
        _ => None,
    }
}

fn root_id(id: &str) -> bool {
    projection_root(id) == Some(id)
}

fn validate_root(id: &str) -> Result<(), ProjectionError> {
    if !root_id(id) {
        return Err(ProjectionError::InvalidInput(format!("{id} is not a root story")));
    }
    Ok(())
}

fn attr_string<'a>(attrs: &'a Attrs, key: &str) -> Option<&'a str> {
    crate::read_types::payload_str(attrs.get(key))
}

fn json_value<T: Serialize>(value: &T) -> Result<Value, ProjectionError> {
    Ok(serde_json::to_value(value)?)
}

fn decode<T: serde::de::DeserializeOwned>(
    value: Option<&Value>,
) -> Result<Option<T>, ProjectionError> {
    value
        .map(|value| T::deserialize(value).map_err(ProjectionError::from))
        .transpose()
}

fn materialize<T: Clone + Serialize + serde::de::DeserializeOwned>(
    value: &T,
    table: &docx_parse::media::MediaTable,
) -> Result<T, ProjectionError> {
    let mut wire = json_value(value)?;
    let mut pending = vec![&mut wire];
    let mut urls = BTreeMap::new();
    let mut changed = false;
    while let Some(value) = pending.pop() {
        match value {
            Value::Array(values) => pending.extend(values.iter_mut()),
            Value::Object(fields) => {
                for (key, value) in fields {
                    if matches!(key.as_str(), "src" | "dataUrl")
                        && let Some(index) = value
                            .as_str()
                            .and_then(docx_parse::media::media_token_index)
                    {
                        let url = match urls.entry(index) {
                            std::collections::btree_map::Entry::Occupied(entry) => entry.into_mut(),
                            std::collections::btree_map::Entry::Vacant(entry) => entry.insert(
                                table
                                    .data_url(index)
                                    .map_err(ProjectionError::SourceParse)?,
                            ),
                        };
                        *value = Value::String(url.clone());
                        changed = true;
                    } else {
                        pending.push(value);
                    }
                }
            }
            _ => {}
        }
    }
    if changed {
        Ok(serde_json::from_value(wire)?)
    } else {
        Ok(value.clone())
    }
}

fn materialize_package(
    document: &mut S9DocumentWire,
    table: &docx_parse::media::MediaTable,
) -> Result<(), ProjectionError> {
    let package = &mut document.package;
    package.document.content = materialize(&package.document.content, table)?;
    package.document.comments = materialize(&package.document.comments, table)?;
    for entries in [&mut package.header_entries, &mut package.footer_entries] {
        for (_, part) in entries.iter_mut().flatten() {
            part.content = materialize(&part.content, table)?;
        }
    }
    for notes in [
        &mut package.footnotes,
        &mut package.endnotes,
        &mut package.footnote_separators,
        &mut package.endnote_separators,
    ] {
        for note in notes.iter_mut().flatten() {
            note.content = materialize(&note.content, table)?;
        }
    }
    let mut urls = BTreeMap::new();
    package.media_entries = table
        .media_map()
        .into_iter()
        .map(|(path, file)| {
            let index =
                docx_parse::media::media_token_index(&file.data_url).expect("media table token");
            let file = match urls.entry(index) {
                std::collections::btree_map::Entry::Occupied(entry) => Arc::clone(entry.get()),
                std::collections::btree_map::Entry::Vacant(entry) => {
                    let mut file = (*file).clone();
                    file.data_url = table
                        .data_url(index)
                        .map_err(ProjectionError::SourceParse)?;
                    file.base64 = file
                        .data_url
                        .split_once(',')
                        .map_or("", |(_, base64)| base64)
                        .to_owned();
                    Arc::clone(entry.insert(Arc::new(file)))
                }
            };
            Ok((path, file))
        })
        .collect::<Result<_, ProjectionError>>()?;
    Ok(())
}

fn validate_input(input: &ProjectionInput) -> Result<(), ProjectionError> {
    if input.host.media_urls.is_some() {
        return Err(ProjectionError::InvalidInput(
            "host.mediaUrls is not supported by projection".to_owned(),
        ));
    }
    let numbers = input
        .comment_ids
        .iter()
        .map(|(number, _)| *number)
        .chain(input.crossing_comment_fallback_ids.iter().copied())
        .chain(
            input
                .revision_ids
                .values()
                .flat_map(|ids| [ids.insertion, ids.deletion])
                .flatten(),
        );
    if numbers.into_iter().any(|number| !number.is_finite()) {
        return Err(ProjectionError::InvalidInput(
            "non-finite OOXML id".to_owned(),
        ));
    }
    Ok(())
}

impl EditingDoc {
    #[doc(hidden)]
    pub fn project_document(
        &self,
        input: &ProjectionInput,
    ) -> Result<ProjectionOutput, ProjectionError> {
        self.project_package(input).map(|(output, _)| output)
    }

    fn project_package(
        &self,
        input: &ProjectionInput,
    ) -> Result<(ProjectionOutput, Arc<S9WireEnvelope>), ProjectionError> {
        validate_input(input)?;
        for id in input.root_stories.iter().flatten() {
            validate_root(id)?;
        }
        let mut cache = self.projection.lock().unwrap();
        let txn = self.doc.transact();
        cache.begin_generation(projection_generation(&txn));
        let aliases = cache.aliases.active_aliases(self, &txn);
        let (mut document, envelope) = {
            let source = cache.source(self)?;
            (source.document(self)?, Arc::clone(&source.envelope))
        };
        apply_host_metadata(&mut document, &input.host);
        let mut projections = BTreeMap::new();
        {
            let mut roots = BTreeMap::new();
            for (id, kind, blocks) in auxiliary_roots(&document.package) {
                let canonical = aliases.get(&id).unwrap_or(&id);
                let overridden = input.host.replaces_auxiliary(Some(kind), &id);
                let source = (overridden || canonical != &id).then_some(blocks);
                let entry = roots
                    .entry(canonical.clone())
                    .or_insert_with(|| (id.clone(), source));
                if &entry.0 != canonical && (canonical == &id || id < entry.0) {
                    *entry = (id, source);
                }
            }
            let source = cache.source(self)?;
            for (id, members) in &self.story_revisions.lock().unwrap().live_projection_roots {
                if members.contains(id.as_str())
                    && !input.host.replaces_auxiliary(source.roots().get(id).copied(), id)
                {
                    roots
                        .entry(id.clone())
                        .or_insert_with(|| (id.clone(), None));
                }
            }
            cache.stories.retain(|id, _| roots.contains_key(id));
            for id in input.root_stories.iter().flatten() {
                cache.stories.remove(aliases.get(id).unwrap_or(id));
            }
            for (id, (_, blocks)) in roots {
                let projection = self.project_root(&mut cache, &txn, &id, input, blocks)?;
                projections.insert(id, projection);
            }
        }
        assemble_stories(&mut document, &projections, &aliases, input.host.sections.is_some())?;
        let mut side_tables = StoryProjection::default();
        for projection in projections.values() {
            side_tables.append_side_tables(projection);
        }
        Ok((
            ProjectionOutput {
                document,
                paragraph_session_keys: side_tables.paragraph_session_keys,
                session_revision_placeholders: side_tables.session_revision_placeholders,
                created_list_sites: side_tables.created_list_sites,
                deferred: side_tables.deferred,
            },
            envelope,
        ))
    }

    #[doc(hidden)]
    pub fn project_story(
        &self,
        story_id: &str,
        input: &ProjectionInput,
    ) -> Result<StoryProjection, ProjectionError> {
        validate_input(input)?;
        validate_root(story_id)?;
        let mut cache = self.projection.lock().unwrap();
        let txn = self.doc.transact();
        cache.begin_generation(projection_generation(&txn));
        let group = cache.aliases.group(self, &txn, story_id);
        let id = group.as_ref().map_or(story_id, |group| group[0].as_str());
        let source = cache.source(self)?;
        let mut alternate = None;
        let mut override_blocks = None;
        let mut retained = false;
        if id != "body" {
            let candidates = std::iter::once(id)
                .chain(group.iter().flat_map(|ids| ids[1..].iter().map(String::as_str)));
            for candidate in candidates {
                let root = source.roots().get(candidate).copied();
                override_blocks = input.host.auxiliary_blocks(candidate, root);
                retained = root.is_some() && !input.host.replaces_auxiliary(root, candidate);
                if override_blocks.is_some() || retained {
                    if override_blocks.is_none() && candidate != id {
                        alternate = Some(source.story(self, candidate)?);
                    }
                    break;
                }
            }
        }
        if !retained
            && override_blocks.is_none()
            && (crate::story_text(&txn, id).is_none()
                || (id != "body"
                    && input.host.replaces_auxiliary(source.roots().get(id).copied(), id)))
        {
            return Err(ProjectionError::StoryRead {
                story_id: story_id.to_owned(),
                error: EditError::StoryNotFound(id.to_owned()),
            });
        }
        let override_blocks =
            override_blocks.or_else(|| alternate.as_ref().map(|story| story.blocks.as_slice()));
        let mut projection = self.project_root(&mut cache, &txn, id, input, override_blocks)?;
        projection.story_id = story_id.to_owned();
        Ok(projection)
    }

    fn project_root(
        &self,
        cache: &mut ProjectionCache,
        txn: &Transaction<'_>,
        id: &str,
        input: &ProjectionInput,
        override_blocks: Option<&[BlockContent]>,
    ) -> Result<StoryProjection, ProjectionError> {
        if let Some(blocks) = override_blocks {
            return Ok(stories_aux_dispatch::project_root(id, blocks));
        }
        let envelope = Arc::clone(&cache.source(self)?.envelope);
        let package = &envelope.document.package;
        let has_sections = input
            .host
            .sections
            .as_ref()
            .map_or(package.document.sections.is_some(), Option::is_some)
            || input.host.final_section_properties.as_ref().map_or(
                package.document.final_section_properties.is_some(),
                Option::is_some,
            );
        if id == "body" {
            let source = self.source.lock().unwrap();
            let source_index = match source.as_ref() {
                Some(SourcePackage::Ready(index)) => Some(index),
                _ => None,
            };
            let same = match (&cache.body_source_index, source_index) {
                (Some(previous), Some(current)) => Arc::ptr_eq(previous, current),
                (None, None) => true,
                _ => false,
            };
            if !same {
                cache.stories.remove(id);
                cache.body_source_index = source_index.cloned();
            }
        }
        let revision = self
            .story_revisions
            .lock()
            .unwrap()
            .projection_roots
            .get(id)
            .copied()
            .unwrap_or_default();
        let markers = self.projection_markers.lock().unwrap();
        let markers_revision = markers.revision(id);
        let same_styles = id == "body"
            && cache
                .styles
                .as_ref()
                .is_some_and(|(styles, _)| *styles == input.host.styles);
        let same_theme = id == "body" && cache.theme.as_ref() == Some(&input.host.theme);
        if (id != "body" || (same_styles && same_theme && cache.sections == Some(has_sections)))
            && let Some(cached) = cache.stories.get(id)
            && cached.revision == revision
            && cached.markers_revision == markers_revision
        {
            return Ok(cached.projection.clone());
        }

        let source = cache.source(self)?.story(self, id)?;
        let projection = if id == "body" {
            cache.stories.remove(id);
            let source_metadata = self.source_metadata();
            if !same_styles {
                let inheritance = if input.host.styles.is_none()
                    && let Some(metadata) = &source_metadata
                {
                    metadata.projection_inheritance()
                } else {
                    let definitions = input
                        .host
                        .styles
                        .as_ref()
                        .unwrap_or(&package.styles)
                        .as_ref()
                        .map(json_value)
                        .transpose()?;
                    RunInheritance::new(definitions.as_ref())
                };
                cache.styles = Some((input.host.styles.clone(), inheritance));
            }
            cache.theme = Some(input.host.theme.clone());
            cache.sections = Some(has_sections);
            let default_theme = docx_parse::theme::Theme::default();
            let theme = input
                .host
                .theme
                .as_ref()
                .map_or(Some(&package.theme), Option::as_ref)
                .unwrap_or(&default_theme);
            let mut context = ProjectionContext {
                doc: self,
                txn,
                input,
                source: &source,
                package_body_has_sections: has_sections,
                source_index: cache.body_source_index.as_deref(),
                source_provenance: source_metadata
                    .as_ref()
                    .map(|metadata| &metadata.read().provenance),
                inheritance: &mut cache.styles.as_mut().expect("style resolver initialized").1,
                theme,
                markers: &markers,
                visiting: BTreeSet::from([id.to_owned()]),
            };
            story::project(&mut context, id, StoryState::default())?
        } else {
            stories_aux_dispatch::project_root(id, &source.blocks)
        };
        cache.stories.insert(
            id.to_owned(),
            CachedStory {
                revision,
                markers_revision,
                projection: projection.clone(),
            },
        );
        Ok(projection)
    }

    #[cfg(debug_assertions)]
    #[doc(hidden)]
    pub fn projection_has_visited_source_story(&self, story_id: &str) -> bool {
        self.projection
            .lock()
            .unwrap()
            .source
            .as_ref()
            .is_some_and(|source| {
                source
                    .as_ref()
                    .is_ok_and(|source| source.stories.contains_key(story_id))
            })
    }

    #[doc(hidden)]
    pub fn project_document_json(&self, input_json: &str) -> Result<String, ProjectionError> {
        #[derive(Serialize)]
        #[serde(rename_all = "camelCase")]
        struct WireOutput {
            #[serde(flatten)]
            envelope: S9WireEnvelope,
            paragraph_session_keys: Vec<String>,
            session_revision_placeholders: BTreeSet<u64>,
            created_list_sites: Vec<CreatedListSite>,
            deferred: Vec<ProjectionDeferral>,
        }

        let input = projection_input_json(input_json)?;
        let (output, source) = self.project_package(&input)?;
        Ok(serde_json::to_string(&WireOutput {
            envelope: S9WireEnvelope {
                wire_version: source.wire_version,
                document: output.document,
                embedded_font_parts: source.embedded_font_parts.clone(),
                font_table_relationships_xml: source.font_table_relationships_xml.clone(),
                canonical_base64: source.canonical_base64.clone(),
                canonical_sha256: source.canonical_sha256.clone(),
            },
            paragraph_session_keys: output.paragraph_session_keys,
            session_revision_placeholders: output.session_revision_placeholders,
            created_list_sites: output.created_list_sites,
            deferred: output.deferred,
        })?)
    }

    #[doc(hidden)]
    pub fn project_story_json(
        &self,
        story_id: &str,
        input_json: &str,
    ) -> Result<String, ProjectionError> {
        let input = projection_input_json(input_json)?;
        let story = self.project_story(story_id, &input)?;
        Ok(serde_json::to_string(&story)?)
    }
}

fn projection_input_json(input_json: &str) -> Result<ProjectionInput, ProjectionError> {
    use crate::json_tree::OrderedValue;

    const MAX_MODELED_JSON_DEPTH: usize = 2 * docx_parse::MAX_NESTING_DEPTH;
    let mut input = crate::json_tree::parse(input_json).map_err(ProjectionError::InvalidInput)?;
    let mut pending = vec![(&mut input, 0)];
    while let Some((node, depth)) = pending.pop() {
        if depth >= MAX_MODELED_JSON_DEPTH
            && matches!(node, OrderedValue::Array(_) | OrderedValue::Object(_))
        {
            return Err(ProjectionError::InvalidInput(format!(
                "modeled content nesting exceeds {MAX_MODELED_JSON_DEPTH} JSON containers"
            )));
        }
        match node {
            OrderedValue::Array(children) => {
                pending.extend(children.iter_mut().map(|child| (child, depth + 1)));
            }
            OrderedValue::Object(entries) => {
                let raw_xml = entries.iter().any(|(key, value)| {
                    key == "type" && matches!(value, OrderedValue::String(kind) if kind == "rawXml")
                });
                if raw_xml {
                    entries
                        .retain(|(key, _)| matches!(key.as_str(), "type" | "xml" | "sourceOrdinal"));
                }
                pending.extend(entries.iter_mut().map(|(_, child)| (child, depth + 1)));
            }
            _ => {}
        }
    }
    serde_json::from_value(input.into_value())
        .map_err(|error| ProjectionError::InvalidInput(error.to_string()))
}

fn projection_generation(txn: &Transaction<'_>) -> Option<String> {
    txn.get_map(crate::identity::SESSION)
        .and_then(|session| crate::map_string(&session, txn, crate::identity::GENERATION))
}

fn assemble_stories(
    document: &mut S9DocumentWire,
    stories: &BTreeMap<String, StoryProjection>,
    aliases: &BTreeMap<String, String>,
    override_sections: bool,
) -> Result<(), ProjectionError> {
    let package = &mut document.package;
    if let Some(body) = stories.get("body") {
        package.document.content.clone_from(&body.blocks);
    }
    if override_sections {
        let len = package.document.content.len();
        if package.document.sections.iter().flatten().any(|section| {
            section.content_start > section.content_end || section.content_end > len
        }) {
            return Err(ProjectionError::InvalidInput(
                "host.sections has an invalid content range".to_owned(),
            ));
        }
    } else if stories.contains_key("body") {
        story::clamp_sections(&mut package.document);
    }
    for (id, part) in package
        .header_entries
        .iter_mut()
        .chain(&mut package.footer_entries)
        .flatten()
    {
        let id = format!("hf:{id}");
        if let Some(story) = stories.get(aliases.get(&id).unwrap_or(&id)) {
            part.content.clone_from(&story.blocks);
        }
    }
    for (prefix, notes) in [
        ("fn", &mut package.footnotes),
        ("en", &mut package.endnotes),
    ] {
        for note in notes.iter_mut().flatten() {
            if let Some(story) = stories.get(&format!("{prefix}:{}", note.id)) {
                note.content.clone_from(&story.blocks);
            }
        }
    }
    Ok(())
}

fn auxiliary_roots(
    package: &S9PackageWire,
) -> impl Iterator<Item = (String, SourceRoot, &[BlockContent])> {
    let furniture = [
        (&package.header_entries, SourceRoot::Header as fn(usize) -> SourceRoot),
        (&package.footer_entries, SourceRoot::Footer as fn(usize) -> SourceRoot),
    ]
    .into_iter()
    .flat_map(|(entries, kind)| {
        entries.iter().flatten().enumerate().map(move |(index, (id, part))| {
            (format!("hf:{id}"), kind(index), part.content.as_slice())
        })
    });
    let notes = [
        ("fn", &package.footnotes, SourceRoot::Footnote as fn(usize) -> SourceRoot),
        ("en", &package.endnotes, SourceRoot::Endnote as fn(usize) -> SourceRoot),
    ]
    .into_iter()
    .flat_map(|(prefix, entries, kind)| {
        entries.iter().flatten().enumerate().map(move |(index, note)| {
            (format!("{prefix}:{}", note.id), kind(index), note.content.as_slice())
        })
    });
    furniture.chain(notes)
}

impl ProjectionHostMetadata {
    fn replaces_auxiliary(&self, root: Option<SourceRoot>, id: &str) -> bool {
        match root {
            Some(SourceRoot::Header(_)) => self.headers.is_some(),
            Some(SourceRoot::Footer(_)) => self.footers.is_some(),
            Some(SourceRoot::Footnote(_)) => self.footnotes.is_some(),
            Some(SourceRoot::Endnote(_)) => self.endnotes.is_some(),
            None if id.starts_with("fn:") => self.footnotes.is_some(),
            None if id.starts_with("en:") => self.endnotes.is_some(),
            None => false,
        }
    }

    fn auxiliary_blocks(&self, id: &str, root: Option<SourceRoot>) -> Option<&[BlockContent]> {
        let (prefix, id) = id.split_once(':')?;
        match prefix {
            "hf" => {
                let header = self
                    .headers
                    .iter()
                    .flatten()
                    .flatten()
                    .find(|(key, _)| key == id);
                if header.is_some()
                    || (self.headers.is_none() && matches!(root, Some(SourceRoot::Header(_))))
                {
                    return header.map(|(_, part)| part.content.as_slice());
                }
                self.footers
                    .iter()
                    .flatten()
                    .flatten()
                    .find(|(key, _)| key == id)
                    .map(|(_, part)| part.content.as_slice())
            }
            "fn" | "en" => {
                let entries = if prefix == "fn" {
                    &self.footnotes
                } else {
                    &self.endnotes
                };
                let id = id.parse::<f64>().ok()?;
                entries
                    .iter()
                    .flatten()
                    .flatten()
                    .find(|note| note.id == id)
                    .map(|note| note.content.as_slice())
            }
            _ => None,
        }
    }
}

fn apply_host_metadata(document: &mut S9DocumentWire, host: &ProjectionHostMetadata) {
    let package = &mut document.package;
    if let Some(styles) = &host.styles {
        package.styles.clone_from(styles);
    }
    if let Some(theme) = &host.theme {
        package.theme = theme.clone().unwrap_or_default();
    }
    if let Some(settings) = &host.settings {
        package.settings = settings.clone().unwrap_or_default();
    }
    if let Some(font_table) = &host.font_table {
        package.font_table = font_table.clone().unwrap_or_default();
    }
    if let Some(relationships) = &host.relationships {
        package.relationship_entries = relationships.clone().unwrap_or_default();
    }
    if let Some(warnings) = &host.warnings {
        document.warnings.clone_from(warnings);
    }
    if let Some(comments) = &host.comments {
        package.document.comments.clone_from(comments);
    }
    if let Some(headers) = &host.headers {
        package.header_entries.clone_from(headers);
    }
    if let Some(footers) = &host.footers {
        package.footer_entries.clone_from(footers);
    }
    if let Some(footnotes) = &host.footnotes {
        package.footnotes.clone_from(footnotes);
    }
    if let Some(endnotes) = &host.endnotes {
        package.endnotes.clone_from(endnotes);
    }
    if let Some(separators) = &host.footnote_separators {
        package.footnote_separators.clone_from(separators);
    }
    if let Some(separators) = &host.endnote_separators {
        package.endnote_separators.clone_from(separators);
    }
    if let Some(sections) = &host.sections {
        package.document.sections.clone_from(sections);
    }
    if let Some(properties) = &host.final_section_properties {
        package.document.final_section_properties.clone_from(properties);
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use serde_json::json;
    use yrs::{Assoc, Map, MapPrelim};

    fn seeded_header() -> EditingDoc {
        let parts = [
            ("[Content_Types].xml", r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/header.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/></Types>"#),
            ("_rels/.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>"#),
            ("word/document.xml", r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r><w:t>Body</w:t></w:r></w:p><w:sectPr><w:headerReference w:type="default" r:id="header"/></w:sectPr></w:body></w:document>"#),
            ("word/_rels/document.xml.rels", r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="header" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header.xml"/></Relationships>"#),
            ("word/header.xml", r#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>"#),
        ]
        .map(|(path, xml)| (path.to_owned(), xml.as_bytes().to_vec()));
        let doc = EditingDoc::new(41);
        crate::seed_from_docx(&doc, &ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
        doc
    }

    #[test]
    fn cold_partial_projection_includes_untouched_header_deferrals() {
        let doc = seeded_header();
        let output = doc
            .project_document(&ProjectionInput {
                root_stories: Some(BTreeSet::from(["body".to_owned()])),
                ..ProjectionInput::default()
            })
            .unwrap();
        assert!(output.document.package.header_entries.is_some());
        assert!(output.deferred.iter().any(|entry| {
            entry.story_id == "hf:header" && entry.kind == "headerFooterStory"
        }));
    }

    #[test]
    fn partial_projection_evicts_deleted_live_roots() {
        let doc = seeded_header();
        doc.create_story("hf:extra", "Extra", "Normal", "left")
            .unwrap();
        let input = ProjectionInput {
            root_stories: Some(BTreeSet::from(["body".to_owned()])),
            ..ProjectionInput::default()
        };
        let full = doc.project_document(&ProjectionInput::default()).unwrap();
        assert!(full.deferred.iter().any(|entry| entry.story_id == "hf:extra"));
        doc.delete_story("hf:extra").unwrap();
        let partial = doc.project_document(&input).unwrap();
        assert!(!partial.deferred.iter().any(|entry| entry.story_id == "hf:extra"));
    }

    #[test]
    fn host_can_remove_headers_from_the_projected_package() {
        let doc = seeded_header();
        let wire = doc
            .project_document_json(r#"{"host":{"headers":null}}"#)
            .unwrap();
        let wire: Value = serde_json::from_str(&wire).unwrap();
        assert!(wire["document"]["package"].get("headerEntries").is_none());
        assert!(!wire["deferred"].as_array().unwrap().iter().any(|entry| {
            entry["storyId"] == "hf:header"
        }));
    }

    #[test]
    fn rejects_unsupported_host_media_urls() {
        let doc = seeded_header();
        let error = doc
            .project_document_json(r#"{"host":{"mediaUrls":{"media:0":"data:image/png;base64,AA=="}}}"#)
            .unwrap_err();
        assert!(matches!(
            error,
            ProjectionError::InvalidInput(message) if message.contains("mediaUrls")
        ));
    }

    #[test]
    fn rejects_deep_modeled_host_blocks_before_decoding() {
        let mut block = r#"{"type":"rawXml","xml":"<leaf/>"}"#.to_owned();
        for _ in 0..=docx_parse::MAX_NESTING_DEPTH {
            block = format!(
                r#"{{"type":"blockSdt","properties":{{"sdtType":"richText"}},"content":[{block}]}}"#
            );
        }
        let input = format!(
            r#"{{"host":{{"comments":[{{"id":1,"author":"Ada","content":[],"blockContent":[{block}]}}]}}}}"#
        );
        let doc = seeded_header();
        for error in [
            doc.project_document_json(&input).unwrap_err(),
            doc.project_story_json("body", &input).unwrap_err(),
        ] {
            assert!(matches!(
                error,
                ProjectionError::InvalidInput(message) if message.contains("modeled content nesting")
            ));
        }
    }

    #[test]
    fn resolves_retained_bytes_only_when_projecting() {
        let doc = EditingDoc::new(41);
        doc.retain_source_docx(Arc::<[u8]>::from(&b"invalid zip"[..]));
        assert!(doc.projection.lock().unwrap().source.is_none());
        assert!(matches!(
            doc.project_story("body", &ProjectionInput::default()),
            Err(ProjectionError::SourceParse(_))
        ));
    }

    #[test]
    fn story_rejects_non_finite_revision_ids() {
        let doc = seeded_header();
        let input = ProjectionInput {
            revision_ids: BTreeMap::from([(
                "revision".to_owned(),
                ProjectionRevisionIds {
                    insertion: Some(f64::NAN),
                    deletion: None,
                },
            )]),
            ..ProjectionInput::default()
        };
        assert!(matches!(
            doc.project_story("body", &input),
            Err(ProjectionError::InvalidInput(_))
        ));
    }

    #[test]
    fn projects_live_language_without_original_paragraph_formatting() {
        let doc = seeded_header();
        let para_id = doc.paragraphs("body").unwrap()[0].para_id.clone();
        doc.set_paragraph_attr(&para_id, "_originalFormatting", Any::Null)
            .unwrap();
        doc.set_paragraph_attr(
            &para_id,
            "language",
            Any::Map(Arc::new(std::collections::HashMap::from([(
                "latin".to_owned(),
                Any::from("fr-FR"),
            )]))),
        )
        .unwrap();
        let wire: Value =
            serde_json::from_str(&doc.project_story_json("body", "{}").unwrap()).unwrap();
        assert_eq!(
            wire["blocks"][0]["formatting"]["runProperties"]["language"]["latin"],
            "fr-FR"
        );
    }

    #[test]
    fn projects_valid_comment_anchors_beside_malformed_ones() {
        let doc = seeded_header();
        let input = ProjectionInput::default();
        assert!(!doc.project_story("body", &input).unwrap().deferred.iter().any(|entry| {
            entry.kind == "commentRanges"
        }));
        {
            let mut txn = doc.doc.transact_mut();
            let body = crate::story_text(&txn, "body").unwrap();
            let start = crate::sticky_index(&body, &txn, 0, Assoc::After).unwrap();
            let end = crate::sticky_index(&body, &txn, 1, Assoc::Before).unwrap();
            let anchor = crate::anchor_value("body", &start, &end);
            let comments = txn.get_map(crate::COMMENTS).unwrap();
            let comment = comments.insert(&mut txn, "live", MapPrelim::default());
            comment.insert(&mut txn, "anchors", Any::Array(Arc::from([anchor, Any::Null])));
        }
        assert!(doc.project_story("body", &input).unwrap().deferred.iter().any(|entry| {
            entry.kind == "commentRanges"
        }));
    }

    #[test]
    fn document_projection_reuses_the_parsed_alias_declaration() {
        let doc = seeded_header();
        doc.set_header_footer_aliases(
            &json!([{
                "isHeader": true,
                "partPath": "word/header.xml",
                "relationshipIds": ["header", "alias"],
            }])
            .to_string(),
        )
        .unwrap();
        let input = ProjectionInput::default();
        doc.project_document(&input).unwrap();
        let parses = crate::header_footer::ALIAS_PARSES.get();
        doc.project_document(&input).unwrap();
        assert_eq!(crate::header_footer::ALIAS_PARSES.get(), parses);
    }
}
