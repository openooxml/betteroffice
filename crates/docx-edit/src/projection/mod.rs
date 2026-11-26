use std::collections::{BTreeMap, BTreeSet};
use std::sync::{Arc, Mutex};

use docx_parse::paragraph::{Paragraph, ParagraphContent};
use docx_parse::s9::{S9SectionWire, S9WireEnvelope};
use docx_parse::{BlockContent, S9DocumentWire};
use serde::{Deserialize, Deserializer, Serialize};
use serde_json::Value;
use yrs::{Any, ReadTxn, Transact, Transaction};

use crate::identity::{SourceIndex, SourcePackage};
use crate::seed::RunInheritance;
use crate::{EditError, EditingDoc, ParagraphProperties, SegmentContent, StorySegment};

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
}

impl std::fmt::Display for ProjectionError {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        match self {
            Self::MissingSource => f.write_str("the DOCX session has no retained source package"),
            Self::SourceParse(error) => write!(f, "could not parse the retained DOCX source: {error}"),
            Self::StoryRead { story_id, error } => write!(f, "could not read projection story {story_id}: {error}"),
            Self::InvalidInput(error) => write!(f, "invalid projection input: {error}"),
        }
    }
}

impl std::error::Error for ProjectionError {}

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
    fn defer(&mut self, family: ProjectionFamily, kind: &str, story_id: &str, paragraph: Option<usize>) {
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
        self.created_list_sites.extend(child.created_list_sites.iter().cloned().map(|mut site| {
            site.paragraph += offset;
            site
        }));
        self.paragraph_session_keys.extend_from_slice(&child.paragraph_session_keys);
        self.session_revision_placeholders.extend(&child.session_revision_placeholders);
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
    roots: BTreeMap<String, SourceRoot>,
    stories: BTreeMap<String, Arc<SourceStory>>,
    materialize_media: bool,
    materialized_document: Option<Arc<S9DocumentWire>>,
}

impl SourceCache {
    fn new(envelope: Arc<S9WireEnvelope>, materialize_media: bool) -> Self {
        let package = &envelope.document.package;
        let mut roots = BTreeMap::new();
        for (header, entries) in [(true, &package.header_entries), (false, &package.footer_entries)] {
            for (index, (id, _)) in entries.iter().flatten().enumerate() {
                let part = if header { SourceRoot::Header(index) } else { SourceRoot::Footer(index) };
                roots.entry(format!("hf:{id}")).or_insert(part);
            }
        }
        for (footnote, entries) in [(true, &package.footnotes), (false, &package.endnotes)] {
            for (index, note) in entries.iter().flatten().enumerate() {
                let (prefix, part) = if footnote {
                    ("fn", SourceRoot::Footnote(index))
                } else {
                    ("en", SourceRoot::Endnote(index))
                };
                roots.entry(format!("{prefix}:{}", note.id)).or_insert(part);
            }
        }
        Self {
            envelope,
            roots,
            stories: BTreeMap::new(),
            materialize_media,
            materialized_document: None,
        }
    }

    fn document(&mut self, doc: &EditingDoc) -> Result<S9DocumentWire, ProjectionError> {
        if let Some(document) = &self.materialized_document {
            return Ok((**document).clone());
        }
        let mut document = self.envelope.document.clone();
        if self.materialize_media && let Some(table) = doc.media_table() && !table.is_empty() {
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
        match self.roots.get(id) {
            Some(SourceRoot::Header(index)) => &package.header_entries.as_ref().unwrap()[*index].1.content,
            Some(SourceRoot::Footer(index)) => &package.footer_entries.as_ref().unwrap()[*index].1.content,
            Some(SourceRoot::Footnote(index)) => &package.footnotes.as_ref().unwrap()[*index].content,
            Some(SourceRoot::Endnote(index)) => &package.endnotes.as_ref().unwrap()[*index].content,
            None => &[],
        }
    }

    fn story(&mut self, doc: &EditingDoc, id: &str) -> Result<Arc<SourceStory>, ProjectionError> {
        if let Some(story) = self.stories.get(id) {
            return Ok(Arc::clone(story));
        }
        let blocks = self.blocks(id);
        let blocks = if self.materialize_media && let Some(table) = doc.media_table()
            && !table.is_empty()
        {
            materialize(&blocks.to_vec(), &table)?
        } else {
            blocks.to_vec()
        };
        let mut source = SourceStory { blocks, ..SourceStory::default() };
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
                    let key = source.para_id.as_ref()
                        .filter(|id| !id.is_empty() && source.repeated_para_id != Some(true))
                        .cloned()
                        .unwrap_or_else(|| format!("{story_id}:p{paragraph}"));
                    self.paragraphs.entry(key).or_insert_with(|| Arc::clone(source));
                    self.has_comment_ranges |= story::has_comment_ranges(block);
                    self.has_revision_ranges |= story::has_revision_ranges(block);
                    paragraph += 1;
                }
                BlockContent::Table(_) => {
                    self.containers.insert(format!("{story_id}:t{table}"), block.clone());
                    table += 1;
                }
                BlockContent::BlockSdt(_) => {
                    self.containers.insert(format!("{story_id}:sdt{sdt}"), block.clone());
                    sdt += 1;
                }
                BlockContent::RawXml(_) => {}
            }
        }
    }

    fn child(&self, id: &str) -> Option<Arc<Self>> {
        let mut children = self.children.lock().unwrap();
        if let Some(child) = children.get(id) {
            return Some(Arc::clone(child));
        }
        let blocks = if let Some(BlockContent::BlockSdt(source)) = self.containers.get(id) {
            &source.content
        } else {
            let table_id = crate::structured::source::cell_table(id)?;
            let BlockContent::Table(table) = self.containers.get(table_id)? else { return None };
            let (row, cell) = id.get(table_id.len() + 2..)?.split_once('c')?;
            let row: usize = row.parse().ok()?;
            let cell: usize = cell.parse().ok()?;
            let mut table_cells = self.table_cells.lock().unwrap();
            let cells = table_cells.entry(table_id.to_owned())
                .or_insert_with(|| crate::seed::source_cells(&json_value(table)));
            let source_cell = *cells.get(row)?.get(cell)?;
            &table.rows.get(row)?.cells.get(source_cell)?.content
        };
        let mut child = Self { blocks: blocks.clone(), ..Self::default() };
        child.index(id);
        let child = Arc::new(child);
        children.insert(id.to_owned(), Arc::clone(&child));
        Some(child)
    }

}

#[derive(Default)]
pub(crate) struct ProjectionCache {
    source: Option<Result<SourceCache, String>>,
    generation: Option<String>,
    aliases: crate::header_footer::ProjectionAliases,
    styles: Option<(Override<docx_parse::styles::StyleDefinitions>, RunInheritance)>,
    theme: Option<Override<docx_parse::theme::Theme>>,
    stories: BTreeMap<String, CachedStory>,
    assembled: BTreeMap<String, StoryProjection>,
}

struct CachedStory {
    revision: u64,
    markers_revision: u64,
    projection: StoryProjection,
}

impl ProjectionCache {
    pub(crate) fn for_source(source: &SourcePackage) -> Self {
        let (bytes, digest) = match source {
            SourcePackage::Ready(index) => {
                if let Some(envelope) = &index.projection_envelope {
                    return Self {
                        source: Some(Ok(SourceCache::new(Arc::clone(envelope), true))),
                        ..Self::default()
                    };
                }
                (index.bytes(), Some(index.package_digest().to_owned()))
            }
            SourcePackage::Pending(bytes, digest) => (bytes.clone(), digest.clone()),
        };
        let digest = digest.unwrap_or_else(|| crate::seed::package_digest(&bytes));
        let source = crate::seed::parse_docx_package_with_digest(&bytes, digest)
            .map(|(envelope, _)| SourceCache::new(Arc::new(envelope), false));
        Self { source: Some(source), ..Self::default() }
    }

    fn source(&mut self) -> Result<&mut SourceCache, ProjectionError> {
        self.source.as_mut().ok_or(ProjectionError::MissingSource)?.as_mut()
            .map_err(|error| ProjectionError::SourceParse(error.clone()))
    }

    pub(crate) fn reset_opening(&mut self) {
        let source = self.source.take();
        *self = Self { source, ..Self::default() };
        if let Some(Ok(source)) = &mut self.source {
            source.stories.clear();
            source.materialized_document = None;
        }
    }

    pub(crate) fn invalidate_stories(&mut self) {
        self.stories.clear();
    }

    fn begin_generation(&mut self, generation: Option<String>) {
        if self.generation != generation {
            self.generation = generation;
            self.stories.clear();
            self.assembled.clear();
            self.styles = None;
            self.theme = None;
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
        let source = self.source.child(id);
        let empty = SourceStory::default();
        let mut child = ProjectionContext {
            doc: self.doc,
            txn: self.txn,
            input: self.input,
            source: source.as_deref().unwrap_or(&empty),
            package_body_has_sections: false,
            markers: self.markers,
            source_index: self.source_index,
            inheritance: self.inheritance,
            theme: self.theme,
            visiting: std::mem::take(&mut self.visiting),
        };
        let result = story::project(&mut child, id, state);
        self.visiting = child.visiting;
        self.visiting.remove(id);
        result.map(Arc::new)
    }

    #[allow(dead_code)]
    fn media_url(&self, token: &str) -> Option<String> {
        self.input.host.media_urls.as_ref().and_then(|urls| urls.get(token)).cloned()
            .or_else(|| self.doc.media_table()?.resolve(token))
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
        let payload = self.payload.iter().map(|(key, value)| (key.clone(), value.clone())).collect();
        crate::target::drawing_from_payload(&payload, image)
    }

    fn payload_json(&self, key: &str) -> Option<Value> {
        attr_string(self.payload, key).and_then(crate::read_types::payload_json)
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

fn attr_string<'a>(attrs: &'a Attrs, key: &str) -> Option<&'a str> {
    crate::read_types::payload_str(attrs.get(key))
}

fn json_value<T: Serialize>(value: &T) -> Value {
    serde_json::to_value(value).unwrap_or_default()
}

fn decode<T: serde::de::DeserializeOwned>(value: Option<&Value>) -> Option<T> {
    serde_json::from_value(value?.clone()).ok()
}

fn materialize<T: Clone + Serialize + serde::de::DeserializeOwned>(
    value: &T,
    table: &docx_parse::media::MediaTable,
) -> Result<T, ProjectionError> {
    fn replace(
        value: &mut Value,
        table: &docx_parse::media::MediaTable,
        urls: &mut BTreeMap<usize, String>,
    ) -> Result<bool, ProjectionError> {
        let mut changed = false;
        match value {
            Value::Array(values) => {
                for value in values {
                    changed |= replace(value, table, urls)?;
                }
            }
            Value::Object(fields) => {
                for (key, value) in fields {
                    if matches!(key.as_str(), "src" | "dataUrl")
                        && let Some(index) = value.as_str().and_then(docx_parse::media::media_token_index)
                        && index < table.len()
                    {
                        let url = match urls.entry(index) {
                            std::collections::btree_map::Entry::Occupied(entry) => entry.into_mut(),
                            std::collections::btree_map::Entry::Vacant(entry) => {
                                entry.insert(table.data_url(index).map_err(ProjectionError::SourceParse)?)
                            }
                        };
                        *value = Value::String(url.clone());
                        changed = true;
                    } else {
                        changed |= replace(value, table, urls)?;
                    }
                }
            }
            _ => {}
        }
        Ok(changed)
    }

    let mut wire = serde_json::to_value(value).map_err(|error| ProjectionError::SourceParse(error.to_string()))?;
    if replace(&mut wire, table, &mut BTreeMap::new())? {
        serde_json::from_value(wire).map_err(|error| ProjectionError::SourceParse(error.to_string()))
    } else {
        Ok(value.clone())
    }
}

fn materialize_package(document: &mut S9DocumentWire, table: &docx_parse::media::MediaTable) -> Result<(), ProjectionError> {
    let package = &mut document.package;
    package.document.content = materialize(&package.document.content, table)?;
    package.document.comments = materialize(&package.document.comments, table)?;
    for entries in [&mut package.header_entries, &mut package.footer_entries] {
        for (_, part) in entries.iter_mut().flatten() {
            part.content = materialize(&part.content, table)?;
        }
    }
    for notes in [&mut package.footnotes, &mut package.endnotes,
        &mut package.footnote_separators, &mut package.endnote_separators]
    {
        for note in notes.iter_mut().flatten() {
            note.content = materialize(&note.content, table)?;
        }
    }
    let mut urls = BTreeMap::new();
    package.media_entries = table.media_map().into_iter().map(|(path, file)| {
        let index = docx_parse::media::media_token_index(&file.data_url).expect("media table token");
        let file = match urls.entry(index) {
            std::collections::btree_map::Entry::Occupied(entry) => Arc::clone(entry.get()),
            std::collections::btree_map::Entry::Vacant(entry) => {
                let mut file = (*file).clone();
                file.data_url = table.data_url(index).map_err(ProjectionError::SourceParse)?;
                file.base64 = file.data_url.split_once(',').map_or("", |(_, base64)| base64).to_owned();
                Arc::clone(entry.insert(Arc::new(file)))
            }
        };
        Ok((path, file))
    }).collect::<Result<_, ProjectionError>>()?;
    Ok(())
}

fn validate_input(input: &ProjectionInput) -> Result<(), ProjectionError> {
    let numbers = input.comment_ids.iter().map(|(number, _)| *number)
        .chain(input.crossing_comment_fallback_ids.iter().copied())
        .chain(input.revision_ids.values().flat_map(|ids| [ids.insertion, ids.deletion]).flatten());
    if numbers.into_iter().any(|number| !number.is_finite()) {
        return Err(ProjectionError::InvalidInput("non-finite OOXML id".to_owned()));
    }
    Ok(())
}

impl EditingDoc {
    #[doc(hidden)]
    pub fn project_document(&self, input: &ProjectionInput) -> Result<ProjectionOutput, ProjectionError> {
        validate_input(input)?;
        let (mut document, requested, aliases, mut assembled) = {
            let mut cache = self.projection.lock().unwrap();
            let txn = self.doc.transact();
            cache.begin_generation(projection_generation(&txn));
            let source = cache.source()?;
            let aliases: BTreeMap<_, _> = self.header_footer_aliases_in(&txn).into_iter().collect();
            let requested: BTreeSet<_> = input.root_stories.clone().unwrap_or_else(|| {
                let mut roots: BTreeSet<_> = source.roots.keys().cloned().collect();
                if let Some(stories) = txn.get_map(crate::STORIES) {
                    roots.extend(crate::live_stories(&stories, &txn).map(|(id, _)| id)
                        .filter(|id| root_id(id)).map(str::to_owned));
                }
                roots
            }).into_iter().map(|id| {
                id.strip_prefix("hf:").and_then(|id| aliases.get(id))
                    .map_or(id, |canonical| format!("hf:{canonical}"))
            }).collect();
            let document = source.document(self)?;
            let assembled = if input.root_stories.is_some() {
                cache.assembled.clone()
            } else {
                BTreeMap::new()
            };
            (document, requested, aliases, assembled)
        };
        for story_id in requested {
            let projection = self.project_story(&story_id, input)?;
            assembled.insert(projection.story_id.clone(), projection);
        }

        apply_host_metadata(&mut document, &input.host);
        assemble_stories(&mut document, &assembled, &aliases);
        let mut side_tables = StoryProjection::default();
        for projection in assembled.values() {
            side_tables.append_side_tables(projection);
        }
        self.projection.lock().unwrap().assembled = assembled;
        Ok(ProjectionOutput {
            document,
            paragraph_session_keys: side_tables.paragraph_session_keys,
            session_revision_placeholders: side_tables.session_revision_placeholders,
            created_list_sites: side_tables.created_list_sites,
            deferred: side_tables.deferred,
        })
    }

    #[doc(hidden)]
    pub fn project_story(&self, story_id: &str, input: &ProjectionInput) -> Result<StoryProjection, ProjectionError> {
        if !root_id(story_id) {
            return Err(ProjectionError::InvalidInput(format!("{story_id} is not a root story")));
        }
        let mut cache = self.projection.lock().unwrap();
        cache.source()?;
        let txn = self.doc.transact();
        cache.begin_generation(projection_generation(&txn));
        let canonical = cache.aliases.canonical(self, &txn, story_id);
        let id = canonical.as_deref().unwrap_or(story_id);
        if crate::story_text(&txn, id).is_none() && (id == "body" || !cache.source()?.roots.contains_key(id)) {
            return Err(ProjectionError::StoryRead {
                story_id: story_id.to_owned(),
                error: EditError::StoryNotFound(id.to_owned()),
            });
        }
        let revision = self.story_revisions.lock().unwrap().projection_roots.get(id).copied().unwrap_or_default();
        let markers = self.projection_markers.lock().unwrap();
        let markers_revision = markers.revision(id);
        let same_styles = cache.styles.as_ref().is_some_and(|(styles, _)| *styles == input.host.styles);
        let same_theme = cache.theme.as_ref() == Some(&input.host.theme);
        if (id != "body" || (same_styles && same_theme))
            && let Some(cached) = cache.stories.get(id)
            && cached.revision == revision && cached.markers_revision == markers_revision
        {
            let mut projection = cached.projection.clone();
            projection.story_id = story_id.to_owned();
            return Ok(projection);
        }

        let (source, envelope) = {
            let cache = cache.source()?;
            (cache.story(self, id)?, Arc::clone(&cache.envelope))
        };
        let package = &envelope.document.package;
        let mut projection = if id == "body" {
            if !same_styles {
                let inheritance = if input.host.styles.is_none()
                    && let Some(metadata) = self.source_metadata()
                {
                    metadata.projection_inheritance()
                } else {
                    let definitions = input.host.styles.as_ref().unwrap_or(&package.styles)
                        .as_ref().map(json_value);
                    RunInheritance::new(definitions.as_ref())
                };
                cache.styles = Some((input.host.styles.clone(), inheritance));
            }
            cache.theme = Some(input.host.theme.clone());
            let default_theme = docx_parse::theme::Theme::default();
            let theme = input.host.theme.as_ref().map_or(Some(&package.theme), Option::as_ref)
                .unwrap_or(&default_theme);
            let source_index = match self.source.lock().unwrap().as_ref() {
                Some(SourcePackage::Ready(index)) => Some(Arc::clone(index)),
                _ => None,
            };
            let mut context = ProjectionContext {
                doc: self,
                txn: &txn,
                input,
                source: &source,
                package_body_has_sections: package.document.sections.is_some()
                    || package.document.final_section_properties.is_some(),
                source_index: source_index.as_deref(),
                inheritance: &mut cache.styles.as_mut().expect("style resolver initialized").1,
                theme,
                markers: &markers,
                visiting: BTreeSet::from([id.to_owned()]),
            };
            story::project(&mut context, id, StoryState::default())?
        } else {
            stories_aux_dispatch::project_root(id, &source.blocks)
        };
        cache.stories.insert(id.to_owned(), CachedStory {
            revision,
            markers_revision,
            projection: projection.clone(),
        });
        projection.story_id = story_id.to_owned();
        Ok(projection)
    }

    #[cfg(debug_assertions)]
    #[doc(hidden)]
    pub fn projection_has_visited_source_story(&self, story_id: &str) -> bool {
        self.projection.lock().unwrap().source().is_ok_and(|source| source.stories.contains_key(story_id))
    }

    #[cfg(feature = "wasm")]
    pub(crate) fn projection_document_json(&self, input_json: &str) -> Result<String, String> {
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
        let output = self.project_document(&input).map_err(|error| error.to_string())?;
        let source = Arc::clone(&self.projection.lock().unwrap().source()
            .map_err(|error| error.to_string())?.envelope);
        serde_json::to_string(&WireOutput {
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
        }).map_err(|error| error.to_string())
    }

    #[cfg(feature = "wasm")]
    pub(crate) fn projection_story_json(&self, story_id: &str, input_json: &str) -> Result<String, String> {
        let input = projection_input_json(input_json)?;
        let story = self.project_story(story_id, &input).map_err(|error| error.to_string())?;
        serde_json::to_string(&story).map_err(|error| error.to_string())
    }
}

#[cfg(feature = "wasm")]
fn projection_input_json(input_json: &str) -> Result<ProjectionInput, String> {
    serde_json::from_str(input_json)
        .map_err(|error| ProjectionError::InvalidInput(error.to_string()).to_string())
}

fn projection_generation(txn: &Transaction<'_>) -> Option<String> {
    txn.get_map(crate::identity::SESSION).and_then(|session| {
        crate::map_string(&session, txn, crate::identity::GENERATION)
    })
}

fn assemble_stories(
    document: &mut S9DocumentWire,
    stories: &BTreeMap<String, StoryProjection>,
    aliases: &BTreeMap<String, String>,
) {
    let package = &mut document.package;
    if let Some(body) = stories.get("body") {
        package.document.content.clone_from(&body.blocks);
        story::clamp_sections(&mut package.document);
    }
    for (id, part) in package.header_entries.iter_mut().chain(&mut package.footer_entries).flatten() {
        let id = aliases.get(id).unwrap_or(id);
        if let Some(story) = stories.get(&format!("hf:{id}")) {
            part.content.clone_from(&story.blocks);
        }
    }
    for (prefix, notes) in [("fn", &mut package.footnotes), ("en", &mut package.endnotes)] {
        for note in notes.iter_mut().flatten() {
            if let Some(story) = stories.get(&format!("{prefix}:{}", note.id)) {
                note.content.clone_from(&story.blocks);
            }
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
}
