use std::collections::{BTreeMap, BTreeSet};
use std::sync::Arc;

use docx_parse::formatting::TextFormatting;
use docx_parse::paragraph::{Paragraph, ParagraphContent};
use docx_parse::s9::{S9SectionWire, S9WireEnvelope};
use docx_parse::{BlockContent, S9DocumentWire};
use serde::{Deserialize, Deserializer, Serialize};
use serde_json::Value;
use yrs::types::ToJson;
use yrs::{Any, ReadTxn, Transact, Transaction};

use crate::identity::{SourceIndex, SourcePackage};
use crate::seed::{RunInheritance, SourceMetadata};
use crate::{EditError, EditingDoc, ParagraphProperties, SegmentContent, StorySegment};

mod fields_dispatch;
mod markers_dispatch;
mod paragraph;
mod runs;
mod stories_aux_dispatch;
mod story;
mod tables_dispatch;

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

#[derive(Clone, Default)]
struct StoryProjection {
    blocks: Vec<BlockContent>,
    paragraph_keys: Vec<String>,
    deferred: Vec<ProjectionDeferral>,
    revision_placeholders: BTreeSet<u64>,
    created_list_sites: Vec<CreatedListSite>,
    dependencies: BTreeSet<String>,
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
        let offset = self.paragraph_keys.len();
        self.created_list_sites.extend(child.created_list_sites.drain(..).map(|mut site| {
            site.paragraph += offset;
            site
        }));
        self.blocks.append(&mut child.blocks);
        self.paragraph_keys.append(&mut child.paragraph_keys);
        self.deferred.append(&mut child.deferred);
        self.revision_placeholders.append(&mut child.revision_placeholders);
        self.dependencies.append(&mut child.dependencies);
    }
}

struct SourceCache {
    envelope: Arc<S9WireEnvelope>,
    paragraphs: BTreeMap<String, Arc<Paragraph>>,
    stories: BTreeMap<String, StoryProjection>,
    containers: BTreeMap<String, StoryProjection>,
}

impl SourceCache {
    fn new(envelope: S9WireEnvelope) -> Self {
        let mut source = Self {
            envelope: Arc::new(envelope),
            paragraphs: BTreeMap::new(),
            stories: BTreeMap::new(),
            containers: BTreeMap::new(),
        };
        let package = source.envelope.document.package.clone();
        source.index_story("body", &package.document.content);
        for (id, part) in package.header_entries.iter().chain(&package.footer_entries).flatten() {
            source.index_story(&format!("hf:{id}"), &part.content);
        }
        for (prefix, notes) in [("fn", package.footnotes), ("en", package.endnotes)] {
            for note in notes.into_iter().flatten() {
                source.index_story(&format!("{prefix}:{}", note.id), &note.content);
            }
        }
        source
    }

    fn index_story(&mut self, story_id: &str, blocks: &[BlockContent]) {
        let mut result = StoryProjection::default();
        let (mut paragraph, mut table, mut sdt) = (0, 0, 0);
        for block in blocks {
            match block {
                BlockContent::Paragraph(source) => {
                    let key = source.para_id.as_ref()
                        .filter(|id| !id.is_empty() && source.repeated_para_id != Some(true))
                        .cloned()
                        .unwrap_or_else(|| format!("{story_id}:p{paragraph}"));
                    self.paragraphs.entry(key.clone()).or_insert_with(|| Arc::clone(source));
                    paragraph += 1;
                }
                BlockContent::Table(source) => {
                    let id = format!("{story_id}:t{table}");
                    let cell_indices = serde_json::to_value(source)
                        .map(|value| crate::seed::source_cells(&value))
                        .unwrap_or_default();
                    for (row, content) in source.rows.iter().enumerate() {
                        let indices = cell_indices.get(row).map(Vec::as_slice).unwrap_or_default();
                        for (cell, content) in content.cells.iter().enumerate() {
                            if let Ok(ordinal) = indices.binary_search(&cell) {
                                self.index_story(&format!("{id}:r{row}c{ordinal}"), &content.content);
                            }
                        }
                    }
                    self.containers.insert(id, StoryProjection {
                        blocks: vec![block.clone()],
                        ..StoryProjection::default()
                    });
                    table += 1;
                }
                BlockContent::BlockSdt(source) => {
                    let id = format!("{story_id}:sdt{sdt}");
                    self.index_story(&id, &source.content);
                    self.containers.insert(id, StoryProjection {
                        blocks: vec![block.clone()],
                        ..StoryProjection::default()
                    });
                    sdt += 1;
                }
                BlockContent::RawXml(_) => {}
            }
            result.blocks.push(block.clone());
        }
        self.stories.insert(story_id.to_owned(), result);
    }
}

#[derive(Default)]
pub(crate) struct ProjectionCache {
    source: Option<SourceCache>,
    revision: u64,
    dependencies: Option<OverlayDependencies>,
    stories: BTreeMap<String, CachedStory>,
    assembled: BTreeMap<String, Arc<StoryProjection>>,
}

#[derive(Clone)]
struct CachedStory {
    paragraph_formatting: Option<Value>,
    run_formatting: Option<Value>,
    projection: Arc<StoryProjection>,
}

impl ProjectionCache {
    pub(crate) fn invalidate_stories(&mut self) {
        self.stories.clear();
    }
}

#[derive(PartialEq)]
struct OverlayDependencies {
    input: ProjectionInput,
    side_maps: [Option<Any>; 5],
    aliases: Vec<(String, String)>,
    generation: Option<String>,
}

struct ProjectionContext<'a, 'doc> {
    doc: &'a EditingDoc,
    txn: &'a Transaction<'doc>,
    input: &'a ProjectionInput,
    source: &'a SourceCache,
    #[allow(dead_code)]
    source_metadata: Option<&'a SourceMetadata>,
    source_index: Option<&'a SourceIndex>,
    inheritance: RunInheritance,
    theme: docx_parse::theme::Theme,
    stories: &'a mut BTreeMap<String, CachedStory>,
    visiting: BTreeSet<String>,
}

impl ProjectionContext<'_, '_> {
    fn project_content_story(
        &mut self,
        id: &str,
        state: StoryState,
    ) -> Result<Arc<StoryProjection>, ProjectionError> {
        if self.visiting.contains(id) {
            return Err(ProjectionError::InvalidInput(format!("cyclic story {id}")));
        }
        if let Some(cached) = self.stories.get(id)
            && cached.paragraph_formatting == state.table_paragraph_formatting
            && cached.run_formatting == state.table_run_formatting
        {
            return Ok(Arc::clone(&cached.projection));
        }
        self.visiting.insert(id.to_owned());
        let paragraph_formatting = state.table_paragraph_formatting.clone();
        let run_formatting = state.table_run_formatting.clone();
        let projection = if root_id(id) && id != "body" {
            stories_aux_dispatch::Hooks.project_root(self, id)
        } else {
            story::project(self, id, state)
        };
        self.visiting.remove(id);
        let projection = Arc::new(projection?);
        self.stories.insert(id.to_owned(), CachedStory {
            paragraph_formatting,
            run_formatting,
            projection: Arc::clone(&projection),
        });
        Ok(projection)
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

    fn project_root(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        _story_id: &str,
    ) -> Result<StoryProjection, ProjectionError> {
        Ok(StoryProjection::default())
    }

    fn finish_document(
        &self,
        _context: &mut ProjectionContext<'_, '_>,
        _selected: &BTreeSet<String>,
        _stories: &BTreeMap<String, Arc<StoryProjection>>,
        _document: &mut S9DocumentWire,
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

fn root_id(id: &str) -> bool {
    id == "body" || id.split_once(':').is_some_and(|(kind, item)| {
        matches!(kind, "hf" | "fn" | "en") && !item.is_empty() && !item.contains(':')
    })
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

fn source_base(doc: &EditingDoc) -> Result<SourceCache, ProjectionError> {
    let (bytes, digest) = {
        let source = doc.source.lock().unwrap();
        match source.as_ref().ok_or(ProjectionError::MissingSource)? {
            SourcePackage::Pending(bytes, digest) => (bytes.clone(), digest.clone()),
            SourcePackage::Ready(index) => (index.bytes(), Some(index.package_digest().to_owned())),
        }
    };
    let digest = digest.unwrap_or_else(|| crate::seed::package_digest(&bytes));
    let (envelope, _) = crate::seed::parse_docx_package_with_digest(&bytes, digest)
        .map_err(ProjectionError::SourceParse)?;
    Ok(SourceCache::new(envelope))
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
        self.project(input, None).map(|(output, _)| output)
    }

    #[doc(hidden)]
    pub fn project_story(&self, story_id: &str, input: &ProjectionInput) -> Result<ProjectionOutput, ProjectionError> {
        self.project(input, Some(story_id)).map(|(output, _)| output)
    }

    #[cfg(feature = "wasm")]
    pub(crate) fn projection_json(&self, story_id: Option<&str>, input_json: &str) -> Result<String, String> {
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

        let input = serde_json::from_str(input_json)
            .map_err(|error| ProjectionError::InvalidInput(error.to_string()).to_string())?;
        let (output, source) = self.project(&input, story_id).map_err(|error| error.to_string())?;
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

    fn project(
        &self,
        input: &ProjectionInput,
        story_id: Option<&str>,
    ) -> Result<(ProjectionOutput, Arc<S9WireEnvelope>), ProjectionError> {
        validate_input(input)?;
        let mut cache = self.projection.lock().unwrap();
        if cache.source.is_none() {
            cache.source = Some(source_base(self)?);
        }
        let source_index = self.source_index();
        let metadata = self.source_metadata();
        let txn = self.doc.transact();
        let requested = match story_id {
            Some(id) => BTreeSet::from([id.to_owned()]),
            None => input.root_stories.clone().unwrap_or_else(|| {
                txn.get_map(crate::STORIES).map_or_else(BTreeSet::new, |stories| {
                    crate::live_stories(&stories, &txn).map(|(id, _)| id)
                        .filter(|id| root_id(id)).map(str::to_owned).collect()
                })
            }),
        };
        let aliases = self.header_footer_aliases_in(&txn);
        let mut selected = BTreeSet::new();
        for id in requested {
            if !root_id(&id) {
                return Err(ProjectionError::InvalidInput(format!("{id} is not a root story")));
            }
            let canonical = id.strip_prefix("hf:").and_then(|relationship| {
                aliases.iter().find(|(alias, _)| alias == relationship)
            }).map_or_else(|| id.clone(), |(_, canonical)| format!("hf:{canonical}"));
            if let Err(error) = crate::story_ref(&txn, &canonical) {
                return Err(ProjectionError::StoryRead {
                    story_id: id.clone(),
                    error,
                });
            }
            selected.insert(canonical);
        }

        let dependencies = OverlayDependencies {
            input: ProjectionInput { root_stories: None, ..input.clone() },
            side_maps: [
                crate::COMMENTS,
                crate::RAW_INLINES,
                crate::section_margins::SECTION_MARGINS,
                crate::part_metadata::PART_METADATA,
                crate::document_properties::DOCUMENT_PROPERTIES,
            ].map(|name| txn.get_map(name).map(|map| map.to_json(&txn))),
            aliases,
            generation: txn.get_map(crate::identity::SESSION).and_then(|session| {
                crate::map_string(&session, &txn, crate::identity::GENERATION)
            }),
        };
        let (revision, changed) = self.stories_changed_since(cache.revision);
        let same_dependencies = cache.dependencies.as_ref() == Some(&dependencies);
        let mut stories = if same_dependencies { cache.stories.clone() } else { BTreeMap::new() };
        stories.retain(|id, cached| !changed.iter().any(|changed| {
            changed == id || cached.projection.dependencies.contains(changed)
                || changed.strip_prefix(id).is_some_and(|suffix| suffix.starts_with(':'))
        }));

        let cached_source = cache.source.as_ref().expect("source cached");
        let mut document = cached_source.envelope.document.clone();
        let same_opening = cache.dependencies.as_ref().is_some_and(|prior| {
            prior.generation == dependencies.generation
        });
        let mut assembled = if story_id.is_none() && same_opening {
            cache.assembled.clone()
        } else {
            BTreeMap::new()
        };
        if story_id.is_none() && input.root_stories.is_none() {
            assembled.retain(|story, _| selected.contains(story));
        }
        let definitions = input.host.styles.as_ref().unwrap_or(&document.package.styles)
            .as_ref().map(json_value);
        let inheritance = metadata.as_deref().filter(|_| input.host.styles.is_none())
            .map_or_else(|| RunInheritance::new(definitions.as_ref()), SourceMetadata::projection_inheritance);
        let theme = input.host.theme.clone().unwrap_or_else(|| Some(document.package.theme.clone()))
            .unwrap_or_default();
        if story_id.is_none() {
            apply_host_metadata(&mut document, &input.host);
        }
        let mut context = ProjectionContext {
            doc: self,
            txn: &txn,
            input,
            source: cached_source,
            source_metadata: metadata.as_deref(),
            source_index: source_index.as_deref(),
            inheritance,
            theme,
            stories: &mut stories,
            visiting: BTreeSet::new(),
        };
        for id in &selected {
            let projection = context.project_content_story(id, StoryState::default())?;
            assembled.insert(id.clone(), projection);
        }
        if let Some(body) = assembled.get("body") {
            document.package.document.content.clone_from(&body.blocks);
            story::clamp_sections(&mut document.package.document);
        }
        for hook in HOOKS {
            hook.finish_document(&mut context, &selected, &assembled, &mut document)?;
        }

        let mut output = ProjectionOutput {
            document,
            paragraph_session_keys: assembled.get("body").map_or_else(Vec::new, |body| body.paragraph_keys.clone()),
            session_revision_placeholders: BTreeSet::new(),
            created_list_sites: Vec::new(),
            deferred: Vec::new(),
        };
        for (id, projection) in &assembled {
            if id != "body" {
                let offset = output.paragraph_session_keys.len();
                output.paragraph_session_keys.extend(projection.paragraph_keys.iter().cloned());
                output.created_list_sites.extend(projection.created_list_sites.iter().cloned().map(|mut site| {
                    site.paragraph += offset;
                    site
                }));
            } else {
                output.created_list_sites.extend_from_slice(&projection.created_list_sites);
            }
            output.session_revision_placeholders.extend(&projection.revision_placeholders);
            output.deferred.extend_from_slice(&projection.deferred);
        }

        let envelope = Arc::clone(&cached_source.envelope);
        cache.revision = revision;
        cache.stories = stories;
        if story_id.is_none() {
            cache.assembled = assembled;
        } else if !same_opening {
            cache.assembled.clear();
        }
        cache.dependencies = Some(dependencies);
        Ok((output, envelope))
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
