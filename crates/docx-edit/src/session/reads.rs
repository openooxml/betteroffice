use serde_json::{Value, json};
use yrs::{ReadTxn, Transact};

use super::args::{attrs_value, find_para_span, json_to_any, parse_inline_format_delta};
use super::{EditSessionCore, EngineError, EngineResult};
use crate::batch::outcome_json;
use crate::{
    EditRefusal, EditRequest, FindTextRequest, OpError, ReadParagraphsRequest, ReadStoriesRequest,
    SegmentContent, TableLocator, TextTarget,
};

impl EditSessionCore {
    /// The session-scoped version token of the committed document state. It
    /// changes on every committed change, local or remote, and when the
    /// document or its retained source is replaced.
    pub fn version(&self) -> String {
        self.engine.doc().version().to_string()
    }

    /// Versioned paragraph texts:
    /// `{"story"?,"paraIds"?,"view":"accepted"|"original"}` ->
    /// `{"ok":true,"version","view","paragraphs":[{"story","paraId","text",
    /// "styleId"?,"atoms":[{"offset","kind"}]}]}`. Each inline atom occupies one
    /// U+FFFC in `text`.
    pub fn read_paragraphs_json(&self, request: &str) -> EngineResult<String> {
        let request: ReadParagraphsRequest =
            serde_json::from_str(request).map_err(EngineError::message)?;
        outcome_json(&self.engine.doc().read_paragraphs(&request)).map_err(EngineError::message)
    }

    /// Paragraph texts of selected stories in one read:
    /// `{"stories"?:"all"|[kind|id…],"byRoot"?,"view","expectVersion"?}` -> `{"ok":true,
    /// "version","view","truncated","stories":[{"story","ok":true,"paragraphs"}|{"story",
    /// "ok":false,"failure"}]}`. Listed entries keep their order; a kind expands in id order.
    pub fn read_stories_json(&self, request: &str) -> EngineResult<String> {
        let request: ReadStoriesRequest =
            serde_json::from_str(request).map_err(EngineError::message)?;
        outcome_json(&self.engine.doc().read_stories(&request)).map_err(EngineError::message)
    }

    /// Every story: `{"ok":true,"version","stories":[{"story","kind","parent"?,"root","part"?,
    /// "uses"?:[{"sectionIndex","variant","shown"?}]}]}`, sorted by id. `shown` is present only
    /// while a complete layout of the current version is retained.
    pub fn list_stories_json(&self) -> EngineResult<String> {
        let mut listed = serde_json::to_value(self.engine.list_stories_with_layout())
            .map_err(EngineError::message)?;
        listed["ok"] = Value::Bool(true);
        Ok(listed.to_string())
    }

    /// Until the matching `end_shared_reads`, committed reads share story projections of each document state.
    pub fn begin_shared_reads(&self) {
        self.engine.doc().begin_shared_reads();
    }

    /// Ends a shared-read scope, dropping shared story projections when the last scope ends.
    pub fn end_shared_reads(&self) {
        self.engine.doc().end_shared_reads();
    }

    /// Exact, case-sensitive, paragraph-local search:
    /// `{"text","within","view","limit"?}` ->
    /// `{"ok":true,"version","matches":[{"text","range"}],"truncated"}`.
    pub fn find_text_json(&self, request: &str) -> EngineResult<String> {
        let request: FindTextRequest =
            serde_json::from_str(request).map_err(EngineError::message)?;
        outcome_json(&self.engine.doc().find_text(&request)).map_err(EngineError::message)
    }

    /// Runs every check of [`EditSessionCore::apply_edits_json`], staging included,
    /// without changing anything: `{"ok":true,"baseVersion","wouldApply","previews"}`.
    pub fn validate_edits_json(&self, request: &str) -> EngineResult<String> {
        let outcome = match self.edit_request(request)? {
            Ok(request) => self
                .engine
                .doc()
                .validate_edits(&request)
                .map_err(EngineError::message)?,
            Err(refusal) => Err(refusal),
        };
        outcome_json(&outcome).map_err(EngineError::message)
    }

    /// Applies an edit batch all-or-nothing:
    /// `{"ok":true,"baseVersion","version","applied","source","changedStories",
    /// "receipts"}`. An applied batch commits one transaction; `history`
    /// `"separate"` makes it exactly one undo step.
    pub fn apply_edits_json(&self, request: &str) -> EngineResult<String> {
        let outcome = match self.edit_request(request)? {
            Ok(request) => self
                .engine
                .doc()
                .apply_edits(&request, &self.undo)
                .map_err(EngineError::message)?,
            Err(refusal) => Err(refusal),
        };
        outcome_json(&outcome).map_err(EngineError::message)
    }

    /// The headings of `story` in document order, classified as the structured export
    /// classifies them: `[{"paraId","heading":{"outlineLevel","source"}}]`.
    pub fn headings_json(&self, story: &str) -> EngineResult<String> {
        let headings: Vec<_> = self
            .engine
            .doc()
            .paragraph_headings(story)
            .map_err(EngineError::message)?
            .into_iter()
            .map(|(para_id, heading)| json!({ "paraId": para_id, "heading": heading }))
            .collect();
        serde_json::to_string(&headings).map_err(EngineError::message)
    }

    /// Resolves a text target and formats it in one call (legacy agent
    /// helpers). `delta_json` is the [`EditSessionCore::format_range`] delta.
    /// Returns `{"ok":true,"version"}` or a refusal.
    pub fn format_text_target_json(
        &self,
        target_json: &str,
        delta_json: &str,
    ) -> EngineResult<String> {
        let target: TextTarget = serde_json::from_str(target_json).map_err(EngineError::message)?;
        let delta = parse_inline_format_delta(delta_json)?;
        let doc = self.engine.doc();
        let outcome = doc
            .format_text_target(&target, &delta)
            .map_err(EngineError::message)?;
        outcome_json(&outcome.map(|()| json!({ "version": doc.version() })))
            .map_err(EngineError::message)
    }

    /// Resolves a text target and anchors a side-map comment over it in one
    /// call (legacy agent helpers). `comment_json`:
    /// `{"id","author","date","body"?}`. Returns `{"ok":true,"version"}` or a
    /// refusal.
    pub fn comment_text_target_json(
        &self,
        target_json: &str,
        comment_json: &str,
    ) -> EngineResult<String> {
        #[derive(serde::Deserialize)]
        #[serde(deny_unknown_fields)]
        struct CommentWire {
            id: String,
            author: String,
            date: String,
            #[serde(default)]
            body: Value,
        }
        let target: TextTarget = serde_json::from_str(target_json).map_err(EngineError::message)?;
        let comment: CommentWire =
            serde_json::from_str(comment_json).map_err(EngineError::message)?;
        let doc = self.engine.doc();
        let outcome = doc
            .comment_text_target(
                &target,
                &comment.id,
                &comment.author,
                &comment.date,
                json_to_any(&comment.body)?,
            )
            .map_err(EngineError::message)?;
        outcome_json(&outcome.map(|()| json!({ "version": doc.version() })))
            .map_err(EngineError::message)
    }

    pub fn geometry_position_outline_json(&self, root: &str) -> EngineResult<String> {
        serde_json::to_string(&self.engine.doc().geometry_position_outline(root))
            .map_err(EngineError::message)
    }

    pub fn resident_geometry_json(&self, root: &str, request: &str) -> EngineResult<String> {
        let request = serde_json::from_str(request).map_err(EngineError::message)?;
        let geometry = self
            .engine
            .doc()
            .resident_geometry(root, &request)
            .map_err(EngineError::message)?;
        serde_json::to_string(&geometry).map_err(EngineError::message)
    }

    pub fn resident_section_revision(&self) -> String {
        let revisions = self.engine.doc().story_revisions.lock().unwrap();
        format!("{}:{}", revisions.structure, revisions.sections)
    }

    pub fn search_text(
        &self,
        query: &str,
        case_sensitive: bool,
        limit: Option<u32>,
    ) -> EngineResult<String> {
        let matches = self
            .engine
            .doc()
            .search_text(query, case_sensitive, limit.map(|value| value as usize))
            .map_err(EngineError::message)?;
        serde_json::to_string(&matches).map_err(EngineError::message)
    }

    /// Whether the document has a story with this id.
    pub fn has_story(&self, story: &str) -> bool {
        let txn = self.engine.doc().yrs_doc().transact();
        crate::story_ref(&txn, story).is_ok()
    }

    /// `{"revision","stories":[…]}`: the current story revision and the sorted
    /// ids of the stories created, edited, or deleted after revision `since`.
    pub fn stories_changed_since(&self, since: f64) -> String {
        let (revision, stories) = self.engine.doc().stories_changed_since(since as u64);
        json!({ "revision": revision, "stories": stories }).to_string()
    }

    /// Every story id in the document, sorted so the order is stable across
    /// replicas.
    pub fn story_ids(&self) -> Vec<String> {
        let txn = self.engine.doc().yrs_doc().transact();
        let Some(stories) = txn.get_map(crate::STORIES) else {
            return Vec::new();
        };
        let mut ids: Vec<String> = crate::live_stories(&stories, &txn)
            .map(|(id, _)| id.to_string())
            .collect();
        ids.sort();
        ids
    }

    /// Story length in UTF-16 units, every embed (pilcrows included) counting
    /// as one. Errors on an unknown story.
    pub fn story_len(&self, story: &str) -> EngineResult<u32> {
        self.engine
            .doc()
            .story_len(story)
            .map_err(EngineError::message)
    }

    /// The story's `canonical-stream-v1` FNV-1a checksum (see
    /// [`crate::canonical`]) as a DECIMAL STRING, because a u64 exceeds the
    /// JavaScript safe-integer range. Two stories with the same authored
    /// content share a checksum even when their paragraph and comment ids
    /// differ. Errors on an unknown story.
    pub fn story_checksum(&self, story: &str) -> EngineResult<String> {
        crate::story_checksum(self.engine.doc(), story)
            .map(|checksum| checksum.to_string())
            .map_err(EngineError::message)
    }

    /// `[{"paraId","text","properties"}, …]` in document order. `text` is the
    /// paragraph's plain text without its pilcrow; `properties` is the
    /// pilcrow's authored property map (`pStyle`, `alignment` and whatever
    /// else has been set on it). Errors on an unknown story.
    pub fn paragraphs(&self, story: &str) -> EngineResult<String> {
        let paragraphs = self
            .engine
            .doc()
            .paragraphs(story)
            .map_err(EngineError::message)?;
        let items = paragraphs
            .into_iter()
            .map(|paragraph| {
                Ok(json!({
                    "paraId": paragraph.para_id,
                    "text": paragraph.text,
                    "properties": attrs_value(&paragraph.properties)?,
                }))
            })
            .collect::<EngineResult<Vec<Value>>>()?;
        serde_json::to_string(&items).map_err(EngineError::message)
    }

    /// Compact paragraph-position projection built in one story traversal:
    /// `[{"paraId","length"}, …]` in document order, where `length` counts the
    /// UTF-16 text and inline embed units before that paragraph's pilcrow :
    /// exactly the `offset` domain of a Loc in it. One call replaces crossing
    /// the boundary once per paragraph. Errors on an unknown story.
    pub fn paragraph_spans(&self, story: &str) -> EngineResult<String> {
        let mut items = Vec::new();
        let mut cursor = 0_u32;
        let mut paragraph_start = 0_u32;
        for segment in self
            .engine
            .doc()
            .story_segments(story)
            .map_err(EngineError::message)?
        {
            match segment.content {
                SegmentContent::Text(text) => {
                    cursor += text.encode_utf16().count() as u32;
                }
                SegmentContent::Pilcrow(properties) => {
                    items.push(json!({
                        "paraId": properties.para_id,
                        "length": cursor - paragraph_start,
                    }));
                    cursor += 1;
                    paragraph_start = cursor;
                }
                SegmentContent::OtherEmbed { .. } => cursor += 1,
            }
        }
        serde_json::to_string(&items).map_err(EngineError::message)
    }

    /// The story as an ordered run of formatted segments: the same view
    /// lowering reads:
    ///
    /// - `{"kind":"text","text","attributes"}`
    /// - `{"kind":"pilcrow","paraId","properties","attributes"}`
    /// - `{"kind":"embed","embedKind","payload","attributes"}`
    ///
    /// A text segment covers one maximal run of identically formatted
    /// characters; each pilcrow and embed is its own segment worth one unit.
    /// `attributes` holds the segment's run marks together with any `ins`/`del`
    /// tracked-change stamps. Errors on an unknown story.
    pub fn story_segments(&self, story: &str) -> EngineResult<String> {
        let segments = self
            .engine
            .doc()
            .story_segments(story)
            .map_err(EngineError::message)?;
        crate::segment_json::segments_json_string(&segments).map_err(EngineError::message)
    }

    /// `story_segments` split after each pilcrow into units, as one hex digest
    /// per unit: `["digest", …]`. Equal digests mean equal segments.
    pub fn story_segment_unit_digests(&self, story: &str) -> EngineResult<String> {
        let units = self
            .engine
            .doc()
            .story_segment_units(story)
            .map_err(EngineError::message)?;
        let digests: Vec<String> = units
            .iter()
            .map(|unit| format!("{:032x}", crate::segments_digest(unit)))
            .collect();
        serde_json::to_string(&digests).map_err(EngineError::message)
    }

    /// The segments of the listed units (indices into
    /// `story_segment_unit_digests`), each as `story_segments` gives them:
    /// `[[segment, …], …]`. Errors on an index past the last unit.
    pub fn story_segment_units(&self, story: &str, units: Vec<u32>) -> EngineResult<String> {
        let requested = self
            .engine
            .doc()
            .story_segment_units_at(story, &units)
            .map_err(EngineError::message)?
            .into_iter()
            .map(|unit| crate::segment_json::segments_json(unit).map_err(EngineError::message))
            .collect::<EngineResult<Vec<Vec<Value>>>>()?;
        serde_json::to_string(&requested).map_err(EngineError::message)
    }

    /// The `payload` of the story's `table_index`-th table embed, as
    /// `story_segments` gives it, or `None` when the story has no such
    /// table. Errors on an unknown story.
    pub fn table_payload(&self, story: &str, table_index: u32) -> EngineResult<Option<String>> {
        match self
            .engine
            .doc()
            .table_payload(&TableLocator::new(story, table_index))
        {
            Ok(payload) => serde_json::to_string(&attrs_value(&payload)?)
                .map(Some)
                .map_err(EngineError::message),
            Err(OpError::UnknownTable { .. }) => Ok(None),
            Err(error) => Err(EngineError::message(error)),
        }
    }

    /// `{"start","end"}`: the paragraph's span in story-global UTF-16 units.
    /// `end` is the index of its own pilcrow, so `end - start` is the
    /// paragraph length and the upper bound of a Loc `offset` in it. Errors
    /// when the paragraph is not in `story`.
    pub fn locate_paragraph(&self, story: &str, para_id: &str) -> EngineResult<String> {
        let span = find_para_span(self.engine.doc(), story, para_id)?;
        Ok(json!({ "start": span.start, "end": span.pilcrow }).to_string())
    }

    /// The index of the first paragraph of `story` that carries `para_id`, in
    /// `paragraph_spans` order, or -1. Errors on an unknown story.
    pub fn paragraph_unit(&self, story: &str, para_id: &str) -> EngineResult<i32> {
        Ok(self
            .engine
            .doc()
            .paragraph_index(story)
            .map_err(EngineError::message)?
            .para_slot(para_id)
            .map_or(-1, |slot| slot as i32))
    }

    /// How many paragraphs of `story` carry `para_id`: 0, 1, or 2 for two or
    /// more. Errors on an unknown story.
    pub fn paragraph_id_count(&self, story: &str, para_id: &str) -> EngineResult<u32> {
        Ok(self
            .engine
            .doc()
            .paragraph_index(story)
            .map_err(EngineError::message)?
            .para_id_count(para_id))
    }

    /// An edit request, or the refusal a step's unrepresentable text earns.
    fn edit_request(&self, json: &str) -> EngineResult<Result<EditRequest, EditRefusal>> {
        Ok(EditRequest::from_json(json)
            .map_err(EngineError::message)?
            .map_err(|failure| EditRefusal {
                version: self.engine.doc().version(),
                failure,
            }))
    }
}
