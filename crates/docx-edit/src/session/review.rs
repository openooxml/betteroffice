use serde_json::{Value, json};
use yrs::Any;

use super::args::{comment_ranges, find_para_span, json_to_any, parse_change_target};
use super::{EditSessionCore, EngineError, EngineResult};
use crate::batch::outcome_json;
use crate::{ChangeKind, EditCtx};

impl EditSessionCore {
    /// Adds a comment anchored to one or more ranges. `ranges_json`:
    /// `[{"story","startPara","startOffset","endPara","endOffset"}, …]`;
    /// `body_json` is any JSON value, stored as given. The anchors are sticky,
    /// so they follow the text they cover, and the comment lives outside the
    /// story rather than as an attribute on it. Receipt: `{"commentId"}`: a
    /// freshly minted id. Errors when `ranges_json` is not an array of
    /// well-formed ranges, or is empty.
    pub fn add_comment(
        &self,
        ranges_json: &str,
        author: &str,
        date: &str,
        body_json: &str,
    ) -> EngineResult<String> {
        let ranges = comment_ranges(self.engine.doc(), ranges_json)?;
        let body = Any::from_json(body_json).map_err(EngineError::message)?;
        let comment_id = self
            .engine
            .doc()
            .add_comment(&ranges, author, date, body)
            .map_err(EngineError::message)?;
        Ok(json!({ "commentId": comment_id }).to_string())
    }

    /// Reanchors an existing comment without changing its metadata.
    pub fn set_comment_ranges(&self, comment_id: &str, ranges_json: &str) -> EngineResult<()> {
        let ranges = comment_ranges(self.engine.doc(), ranges_json)?;
        self.engine
            .doc()
            .set_comment_ranges(comment_id, &ranges)
            .map_err(EngineError::message)
    }

    /// Initializes imported metadata, creating an unanchored comment when absent.
    pub fn seed_comment_metadata_json(
        &self,
        comment_id: &str,
        patch_json: &str,
    ) -> EngineResult<()> {
        let patch: Value = serde_json::from_str(patch_json).map_err(EngineError::message)?;
        let fields = patch
            .as_object()
            .ok_or_else(|| EngineError::message("comment metadata must be an object"))?;
        let patch = fields
            .iter()
            .map(|(key, value)| Ok((key.clone(), json_to_any(value)?)))
            .collect::<EngineResult<_>>()?;
        self.engine
            .doc()
            .seed_comment_metadata(comment_id, patch)
            .map_err(EngineError::message)
    }

    pub fn update_comment_metadata_json(
        &self,
        comment_id: &str,
        patch_json: &str,
    ) -> EngineResult<()> {
        let patch: Value = serde_json::from_str(patch_json).map_err(EngineError::message)?;
        let fields = patch
            .as_object()
            .ok_or_else(|| EngineError::message("comment metadata must be an object"))?;
        let patch = fields
            .iter()
            .map(|(key, value)| Ok((key.clone(), json_to_any(value)?)))
            .collect::<EngineResult<_>>()?;
        self.engine
            .doc()
            .update_comment_metadata(comment_id, patch)
            .map_err(EngineError::message)
    }

    /// Accepts tracked changes: pending insertions become plain content,
    /// pending deletions are carried out; `pPrIns` marks clear so the split
    /// stays, `pPrDel` marks join with the following paragraph, whose own
    /// properties survive.
    ///
    /// `target_json` is either `{"revisionId": string}`, resolving that one
    /// coalesced revision wherever it appears in any story, or
    /// `{"story","startPara","startOffset","endPara","endOffset"}`, resolving
    /// every tracked change overlapping that range regardless of revision id.
    /// Receipt: `{"revisionIds": [string, …]}`: the ids actually resolved, in
    /// resolution order and deduplicated. Resolving applies a revision rather
    /// than authoring one, so it never stamps a new revision and ignores
    /// suggesting mode. Errors on an empty range and on a `revisionId` that
    /// matches nothing.
    pub fn accept_change(&self, target_json: &str) -> EngineResult<String> {
        let target = parse_change_target(self.engine.doc(), target_json)?;
        let ctx = EditCtx::local(String::new(), String::new());
        let receipt = self
            .engine
            .doc()
            .accept_change(&ctx, &target)
            .map_err(EngineError::message)?;
        Ok(json!({ "revisionIds": receipt.revision_ids }).to_string())
    }

    /// Rejects tracked changes: the inverse of
    /// [`EditSessionCore::accept_change`]: pending insertions roll back, pending
    /// deletions restore their text; `pPrIns` marks join back with the
    /// following paragraph, `pPrDel` marks clear so the split stays, and a
    /// `pPrChange` restores the paragraph's previous properties. Same target,
    /// receipt and error contract.
    pub fn reject_change(&self, target_json: &str) -> EngineResult<String> {
        let target = parse_change_target(self.engine.doc(), target_json)?;
        let ctx = EditCtx::local(String::new(), String::new());
        let receipt = self
            .engine
            .doc()
            .reject_change(&ctx, &target)
            .map_err(EngineError::message)?;
        Ok(json!({ "revisionIds": receipt.revision_ids }).to_string())
    }

    /// Resolves tracked changes by revision id outside undo history:
    /// `{"accept":[string, …],"reject":[string, …]}` -> `{"ok":true,"revisionIds":[string, …]}`,
    /// the ids resolved, or `{"ok":false,"version","failure"}`. An id that matches nothing is
    /// skipped. See [`crate::EditingDoc::settle_revisions`].
    pub fn settle_revisions_json(&self, request: &str) -> EngineResult<String> {
        #[derive(serde::Deserialize)]
        struct Settle {
            #[serde(default)]
            accept: Vec<String>,
            #[serde(default)]
            reject: Vec<String>,
        }
        let request: Settle = serde_json::from_str(request).map_err(EngineError::message)?;
        let outcome = self
            .engine
            .doc()
            .settle_revisions(&request.accept, &request.reject, &self.undo)
            .map_err(EngineError::message)?
            .map(|resolved| json!({ "revisionIds": resolved }));
        outcome_json(&outcome).map_err(EngineError::message)
    }

    pub fn proposal_revision_ranges_json(&self, ids_json: &str) -> EngineResult<String> {
        let ids = match serde_json::from_str::<Vec<String>>(ids_json) {
            Ok(ids) => ids,
            Err(_) => return Ok("\"legacy\"".to_owned()),
        };
        serde_json::to_string(&self.engine.doc().owned_revision_ranges(&ids))
            .map_err(EngineError::message)
    }

    /// Every pending tracked change across all stories, in deterministic
    /// story-then-position order:
    /// `[{"revisionId","author","date","kind","story","preview",
    /// "range":{"story","start":{"paraId","offset"},"end":{…}}}, …]`. `kind`
    /// is one of `"insertion"`, `"deletion"`, `"pPrIns"`, `"pPrDel"`,
    /// `"pPrChange"`, `"trIns"`, `"trDel"`, `"tableIns"` or `"tableDel"`.
    /// `preview` holds the first few characters of the affected text, and is
    /// empty for every structural kind, which covers no text.
    pub fn list_revisions(&self) -> EngineResult<String> {
        let revisions = self
            .engine
            .doc()
            .list_revisions()
            .map_err(EngineError::message)?;
        let items: Vec<Value> = revisions
            .into_iter()
            .map(|revision| {
                let kind = match revision.change.kind {
                    ChangeKind::Insertion => "insertion",
                    ChangeKind::Deletion => "deletion",
                    ChangeKind::ParagraphMarkInsertion => "pPrIns",
                    ChangeKind::ParagraphMarkDeletion => "pPrDel",
                    ChangeKind::ParagraphPropertiesChanged => "pPrChange",
                    ChangeKind::TableRowInsertion => "trIns",
                    ChangeKind::TableRowDeletion => "trDel",
                    ChangeKind::TableInsertion => "tableIns",
                    ChangeKind::TableDeletion => "tableDel",
                };
                json!({
                    "revisionId": revision.change.revision_id,
                    "author": revision.change.author,
                    "date": revision.change.date,
                    "kind": kind,
                    "story": revision.story,
                    "preview": revision.preview,
                    "range": {
                        "story": revision.change.range.start.story,
                        "start": {
                            "paraId": revision.change.range.start.para,
                            "offset": revision.change.range.start.offset,
                        },
                        "end": {
                            "paraId": revision.change.range.end.para,
                            "offset": revision.change.range.end.offset,
                        },
                    },
                })
            })
            .collect();
        serde_json::to_string(&items).map_err(EngineError::message)
    }

    /// Author and date stamps for the requested revision ids.
    pub fn revision_stamps_json(&self, ids_json: &str) -> EngineResult<String> {
        let ids: Vec<String> = serde_json::from_str(ids_json).map_err(EngineError::message)?;
        let stamps = self
            .engine
            .doc()
            .revision_stamps(&ids)
            .map_err(EngineError::message)?;
        let items: serde_json::Map<String, Value> = stamps
            .into_iter()
            .map(|(id, stamps)| {
                let stamps: Vec<Value> = stamps
                    .into_iter()
                    .map(|(author, date)| json!({ "author": author, "date": date }))
                    .collect();
                (id, Value::Array(stamps))
            })
            .collect();
        serde_json::to_string(&items).map_err(EngineError::message)
    }

    /// Ids of the tracked changes of `story` whose range holds the caret at
    /// `para_id`/`offset`, in `list_revisions` order: `["id", …]`. The caret
    /// and both range ends resolve through their paragraph's span start as
    /// `locate_paragraph` gives it; a range whose paragraph does not resolve
    /// holds nothing. Reads only `story`, unlike `list_revisions`.
    pub fn revision_ids_at(&self, story: &str, para_id: &str, offset: u32) -> EngineResult<String> {
        let doc = self.engine.doc();
        find_para_span(doc, story, para_id)?;
        let ids = doc
            .revision_ids_at(story, para_id, offset)
            .map_err(EngineError::message)?;
        serde_json::to_string(&ids).map_err(EngineError::message)
    }

    /// Every comment the session holds, sorted by id:
    /// `[{"id","author","date","done","parentId","body"}, …]`, `parentId`
    /// null for a top-level comment and `body` the JSON value it was given.
    pub fn list_comments(&self) -> EngineResult<String> {
        let comments = self
            .engine
            .doc()
            .list_comments()
            .map_err(EngineError::message)?;
        let mut items = Vec::with_capacity(comments.len());
        for comment in comments {
            items.push(json!({
                "id": comment.id,
                "author": comment.author,
                "date": comment.date,
                "done": comment.done,
                "parentId": comment.parent_id,
                "body": serde_json::to_value(&comment.body).map_err(EngineError::message)?,
            }));
        }
        serde_json::to_string(&items).map_err(EngineError::message)
    }

    /// Where a comment's sticky anchors currently sit:
    /// `[{"story","start","end"}, …]`, one entry per anchored range, in
    /// story-global UTF-16 units. Errors on an unknown comment id and when an
    /// anchor no longer resolves.
    pub fn resolve_comment(&self, comment_id: &str) -> EngineResult<String> {
        let anchors = self
            .engine
            .doc()
            .resolve_comment(comment_id)
            .map_err(EngineError::message)?;
        let items: Vec<Value> = anchors
            .into_iter()
            .map(
                |anchor| json!({ "story": anchor.story, "start": anchor.start, "end": anchor.end }),
            )
            .collect();
        serde_json::to_string(&items).map_err(EngineError::message)
    }
}
