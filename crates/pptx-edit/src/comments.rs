use pptx_parse::{CommentFlavor, PptxPackage};
use sha2::{Digest, Sha256};
use yrs::{Any, Map, MapPrelim, MapRef, ReadTxn, Transact, TransactionMut, WriteTxn};

use crate::deck::{map_bool, map_number, map_string, required_map};
use crate::{
    BOOTSTRAP_CLIENT_ID, COMMENTS, CommentReceipt, CommentSnapshot, DeckSession, EditCtx,
    EditError, EditResult, META, MIGRATE_ORIGIN, SLIDES, doc_with_client_id, hydrate_doc,
};

pub(crate) fn seeded_comment_id(index: usize, source_id: &str) -> String {
    format!("comment:{index}:{source_id}")
}

pub(crate) fn seed_comments(
    txn: &mut TransactionMut<'_>,
    package: &PptxPackage,
    slide_id_by_part: &dyn Fn(&str) -> Option<String>,
) -> EditResult<()> {
    let comments = txn.get_or_insert_map(COMMENTS);
    for (index, comment) in package.comments.iter().enumerate() {
        let Some(slide_id) = slide_id_by_part(&comment.slide_part_path) else {
            continue;
        };
        let author = package
            .comment_authors
            .iter()
            .find(|author| author.id == comment.author_id);
        let id = seeded_comment_id(index, &comment.id);
        let entry = comments.insert(txn, id.as_str(), MapPrelim::default());
        entry.insert(txn, "id", id.as_str());
        entry.insert(txn, "slideId", slide_id.as_str());
        entry.insert(
            txn,
            "author",
            author.map(|author| author.name.as_str()).unwrap_or(""),
        );
        entry.insert(
            txn,
            "initials",
            author.map(|author| author.initials.as_str()).unwrap_or(""),
        );
        entry.insert(txn, "text", comment.text.as_str());
        if let Some(created) = &comment.created {
            entry.insert(txn, "created", created.as_str());
        }
        entry.insert(txn, "x", comment.x_emu as f64);
        entry.insert(txn, "y", comment.y_emu as f64);
        if let Some(parent) = &comment.parent_id {
            let parent_index = package
                .comments
                .iter()
                .position(|candidate| &candidate.id == parent);
            if let Some(parent_index) = parent_index {
                entry.insert(
                    txn,
                    "parentId",
                    seeded_comment_id(parent_index, parent).as_str(),
                );
            }
        }
        entry.insert(
            txn,
            "resolved",
            comment.status.as_deref() == Some("resolved"),
        );
    }
    Ok(())
}

pub(crate) fn import_source_comments(
    session: &DeckSession,
    source: &PptxPackage,
) -> EditResult<()> {
    {
        let txn = session.doc.transact();
        let meta = required_map(&txn, META)?;
        if map_bool(&meta, &txn, "commentsPendingSource") != Some(true) {
            return Ok(());
        }
    }
    let bootstrap = doc_with_client_id(BOOTSTRAP_CLIENT_ID);
    hydrate_doc(&bootstrap, &session.encode_state_as_update_v1())?;
    let slide_ids: std::collections::HashMap<_, _> = source
        .presentation
        .slides
        .iter()
        .enumerate()
        .map(|(index, slide)| {
            (
                slide.part_path.as_str(),
                format!("slide:{index}:{}", slide.id),
            )
        })
        .collect();
    seed_comments(
        &mut bootstrap.transact_mut_with(MIGRATE_ORIGIN),
        source,
        &|part| slide_ids.get(part).cloned(),
    )?;
    let update = bootstrap
        .transact()
        .encode_diff_v1(&session.doc.transact().state_vector());
    hydrate_doc(&session.doc, &update)?;
    let mut package = session.package().clone();
    package.comments = source.comments.clone();
    package.comment_authors = source.comment_authors.clone();
    package.comment_flavor = source.comment_flavor;
    let mut txn = session.doc.transact_mut_with(MIGRATE_ORIGIN);
    let meta = required_map(&txn, META)?;
    meta.insert(
        &mut txn,
        "packageJson",
        Any::Buffer(std::sync::Arc::from(
            serde_json::to_vec(&package).map_err(|error| EditError::Json(error.to_string()))?,
        )),
    );
    meta.insert(
        &mut txn,
        "commentFlavor",
        flavor_key(source.comment_flavor.unwrap_or_default()),
    );
    meta.insert(&mut txn, "commentsPendingSource", false);
    Ok(())
}

pub(crate) fn snapshot_comments<T: ReadTxn>(txn: &T) -> EditResult<Vec<CommentSnapshot>> {
    let comments = required_map(txn, COMMENTS)?;
    let slides = required_map(txn, SLIDES)?;
    let mut output = Vec::new();
    for (id, value) in comments.iter(txn) {
        let Ok(entry) = value.cast::<MapRef>() else {
            return Err(EditError::InvalidState(format!(
                "comment {id} is not a map"
            )));
        };
        let slide_id = map_string(&entry, txn, "slideId").unwrap_or_default();
        if !slides.contains_key(txn, &slide_id) {
            continue;
        }
        output.push(CommentSnapshot {
            id: id.to_owned(),
            slide_id,
            author: map_string(&entry, txn, "author").unwrap_or_default(),
            initials: map_string(&entry, txn, "initials").unwrap_or_default(),
            text: map_string(&entry, txn, "text").unwrap_or_default(),
            created: map_string(&entry, txn, "created"),
            x_emu: map_number(&entry, txn, "x").unwrap_or(0.0) as i64,
            y_emu: map_number(&entry, txn, "y").unwrap_or(0.0) as i64,
            parent_id: live_parent(&comments, &entry, txn),
            resolved: map_bool(&entry, txn, "resolved").unwrap_or(false),
        });
    }
    output.sort_by(|left, right| {
        left.created
            .cmp(&right.created)
            .then_with(|| left.id.cmp(&right.id))
    });
    Ok(output)
}

/// Promotes alternating levels so undo cannot create nested replies.
fn live_parent<T: ReadTxn>(comments: &MapRef, entry: &MapRef, txn: &T) -> Option<String> {
    let parent_id = map_string(entry, txn, "parentId")?;
    let mut next = Some(parent_id.clone());
    let mut seen = std::collections::HashSet::new();
    let mut depth = 0;
    while let Some(id) = next {
        if !seen.insert(id.clone()) || depth == 128 {
            return None;
        }
        let Some(parent) = comments
            .get(txn, &id)
            .and_then(|value| value.cast::<MapRef>().ok())
        else {
            break;
        };
        depth += 1;
        next = map_string(&parent, txn, "parentId");
    }
    (depth % 2 == 1).then_some(parent_id)
}

pub(crate) fn snapshot_flavor<T: ReadTxn>(txn: &T) -> EditResult<CommentFlavor> {
    let meta = required_map(txn, META)?;
    Ok(match map_string(&meta, txn, "commentFlavor").as_deref() {
        Some("modern") => CommentFlavor::Modern,
        _ => CommentFlavor::Legacy,
    })
}

pub(crate) fn flavor_key(flavor: CommentFlavor) -> &'static str {
    match flavor {
        CommentFlavor::Legacy => "legacy",
        CommentFlavor::Modern => "modern",
    }
}

pub(crate) fn derived_guid(seed: &str) -> String {
    let digest = Sha256::digest(seed.as_bytes());
    let hex: String = digest
        .iter()
        .take(16)
        .map(|byte| format!("{byte:02X}"))
        .collect();
    format!(
        "{{{}-{}-{}-{}-{}}}",
        &hex[0..8],
        &hex[8..12],
        &hex[12..16],
        &hex[16..20],
        &hex[20..32]
    )
}

impl DeckSession {
    #[allow(clippy::too_many_arguments)]
    pub fn add_comment(
        &self,
        context: &EditCtx,
        slide_id: &str,
        author: &str,
        initials: &str,
        text: &str,
        created: &str,
        x_emu: i64,
        y_emu: i64,
    ) -> EditResult<CommentReceipt> {
        validate_fields(author, initials, text, created)?;
        if text.is_empty() {
            return Err(EditError::InvalidComment("comment text is empty".into()));
        }
        self.add_undo_barrier();
        let comment_id = self.next_id("comment");
        let mut txn = self.transact_for(context);
        crate::deck::slide_ref(&txn, slide_id)?;
        let comments = required_map(&txn, COMMENTS)?;
        let entry = comments.insert(&mut txn, comment_id.as_str(), MapPrelim::default());
        entry.insert(&mut txn, "id", comment_id.as_str());
        entry.insert(&mut txn, "slideId", slide_id);
        entry.insert(&mut txn, "author", author);
        entry.insert(&mut txn, "initials", initials);
        entry.insert(&mut txn, "text", text);
        entry.insert(&mut txn, "created", created);
        entry.insert(&mut txn, "x", x_emu as f64);
        entry.insert(&mut txn, "y", y_emu as f64);
        entry.insert(&mut txn, "resolved", false);
        drop(txn);
        self.add_undo_barrier();
        Ok(CommentReceipt {
            comment_id,
            slide_id: slide_id.to_owned(),
            parent_id: None,
            resolved: false,
        })
    }

    pub fn reply_to_comment(
        &self,
        context: &EditCtx,
        comment_id: &str,
        author: &str,
        initials: &str,
        text: &str,
        created: &str,
    ) -> EditResult<CommentReceipt> {
        validate_fields(author, initials, text, created)?;
        if text.is_empty() {
            return Err(EditError::InvalidComment("reply text is empty".into()));
        }
        self.add_undo_barrier();
        let reply_id = self.next_id("comment");
        let mut txn = self.transact_for(context);
        require_modern(&txn)?;
        let comments = required_map(&txn, COMMENTS)?;
        let parent = comment_ref(&comments, &txn, comment_id)?;
        if live_parent(&comments, &parent, &txn).is_some() {
            return Err(EditError::InvalidComment(
                "replies cannot be nested below a reply".into(),
            ));
        }
        let slide_id = map_string(&parent, &txn, "slideId").unwrap_or_default();
        let entry = comments.insert(&mut txn, reply_id.as_str(), MapPrelim::default());
        entry.insert(&mut txn, "id", reply_id.as_str());
        entry.insert(&mut txn, "slideId", slide_id.as_str());
        entry.insert(&mut txn, "author", author);
        entry.insert(&mut txn, "initials", initials);
        entry.insert(&mut txn, "text", text);
        entry.insert(&mut txn, "created", created);
        entry.insert(&mut txn, "x", 0.0);
        entry.insert(&mut txn, "y", 0.0);
        entry.insert(&mut txn, "parentId", comment_id);
        entry.insert(&mut txn, "resolved", false);
        drop(txn);
        self.add_undo_barrier();
        Ok(CommentReceipt {
            comment_id: reply_id,
            slide_id,
            parent_id: Some(comment_id.to_owned()),
            resolved: false,
        })
    }

    pub fn set_comment_status(
        &self,
        context: &EditCtx,
        comment_id: &str,
        resolved: bool,
    ) -> EditResult<CommentReceipt> {
        self.add_undo_barrier();
        let mut txn = self.transact_for(context);
        require_modern(&txn)?;
        let comments = required_map(&txn, COMMENTS)?;
        let entry = comment_ref(&comments, &txn, comment_id)?;
        let slide_id = map_string(&entry, &txn, "slideId").unwrap_or_default();
        let parent_id = live_parent(&comments, &entry, &txn);
        entry.insert(&mut txn, "resolved", resolved);
        drop(txn);
        self.add_undo_barrier();
        Ok(CommentReceipt {
            comment_id: comment_id.to_owned(),
            slide_id,
            parent_id,
            resolved,
        })
    }

    pub fn remove_comment(
        &self,
        context: &EditCtx,
        comment_id: &str,
    ) -> EditResult<CommentReceipt> {
        self.add_undo_barrier();
        let mut txn = self.transact_for(context);
        let comments = required_map(&txn, COMMENTS)?;
        let entry = comment_ref(&comments, &txn, comment_id)?;
        let slide_id = map_string(&entry, &txn, "slideId").unwrap_or_default();
        let parent_id = live_parent(&comments, &entry, &txn);
        let resolved = map_bool(&entry, &txn, "resolved").unwrap_or(false);
        let replies: Vec<String> = comments
            .iter(&txn)
            .filter_map(|(id, value)| {
                let entry = value.cast::<MapRef>().ok()?;
                (map_string(&entry, &txn, "parentId").as_deref() == Some(comment_id))
                    .then(|| id.to_owned())
            })
            .collect();
        for reply in replies {
            comments.remove(&mut txn, &reply);
        }
        comments.remove(&mut txn, comment_id);
        drop(txn);
        self.add_undo_barrier();
        Ok(CommentReceipt {
            comment_id: comment_id.to_owned(),
            slide_id,
            parent_id,
            resolved,
        })
    }

    pub fn set_comment_flavor(
        &self,
        context: &EditCtx,
        flavor: CommentFlavor,
    ) -> EditResult<CommentFlavor> {
        self.add_undo_barrier();
        let mut txn = self.transact_for(context);
        let comments = required_map(&txn, COMMENTS)?;
        if comments.len(&txn) > 0 {
            return Err(EditError::InvalidComment(
                "the comment flavour is fixed once a deck has comments".into(),
            ));
        }
        let meta = required_map(&txn, META)?;
        meta.insert(&mut txn, "commentFlavor", flavor_key(flavor));
        drop(txn);
        self.add_undo_barrier();
        Ok(flavor)
    }

    pub fn comments(&self) -> EditResult<Vec<CommentSnapshot>> {
        snapshot_comments(&self.doc.transact())
    }

    pub fn comment_flavor(&self) -> EditResult<CommentFlavor> {
        snapshot_flavor(&self.doc.transact())
    }
}

fn validate_fields(author: &str, initials: &str, text: &str, created: &str) -> EditResult<()> {
    for value in [author, initials, text, created] {
        crate::model::validate_xml_text(value)?;
    }
    Ok(())
}

fn require_modern<T: ReadTxn>(txn: &T) -> EditResult<()> {
    match snapshot_flavor(txn)? {
        CommentFlavor::Modern => Ok(()),
        CommentFlavor::Legacy => Err(EditError::InvalidComment(
            "legacy comments carry no replies or status; switch the deck to modern comments".into(),
        )),
    }
}

fn comment_ref<T: ReadTxn>(comments: &MapRef, txn: &T, comment_id: &str) -> EditResult<MapRef> {
    comments
        .get(txn, comment_id)
        .and_then(|value| value.cast::<MapRef>().ok())
        .ok_or_else(|| EditError::CommentNotFound(comment_id.to_owned()))
}
