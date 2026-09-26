//! Staging shared by proposals and edit batches: a private replica of the live deck that shares
//! its client id, package and id allocation, and the one-transaction adoption of what it staged.

use std::sync::atomic::Ordering;

use yrs::{ReadTxn, StateVector, Transact, Update};

use crate::{
    DeckSession, DeckSnapshot, EditError, EditResult, decode_update_v1, doc_with_client_id,
    hydrate_doc,
};

/// Origin of adoptions that stay out of local undo history.
const HOST_ORIGIN: &str = "pptx:host";

#[cfg(test)]
thread_local! {
    /// Makes rehearsals on this thread report divergence, to exercise that failure.
    pub(crate) static FORCE_DIVERGENCE: std::cell::Cell<bool> = const { std::cell::Cell::new(false) };
}

/// How an adoption enters local undo history.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum Adoption {
    /// A tracked step bounded only where the capture mode closes groups by itself.
    Proposal,
    /// Exactly one tracked step, whatever the capture mode.
    Tracked,
    /// No undo step; existing entries stay and surrounding groups do not merge across it.
    Untracked,
}

fn pending<T: ReadTxn>(txn: &T) -> bool {
    txn.store().pending_update().is_some() || txn.store().pending_ds().is_some()
}

impl DeckSession {
    /// A private replica of the committed state that allocates ids where this session would.
    pub(crate) fn stage(&self) -> EditResult<DeckSession> {
        let doc = doc_with_client_id(self.client_id);
        hydrate_doc(&doc, &self.state_update_v1())?;
        DeckSession::assemble(
            doc,
            self.client_id,
            self.id_counter.load(Ordering::Relaxed),
            self.package.clone(),
            0,
        )
    }

    pub(crate) fn validated_snapshot(&self) -> EditResult<DeckSnapshot> {
        crate::deck::validated_snapshot(&self.doc, &self.package)
    }

    /// What this replica holds beyond `base`, encoded and predecoded.
    pub(crate) fn staged_update(&self, base: &StateVector) -> EditResult<(Vec<u8>, Update)> {
        let bytes = self.doc.transact().encode_diff_v1(base);
        let update = decode_update_v1(&bytes).map_err(EditError::InvalidUpdate)?;
        Ok((bytes, update))
    }

    /// Whether updates wait on missing dependencies; adopting next to them could integrate
    /// more than the staged edits.
    pub(crate) fn has_pending_updates(&self) -> bool {
        pending(&self.doc.transact())
    }

    /// Integrates `update` into an untouched replica of `base` and checks it reproduces the
    /// stage, so adopting it cannot meet a failure the rehearsal did not.
    pub(crate) fn rehearse(
        &self,
        base: &[u8],
        update: &[u8],
        stage: &DeckSession,
        staged: &DeckSnapshot,
    ) -> EditResult<()> {
        let rehearsal = doc_with_client_id(self.client_id);
        hydrate_doc(&rehearsal, base)?;
        rehearsal
            .transact_mut()
            .apply_update(decode_update_v1(update).map_err(EditError::InvalidUpdate)?)
            .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        let diverged = pending(&rehearsal.transact())
            || rehearsal.transact().state_vector() != stage.doc.transact().state_vector()
            || crate::deck::validated_snapshot(&rehearsal, &self.package)? != *staged;
        #[cfg(test)]
        let diverged = diverged || FORCE_DIVERGENCE.get();
        if diverged {
            return Err(EditError::InvalidUpdate(
                "the rehearsed adoption diverged from the staged edits".to_owned(),
            ));
        }
        Ok(())
    }

    /// Integrates a stage's update as one local transaction and takes over its id allocation.
    pub(crate) fn adopt(
        &self,
        stage: &DeckSession,
        update: Update,
        adoption: Adoption,
    ) -> EditResult<()> {
        let boundary = || match adoption {
            Adoption::Proposal => self.automatic_undo_barrier(),
            Adoption::Tracked | Adoption::Untracked => self.add_undo_barrier(),
        };
        boundary();
        match adoption {
            Adoption::Untracked => self.doc.transact_mut_with(HOST_ORIGIN),
            Adoption::Proposal | Adoption::Tracked => self.doc.transact_mut_with(self.client_id),
        }
        .apply_update(update)
        .map_err(|error| EditError::InvalidUpdate(error.to_string()))?;
        self.id_counter
            .store(stage.id_counter.load(Ordering::Relaxed), Ordering::Relaxed);
        boundary();
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use std::cell::Cell;
    use std::rc::Rc;
    use std::sync::atomic::Ordering;

    use super::FORCE_DIVERGENCE;
    use crate::{
        DeckSession, EditCtx, EditError, EditRequest, ProposalEdit, ProposalRequest, ReadRequest,
        TextStyle,
    };

    const DECK: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");

    #[test]
    fn a_diverging_rehearsal_fails_without_touching_the_deck() {
        let session = DeckSession::open(DECK, 95).unwrap();
        let read = session
            .read_content(&ReadRequest::default())
            .unwrap()
            .unwrap();
        let story = &read.stories[0];
        session
            .insert_text(
                &EditCtx::local("human"),
                &story.story_id,
                0,
                "X",
                &TextStyle::default(),
            )
            .unwrap();
        assert!(session.undo());
        session.next_id("probe");
        let proposal = session
            .propose(ProposalRequest {
                agent_id: "agent".into(),
                note: None,
                edits: vec![ProposalEdit::SetSlideNotes {
                    slide_id: story.slide_id.clone(),
                    text: "Proposed".into(),
                }],
            })
            .unwrap();
        let events = Rc::new(Cell::new(0));
        let counter = Rc::clone(&events);
        let _observer = session
            .observe_update_v1(move |_| counter.set(counter.get() + 1))
            .unwrap();
        let (state, version, ids) = (
            session.encode_state_as_update_v1(),
            session.version(),
            session.id_counter.load(Ordering::Relaxed),
        );
        let request: EditRequest = serde_json::from_value(serde_json::json!({
            "expectVersion": version,
            "steps": [{"op": "insertText", "at": "start", "text": "Draft: ", "target": {
                "kind": "range", "slideId": story.slide_id, "shapeId": story.shape_id,
                "storyId": story.story_id, "start": 0, "end": 0,
            }}],
        }))
        .unwrap();

        FORCE_DIVERGENCE.set(true);
        let validated = session.validate_edits(&request).map(|_| ());
        let applied = session.apply_edits(&request).map(|_| ());
        FORCE_DIVERGENCE.set(false);
        for outcome in [validated, applied] {
            assert!(
                matches!(&outcome, Err(EditError::InvalidUpdate(message)) if message.contains("diverged")),
                "{outcome:?}"
            );
        }
        assert_eq!(session.encode_state_as_update_v1(), state);
        assert_eq!(session.version(), version);
        assert_eq!(session.id_counter.load(Ordering::Relaxed), ids);
        assert!(!session.can_undo());
        assert!(session.can_redo());
        assert_eq!(session.proposals().unwrap(), [proposal]);
        assert_eq!(events.get(), 0);

        assert!(session.apply_edits(&request).unwrap().unwrap().applied);
        assert_eq!(events.get(), 1);
    }
}
