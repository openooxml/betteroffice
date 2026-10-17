//! yrs cleans up redundant formatting inside the transaction that integrates a remote update,
//! so those deletions ride in an update observers label remote, which no peer is sent. They
//! depend on what this replica held at the time, so peers may never make them on their own.

use std::sync::{Arc, Mutex};

use yrs::encoding::write::Write;
use yrs::error::UpdateError;
use yrs::updates::encoder::{Encode, Encoder, EncoderV1};
use yrs::{IdSet, ReadTxn, TransactionMut, Update};

/// Deletions carried by peers, caused by integrating their structs, or held pending when
/// the integration opened.
#[derive(Clone, Default)]
pub(crate) struct RemoteDeletions(Arc<Mutex<Option<IdSet>>>);

impl RemoteDeletions {
    /// Opens `txn` as a remote integration. Holding `txn` keeps every other transaction out
    /// until the integration has committed and cleared its carried set on drop.
    pub(crate) fn integrate<'doc>(&self, txn: TransactionMut<'doc>) -> RemoteIntegration<'_, 'doc> {
        let pending = ReadTxn::store(&txn)
            .pending_ds()
            .cloned()
            .unwrap_or_default();
        *self.0.lock().unwrap() = Some(pending);
        RemoteIntegration {
            deletions: self,
            txn,
        }
    }

    /// What `txn` deleted beyond the carried deletions, as a delete-only v1 update: `None`
    /// outside a remote integration or when it deleted nothing else.
    pub(crate) fn residual_update_v1(&self, txn: &TransactionMut) -> Option<Vec<u8>> {
        let carried = self.0.lock().unwrap();
        let residual = txn.delete_set().diff(carried.as_ref()?);
        (!residual.is_empty()).then(|| delete_only_update_v1(&residual))
    }
}

/// One remote transaction; any number of peer updates may be carried into it.
pub(crate) struct RemoteIntegration<'a, 'doc> {
    deletions: &'a RemoteDeletions,
    txn: TransactionMut<'doc>,
}

impl RemoteIntegration<'_, '_> {
    /// Adds deletions a peer sent, so they are not republished as this replica's own.
    pub(crate) fn carry(&mut self, deletions: &IdSet) {
        if let Some(carried) = self.deletions.0.lock().unwrap().as_mut() {
            carried.merge_with(deletions.clone());
        }
    }

    pub(crate) fn apply_update(&mut self, update: Update) -> Result<(), UpdateError> {
        self.txn.apply_update(update)
    }
}

impl Drop for RemoteIntegration<'_, '_> {
    fn drop(&mut self) {
        self.txn.commit();
        *self.deletions.0.lock().unwrap() = None;
    }
}

pub(crate) fn apply_staged_update(
    txn: &mut TransactionMut,
    update: Update,
) -> Result<IdSet, UpdateError> {
    let mut carried = update.delete_set().clone();
    txn.apply_update(update)?;
    carried.merge_with(txn.delete_set().clone());
    Ok(carried)
}

fn delete_only_update_v1(deletions: &IdSet) -> Vec<u8> {
    let mut encoder = EncoderV1::new();
    encoder.write_var(0_u32);
    deletions.encode(&mut encoder);
    encoder.to_vec()
}

#[cfg(test)]
mod tests {
    use super::*;
    use yrs::types::Attrs;
    use yrs::updates::decoder::Decode;
    use yrs::{Any, ClientID, Doc, GetString, Options, StateVector, Subscription, Text, Transact};

    fn replica(client: u64) -> Doc {
        Doc::with_options(Options::with_client_id(ClientID::new(client)))
    }

    fn edit(doc: &Doc, change: impl FnOnce(&mut TransactionMut, &yrs::TextRef)) -> Vec<u8> {
        let text = doc.get_or_insert_text("t");
        let mut txn = doc.transact_mut();
        change(&mut txn, &text);
        txn.encode_update_v1()
    }

    fn published(
        doc: &Doc,
        deletions: &RemoteDeletions,
    ) -> (Subscription, Arc<Mutex<Vec<Vec<u8>>>>) {
        let sink = Arc::new(Mutex::new(Vec::new()));
        let (residuals, deletions) = (Arc::clone(&sink), deletions.clone());
        let subscription = doc
            .observe_update_v1(move |txn, _| {
                residuals
                    .lock()
                    .unwrap()
                    .extend(deletions.residual_update_v1(txn));
            })
            .unwrap();
        (subscription, sink)
    }

    fn integrate(doc: &Doc, deletions: &RemoteDeletions, update: &[u8]) {
        let update = Update::decode_v1(update).unwrap();
        let mut integration = deletions.integrate(doc.transact_mut());
        integration.carry(update.delete_set());
        integration.apply_update(update).unwrap();
    }

    fn text(doc: &Doc) -> String {
        doc.get_or_insert_text("t").get_string(&doc.transact())
    }

    #[test]
    fn residual_update_deletes_exactly_the_formatting_cleanup() {
        let origin = replica(7);
        edit(&origin, |txn, text| text.insert(txn, 0, "abcd"));
        let state = origin
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let (peer, behind) = (replica(8), replica(9));
        for doc in [&peer, &behind] {
            let state = Update::decode_v1(&state).unwrap();
            doc.transact_mut().apply_update(state).unwrap();
        }
        let bold = || Attrs::from([(Arc::from("b"), Any::Bool(true))]);
        let theirs = edit(&origin, |txn, text| text.format(txn, 0, 2, bold()));
        let ours = edit(&peer, |txn, text| text.format(txn, 0, 2, bold()));

        let deletions = RemoteDeletions::default();
        let (_subscription, residuals) = published(&peer, &deletions);
        let expected = Arc::new(Mutex::new(IdSet::new()));
        let deleted = Arc::clone(&expected);
        let _observer = peer
            .observe_update_v1(move |txn, _| {
                deleted.lock().unwrap().merge_with(txn.delete_set().clone());
            })
            .unwrap();
        assert!(Update::decode_v1(&theirs).unwrap().delete_set().is_empty());
        integrate(&peer, &deletions, &theirs);
        assert!(deletions.residual_update_v1(&peer.transact_mut()).is_none());

        let residuals = residuals.lock().unwrap().clone();
        assert_eq!(residuals.len(), 1);
        assert_eq!(residuals[0][0], 0, "a residual update carries no structs");
        let residual = Update::decode_v1(&residuals[0]).unwrap();
        assert!(residual.delete_set() == &*expected.lock().unwrap());
        for update in [&ours, &theirs, &residuals[0]] {
            let update = Update::decode_v1(update).unwrap();
            behind.transact_mut().apply_update(update).unwrap();
        }
        let encoded = |doc: &Doc| {
            doc.transact()
                .encode_state_as_update_v1(&StateVector::default())
        };
        assert_eq!(encoded(&behind), encoded(&peer));
    }

    #[test]
    fn staged_formatting_cleanup_is_republished() {
        let origin = replica(7);
        let state = edit(&origin, |txn, text| text.insert(txn, 0, "abcd"));
        let peer = replica(8);
        peer.transact_mut()
            .apply_update(Update::decode_v1(&state).unwrap())
            .unwrap();
        let bold = || Attrs::from([(Arc::from("b"), Any::Bool(true))]);
        let theirs = edit(&origin, |txn, text| text.format(txn, 0, 2, bold()));
        edit(&peer, |txn, text| text.format(txn, 0, 2, bold()));
        let staged = replica(9);
        let state = peer
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        staged
            .transact_mut()
            .apply_update(Update::decode_v1(&state).unwrap())
            .unwrap();
        let carried = apply_staged_update(
            &mut staged.transact_mut(),
            Update::decode_v1(&theirs).unwrap(),
        )
        .unwrap();
        assert!(carried.is_empty());
        let diff = staged
            .transact()
            .encode_state_as_update_v1(&peer.transact().state_vector());
        let diff = Update::decode_v1(&diff).unwrap();
        let expected = diff.delete_set().clone();
        assert!(!expected.is_empty());
        let deletions = RemoteDeletions::default();
        let (_subscription, residuals) = published(&peer, &deletions);
        let mut integration = deletions.integrate(peer.transact_mut());
        integration.carry(&carried);
        integration.apply_update(diff).unwrap();
        drop(integration);
        let residuals = residuals.lock().unwrap();
        assert_eq!(residuals.len(), 1);
        assert_eq!(residuals[0][0], 0);
        assert!(Update::decode_v1(&residuals[0]).unwrap().delete_set() == &expected);
    }

    #[test]
    fn a_deletion_that_arrives_before_its_struct_is_not_republished() {
        let origin = replica(7);
        let insert = edit(&origin, |txn, text| text.insert(txn, 0, "abcd"));
        let delete = edit(&origin, |txn, text| text.remove_range(txn, 1, 1));
        let peer = replica(8);
        let deletions = RemoteDeletions::default();
        let (_subscription, residuals) = published(&peer, &deletions);

        integrate(&peer, &deletions, &delete);
        assert!(ReadTxn::store(&peer.transact()).pending_ds().is_some());
        integrate(&peer, &deletions, &insert);

        assert_eq!(text(&peer), "acd");
        assert!(residuals.lock().unwrap().is_empty());
    }

    #[test]
    fn deletions_held_with_a_pending_struct_are_not_republished() {
        let origin = replica(7);
        let first = edit(&origin, |txn, text| text.insert(txn, 0, "ab"));
        let second = edit(&origin, |txn, text| {
            text.insert(txn, 2, "cd");
            text.remove_range(txn, 2, 1);
        });
        let peer = replica(8);
        let deletions = RemoteDeletions::default();
        let (_subscription, residuals) = published(&peer, &deletions);

        integrate(&peer, &deletions, &second);
        assert!(ReadTxn::store(&peer.transact()).pending_update().is_some());
        integrate(&peer, &deletions, &first);

        assert_eq!(text(&peer), "abd");
        assert!(residuals.lock().unwrap().is_empty());
    }
}
