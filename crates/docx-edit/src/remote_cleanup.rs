//! yrs cleans up redundant formatting inside the transaction that integrates a remote update,
//! so those deletions ride in an update observers label remote, which no peer is sent. They
//! depend on what this replica held at the time, so peers may never make them on their own.

use std::sync::{Arc, Mutex};

use yrs::encoding::write::Write;
use yrs::error::UpdateError;
use yrs::updates::encoder::{Encode, Encoder, EncoderV1};
use yrs::{IdSet, ReadTxn, TransactionMut, Update};

/// Deletions an open remote integration must not republish: what its updates carried, the
/// deletions earlier remote integrations left pending, and what applying its updates deleted,
/// which every peer that integrates them deletes too. Deletions a local apply left pending are
/// not carried, so they are published once they apply.
#[derive(Clone, Default)]
pub(crate) struct RemoteDeletions(Arc<Mutex<Carried>>);

#[derive(Default)]
struct Carried {
    /// Set only while a remote integration is open.
    open: Option<IdSet>,
    /// Carried deletions yrs still held pending when their integration committed.
    pending: IdSet,
}

impl RemoteDeletions {
    /// Opens `txn` as a remote integration. Holding `txn` keeps every other transaction out
    /// until the integration has committed and cleared its carried set on drop.
    pub(crate) fn integrate<'doc>(&self, txn: TransactionMut<'doc>) -> RemoteIntegration<'_, 'doc> {
        let mut carried = self.0.lock().unwrap();
        let held = still_pending(&txn, &carried.pending);
        let local = ReadTxn::store(&txn)
            .pending_ds()
            .map_or_else(IdSet::new, |pending| pending.diff(&held));
        carried.open = Some(held);
        drop(carried);
        RemoteIntegration {
            deletions: self,
            txn,
            local,
        }
    }

    /// What `txn` deleted beyond the carried deletions, as a delete-only v1 update: `None`
    /// outside a remote integration or when it deleted nothing else.
    pub(crate) fn residual_update_v1(&self, txn: &TransactionMut) -> Option<Vec<u8>> {
        let carried = self.0.lock().unwrap();
        let residual = txn.delete_set().diff(carried.open.as_ref()?);
        (!residual.is_empty()).then(|| delete_only_update_v1(&residual))
    }
}

/// One remote transaction; any number of peer updates may be carried into it.
pub(crate) struct RemoteIntegration<'a, 'doc> {
    deletions: &'a RemoteDeletions,
    txn: TransactionMut<'doc>,
    /// Deletions a local apply left pending when the integration opened.
    local: IdSet,
}

impl RemoteIntegration<'_, '_> {
    /// Adds deletions a peer sent, so they are not republished as this replica's own.
    pub(crate) fn carry(&mut self, deletions: &IdSet) {
        if let Some(carried) = self.deletions.0.lock().unwrap().open.as_mut() {
            carried.merge_with(deletions.clone());
        }
    }

    /// Leaves only the formatting cleanup yrs runs at commit and the local deletions applying
    /// released to be republished.
    pub(crate) fn apply_update(&mut self, update: Update) -> Result<(), UpdateError> {
        self.txn.apply_update(update)?;
        let deleted = self.txn.delete_set().diff(&self.local);
        self.carry(&deleted);
        Ok(())
    }
}

impl Drop for RemoteIntegration<'_, '_> {
    fn drop(&mut self) {
        self.txn.commit();
        let mut carried = self.deletions.0.lock().unwrap();
        let open = carried.open.take().unwrap_or_default();
        carried.pending = still_pending(&self.txn, &open);
    }
}

/// The part of `deletions` that yrs still holds pending in `txn`'s document.
fn still_pending(txn: &TransactionMut, deletions: &IdSet) -> IdSet {
    ReadTxn::store(txn)
        .pending_ds()
        .map_or_else(IdSet::new, |pending| pending.intersect(deletions))
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
    use yrs::{
        Any, ClientID, Doc, GetString, ID, Map, Options, StateVector, Subscription, Text, Transact,
    };

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
        integrate(&peer, &deletions, &theirs);
        assert!(deletions.residual_update_v1(&peer.transact_mut()).is_none());

        let residuals = residuals.lock().unwrap().clone();
        assert_eq!(residuals.len(), 1);
        assert_eq!(residuals[0][0], 0, "a residual update carries no structs");
        let residual = Update::decode_v1(&residuals[0]).unwrap();
        let mut expected = IdSet::new();
        expected.insert(ID::new(ClientID::new(7), 4), 1);
        assert!(residual.delete_set() == &expected);
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

    #[test]
    fn a_held_local_deletion_is_published_beside_held_remote_ones() {
        let origin = replica(7);
        let insert = edit(&origin, |txn, text| text.insert(txn, 0, "abcd"));
        let worker = replica(8);
        let state = Update::decode_v1(&insert).unwrap();
        worker.transact_mut().apply_update(state).unwrap();
        let local = edit(&worker, |txn, text| text.remove_range(txn, 1, 1));
        let remote = edit(&origin, |txn, text| text.remove_range(txn, 3, 1));
        let main = replica(9);
        let deletions = RemoteDeletions::default();
        let (_subscription, residuals) = published(&main, &deletions);

        integrate(&main, &deletions, &remote);
        let local = Update::decode_v1(&local).unwrap();
        let expected = local.delete_set().clone();
        main.transact_mut_with("worker")
            .apply_update(local)
            .unwrap();
        integrate(&main, &deletions, &insert);

        assert_eq!(text(&main), "ac");
        let residuals = residuals.lock().unwrap().clone();
        assert_eq!(residuals.len(), 1);
        assert!(Update::decode_v1(&residuals[0]).unwrap().delete_set() == &expected);
    }

    #[test]
    fn a_released_local_deletion_is_published_without_the_entry_a_remote_value_replaced() {
        let origin = replica(7);
        let insert = edit(&origin, |txn, text| text.insert(txn, 0, "abcd"));
        let worker = replica(8);
        worker
            .transact_mut()
            .apply_update(Update::decode_v1(&insert).unwrap())
            .unwrap();
        let local =
            Update::decode_v1(&edit(&worker, |txn, text| text.remove_range(txn, 1, 1))).unwrap();
        let expected = local.delete_set().clone();
        let set = |doc: &Doc, value: &str| {
            let map = doc.get_or_insert_map("m");
            let mut txn = doc.transact_mut();
            map.insert(&mut txn, "k", value);
            txn.encode_update_v1()
        };
        let main = replica(6);
        let ours = set(&main, "ours");
        set(&origin, "theirs");
        let remote = origin
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let deletions = RemoteDeletions::default();
        let (_subscription, residuals) = published(&main, &deletions);

        main.transact_mut_with("worker")
            .apply_update(local)
            .unwrap();
        integrate(&main, &deletions, &remote);

        let residuals = residuals.lock().unwrap().clone();
        assert_eq!(residuals.len(), 1);
        assert!(Update::decode_v1(&residuals[0]).unwrap().delete_set() == &expected);
        let behind = replica(9);
        let value = |doc: &Doc| {
            let map = doc.get_or_insert_map("m");
            let txn = doc.transact();
            map.get(&txn, "k").map(|value| value.to_string(&txn))
        };
        for update in [&ours, &residuals[0]] {
            let update = Update::decode_v1(update).unwrap();
            behind.transact_mut().apply_update(update).unwrap();
        }
        assert_eq!(value(&behind).as_deref(), Some("ours"));
        behind
            .transact_mut()
            .apply_update(Update::decode_v1(&remote).unwrap())
            .unwrap();
        assert_eq!((text(&behind), value(&behind)), (text(&main), value(&main)));
        assert_eq!(text(&main), "acd");
    }

    #[test]
    fn a_batch_publishes_only_local_deletions_unblocked_by_its_second_update_once() {
        let origin = replica(7);
        let insert = edit(&origin, |txn, text| text.insert(txn, 0, "abcd"));
        let worker = replica(8);
        worker
            .transact_mut()
            .apply_update(Update::decode_v1(&insert).unwrap())
            .unwrap();
        let local = edit(&worker, |txn, text| text.remove_range(txn, 1, 1));
        let remote = edit(&origin, |txn, text| text.remove_range(txn, 3, 1));
        let main = crate::EditingDoc::new(9);
        let local = Update::decode_v1(&local).unwrap();
        let expected = local.delete_set().clone();
        main.integrate_update(local, crate::UpdateOrigin::Local)
            .unwrap();
        assert!(ReadTxn::store(&main.doc.transact()).pending_ds().is_some());

        let updates = Arc::new(Mutex::new(Vec::new()));
        let published = Arc::clone(&updates);
        let _subscription = main
            .observe_update_v1(move |update, remote| {
                published.lock().unwrap().push((update.to_vec(), remote));
            })
            .unwrap();
        let chunks = [remote, insert];
        main.apply_update_chunks_v1(&chunks).unwrap();

        assert_eq!(text(&main.doc), "ac");
        assert!(ReadTxn::store(&main.doc.transact()).pending_ds().is_none());
        {
            let updates = updates.lock().unwrap();
            assert_eq!(updates.len(), 2);
            assert!(updates[0].1);
            assert!(!updates[1].1);
            assert_eq!(updates[1].0[0], 0);
            let residual = Update::decode_v1(&updates[1].0).unwrap();
            assert!(residual.delete_set() == &expected);
            worker
                .transact_mut()
                .apply_update(Update::decode_v1(&chunks[0]).unwrap())
                .unwrap();
            worker.transact_mut().apply_update(residual).unwrap();
            assert_eq!(text(&worker), "ac");
        }

        main.apply_update_chunks_v1(&chunks).unwrap();
        assert_eq!(updates.lock().unwrap().len(), 2);
    }

    #[test]
    fn batched_and_sequential_remote_applies_forward_the_same_deletions() {
        for delete_first in [true, false] {
            let origin = replica(7);
            let insert = edit(&origin, |txn, text| text.insert(txn, 0, "abcd"));
            let worker = replica(8);
            worker
                .transact_mut()
                .apply_update(Update::decode_v1(&insert).unwrap())
                .unwrap();
            let local = edit(&worker, |txn, text| text.remove_range(txn, 1, 1));
            let remote = edit(&origin, |txn, text| text.remove_range(txn, 3, 1));
            let expected = Update::decode_v1(&local).unwrap().delete_set().clone();
            let batch = crate::EditingDoc::new(9);
            let sequential = crate::EditingDoc::new(10);
            let mut subscriptions = Vec::new();
            let mut publications = Vec::new();
            for doc in [&batch, &sequential] {
                doc.integrate_update(
                    Update::decode_v1(&local).unwrap(),
                    crate::UpdateOrigin::Local,
                )
                .unwrap();
                assert!(ReadTxn::store(&doc.doc.transact()).pending_ds() == Some(&expected));

                let updates = Arc::new(Mutex::new(Vec::new()));
                let published = Arc::clone(&updates);
                subscriptions.push(
                    doc.observe_update_v1(move |update, remote| {
                        if !remote {
                            published.lock().unwrap().push(update.to_vec());
                        }
                    })
                    .unwrap(),
                );
                publications.push(updates);
            }

            let chunks = if delete_first {
                [remote, insert]
            } else {
                [insert, remote]
            };
            batch.apply_update_chunks_v1(&chunks).unwrap();
            sequential.apply_update_v1(&chunks[0]).unwrap();
            if delete_first {
                let carried = Update::decode_v1(&chunks[0]).unwrap();
                let mut pending = expected.clone();
                pending.merge_with(carried.delete_set().clone());
                assert!(ReadTxn::store(&sequential.doc.transact()).pending_ds() == Some(&pending));
                assert!(publications[1].lock().unwrap().is_empty());
            }
            sequential.apply_update_v1(&chunks[1]).unwrap();

            assert_eq!(text(&batch.doc), "ac");
            assert_eq!(text(&batch.doc), text(&sequential.doc));
            let batch_txn = batch.doc.transact();
            let sequential_txn = sequential.doc.transact();
            let batch_pending = ReadTxn::store(&batch_txn).pending_ds();
            let sequential_pending = ReadTxn::store(&sequential_txn).pending_ds();
            assert!(batch_pending == sequential_pending);
            assert!(batch_pending.is_none());
            drop((batch_txn, sequential_txn));

            batch.apply_update_chunks_v1(&chunks).unwrap();
            for chunk in &chunks {
                sequential.apply_update_v1(chunk).unwrap();
            }

            let mut unions = Vec::new();
            for updates in publications {
                let updates = updates.lock().unwrap();
                let mut union = IdSet::new();
                for update in updates.iter() {
                    assert_eq!(update[0], 0, "a local publication carries no structs");
                    let update = Update::decode_v1(update).unwrap();
                    let deletions = update.delete_set();
                    assert!(!deletions.is_empty());
                    assert!(
                        union.intersect(deletions).is_empty(),
                        "a deletion was published twice"
                    );
                    union.merge_with(deletions.clone());
                }
                assert!(
                    union == expected,
                    "only the unblocked local deletion is published"
                );
                unions.push(union);
            }
            assert!(
                unions[0] == unions[1],
                "batch and sequential publications differ"
            );
        }
    }
}
