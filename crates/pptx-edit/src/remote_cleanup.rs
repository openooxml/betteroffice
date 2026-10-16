//! yrs cleans up redundant formatting inside the transaction that integrates a remote update,
//! so those deletions ride in an update observers label remote, which no peer is sent. They
//! depend on what this replica held at the time, so peers may never make them on their own.

use std::sync::{Arc, Mutex};

use yrs::encoding::write::Write;
use yrs::updates::encoder::{Encode, Encoder, EncoderV1};
use yrs::{IdSet, TransactionMut};

/// The delete set of the remote update being integrated, held only while it commits.
#[derive(Clone, Default)]
pub(crate) struct RemoteDeletions(Arc<Mutex<Option<IdSet>>>);

impl RemoteDeletions {
    /// Holds `carried` as the integrating update's deletions until the guard drops.
    pub(crate) fn carry(&self, carried: &IdSet) -> CarriedDeletions<'_> {
        *self.0.lock().unwrap() = Some(carried.clone());
        CarriedDeletions(self)
    }

    /// What `txn` deleted beyond the carried deletions, as a delete-only v1 update: `None`
    /// outside a remote integration or when it deleted nothing else.
    pub(crate) fn residual_update_v1(&self, txn: &TransactionMut) -> Option<Vec<u8>> {
        let carried = self.0.lock().unwrap();
        let residual = txn.delete_set().diff(carried.as_ref()?);
        (!residual.is_empty()).then(|| delete_only_update_v1(&residual))
    }
}

pub(crate) struct CarriedDeletions<'a>(&'a RemoteDeletions);

impl Drop for CarriedDeletions<'_> {
    fn drop(&mut self) {
        *self.0.0.lock().unwrap() = None;
    }
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
    use yrs::updates::decoder::Decode;
    use yrs::{
        ClientID, Doc, GetString, ID, Options, ReadTxn, StateVector, Text, Transact, Update,
    };

    #[test]
    fn residual_update_deletes_exactly_the_residual() {
        let doc = Doc::with_options(Options::with_client_id(ClientID::new(7)));
        let text = doc.get_or_insert_text("t");
        text.insert(&mut doc.transact_mut(), 0, "abcd");
        let peer = Doc::with_options(Options::with_client_id(ClientID::new(8)));
        let peer_text = peer.get_or_insert_text("t");
        let state = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let state = Update::decode_v1(&state).unwrap();
        peer.transact_mut().apply_update(state).unwrap();

        let deletions = RemoteDeletions::default();
        let mut carried = IdSet::new();
        carried.insert(ID::new(ClientID::new(7), 0), 1);
        let residual = {
            let _carried = deletions.carry(&carried);
            let mut txn = doc.transact_mut();
            text.remove_range(&mut txn, 0, 2);
            deletions.residual_update_v1(&txn).unwrap()
        };
        assert!(deletions.residual_update_v1(&doc.transact_mut()).is_none());

        let residual = Update::decode_v1(&residual).unwrap();
        peer.transact_mut().apply_update(residual).unwrap();
        assert_eq!(peer_text.get_string(&peer.transact()), "acd");
    }
}
