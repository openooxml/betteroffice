//! The base of the next edit batch, prepared ahead of it: the deterministic encoding of the
//! committed state, and a replica decoded from it that the batch's rehearsal starts from.

use std::sync::Arc;
use std::sync::atomic::Ordering;

use yrs::{ReadTxn, StateVector, Transact};

use crate::{EditResult, EditingDoc, deterministic};

/// The committed state a prepared base encodes: every mutable transaction since moves it.
#[derive(Clone, PartialEq)]
pub(crate) struct StagingKey {
    nonce: u64,
    epoch: u64,
    transactions: u64,
    state_vector: StateVector,
}

pub(crate) struct StagingBase {
    key: StagingKey,
    bytes: Arc<Vec<u8>>,
    replica: Option<Box<EditingDoc>>,
}

impl EditingDoc {
    /// The key of the state `txn` reads; `None` while updates wait to integrate.
    fn staging_key<T: ReadTxn>(&self, txn: &T) -> Option<StagingKey> {
        let store = txn.store();
        if store.pending_update().is_some() || store.pending_ds().is_some() {
            return None;
        }
        Some(StagingKey {
            nonce: self.version_nonce.load(Ordering::Relaxed),
            epoch: self.epoch.load(Ordering::Relaxed),
            transactions: self.transactions.load(Ordering::Relaxed),
            state_vector: txn.state_vector(),
        })
    }

    /// Encodes the committed state for the next edit batch when it takes at most `max_bytes`.
    /// Returns whether a base of the current state is held. Opt-in: the base stays resident until
    /// a batch uses it, the document changes and a batch drops it, or
    /// [`Self::clear_staging_base`].
    pub fn prepare_staging_base_bytes(&self, max_bytes: usize) -> bool {
        let txn = self.yrs_doc().transact();
        let Some(key) = self.staging_key(&txn) else {
            self.clear_staging_base();
            return false;
        };
        {
            let mut held = self.staging_base.lock().unwrap();
            if held.as_ref().is_some_and(|base| base.key == key) {
                return true;
            }
            *held = None;
        }
        let bytes = deterministic::encode_state_as_update_v1(&txn, &StateVector::default());
        if bytes.len() > max_bytes {
            return false;
        }
        *self.staging_base.lock().unwrap() = Some(StagingBase {
            key,
            bytes: Arc::new(bytes),
            replica: None,
        });
        true
    }

    /// Decodes the bytes [`Self::prepare_staging_base_bytes`] holds into the replica the next
    /// batch rehearses on. Returns whether a replica of the current state is held.
    pub fn prepare_staging_base_replica(&self) -> EditResult<bool> {
        let (key, bytes) = {
            let txn = self.yrs_doc().transact();
            let key = self.staging_key(&txn);
            let mut held = self.staging_base.lock().unwrap();
            match (held.as_ref(), key) {
                (Some(base), Some(key)) if base.key == key => {
                    if base.replica.is_some() {
                        return Ok(true);
                    }
                    (key, Arc::clone(&base.bytes))
                }
                _ => {
                    *held = None;
                    return Ok(false);
                }
            }
        };
        let replica = EditingDoc::new(self.client_id);
        replica.apply_verbatim_v1(&bytes)?;
        let txn = self.yrs_doc().transact();
        let current = self.staging_key(&txn);
        let mut held = self.staging_base.lock().unwrap();
        match held.as_mut() {
            Some(base) if base.key == key && current.as_ref() == Some(&key) => {
                base.replica = Some(Box::new(replica));
                Ok(true)
            }
            _ => Ok(false),
        }
    }

    /// Drops a prepared staging base.
    pub fn clear_staging_base(&self) {
        let base = self.staging_base.lock().unwrap().take();
        drop(base);
    }

    /// Whether a prepared base and its replica match the committed state.
    #[doc(hidden)]
    pub fn staging_base_ready(&self) -> bool {
        let txn = self.yrs_doc().transact();
        let key = self.staging_key(&txn);
        self.staging_base
            .lock()
            .unwrap()
            .as_ref()
            .is_some_and(|base| base.replica.is_some() && key.as_ref() == Some(&base.key))
    }

    /// The prepared bytes when they encode exactly the state `txn` reads; a prepared base of
    /// any other state is dropped.
    pub(crate) fn prepared_staging_bytes<T: ReadTxn>(&self, txn: &T) -> Option<Arc<Vec<u8>>> {
        let mut held = self.staging_base.lock().unwrap();
        let base = held.as_ref()?;
        if self.staging_key(txn).as_ref() == Some(&base.key) {
            return Some(Arc::clone(&base.bytes));
        }
        *held = None;
        None
    }

    /// Takes the prepared base out, returning its replica when one was decoded.
    pub(crate) fn take_staging_replica(&self) -> Option<EditingDoc> {
        let base = self.staging_base.lock().unwrap().take()?;
        base.replica.map(|replica| *replica)
    }
}

#[cfg(test)]
mod tests {
    use crate::EditingDoc;

    #[test]
    fn version_rotation_retires_the_prepared_staging_base() {
        let doc = EditingDoc::new(7001);
        doc.create_story("body", "text", "Normal", "left").unwrap();
        assert!(doc.prepare_staging_base_bytes(usize::MAX));
        assert!(doc.prepare_staging_base_replica().unwrap());
        assert!(doc.staging_base_ready());
        let version = doc.version();
        let state = doc.encode_state_as_update_v1();
        doc.rotate_version(0);
        assert_ne!(doc.version(), version);
        assert!(doc.staging_base.lock().unwrap().is_none());
        assert!(!doc.staging_base_ready());
        assert!(!doc.prepare_staging_base_replica().unwrap());
        assert_eq!(doc.encode_state_as_update_v1(), state);
    }
}
