use std::collections::{HashMap, HashSet};
use std::hash::Hash;
use std::mem::size_of;

pub(super) struct MapGrowth<K, V> {
    pending: Option<MapMove<K, V>>,
}

struct MapMove<K, V> {
    entries: std::collections::hash_map::IntoIter<K, V>,
    next: Option<(K, V)>,
    storage: HashMap<K, V>,
}

impl<K, V> Default for MapGrowth<K, V> {
    fn default() -> Self {
        Self { pending: None }
    }
}

impl<K: Eq + Hash, V> MapGrowth<K, V> {
    pub(super) fn ensure(
        &mut self,
        target: &mut HashMap<K, V>,
        key_bytes: impl Fn(&K) -> usize,
        max_bytes: usize,
    ) -> Result<Option<usize>, String> {
        if self.pending.is_none() {
            if target.len() < target.capacity() {
                return Ok(None);
            }
            let mut storage = HashMap::new();
            storage
                .try_reserve(target.len().saturating_mul(2).max(1))
                .map_err(|_| "cannot allocate snapshot graph storage".to_owned())?;
            if target.is_empty() {
                *target = storage;
                return Ok(None);
            }
            self.pending = Some(MapMove {
                next: None,
                entries: std::mem::take(target).into_iter(),
                storage,
            });
        }
        let pending = self.pending.as_mut().unwrap();
        if pending.next.is_none() {
            let bytes = super::snapshot::admit(size_of::<(K, V)>(), max_bytes)?;
            pending.next = pending.entries.next();
            return Ok(Some(bytes));
        }
        let (key, _) = pending.next.as_ref().unwrap();
        let bytes = super::snapshot::admit(
            size_of::<(K, V)>().saturating_add(key_bytes(key)),
            max_bytes,
        )?;
        let (key, value) = pending.next.take().unwrap();
        pending.storage.insert(key, value);
        if pending.entries.len() == 0 {
            *target = self.pending.take().unwrap().storage;
        }
        Ok(Some(bytes))
    }
}

pub(super) struct SetGrowth<K> {
    pending: Option<SetMove<K>>,
}

struct SetMove<K> {
    entries: std::collections::hash_set::IntoIter<K>,
    storage: HashSet<K>,
}

impl<K> Default for SetGrowth<K> {
    fn default() -> Self {
        Self { pending: None }
    }
}

impl<K: Eq + Hash> SetGrowth<K> {
    pub(super) fn ensure(
        &mut self,
        target: &mut HashSet<K>,
        max_bytes: usize,
    ) -> Result<Option<usize>, String> {
        if self.pending.is_none() {
            if target.len() < target.capacity() {
                return Ok(None);
            }
            let mut storage = HashSet::new();
            storage
                .try_reserve(target.len().saturating_mul(2).max(1))
                .map_err(|_| "cannot allocate snapshot graph storage".to_owned())?;
            if target.is_empty() {
                *target = storage;
                return Ok(None);
            }
            self.pending = Some(SetMove {
                entries: std::mem::take(target).into_iter(),
                storage,
            });
        }
        let bytes = super::snapshot::admit(size_of::<K>().saturating_mul(2), max_bytes)?;
        let pending = self.pending.as_mut().unwrap();
        pending.storage.insert(pending.entries.next().unwrap());
        if pending.entries.len() == 0 {
            *target = self.pending.take().unwrap().storage;
        }
        Ok(Some(bytes))
    }
}
