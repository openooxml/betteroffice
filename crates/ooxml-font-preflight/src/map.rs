use std::collections::BTreeMap;

#[derive(Clone, Default)]
pub(crate) struct Map<V>(BTreeMap<String, V>);

impl<V> Map<V> {
    pub(crate) fn new() -> Self {
        Self(BTreeMap::new())
    }

    pub(crate) fn get(&self, key: &str) -> Option<&V> {
        self.0.get(key)
    }

    pub(crate) fn get_mut(&mut self, key: &str) -> Option<&mut V> {
        self.0.get_mut(key)
    }

    pub(crate) fn contains_key(&self, key: &str) -> bool {
        self.0.contains_key(key)
    }

    pub(crate) fn insert(&mut self, key: String, value: V) {
        self.0.insert(key, value);
    }

    pub(crate) fn iter(&self) -> impl Iterator<Item = (&String, &V)> {
        self.0.iter()
    }

    pub(crate) fn keys(&self) -> impl Iterator<Item = &String> {
        self.0.keys()
    }
}
