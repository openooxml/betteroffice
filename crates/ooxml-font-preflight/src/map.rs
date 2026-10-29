#[derive(Clone, Default)]
pub(crate) struct Map<V>(Vec<(String, V)>);

impl<V> Map<V> {
    pub(crate) fn new() -> Self {
        Self(Vec::new())
    }

    fn index(&self, key: &str) -> Result<usize, usize> {
        self.0.binary_search_by(|(name, _)| name.as_str().cmp(key))
    }

    pub(crate) fn get(&self, key: &str) -> Option<&V> {
        let index = self.index(key).ok()?;
        Some(&self.0[index].1)
    }

    pub(crate) fn get_mut(&mut self, key: &str) -> Option<&mut V> {
        let index = self.index(key).ok()?;
        Some(&mut self.0[index].1)
    }

    pub(crate) fn contains_key(&self, key: &str) -> bool {
        self.index(key).is_ok()
    }

    pub(crate) fn insert(&mut self, key: String, value: V) {
        match self.index(&key) {
            Ok(index) => self.0[index].1 = value,
            Err(index) => self.0.insert(index, (key, value)),
        }
    }

    pub(crate) fn iter(&self) -> impl Iterator<Item = &(String, V)> {
        self.0.iter()
    }

    pub(crate) fn keys(&self) -> impl Iterator<Item = &String> {
        self.0.iter().map(|(key, _)| key)
    }
}
