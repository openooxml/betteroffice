use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use yrs::types::text::YChange;
use yrs::{Any, Array, ClientID, Map, Out, ReadTxn, Text};

/// Draws a replica's `{client_id}:{counter}` ids.
pub(crate) struct SessionIds {
    client_id: u64,
    counter: AtomicU64,
    floored: AtomicBool,
}

impl SessionIds {
    pub(crate) fn new(client_id: u64) -> Self {
        Self {
            client_id,
            counter: AtomicU64::new(0),
            floored: AtomicBool::new(false),
        }
    }

    /// The next id. The first draw after construction or [`Self::refloor`] first moves the
    /// counter past every `{client_id}:N` the document read through `txn` holds.
    pub(crate) fn draw<T: ReadTxn>(&self, txn: &T) -> String {
        if !self.floored.swap(true, Ordering::Relaxed) {
            self.counter.fetch_max(self.floor(txn), Ordering::Relaxed);
        }
        let counter = self
            .counter
            .fetch_update(Ordering::Relaxed, Ordering::Relaxed, |n| n.checked_add(1))
            .expect("session id counter exhausted");
        format!("{}:{counter}", self.client_id)
    }

    /// Makes the next draw floor the counter again.
    pub(crate) fn refloor(&self) {
        self.floored.store(false, Ordering::Relaxed);
    }

    /// Continues from `other`'s counter and floor.
    pub(crate) fn adopt(&self, other: &SessionIds) {
        self.counter
            .store(other.counter.load(Ordering::Relaxed), Ordering::Relaxed);
        self.floored
            .store(other.floored.load(Ordering::Relaxed), Ordering::Relaxed);
    }

    #[cfg(test)]
    pub(crate) fn drawn(&self) -> u64 {
        self.counter.load(Ordering::Relaxed)
    }

    /// One past the highest `N` of a `{client_id}:N` in a live map key, text attribute or value;
    /// 0 without reading when this client has written nothing to the document.
    fn floor<T: ReadTxn>(&self, txn: &T) -> u64 {
        if txn.state_vector().get(&ClientID::new(self.client_id)) == 0 {
            return 0;
        }
        let prefix = format!("{}:", self.client_id);
        let mut floor = 0;
        let mut pending: Vec<Out> = txn.root_refs().map(|(_, value)| value).collect();
        while let Some(value) = pending.pop() {
            match value {
                Out::Any(any) => note_any(&any, &prefix, &mut floor),
                Out::YMap(map) => {
                    for (key, value) in map.iter(txn) {
                        note(key, &prefix, &mut floor);
                        pending.push(value);
                    }
                }
                Out::YText(text) => {
                    for chunk in text.diff(txn, YChange::identity) {
                        if let Some(attrs) = &chunk.attributes {
                            for (key, value) in attrs.iter() {
                                note(key, &prefix, &mut floor);
                                note_any(value, &prefix, &mut floor);
                            }
                        }
                        if !matches!(chunk.insert, Out::Any(Any::String(_))) {
                            pending.push(chunk.insert);
                        }
                    }
                }
                Out::YArray(array) => pending.extend(array.iter(txn)),
                _ => {}
            }
        }
        floor
    }
}

fn note_any(any: &Any, prefix: &str, floor: &mut u64) {
    let mut pending = vec![any];
    while let Some(any) = pending.pop() {
        match any {
            Any::String(text) => note(text, prefix, floor),
            Any::Array(items) => pending.extend(items.iter()),
            Any::Map(entries) => {
                for (key, value) in entries.iter() {
                    note(key, prefix, floor);
                    pending.push(value);
                }
            }
            _ => {}
        }
    }
}

const MAX_FLOOR: u64 = 1 << 53;

/// Raises `floor` past each `N` below [`MAX_FLOOR`] in `text` written as `prefix` then digits,
/// where `prefix` does not continue a longer number.
fn note(text: &str, prefix: &str, floor: &mut u64) {
    for (at, _) in text.match_indices(prefix) {
        if text[..at].ends_with(|c: char| c.is_ascii_digit()) {
            continue;
        }
        let digits = &text[at + prefix.len()..];
        let end = digits
            .find(|c: char| !c.is_ascii_digit())
            .unwrap_or(digits.len());
        if let Ok(n) = digits[..end].parse::<u64>()
            && n < MAX_FLOOR
        {
            *floor = (*floor).max(n + 1);
        }
    }
}

#[cfg(test)]
mod tests {
    use std::sync::atomic::Ordering;

    use yrs::{Any, ClientID, Doc, Map, Options, Transact};

    use super::{SessionIds, note};

    fn floor(text: &str) -> u64 {
        let mut floor = 0;
        note(text, "41:", &mut floor);
        floor
    }

    #[test]
    fn only_this_clients_prefix_counts() {
        assert_eq!(floor("41:3"), 4);
        assert_eq!(floor("141:3"), 0);
        assert_eq!(floor("body:repeat:41:7"), 8);
        assert_eq!(floor("41:2:repeat:41:9"), 10);
        assert_eq!(floor("41:x"), 0);
        assert_eq!(floor("x41:5"), 6);
    }

    #[test]
    fn numbers_at_the_cap_never_raise_the_floor() {
        assert_eq!(floor("41:9007199254740991"), 1 << 53);
        assert_eq!(floor("41:9007199254740992"), 0);
        assert_eq!(floor("41:18446744073709551615"), 0);
    }

    fn drawn_after(body: &str) -> (String, String) {
        let doc = Doc::with_options(Options::with_client_id(ClientID::new(41)));
        let comments = doc.get_or_insert_map("comments");
        {
            let mut txn = doc.transact_mut();
            comments.insert(&mut txn, "41:0", Any::String(body.into()));
        }
        let ids = SessionIds::new(41);
        let txn = doc.transact();
        (ids.draw(&txn), ids.draw(&txn))
    }

    #[test]
    fn a_comment_body_never_makes_a_draw_repeat_a_live_id() {
        let (first, second) = drawn_after("41:18446744073709551615");
        assert_eq!((first.as_str(), second.as_str()), ("41:1", "41:2"));
        let (first, second) = drawn_after("41:9007199254740991");
        assert_eq!(
            (first.as_str(), second.as_str()),
            ("41:9007199254740992", "41:9007199254740993")
        );
        let (first, second) = drawn_after("41:5");
        assert_eq!((first.as_str(), second.as_str()), ("41:6", "41:7"));
    }

    #[test]
    #[should_panic(expected = "session id counter exhausted")]
    fn the_counter_fails_loudly_instead_of_wrapping() {
        let doc = Doc::with_options(Options::with_client_id(ClientID::new(41)));
        let ids = SessionIds::new(41);
        ids.floored.store(true, Ordering::Relaxed);
        ids.counter.store(u64::MAX, Ordering::Relaxed);
        ids.draw(&doc.transact());
    }
}
