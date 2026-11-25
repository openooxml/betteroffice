use std::collections::HashSet;
use std::ops::Range;
use std::sync::Mutex;
use std::sync::atomic::{AtomicBool, AtomicU64, Ordering};

use yrs::types::text::YChange;
use yrs::{Any, Array, ClientID, Map, Out, ReadTxn, Text};

/// Draws a replica's `{client_id}:{counter}` ids.
pub(crate) struct SessionIds {
    client_id: u64,
    counter: AtomicU64,
    floored: AtomicBool,
    unused: AtomicU64,
    drawn: Mutex<Vec<Range<u64>>>,
}

impl SessionIds {
    pub(crate) fn new(client_id: u64) -> Self {
        Self {
            client_id,
            counter: AtomicU64::new(0),
            floored: AtomicBool::new(false),
            unused: AtomicU64::new(0),
            drawn: Mutex::new(Vec::new()),
        }
    }

    /// The next id. The first draw after construction or [`Self::refloor`] first moves the
    /// counter past every `{client_id}:N` the document read through `txn` holds.
    pub(crate) fn draw<T: ReadTxn>(&self, txn: &T) -> String {
        if !self.floored.swap(true, Ordering::Relaxed) {
            self.counter.fetch_max(self.floor(txn), Ordering::Relaxed);
        }
        let mut counter = self.counter.load(Ordering::Relaxed);
        loop {
            if counter >= MAX_FLOOR {
                return self.draw_unused(txn);
            }
            match self.counter.compare_exchange_weak(
                counter,
                counter + 1,
                Ordering::Relaxed,
                Ordering::Relaxed,
            ) {
                Ok(_) => break,
                Err(current) => counter = current,
            }
        }
        let mut drawn = self.drawn.lock().unwrap();
        match drawn.last_mut() {
            Some(run) if run.end == counter => run.end += 1,
            _ => drawn.push(counter..counter + 1),
        }
        drop(drawn);
        format!("{}:{counter}", self.client_id)
    }

    /// The smallest `{client_id}:N` from the last one drawn here that no live id uses and this
    /// session has not drawn, for when the floor has reached the cap.
    fn draw_unused<T: ReadTxn>(&self, txn: &T) -> String {
        let mut used = HashSet::new();
        self.scan(txn, &mut |n| {
            used.insert(n);
        });
        let drawn = self.drawn.lock().unwrap();
        let mut next = self.unused.load(Ordering::Relaxed);
        loop {
            if let Some(run) = drawn.iter().find(|run| run.contains(&next)) {
                next = run.end;
            } else if used.contains(&next) {
                next += 1;
            } else {
                break;
            }
        }
        drop(drawn);
        assert!(next < MAX_FLOOR, "session id counter exhausted");
        self.unused.store(next + 1, Ordering::Relaxed);
        format!("{}:{next}", self.client_id)
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
        self.unused
            .store(other.unused.load(Ordering::Relaxed), Ordering::Relaxed);
        let drawn = other.drawn.lock().unwrap().clone();
        *self.drawn.lock().unwrap() = drawn;
    }

    #[cfg(test)]
    pub(crate) fn drawn(&self) -> u64 {
        self.counter.load(Ordering::Relaxed)
    }

    /// One past the highest `N` of a `{client_id}:N` in a live map key, text attribute or value;
    /// 0 without reading when this client has written nothing to the document.
    fn floor<T: ReadTxn>(&self, txn: &T) -> u64 {
        let mut floor = 0;
        self.scan(txn, &mut |n| floor = floor.max(n + 1));
        floor
    }

    /// Each `N` below [`MAX_FLOOR`] of a `{client_id}:N` in a live map key, text attribute or
    /// value; none without reading when this client has written nothing to the document.
    fn scan<T: ReadTxn>(&self, txn: &T, seen: &mut dyn FnMut(u64)) {
        if txn.state_vector().get(&ClientID::new(self.client_id)) == 0 {
            return;
        }
        let prefix = format!("{}:", self.client_id);
        let mut pending: Vec<Out> = txn.root_refs().map(|(_, value)| value).collect();
        while let Some(value) = pending.pop() {
            match value {
                Out::Any(any) => note_any(&any, &prefix, seen),
                Out::YMap(map) => {
                    for (key, value) in map.iter(txn) {
                        note(key, &prefix, seen);
                        pending.push(value);
                    }
                }
                Out::YText(text) => {
                    for chunk in text.diff(txn, YChange::identity) {
                        if let Some(attrs) = &chunk.attributes {
                            for (key, value) in attrs.iter() {
                                note(key, &prefix, seen);
                                note_any(value, &prefix, seen);
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
    }
}

fn note_any(any: &Any, prefix: &str, seen: &mut dyn FnMut(u64)) {
    let mut pending = vec![any];
    while let Some(any) = pending.pop() {
        match any {
            Any::String(text) => note(text, prefix, seen),
            Any::Array(items) => pending.extend(items.iter()),
            Any::Map(entries) => {
                for (key, value) in entries.iter() {
                    note(key, prefix, seen);
                    pending.push(value);
                }
            }
            _ => {}
        }
    }
}

const MAX_FLOOR: u64 = 1 << 53;

/// Passes each `N` below [`MAX_FLOOR`] in `text` written as `prefix` then digits to `seen`,
/// where `prefix` does not continue a longer number.
fn note(text: &str, prefix: &str, seen: &mut dyn FnMut(u64)) {
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
            seen(n);
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
        note(text, "41:", &mut |n| floor = floor.max(n + 1));
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
        let (first, second) = drawn_after("41:9007199254740989");
        assert_eq!(
            (first.as_str(), second.as_str()),
            ("41:9007199254740990", "41:9007199254740991")
        );
        assert_eq!(floor(&second), 1 << 53);
        let (first, second) = drawn_after("41:5");
        assert_eq!((first.as_str(), second.as_str()), ("41:6", "41:7"));
    }

    #[test]
    fn a_floor_at_the_cap_draws_the_smallest_unused_ids() {
        let (first, second) = drawn_after("41:9007199254740991 41:1");
        assert_eq!((first.as_str(), second.as_str()), ("41:2", "41:3"));
    }

    #[test]
    fn the_fallback_skips_ids_this_session_drew() {
        let doc = Doc::with_options(Options::with_client_id(ClientID::new(41)));
        let comments = doc.get_or_insert_map("comments");
        {
            let mut txn = doc.transact_mut();
            comments.insert(&mut txn, "41:0", Any::String("note".into()));
        }
        let ids = SessionIds::new(41);
        let txn = doc.transact();
        assert_eq!(ids.draw(&txn), "41:1");
        ids.counter.store(super::MAX_FLOOR, Ordering::Relaxed);
        assert_eq!(ids.draw(&txn), "41:2");
        assert_eq!(ids.draw(&txn), "41:3");
    }

    #[test]
    fn the_fallback_never_reissues_an_id_seen_live_this_session() {
        let doc = Doc::with_options(Options::with_client_id(ClientID::new(41)));
        let comments = doc.get_or_insert_map("comments");
        {
            let mut txn = doc.transact_mut();
            comments.insert(&mut txn, "41:1", Any::String("old".into()));
            comments.insert(&mut txn, "41:5", Any::String("41:9007199254740991".into()));
        }
        let ids = SessionIds::new(41);
        let first = ids.draw(&doc.transact());
        {
            let mut txn = doc.transact_mut();
            comments.remove(&mut txn, "41:1");
        }
        let txn = doc.transact();
        let later = [ids.draw(&txn), ids.draw(&txn)];
        assert_eq!(first, "41:0");
        assert!(!later.contains(&"41:1".to_owned()), "{later:?}");
        assert!(!later.contains(&first), "{later:?}");
    }

    #[test]
    fn the_counter_never_wraps() {
        let doc = Doc::with_options(Options::with_client_id(ClientID::new(41)));
        let ids = SessionIds::new(41);
        ids.floored.store(true, Ordering::Relaxed);
        ids.counter.store(u64::MAX, Ordering::Relaxed);
        assert_eq!(ids.draw(&doc.transact()), "41:0");
    }
}
