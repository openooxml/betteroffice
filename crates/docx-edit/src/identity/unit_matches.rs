use std::collections::HashMap;
use std::hash::Hash;

use super::UnitState;

pub(super) fn row_splice_units<'a>(
    before: &'a [UnitState],
    after: &'a [UnitState],
) -> impl Iterator<Item = (&'a UnitState, &'a UnitState)> {
    matching_by_key(before, after, |unit| unit.key.as_deref())
}

fn matching_by_key<'a, T, K: Eq + Hash + 'a>(
    before: &'a [T],
    after: &'a [T],
    key: impl Fn(&'a T) -> K + Copy + 'a,
) -> impl Iterator<Item = (&'a T, &'a T)> {
    let indexed = (after.len() > 8 && !before.is_empty()).then(|| {
        let mut indexed = HashMap::with_capacity(after.len());
        for now in after {
            indexed.entry(key(now)).or_insert(now);
        }
        indexed
    });
    before.iter().filter_map(move |was| {
        let now = match &indexed {
            Some(indexed) => indexed.get(&key(was)).copied(),
            None => after.iter().find(|now| key(was) == key(now)),
        };
        now.map(|now| (was, now))
    })
}

#[cfg(test)]
mod tests {
    use std::cell::Cell;
    use std::hash::{Hash, Hasher};

    use super::matching_by_key;

    struct CountedKey<'a>(usize, &'a Cell<usize>);

    impl PartialEq for CountedKey<'_> {
        fn eq(&self, other: &Self) -> bool {
            self.1.set(self.1.get() + 1);
            self.0 == other.0
        }
    }

    impl Eq for CountedKey<'_> {}

    impl Hash for CountedKey<'_> {
        fn hash<H: Hasher>(&self, state: &mut H) {
            self.0.hash(state);
        }
    }

    #[test]
    fn no_edit_row_matching_has_a_linear_comparison_budget() {
        let comparisons = Cell::new(0);
        let paragraphs: Vec<_> = (0..1024).collect();
        let matched =
            matching_by_key(&paragraphs, &paragraphs, |id| CountedKey(*id, &comparisons)).count();
        assert_eq!(matched, paragraphs.len());
        assert!(
            comparisons.get() <= paragraphs.len() * 8,
            "{} key comparisons for {} unchanged paragraphs",
            comparisons.get(),
            paragraphs.len()
        );
    }

    #[test]
    fn matches_the_first_occurrence_and_keeps_source_order() {
        let before = [
            (Some("b"), 0),
            (None, 1),
            (Some("a"), 2),
            (Some("missing"), 3),
        ];
        let after = [
            (Some("a"), 4),
            (Some("b"), 5),
            (Some("b"), 6),
            (None, 7),
            (None, 8),
        ];
        for padding in [0, 32] {
            let mut after = after.to_vec();
            after.extend(std::iter::repeat_n((Some("unrelated"), 9), padding));
            let matched: Vec<_> = matching_by_key(&before, &after, |entry| entry.0)
                .map(|(was, now)| (was.1, now.1))
                .collect();
            assert_eq!(matched, [(0, 5), (1, 7), (2, 4)]);
        }
    }
}
