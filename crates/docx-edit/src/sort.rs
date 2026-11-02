//! Stable sorts sharing one instantiation of the slice sort.

use std::cmp::Ordering;

pub(crate) fn sort_by<T>(items: &mut [T], compare: &dyn Fn(&T, &T) -> Ordering) {
    let order = sorted_indices(items.len(), &|a, b| compare(&items[a], &items[b]));
    permute(items, &order);
}

pub(crate) fn sort_by_key<T, K: Ord>(items: &mut [T], key: &dyn Fn(&T) -> K) {
    sort_by(items, &|a, b| key(a).cmp(&key(b)));
}

pub(crate) fn sort<T: Ord>(items: &mut [T]) {
    sort_by(items, &|a, b| a.cmp(b));
}

fn sorted_indices(len: usize, compare: &dyn Fn(usize, usize) -> Ordering) -> Vec<usize> {
    let mut order: Vec<usize> = (0..len).collect();
    order.sort_by(|&a, &b| compare(a, b));
    order
}

/// Reorders `items` so position `i` holds what was at `order[i]`.
fn permute<T>(items: &mut [T], order: &[usize]) {
    let mut done = vec![false; order.len()];
    for start in 0..order.len() {
        let mut current = start;
        while !done[current] {
            done[current] = true;
            let next = order[current];
            if next == start {
                break;
            }
            items.swap(current, next);
            current = next;
        }
    }
}

#[cfg(test)]
mod tests {
    #[test]
    fn matches_the_slice_sort_including_ties() {
        let mut seed = 0x2545_f491_4f6c_dd1du64;
        for len in 0..200 {
            let items: Vec<(u8, usize)> = (0..len)
                .map(|index| {
                    seed ^= seed << 13;
                    seed ^= seed >> 7;
                    seed ^= seed << 17;
                    ((seed % 7) as u8, index)
                })
                .collect();
            let mut expected = items.clone();
            expected.sort_by_key(|item| item.0);
            let mut actual = items.clone();
            super::sort_by_key(&mut actual, &|item: &(u8, usize)| item.0);
            assert_eq!(actual, expected);
            let mut expected = items.clone();
            expected.sort();
            let mut actual = items;
            super::sort(&mut actual);
            assert_eq!(actual, expected);
        }
    }
}
