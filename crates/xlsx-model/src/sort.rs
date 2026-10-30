use std::cmp::Ordering;

#[inline(never)]
fn sorted_indices(len: usize, compare: &mut dyn FnMut(usize, usize) -> Ordering) -> Vec<u32> {
    let len = u32::try_from(len).expect("shared sort length exceeds u32::MAX");
    let mut indices: Vec<u32> = (0..len).collect();
    indices.sort_by(|left, right| compare(*left as usize, *right as usize));
    indices
}

pub fn sort_by_shared<T>(values: &mut [T], compare: impl Fn(&T, &T) -> Ordering) {
    if values.len() < 2 || std::mem::size_of::<T>() == 0 {
        return;
    }
    let mut indices = sorted_indices(values.len(), &mut |left, right| {
        compare(&values[left], &values[right])
    });
    for start in 0..indices.len() {
        let mut current = start;
        while indices[current] as usize != start {
            let next = indices[current] as usize;
            values.swap(current, next);
            indices[current] = current as u32;
            current = next;
        }
        indices[current] = current as u32;
    }
}

#[cfg(test)]
mod tests {
    use super::sort_by_shared;

    #[test]
    fn empty() {
        let mut values: Vec<String> = Vec::new();
        sort_by_shared(&mut values, |_, _| panic!("empty comparator"));
        assert!(values.is_empty());
    }

    #[test]
    fn single() {
        let mut values = vec![String::from("only")];
        sort_by_shared(&mut values, |_, _| panic!("single comparator"));
        assert_eq!(values, ["only"]);
    }

    #[test]
    fn stable() {
        let mut values = vec![(2, 'a'), (1, 'b'), (2, 'c'), (1, 'd'), (2, 'e')];
        sort_by_shared(&mut values, |left, right| left.0.cmp(&right.0));
        assert_eq!(values, [(1, 'b'), (1, 'd'), (2, 'a'), (2, 'c'), (2, 'e')]);
        sort_by_shared(&mut values, |_, _| std::cmp::Ordering::Equal);
        assert_eq!(values, [(1, 'b'), (1, 'd'), (2, 'a'), (2, 'c'), (2, 'e')]);
    }

    #[test]
    fn large_random_matches_standard_sort() {
        let ascending = |left: &(u64, usize), right: &(u64, usize)| left.0.cmp(&right.0);
        let descending = |left: &(u64, usize), right: &(u64, usize)| ascending(right, left);
        let mut seed = 0x6a09_e667_f3bc_c909_u64;
        let input: Vec<_> = (0..16_384_usize)
            .map(|index| {
                seed = seed.wrapping_mul(6_364_136_223_846_793_005).wrapping_add(1);
                ((seed >> 32) % 127, index)
            })
            .collect();
        for len in [2, 3, 8, 64, 257, input.len()] {
            let mut actual = input[..len].to_vec();
            let mut expected = actual.clone();
            expected.sort_by(ascending);
            sort_by_shared(&mut actual, ascending);
            assert_eq!(actual, expected);
            expected.sort_by(descending);
            sort_by_shared(&mut actual, descending);
            assert_eq!(actual, expected);
        }
    }

    #[test]
    fn permutes_non_copy_values() {
        #[derive(Debug, PartialEq)]
        struct Entry(u32, String);

        let mut values: Vec<_> = [3, 0, 4, 1, 2, 5]
            .into_iter()
            .map(|key| Entry(key, format!("entry {key}")))
            .collect();
        let pointers: Vec<_> = (0..6)
            .map(|key| {
                values
                    .iter()
                    .find(|entry| entry.0 == key)
                    .unwrap()
                    .1
                    .as_ptr()
            })
            .collect();
        sort_by_shared(&mut values, |left, right| left.0.cmp(&right.0));
        for (index, entry) in values.iter().enumerate() {
            assert_eq!(entry.0, index as u32);
            assert_eq!(entry.1, format!("entry {index}"));
            assert_eq!(entry.1.as_ptr(), pointers[index]);
        }
    }

    #[test]
    fn comparator_panic_leaves_values_untouched() {
        let mut values = vec![String::from("c"), String::from("a"), String::from("b")];
        let result = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            sort_by_shared(&mut values, |_, _| panic!("comparison failed"));
        }));
        assert!(result.is_err());
        assert_eq!(values, ["c", "a", "b"]);
    }
}
