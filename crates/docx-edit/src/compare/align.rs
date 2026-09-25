//! One-to-one alignment of body paragraphs between unchanged structural boundaries: unique texts
//! anchor it, crossing anchors are moves, and larger gaps pair only on a unique best match.

use std::collections::HashMap;
use std::ops::Range;

use unicode_segmentation::UnicodeSegmentation;

use super::options::MAX_GAP_PARAGRAPHS;

/// Work the alignment and diff may still do.
pub(crate) struct Budget {
    pub alignment_cells: usize,
    pub diff_cells: usize,
}

/// Which budget ran out.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub(crate) enum Exhausted {
    Alignment,
    Diff,
    Gap,
}

impl Budget {
    pub fn charge_alignment(&mut self, cells: usize) -> Result<(), Exhausted> {
        self.alignment_cells = self
            .alignment_cells
            .checked_sub(cells)
            .ok_or(Exhausted::Alignment)?;
        Ok(())
    }

    pub fn charge_diff(&mut self, cells: usize) -> Result<(), Exhausted> {
        self.diff_cells = self.diff_cells.checked_sub(cells).ok_or(Exhausted::Diff)?;
        Ok(())
    }
}

/// Why paragraphs could not be paired. Indices are positions in the aligned slices.
#[derive(Clone, Debug, Eq, PartialEq)]
pub(crate) enum Issue {
    /// Unique matching texts in a different order.
    Moved { original: usize, revised: usize },
    /// More paragraphs on one side of an unmatched stretch than the other.
    Count {
        original: Range<usize>,
        revised: Range<usize>,
    },
    /// Several paragraphs on each side without evidence for one correspondence.
    Ambiguous {
        original: Range<usize>,
        revised: Range<usize>,
    },
}

#[derive(Debug, Default, Eq, PartialEq)]
pub(crate) struct Alignment {
    pub pairs: Vec<(usize, usize)>,
    pub issues: Vec<Issue>,
}

fn key(text: &str) -> String {
    text.split_whitespace().collect::<Vec<_>>().join(" ")
}

/// Aligns `original` with `revised` paragraph texts.
pub(crate) fn align(
    original: &[&str],
    revised: &[&str],
    budget: &mut Budget,
) -> Result<Alignment, Exhausted> {
    let keys_a: Vec<String> = original.iter().map(|text| key(text)).collect();
    let keys_b: Vec<String> = revised.iter().map(|text| key(text)).collect();
    let mut counts: HashMap<&str, (usize, usize, usize, usize)> = HashMap::new();
    for (index, key) in keys_a.iter().enumerate() {
        let entry = counts.entry(key).or_default();
        entry.0 += 1;
        entry.2 = index;
    }
    for (index, key) in keys_b.iter().enumerate() {
        let entry = counts.entry(key).or_default();
        entry.1 += 1;
        entry.3 = index;
    }
    let mut anchors: Vec<(usize, usize)> = counts
        .values()
        .filter(|(in_a, in_b, _, _)| *in_a == 1 && *in_b == 1)
        .map(|(_, _, i, j)| (*i, *j))
        .collect();
    anchors.sort_unstable();
    let kept = increasing_by_revised(&anchors);
    let mut alignment = Alignment::default();
    let mut moved_a = vec![false; original.len()];
    let mut moved_b = vec![false; revised.len()];
    for (index, &(i, j)) in anchors.iter().enumerate() {
        if !kept.contains(&index) {
            alignment.issues.push(Issue::Moved {
                original: i,
                revised: j,
            });
            moved_a[i] = true;
            moved_b[j] = true;
        }
    }
    let (mut next_a, mut next_b) = (0, 0);
    let boundaries = kept
        .iter()
        .map(|index| Some(anchors[*index]))
        .chain(std::iter::once(None));
    for boundary in boundaries {
        let (end_a, end_b) = boundary.unwrap_or((original.len(), revised.len()));
        let gap_a: Vec<usize> = (next_a..end_a).filter(|i| !moved_a[*i]).collect();
        let gap_b: Vec<usize> = (next_b..end_b).filter(|j| !moved_b[*j]).collect();
        resolve_gap(
            &gap_a,
            &gap_b,
            (&keys_a, &keys_b),
            (original, revised),
            budget,
            &mut alignment,
        )?;
        if let Some(anchor) = boundary {
            alignment.pairs.push(anchor);
            (next_a, next_b) = (anchor.0 + 1, anchor.1 + 1);
        }
    }
    alignment.pairs.sort_unstable();
    Ok(alignment)
}

/// Positions in `anchors` (sorted by original index) of a longest subsequence whose revised
/// indices increase.
fn increasing_by_revised(anchors: &[(usize, usize)]) -> Vec<usize> {
    let mut tails: Vec<usize> = Vec::new();
    let mut previous = vec![usize::MAX; anchors.len()];
    for (index, &(_, j)) in anchors.iter().enumerate() {
        let position = tails.partition_point(|&tail| anchors[tail].1 < j);
        if position > 0 {
            previous[index] = tails[position - 1];
        }
        if position == tails.len() {
            tails.push(index);
        } else {
            tails[position] = index;
        }
    }
    let mut kept = Vec::with_capacity(tails.len());
    let mut cursor = tails.last().copied();
    while let Some(index) = cursor {
        kept.push(index);
        cursor = (previous[index] != usize::MAX).then_some(previous[index]);
    }
    kept.reverse();
    kept
}

fn words(text: &str) -> HashMap<&str, usize> {
    let mut counts = HashMap::new();
    for word in text.unicode_words() {
        *counts.entry(word).or_default() += 1;
    }
    counts
}

fn shared_words(left: &HashMap<&str, usize>, right: &HashMap<&str, usize>) -> usize {
    left.iter()
        .map(|(word, count)| (*count).min(right.get(word).copied().unwrap_or(0)))
        .sum()
}

fn span(indices: &[usize]) -> Range<usize> {
    match (indices.first(), indices.last()) {
        (Some(first), Some(last)) => *first..*last + 1,
        _ => 0..0,
    }
}

fn resolve_gap(
    gap_a: &[usize],
    gap_b: &[usize],
    (keys_a, keys_b): (&[String], &[String]),
    (texts_a, texts_b): (&[&str], &[&str]),
    budget: &mut Budget,
    alignment: &mut Alignment,
) -> Result<(), Exhausted> {
    let prefix = gap_a
        .iter()
        .zip(gap_b)
        .take_while(|(i, j)| keys_a[**i] == keys_b[**j])
        .count();
    let suffix = gap_a[prefix..]
        .iter()
        .rev()
        .zip(gap_b[prefix..].iter().rev())
        .take_while(|(i, j)| keys_a[**i] == keys_b[**j])
        .count();
    for (i, j) in gap_a[..prefix].iter().zip(&gap_b[..prefix]) {
        alignment.pairs.push((*i, *j));
    }
    for (i, j) in gap_a[gap_a.len() - suffix..]
        .iter()
        .zip(&gap_b[gap_b.len() - suffix..])
    {
        alignment.pairs.push((*i, *j));
    }
    let core_a = &gap_a[prefix..gap_a.len() - suffix];
    let core_b = &gap_b[prefix..gap_b.len() - suffix];
    match (core_a.len(), core_b.len()) {
        (0, 0) => return Ok(()),
        (1, 1) => {
            alignment.pairs.push((core_a[0], core_b[0]));
            return Ok(());
        }
        (0, _) | (_, 0) => {
            alignment.issues.push(Issue::Count {
                original: span(core_a),
                revised: span(core_b),
            });
            return Ok(());
        }
        (n, m) if n > MAX_GAP_PARAGRAPHS || m > MAX_GAP_PARAGRAPHS => {
            return Err(Exhausted::Gap);
        }
        _ => {}
    }
    let (n, m) = (core_a.len(), core_b.len());
    let bags_a: Vec<_> = core_a.iter().map(|i| words(texts_a[*i])).collect();
    let bags_b: Vec<_> = core_b.iter().map(|j| words(texts_b[*j])).collect();
    let mut weights = vec![0usize; n * m];
    for (i, left) in bags_a.iter().enumerate() {
        for (j, right) in bags_b.iter().enumerate() {
            budget.charge_diff(left.len() + right.len() + 1)?;
            weights[i * m + j] = shared_words(left, right);
        }
    }
    budget.charge_alignment(2 * (n + 1) * (m + 1))?;
    let (matches, unique) = best_alignment(n, m, &weights);
    if !unique {
        alignment.issues.push(Issue::Ambiguous {
            original: span(core_a),
            revised: span(core_b),
        });
        return Ok(());
    }
    let (mut next_i, mut next_j) = (0, 0);
    for boundary in matches
        .iter()
        .copied()
        .map(Some)
        .chain(std::iter::once(None))
    {
        let (end_i, end_j) = boundary.unwrap_or((n, m));
        let sub_a = &core_a[next_i..end_i];
        let sub_b = &core_b[next_j..end_j];
        match (sub_a.len(), sub_b.len()) {
            (0, 0) => {}
            (1, 1) => alignment.pairs.push((sub_a[0], sub_b[0])),
            (x, y) if x == y => alignment.issues.push(Issue::Ambiguous {
                original: span(sub_a),
                revised: span(sub_b),
            }),
            _ => alignment.issues.push(Issue::Count {
                original: span(sub_a),
                revised: span(sub_b),
            }),
        }
        if let Some((i, j)) = boundary {
            alignment.pairs.push((core_a[i], core_b[j]));
            (next_i, next_j) = (i + 1, j + 1);
        }
    }
    Ok(())
}

/// The highest-weight monotone matching over pairs with positive weight, and whether it is the
/// only matching with that weight. Skips of the original side never follow skips of the revised
/// side, so each matching is one path.
fn best_alignment(n: usize, m: usize, weights: &[usize]) -> (Vec<(usize, usize)>, bool) {
    let width = m + 1;
    let at = |i: usize, j: usize, s: usize| (i * width + j) * 2 + s;
    let mut best = vec![0usize; (n + 1) * width * 2];
    let mut count = vec![0u8; (n + 1) * width * 2];
    for i in (0..=n).rev() {
        for j in (0..=m).rev() {
            for s in 0..2 {
                if i == n && j == m {
                    count[at(i, j, s)] = 1;
                    continue;
                }
                let mut options: Vec<(usize, u8)> = Vec::with_capacity(3);
                if i < n && j < m && weights[i * m + j] > 0 {
                    options.push((
                        weights[i * m + j] + best[at(i + 1, j + 1, 0)],
                        count[at(i + 1, j + 1, 0)],
                    ));
                }
                if i < n && s == 0 {
                    options.push((best[at(i + 1, j, 0)], count[at(i + 1, j, 0)]));
                }
                if j < m {
                    options.push((best[at(i, j + 1, 1)], count[at(i, j + 1, 1)]));
                }
                let top = options.iter().map(|(score, _)| *score).max().unwrap_or(0);
                best[at(i, j, s)] = top;
                count[at(i, j, s)] = options
                    .iter()
                    .filter(|(score, _)| *score == top)
                    .fold(0u8, |total, (_, ways)| total.saturating_add(*ways).min(2));
            }
        }
    }
    let mut matches = Vec::new();
    let (mut i, mut j, mut s) = (0, 0, 0);
    while i < n || j < m {
        let here = best[at(i, j, s)];
        if i < n
            && j < m
            && weights[i * m + j] > 0
            && weights[i * m + j] + best[at(i + 1, j + 1, 0)] == here
        {
            matches.push((i, j));
            (i, j, s) = (i + 1, j + 1, 0);
        } else if i < n && s == 0 && best[at(i + 1, j, 0)] == here {
            i += 1;
        } else {
            (j, s) = (j + 1, 1);
        }
    }
    (matches, count[at(0, 0, 0)] == 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn budget() -> Budget {
        Budget {
            alignment_cells: 250_000,
            diff_cells: 4_000_000,
        }
    }

    fn run(original: &[&str], revised: &[&str]) -> Alignment {
        align(original, revised, &mut budget()).unwrap()
    }

    #[test]
    fn unchanged_and_edited_paragraphs_pair_in_order() {
        let alignment = run(
            &["Title", "The quick fox", "", "Closing words"],
            &["Title", "The slow fox", "", "Closing words today"],
        );
        assert_eq!(alignment.pairs, vec![(0, 0), (1, 1), (2, 2), (3, 3)]);
        assert!(alignment.issues.is_empty());
    }

    #[test]
    fn one_paragraph_between_anchors_is_a_replacement_even_without_shared_words() {
        let alignment = run(&["A", "old text", "C"], &["A", "entirely new", "C"]);
        assert_eq!(alignment.pairs, vec![(0, 0), (1, 1), (2, 2)]);
        assert!(alignment.issues.is_empty());
    }

    #[test]
    fn reordered_unique_paragraphs_are_moves() {
        let alignment = run(&["one", "two", "three"], &["two", "one", "three"]);
        assert_eq!(
            alignment.issues,
            vec![Issue::Moved {
                original: 0,
                revised: 1
            }]
        );
    }

    #[test]
    fn added_or_removed_paragraphs_are_count_issues() {
        let added = run(&["A", "C"], &["A", "B", "C"]);
        assert_eq!(
            added.issues,
            vec![Issue::Count {
                original: 0..0,
                revised: 1..2
            }]
        );
        let removed = run(&["A", "", "C"], &["A", "C"]);
        assert_eq!(
            removed.issues,
            vec![Issue::Count {
                original: 1..2,
                revised: 0..0
            }]
        );
    }

    #[test]
    fn several_unrelated_replacements_are_ambiguous() {
        let alignment = run(
            &["A", "alpha beta", "gamma delta", "Z"],
            &["A", "one two", "three four", "Z"],
        );
        assert_eq!(
            alignment.issues,
            vec![Issue::Ambiguous {
                original: 1..3,
                revised: 1..3
            }]
        );
    }

    #[test]
    fn edited_neighbours_with_shared_words_pair_positionally() {
        let alignment = run(
            &["A", "the first clause here", "the second clause there", "Z"],
            &[
                "A",
                "the first clause now",
                "the second clause elsewhere",
                "Z",
            ],
        );
        assert!(alignment.issues.is_empty());
        assert_eq!(alignment.pairs, vec![(0, 0), (1, 1), (2, 2), (3, 3)]);
    }

    #[test]
    fn repeated_texts_pair_at_the_gap_ends() {
        let alignment = run(&["x", "x", "old"], &["x", "x", "new"]);
        assert_eq!(alignment.pairs, vec![(0, 0), (1, 1), (2, 2)]);
        assert!(alignment.issues.is_empty());
    }

    #[test]
    fn oversized_gaps_exhaust_the_budget() {
        let original: Vec<String> = (0..70).map(|index| format!("a {index} x")).collect();
        let revised: Vec<String> = (0..70).map(|index| format!("b {index} x")).collect();
        let original: Vec<&str> = original.iter().map(String::as_str).collect();
        let revised: Vec<&str> = revised.iter().map(String::as_str).collect();
        assert_eq!(
            align(&original, &revised, &mut budget()),
            Err(Exhausted::Gap)
        );
        let mut tight = Budget {
            alignment_cells: 10,
            diff_cells: 4_000_000,
        };
        assert_eq!(
            align(
                &["p", "a b", "c d", "q"],
                &["p", "a e", "c f", "q"],
                &mut tight
            ),
            Err(Exhausted::Alignment)
        );
    }
}
