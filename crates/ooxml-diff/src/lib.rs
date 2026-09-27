//! Bounded longest-common-subsequence diff over token slices, returning index-range hunks.

use std::fmt;
use std::ops::Range;

#[derive(Clone, Copy, Debug, Eq, Hash, PartialEq)]
pub enum DiffKind {
    Equal,
    Delete,
    Insert,
}

/// What an over-budget diff does.
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum LimitFallback {
    /// Refuse with [`DiffLimitExceeded`].
    Error,
    /// Report the untrimmed middle as one deletion followed by one insertion.
    ReplaceMiddle,
}

/// One run of equal, deleted or inserted tokens. `old` is empty for insertions and `new` for
/// deletions; equal hunks cover the same number of tokens on both sides.
#[derive(Clone, Debug, Eq, PartialEq)]
pub struct DiffHunk {
    pub kind: DiffKind,
    pub old: Range<usize>,
    pub new: Range<usize>,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DiffLimits {
    /// Tokens either input may hold.
    pub max_tokens: usize,
    /// Cells of the `(old + 1) * (new + 1)` table left after trimming.
    pub max_lcs_cells: usize,
    pub fallback: LimitFallback,
}

impl DiffLimits {
    pub const fn new(max_tokens: usize, max_lcs_cells: usize, fallback: LimitFallback) -> Self {
        Self {
            max_tokens,
            max_lcs_cells,
            fallback,
        }
    }
}

#[derive(Clone, Debug, Eq, PartialEq)]
pub struct TokenDiff {
    /// Ordered hunks covering both inputs; adjacent hunks differ in kind.
    pub hunks: Vec<DiffHunk>,
    /// The middle was replaced wholesale under [`LimitFallback::ReplaceMiddle`].
    pub coarsened: bool,
    /// Table cells built, zero when trimming or the fallback left nothing to build.
    pub cells_used: usize,
}

/// A diff that would exceed its [`DiffLimits`] under [`LimitFallback::Error`].
#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub struct DiffLimitExceeded {
    /// Cells the table would need, saturated at `usize::MAX`.
    pub cells_required: usize,
}

impl fmt::Display for DiffLimitExceeded {
    fn fmt(&self, f: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(
            f,
            "the diff needs {} table cells, more than its limit",
            self.cells_required
        )
    }
}

impl std::error::Error for DiffLimitExceeded {}

struct Hunks(Vec<DiffHunk>);

impl Hunks {
    fn push(&mut self, kind: DiffKind, old: usize, new: usize) {
        let (old_len, new_len) = match kind {
            DiffKind::Equal => (1, 1),
            DiffKind::Delete => (1, 0),
            DiffKind::Insert => (0, 1),
        };
        self.push_run(kind, old..old + old_len, new..new + new_len);
    }

    fn push_run(&mut self, kind: DiffKind, old: Range<usize>, new: Range<usize>) {
        if old.is_empty() && new.is_empty() {
            return;
        }
        if let Some(last) = self.0.last_mut()
            && last.kind == kind
            && last.old.end == old.start
            && last.new.end == new.start
        {
            last.old.end = old.end;
            last.new.end = new.end;
            return;
        }
        self.0.push(DiffHunk { kind, old, new });
    }
}

/// Diffs `old` against `new` within `limits`. Tokens only need `==`; one that is unequal to
/// itself never matches.
pub fn diff_tokens<T: PartialEq>(
    old: &[T],
    new: &[T],
    limits: DiffLimits,
) -> Result<TokenDiff, DiffLimitExceeded> {
    let prefix = old.iter().zip(new).take_while(|(a, b)| a == b).count();
    let suffix = old[prefix..]
        .iter()
        .rev()
        .zip(new[prefix..].iter().rev())
        .take_while(|(a, b)| a == b)
        .count();
    let (old_end, new_end) = (old.len() - suffix, new.len() - suffix);
    let a = &old[prefix..old_end];
    let b = &new[prefix..new_end];
    let mut hunks = Hunks(Vec::new());
    hunks.push_run(DiffKind::Equal, 0..prefix, 0..prefix);
    let cells = a
        .len()
        .checked_add(1)
        .zip(b.len().checked_add(1))
        .and_then(|(rows, columns)| rows.checked_mul(columns));
    let over = old.len() > limits.max_tokens
        || new.len() > limits.max_tokens
        || cells.is_none_or(|cells| cells > limits.max_lcs_cells);
    let table = if over {
        None
    } else if a.is_empty() || b.is_empty() {
        Some(Vec::new())
    } else {
        cells.and_then(|cells| lcs_table(a, b, cells))
    };
    let mut coarsened = false;
    let mut cells_used = 0;
    match table {
        Some(lengths) => {
            cells_used = lengths.len();
            let width = b.len() + 1;
            let (mut i, mut j) = (0, 0);
            while i < a.len() || j < b.len() {
                if i < a.len() && j < b.len() && a[i] == b[j] {
                    hunks.push(DiffKind::Equal, prefix + i, prefix + j);
                    i += 1;
                    j += 1;
                } else if i < a.len()
                    && (j == b.len() || lengths[(i + 1) * width + j] >= lengths[i * width + j + 1])
                {
                    hunks.push(DiffKind::Delete, prefix + i, prefix + j);
                    i += 1;
                } else {
                    hunks.push(DiffKind::Insert, prefix + i, prefix + j);
                    j += 1;
                }
            }
        }
        None if limits.fallback == LimitFallback::Error => {
            return Err(DiffLimitExceeded {
                cells_required: cells.unwrap_or(usize::MAX),
            });
        }
        None => {
            coarsened = !a.is_empty() || !b.is_empty();
            hunks.push_run(DiffKind::Delete, prefix..old_end, prefix..prefix);
            hunks.push_run(DiffKind::Insert, old_end..old_end, prefix..new_end);
        }
    }
    hunks.push_run(DiffKind::Equal, old_end..old.len(), new_end..new.len());
    Ok(TokenDiff {
        hunks: hunks.0,
        coarsened,
        cells_used,
    })
}

/// Suffix LCS lengths, or `None` when the table cannot be allocated.
fn lcs_table<T: PartialEq>(a: &[T], b: &[T], cells: usize) -> Option<Vec<u32>> {
    let mut lengths: Vec<u32> = Vec::new();
    lengths.try_reserve_exact(cells).ok()?;
    lengths.resize(cells, 0);
    let width = b.len() + 1;
    for i in (0..a.len()).rev() {
        for j in (0..b.len()).rev() {
            lengths[i * width + j] = if a[i] == b[j] {
                lengths[(i + 1) * width + j + 1] + 1
            } else {
                lengths[(i + 1) * width + j].max(lengths[i * width + j + 1])
            };
        }
    }
    Some(lengths)
}

#[cfg(test)]
mod tests {
    use super::*;

    const EXACT: DiffLimits = DiffLimits::new(usize::MAX, 250_000, LimitFallback::Error);

    fn hunk(kind: DiffKind, old: Range<usize>, new: Range<usize>) -> DiffHunk {
        DiffHunk { kind, old, new }
    }

    fn chars(text: &str) -> Vec<char> {
        text.chars().collect()
    }

    #[test]
    fn equal_inputs_are_one_equal_hunk() {
        let diff = diff_tokens(&chars("abc"), &chars("abc"), EXACT).unwrap();
        assert_eq!(diff.hunks, vec![hunk(DiffKind::Equal, 0..3, 0..3)]);
        assert_eq!(diff.cells_used, 0);
        assert!(!diff.coarsened);
    }

    #[test]
    fn empty_inputs_have_no_hunks() {
        let diff = diff_tokens::<char>(&[], &[], EXACT).unwrap();
        assert!(diff.hunks.is_empty());
        let inserted = diff_tokens(&[], &chars("ab"), EXACT).unwrap();
        assert_eq!(inserted.hunks, vec![hunk(DiffKind::Insert, 0..0, 0..2)]);
        let deleted = diff_tokens(&chars("ab"), &[], EXACT).unwrap();
        assert_eq!(deleted.hunks, vec![hunk(DiffKind::Delete, 0..2, 0..0)]);
    }

    #[test]
    fn trims_prefix_and_suffix_around_a_replacement() {
        let diff = diff_tokens(&chars("abXcd"), &chars("abYZcd"), EXACT).unwrap();
        assert_eq!(
            diff.hunks,
            vec![
                hunk(DiffKind::Equal, 0..2, 0..2),
                hunk(DiffKind::Delete, 2..3, 2..2),
                hunk(DiffKind::Insert, 3..3, 2..4),
                hunk(DiffKind::Equal, 3..5, 4..6),
            ]
        );
        assert_eq!(diff.cells_used, 2 * 3);
    }

    #[test]
    fn ties_delete_before_inserting() {
        let diff = diff_tokens(&chars("ab"), &chars("ba"), EXACT).unwrap();
        assert_eq!(
            diff.hunks,
            vec![
                hunk(DiffKind::Delete, 0..1, 0..0),
                hunk(DiffKind::Equal, 1..2, 0..1),
                hunk(DiffKind::Insert, 2..2, 1..2),
            ]
        );
    }

    #[test]
    fn hunks_cover_both_inputs_in_order() {
        let old = chars("the quick brown fox");
        let new = chars("a quick red fox jumps");
        let diff = diff_tokens(&old, &new, EXACT).unwrap();
        let (mut i, mut j) = (0, 0);
        for hunk in &diff.hunks {
            assert_eq!((hunk.old.start, hunk.new.start), (i, j));
            if hunk.kind == DiffKind::Equal {
                assert_eq!(old[hunk.old.clone()], new[hunk.new.clone()]);
            }
            i = hunk.old.end;
            j = hunk.new.end;
        }
        assert_eq!((i, j), (old.len(), new.len()));
        assert!(
            diff.hunks
                .windows(2)
                .all(|pair| pair[0].kind != pair[1].kind
                    || pair[0].old.end != pair[1].old.start
                    || pair[0].new.end != pair[1].new.start)
        );
    }

    #[test]
    fn over_budget_errors_or_replaces_the_middle() {
        let old = chars("pXXXXs");
        let new = chars("pYYYYs");
        let tight = DiffLimits::new(usize::MAX, 24, LimitFallback::Error);
        assert_eq!(
            diff_tokens(&old, &new, tight),
            Err(DiffLimitExceeded { cells_required: 25 })
        );
        let coarse = diff_tokens(
            &old,
            &new,
            DiffLimits::new(usize::MAX, 24, LimitFallback::ReplaceMiddle),
        )
        .unwrap();
        assert!(coarse.coarsened);
        assert_eq!(coarse.cells_used, 0);
        assert_eq!(
            coarse.hunks,
            vec![
                hunk(DiffKind::Equal, 0..1, 0..1),
                hunk(DiffKind::Delete, 1..5, 1..1),
                hunk(DiffKind::Insert, 5..5, 1..5),
                hunk(DiffKind::Equal, 5..6, 5..6),
            ]
        );
        let fits = DiffLimits::new(usize::MAX, 25, LimitFallback::Error);
        assert_eq!(diff_tokens(&old, &new, fits).unwrap().cells_used, 25);
    }

    #[test]
    fn token_limit_applies_before_trimming() {
        let limits = DiffLimits::new(2, usize::MAX, LimitFallback::Error);
        assert!(diff_tokens(&chars("abc"), &chars("abc"), limits).is_err());
        let replaced = diff_tokens(
            &chars("abc"),
            &chars("abd"),
            DiffLimits::new(2, usize::MAX, LimitFallback::ReplaceMiddle),
        )
        .unwrap();
        assert!(replaced.coarsened);
    }
}
