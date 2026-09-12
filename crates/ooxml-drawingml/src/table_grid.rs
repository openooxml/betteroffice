//! Table column-width geometry shared by the docx and pptx layout engines.

/// Makes `column_widths` exactly `col_count` long with every entry positive.
/// Fills gaps only: a caller scales a total that overruns `target_width`.
pub fn normalize_table_column_widths(
    column_widths: &[f64],
    col_count: usize,
    target_width: f64,
) -> Vec<f64> {
    if col_count == 0 {
        return Vec::new();
    }

    let even_width = if target_width > 0.0 {
        target_width / col_count as f64
    } else {
        0.0
    };

    if column_widths.is_empty() {
        return vec![even_width; col_count];
    }

    let mut normalized: Vec<f64> = column_widths.iter().copied().take(col_count).collect();
    let missing_columns = col_count - normalized.len();
    if missing_columns > 0 {
        let existing_positive: Vec<f64> = normalized.iter().copied().filter(|w| *w > 0.0).collect();
        let fallback_width = if !existing_positive.is_empty() {
            existing_positive.iter().fold(0.0, |sum, w| sum + w) / existing_positive.len() as f64
        } else {
            even_width
        };
        normalized.extend(std::iter::repeat_n(fallback_width, missing_columns));
    }

    let positive_total = normalized
        .iter()
        .fold(0.0, |sum, &w| sum + if w > 0.0 { w } else { 0.0 });
    let non_positive_count = normalized.iter().filter(|&&w| w <= 0.0).count();

    if positive_total <= 0.0 {
        return vec![even_width; col_count];
    }
    if non_positive_count == 0 {
        return normalized;
    }

    let remaining_width = (target_width - positive_total).max(0.0);
    let fallback_width = if remaining_width > 0.0 {
        remaining_width / non_positive_count as f64
    } else {
        positive_total / std::cmp::max(1, col_count - non_positive_count) as f64
    };

    normalized
        .into_iter()
        .map(|w| if w > 0.0 { w } else { fallback_width })
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn empty_array_returns_evenly_split_target_width() {
        assert_eq!(
            normalize_table_column_widths(&[], 3, 300.0),
            vec![100.0, 100.0, 100.0]
        );
    }

    #[test]
    fn missing_trailing_columns_inherit_average_of_existing_positives() {
        assert_eq!(
            normalize_table_column_widths(&[100.0, 200.0], 4, 1000.0),
            vec![100.0, 200.0, 150.0, 150.0]
        );
    }

    #[test]
    fn zero_negative_widths_split_the_leftover_target_evenly() {
        let out = normalize_table_column_widths(&[100.0, 0.0, 100.0, -5.0], 4, 400.0);
        assert_eq!(out[0], 100.0);
        assert_eq!(out[2], 100.0);
        assert!((out[1] - 100.0).abs() < 1e-5);
        assert!((out[3] - 100.0).abs() < 1e-5);
    }

    #[test]
    fn all_zero_returns_even_split_of_target() {
        assert_eq!(
            normalize_table_column_widths(&[0.0, 0.0, 0.0], 3, 300.0),
            vec![100.0, 100.0, 100.0]
        );
    }
}
