use std::collections::HashSet;

use xlsx_model::Workbook;
use xlsx_model::numfmt::builtin_format_code;

#[derive(Clone, Default, PartialEq, Eq)]
pub(crate) struct CurrencyCoercion {
    currencies: Vec<char>,
}

#[cfg(test)]
thread_local! {
    static EVIDENCE_SCANS: std::cell::Cell<usize> = const { std::cell::Cell::new(0) };
}

impl CurrencyCoercion {
    pub(crate) fn from_workbook(wb: &Workbook) -> Self {
        let mut used = std::collections::BTreeSet::from([0]);
        for sheet in &wb.sheets {
            used.extend(sheet.row_styles.values().copied());
            used.extend(sheet.col_style_runs().iter().map(|column| column.xf));
            used.extend(sheet.used_cell_styles());
        }
        let used: Vec<u32> = used.into_iter().collect();
        Self {
            currencies: wb
                .styles
                .currency_symbols_for(&used, || Self::scan(wb, &used).currencies),
        }
    }

    pub(crate) fn symbols(&self) -> &[char] {
        &self.currencies
    }

    fn scan(wb: &Workbook, used: &[u32]) -> Self {
        #[cfg(test)]
        EVIDENCE_SCANS.set(EVIDENCE_SCANS.get() + 1);
        let mut result = Self::default();
        let custom: HashSet<u16> = wb.styles.num_fmts.iter().map(|(id, _)| *id).collect();
        for (_, code) in &wb.styles.num_fmts {
            result.format_symbols(code);
        }
        for &index in used {
            if let Some(id) = wb
                .styles
                .cell_xfs
                .get(index as usize)
                .and_then(|xf| xf.num_fmt_id)
                && !custom.contains(&id)
                && let Some(code) = builtin_format_code(id)
            {
                result.format_symbols(code);
            }
        }
        result.currencies.sort_unstable();
        result
    }

    fn format_symbols(&mut self, code: &str) {
        let mut chars = code.chars();
        while let Some(ch) = chars.next() {
            match ch {
                '"' => {
                    for literal in chars.by_ref() {
                        if literal == '"' {
                            break;
                        }
                        self.currency_symbol(literal);
                    }
                }
                '[' => {
                    let body: String = chars.by_ref().take_while(|ch| *ch != ']').collect();
                    if let Some(currency) = body.strip_prefix('$') {
                        for literal in currency.split('-').next().unwrap_or("").chars() {
                            self.currency_symbol(literal);
                        }
                    }
                }
                '\\' => {
                    if let Some(literal) = chars.next() {
                        self.currency_symbol(literal);
                    }
                }
                '_' | '*' => {
                    chars.next();
                }
                _ => self.currency_symbol(ch),
            }
        }
    }

    fn currency_symbol(&mut self, ch: char) {
        if matches!(
            ch,
            '$' | '£' | '¥' | '¢' | '\u{20a0}'
                ..='\u{20c0}' | '৳' | '฿' | '៛' | '﷼' | '￥' | '＄' | '￡' | '￦'
        ) && !self.currencies.contains(&ch)
        {
            self.currencies.push(ch);
        }
    }

    pub(crate) fn currency_number(&self, raw: &str) -> Option<f64> {
        let raw = raw.trim();
        let (body, accounting) = match raw.strip_prefix('(') {
            Some(body) => (body.strip_suffix(')')?.trim(), true),
            None => (raw, false),
        };
        let (body, leading_sign) = sign(body);
        let (body, trailing_sign) = if let Some(symbol) = body.chars().next()
            && self.currencies.contains(&symbol)
        {
            sign(body[symbol.len_utf8()..].trim())
        } else {
            let symbol = body.chars().last()?;
            if !self.currencies.contains(&symbol) {
                return None;
            }
            (body[..body.len() - symbol.len_utf8()].trim(), None)
        };
        if (accounting && (leading_sign.is_some() || trailing_sign.is_some()))
            || (leading_sign.is_some() && trailing_sign.is_some())
        {
            return None;
        }
        // currency plus percent stays text until office evidence settles coercion
        let (integer, decimal) = body.split_once('.').unwrap_or((body, ""));
        if integer.is_empty()
            || !decimal.bytes().all(|ch| ch.is_ascii_digit())
            || (body.contains('.') && decimal.is_empty())
        {
            return None;
        }
        let mut groups = integer.split(',');
        let first = groups.next()?;
        if first.is_empty()
            || !first.bytes().all(|ch| ch.is_ascii_digit())
            || (integer.contains(',') && first.len() > 3)
            || groups.any(|group| group.len() != 3 || !group.bytes().all(|ch| ch.is_ascii_digit()))
        {
            return None;
        }
        let magnitude: f64 = body.replace(',', "").parse().ok()?;
        let sign = if accounting {
            -1.0
        } else {
            leading_sign.or(trailing_sign).unwrap_or(1.0)
        };
        magnitude.is_finite().then_some(sign * magnitude)
    }
}

fn sign(raw: &str) -> (&str, Option<f64>) {
    if let Some(body) = raw.strip_prefix('-') {
        (body, Some(-1.0))
    } else if let Some(body) = raw.strip_prefix('+') {
        (body, Some(1.0))
    } else {
        (raw, None)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use xlsx_model::{Cell, CellRef, Sheet, Xf};

    #[test]
    fn format_edits_that_keep_the_evidence_rerun_no_formulas() {
        let mut wb = Workbook::default();
        wb.styles.num_fmts.push((164, "£0.00".into()));
        wb.styles.cell_xfs = vec![
            Xf::default(),
            Xf {
                num_fmt_id: Some(164),
                ..Xf::default()
            },
            Xf {
                num_fmt_id: Some(2),
                ..Xf::default()
            },
        ];
        let mut sheet = Sheet::new("Sheet1");
        for row in 0..200 {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    formula: Some("\"£12\"+1".into()),
                    ..Cell::default()
                },
            );
        }
        let at = CellRef::new(0, 3);
        sheet.set_cell(
            at,
            Cell {
                value: xlsx_model::CellValue::Number { value: 1.0 },
                style: Some(1),
                ..Cell::default()
            },
        );
        wb.sheets.push(sheet);
        let (mut graph, _) = crate::rebuild_and_recalc_all(&mut wb, None);
        let seed = [(xlsx_model::SheetId(0), at)];
        for style in [Some(2), Some(1), None] {
            wb.sheets[0].cell_mut(at).unwrap().style = style;
            let (_, recomputed) =
                crate::recalc_after_with_recomputed(&mut wb, &mut graph, &seed, None, None);
            assert!(recomputed.is_empty(), "{style:?}");
        }
        // a currency the workbook did not use before changes the evidence
        wb.styles.num_fmts.push((165, "€0.00".into()));
        let (_, recomputed) =
            crate::recalc_after_with_recomputed(&mut wb, &mut graph, &[], None, None);
        assert_eq!(recomputed.len(), 200);
    }

    #[test]
    fn a_graph_without_recorded_evidence_adopts_it_without_a_full_recalc() {
        let mut wb = Workbook::default();
        let mut sheet = Sheet::new("Sheet1");
        for row in 0..10 {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    formula: Some("1+1".into()),
                    ..Cell::default()
                },
            );
        }
        sheet.set_cell(
            CellRef::new(0, 1),
            Cell {
                value: xlsx_model::CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
        sheet.set_cell(
            CellRef::new(0, 2),
            Cell {
                formula: Some("B1+1".into()),
                ..Cell::default()
            },
        );
        wb.sheets.push(sheet);
        let (mut graph, _) = crate::rebuild_and_recalc_all(&mut wb, None);
        *graph.currency_coercion.lock().unwrap() = None;
        let sheet = xlsx_model::SheetId(0);
        let (_, recomputed) = crate::recalc_after_with_recomputed(
            &mut wb,
            &mut graph,
            &[(sheet, CellRef::new(0, 1))],
            None,
            None,
        );
        assert_eq!(recomputed, vec![(sheet, CellRef::new(0, 2))]);
    }

    #[test]
    fn stepwise_graphs_record_the_evidence_they_were_built_over() {
        let mut wb = Workbook::default();
        // a built-in currency format is evidence only once a cell uses it
        wb.styles.cell_xfs = vec![
            Xf::default(),
            Xf {
                num_fmt_id: Some(7),
                ..Xf::default()
            },
        ];
        let mut sheet = Sheet::new("Sheet1");
        for (col, formula) in [(0, "\"$12\"+0"), (1, "A1+1")] {
            sheet.set_cell(
                CellRef::new(0, col),
                Cell {
                    formula: Some(formula.into()),
                    ..Cell::default()
                },
            );
        }
        let at = CellRef::new(0, 3);
        sheet.set_cell(
            at,
            Cell {
                value: xlsx_model::CellValue::Number { value: 1.0 },
                ..Cell::default()
            },
        );
        wb.sheets.push(sheet);
        crate::rebuild_and_recalc_all(&mut wb, None);
        let sheet = xlsx_model::SheetId(0);
        let error = xlsx_model::CellValue::Error {
            value: xlsx_model::ErrorValue::Value,
        };
        assert_eq!(wb.sheets[0].cell(CellRef::new(0, 1)).unwrap().value, error);
        let mut builder = crate::graph::DepGraphBuilder::new();
        while !builder.advance(&wb, std::num::NonZeroUsize::new(1).unwrap()) {}
        let mut graph = builder.finish().unwrap();
        // the first edit after the build adds the evidence
        wb.sheets[0].cell_mut(at).unwrap().style = Some(1);
        crate::recalc_after(&mut wb, &mut graph, &[(sheet, at)], None);
        assert_eq!(
            wb.sheets[0].cell(CellRef::new(0, 0)).unwrap().value,
            xlsx_model::CellValue::Number { value: 12.0 }
        );
        assert_eq!(
            wb.sheets[0].cell(CellRef::new(0, 1)).unwrap().value,
            xlsx_model::CellValue::Number { value: 13.0 }
        );
    }

    #[test]
    fn repeated_recalcs_reuse_currency_evidence() {
        let mut wb = Workbook::default();
        wb.styles.num_fmts.push((164, "£0.00".into()));
        let mut sheet = Sheet::new("Sheet1");
        sheet.set_cell(
            CellRef::new(0, 0),
            Cell {
                formula: Some("\"£12\"+1".into()),
                ..Cell::default()
            },
        );
        wb.sheets.push(sheet);
        EVIDENCE_SCANS.set(0);
        let (mut graph, _) = crate::rebuild_and_recalc_all(&mut wb, None);
        crate::recalc_after(&mut wb, &mut graph, &[], None);
        crate::rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(EVIDENCE_SCANS.get(), 1);
        wb.styles.num_fmts[0].1 = "€0.00".into();
        crate::recalc_after(&mut wb, &mut graph, &[], None);
        assert_eq!(EVIDENCE_SCANS.get(), 2);
    }

    #[test]
    fn unused_builtin_styles_do_not_repeat_cell_scans() {
        let mut wb = Workbook::default();
        wb.styles.cell_xfs = vec![
            Xf::default(),
            Xf {
                num_fmt_id: Some(7),
                ..Xf::default()
            },
        ];
        let mut sheet = Sheet::new("Sheet1");
        for row in 0..1000 {
            sheet.set_cell(
                CellRef::new(row, 0),
                Cell {
                    value: xlsx_model::CellValue::Number { value: 1.0 },
                    ..Cell::default()
                },
            );
        }
        wb.sheets.push(sheet);
        EVIDENCE_SCANS.set(0);
        crate::rebuild_and_recalc_all(&mut wb, None);
        crate::rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(EVIDENCE_SCANS.get(), 1);
        wb.sheets[0].cell_mut(CellRef::new(999, 0)).unwrap().style = Some(1);
        crate::rebuild_and_recalc_all(&mut wb, None);
        assert_eq!(EVIDENCE_SCANS.get(), 2);
        assert_eq!(
            CurrencyCoercion::from_workbook(&wb).currency_number("$12"),
            Some(12.0)
        );
    }
}
