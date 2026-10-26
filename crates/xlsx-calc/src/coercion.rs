use std::collections::HashSet;

use xlsx_model::Workbook;
use xlsx_model::numfmt::builtin_format_code;

#[derive(Default)]
pub(crate) struct CurrencyCoercion {
    currencies: Vec<char>,
}

impl CurrencyCoercion {
    pub(crate) fn from_workbook(wb: &Workbook) -> Self {
        let mut result = Self::default();
        for (_, code) in &wb.styles.num_fmts {
            result.format_symbols(code);
        }
        // built-in formats count only when a cell, row or column uses them;
        // the cell scan runs only for those that would add a symbol
        let mut pending: HashSet<u32> = HashSet::new();
        for (index, xf) in wb.styles.cell_xfs.iter().enumerate() {
            let Some(id) = xf.num_fmt_id else { continue };
            if wb.styles.num_fmts.iter().any(|(custom, _)| *custom == id) {
                continue;
            }
            let Some(code) = builtin_format_code(id) else {
                continue;
            };
            let mut symbols = Self::default();
            symbols.format_symbols(code);
            if symbols
                .currencies
                .iter()
                .any(|ch| !result.currencies.contains(ch))
            {
                pending.insert(index as u32);
            }
        }
        let mut used = Vec::new();
        let mut take = |style: u32| {
            if pending.remove(&style) {
                used.push(style);
            }
            pending.is_empty()
        };
        if !take(0) {
            'sheets: for sheet in &wb.sheets {
                let rows = sheet.row_styles.values().copied();
                let columns = sheet.col_styles.iter().map(|column| column.xf);
                let cells = sheet.iter_cells().filter_map(|(_, cell)| cell.style);
                for style in rows.chain(columns).chain(cells) {
                    if take(style) {
                        break 'sheets;
                    }
                }
            }
        }
        for index in used {
            let code = wb
                .styles
                .cell_xfs
                .get(index as usize)
                .and_then(|xf| xf.num_fmt_id)
                .and_then(builtin_format_code);
            if let Some(code) = code {
                result.format_symbols(code);
            }
        }
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
