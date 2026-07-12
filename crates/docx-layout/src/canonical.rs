use crate::types::Layout;
use serde_json::Value;

const GOLDEN_FACTOR: f64 = 1000.0;

/// Keys excluded from the canonical form (derived-redundant data).
const OMITTED_KEYS: [&str; 2] = ["resolvedLines", "checkpoints"];

/// Rounds to the nearest integer, with ties toward positive infinity.
fn round_ties_positive(x: f64) -> f64 {
    let floor = x.floor();
    if x - floor >= 0.5 { floor + 1.0 } else { floor }
}

fn round_number(n: f64) -> f64 {
    if !n.is_finite() {
        return n;
    }
    let rounded = round_ties_positive(n * GOLDEN_FACTOR) / GOLDEN_FACTOR;
    if rounded == 0.0 { 0.0 } else { rounded }
}

fn format_canonical_number(n: f64) -> String {
    if !n.is_finite() {
        return "null".to_string();
    }
    let n = if n == 0.0 { 0.0 } else { n };
    let abs = n.abs();
    if n == 0.0 || (1e-6..1e21).contains(&abs) {
        return format!("{n}");
    }
    let s = format!("{n:e}");
    match s.split_once('e') {
        Some((mantissa, exp)) if !exp.starts_with('-') => format!("{mantissa}e+{exp}"),
        _ => s,
    }
}

/// JSON.stringify string escaping: the two mandatory escapes, the short
/// control escapes, and `\u00xx` (lowercase hex) for remaining C0 controls.
/// Everything else passes through as UTF-8.
fn write_json_string(out: &mut String, s: &str) {
    out.push('"');
    for ch in s.chars() {
        match ch {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\u{8}' => out.push_str("\\b"),
            '\t' => out.push_str("\\t"),
            '\n' => out.push_str("\\n"),
            '\u{c}' => out.push_str("\\f"),
            '\r' => out.push_str("\\r"),
            c if (c as u32) < 0x20 => {
                out.push_str(&format!("\\u{:04x}", c as u32));
            }
            c => out.push(c),
        }
    }
    out.push('"');
}

/// Pretty-print like `JSON.stringify(value, null, 2)`. Object keys arrive
/// pre-sorted: serde_json's default `Map` is BTreeMap-backed, and byte order
/// equals `Object.keys().sort()` for these ASCII keys.
fn write_value(out: &mut String, value: &Value, depth: usize) {
    match value {
        Value::Null => out.push_str("null"),
        Value::Bool(b) => out.push_str(if *b { "true" } else { "false" }),
        Value::Number(num) => {
            let n = num.as_f64().unwrap_or(f64::NAN);
            out.push_str(&format_canonical_number(round_number(n)));
        }
        Value::String(s) => write_json_string(out, s),
        Value::Array(items) => {
            if items.is_empty() {
                out.push_str("[]");
                return;
            }
            out.push_str("[\n");
            for (i, item) in items.iter().enumerate() {
                if i > 0 {
                    out.push_str(",\n");
                }
                push_indent(out, depth + 1);
                write_value(out, item, depth + 1);
            }
            out.push('\n');
            push_indent(out, depth);
            out.push(']');
        }
        Value::Object(map) => {
            let entries: Vec<(&String, &Value)> = map
                .iter()
                .filter(|(key, _)| !OMITTED_KEYS.contains(&key.as_str()))
                .collect();
            if entries.is_empty() {
                out.push_str("{}");
                return;
            }
            out.push_str("{\n");
            for (i, (key, child)) in entries.iter().enumerate() {
                if i > 0 {
                    out.push_str(",\n");
                }
                push_indent(out, depth + 1);
                write_json_string(out, key);
                out.push_str(": ");
                write_value(out, child, depth + 1);
            }
            out.push('\n');
            push_indent(out, depth);
            out.push('}');
        }
    }
}

fn push_indent(out: &mut String, depth: usize) {
    for _ in 0..depth {
        out.push_str("  ");
    }
}

pub fn serialize_layout(layout: &Layout) -> String {
    let value = serde_json::to_value(layout).expect("Layout serializes to JSON");
    let mut out = String::new();
    write_value(&mut out, &value, 0);
    out.push('\n');
    out
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn rounds_like_round_ties_positive() {
        assert_eq!(round_ties_positive(2.5), 3.0);
        assert_eq!(round_ties_positive(-2.5), -2.0);
        assert_eq!(round_ties_positive(2.4), 2.0);
        assert_eq!(round_ties_positive(-2.6), -3.0);
    }

    #[test]
    fn formats_numbers_canonically() {
        assert_eq!(format_canonical_number(96.0), "96");
        assert_eq!(format_canonical_number(0.5), "0.5");
        assert_eq!(format_canonical_number(-0.0), "0");
        assert_eq!(format_canonical_number(42.667), "42.667");
        assert_eq!(format_canonical_number(1e21), "1e+21");
    }

    #[test]
    fn round_collapses_negative_zero() {
        assert_eq!(format_canonical_number(round_number(-0.0001)), "0");
    }
}
