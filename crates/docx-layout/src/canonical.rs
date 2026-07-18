//! Deterministic canonicalizer for a `Layout` tree — port of
//! `packages/core/src/layout/pagination/__golden__/serializeLayout.ts`.
//!
//! Produces the exact byte sequence the TS canonicalizer produces: sorted
//! keys, derived fields (`resolvedLines`, `checkpoints`) omitted, every
//! number rounded with ECMAScript `Math.round(n * 1000) / 1000` semantics and
//! printed with ECMAScript `Number::toString` formatting, pretty-printed like
//! `JSON.stringify(value, null, 2)`, trailing newline included.

use crate::types::Layout;
use serde_json::Value;

/// TS `GOLDEN_PRECISION` — decimal places kept for every number.
const GOLDEN_FACTOR: f64 = 1000.0; // 10 ** 3

/// Keys excluded from the canonical form (derived-redundant data).
const OMITTED_KEYS: [&str; 2] = ["resolvedLines", "checkpoints"];

/// ECMAScript `Math.round`: nearest integral value, ties toward +Infinity.
/// (`f64::round` ties away from zero — wrong for negative halves.) The
/// fractional part `x - floor(x)` is exactly representable, so the comparison
/// is exact.
fn js_math_round(x: f64) -> f64 {
    let floor = x.floor();
    if x - floor >= 0.5 { floor + 1.0 } else { floor }
}

/// TS `roundNumber` — round to golden precision, collapsing -0 to 0.
fn round_number(n: f64) -> f64 {
    if !n.is_finite() {
        return n;
    }
    let rounded = js_math_round(n * GOLDEN_FACTOR) / GOLDEN_FACTOR;
    if rounded == 0.0 { 0.0 } else { rounded }
}

/// ECMAScript `Number::toString(10)` for the canonical range. Rust's `f64`
/// `Display` is the same shortest-round-trip decimal without an exponent, so
/// it matches JS everywhere JS uses plain decimal notation
/// (`1e-6 <= |x| < 1e21`, integers, zero). Outside that range JS switches to
/// exponential — replicated from Rust's `LowerExp` plus JS's explicit `+`.
fn format_js_number(n: f64) -> String {
    if !n.is_finite() {
        // JSON.stringify(NaN / Infinity) === "null"
        return "null".to_string();
    }
    let n = if n == 0.0 { 0.0 } else { n }; // -0 prints as "0"
    let abs = n.abs();
    if n == 0.0 || (1e-6..1e21).contains(&abs) {
        return format!("{n}");
    }
    // exponential form: JS writes a '+' for positive exponents ("1e+21")
    let s = format!("{n:e}");
    match s.split_once('e') {
        Some((mantissa, exp)) if !exp.starts_with('-') => format!("{mantissa}e+{exp}"),
        _ => s,
    }
}

/// JSON.stringify string escaping: the two mandatory escapes, the short
/// control escapes, and `\u00xx` (lowercase hex) for remaining C0 controls.
/// Everything else passes through as UTF-8.
fn write_js_string(out: &mut String, s: &str) {
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
            out.push_str(&format_js_number(round_number(n)));
        }
        Value::String(s) => write_js_string(out, s),
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
                write_js_string(out, key);
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

/// TS `serializeLayout` — canonical golden string (trailing newline included).
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
    fn rounds_like_js_math_round() {
        assert_eq!(js_math_round(2.5), 3.0);
        assert_eq!(js_math_round(-2.5), -2.0); // JS ties toward +Infinity
        assert_eq!(js_math_round(2.4), 2.0);
        assert_eq!(js_math_round(-2.6), -3.0);
    }

    #[test]
    fn formats_numbers_like_js() {
        assert_eq!(format_js_number(96.0), "96");
        assert_eq!(format_js_number(0.5), "0.5");
        assert_eq!(format_js_number(-0.0), "0");
        assert_eq!(format_js_number(42.667), "42.667");
        assert_eq!(format_js_number(1e21), "1e+21");
    }

    #[test]
    fn round_collapses_negative_zero() {
        assert_eq!(format_js_number(round_number(-0.0001)), "0");
    }
}
