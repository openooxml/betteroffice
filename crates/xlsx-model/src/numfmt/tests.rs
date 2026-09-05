use super::*;
use crate::value::ErrorValue;

fn num(v: f64) -> CellValue {
    CellValue::Number { value: v }
}

fn fv(v: f64, code: &str) -> String {
    format_value(&num(v), code, DateSystem::V1900).text
}

fn fv04(v: f64, code: &str) -> String {
    format_value(&num(v), code, DateSystem::V1904).text
}

fn color(v: f64, code: &str) -> Option<String> {
    format_value(&num(v), code, DateSystem::V1900).color
}

fn time_serial(secs: f64) -> f64 {
    secs / 86_400.0
}

#[test]
fn builtin_table_ids() {
    assert_eq!(builtin_format_code(0), Some("General"));
    assert_eq!(builtin_format_code(1), Some("0"));
    assert_eq!(builtin_format_code(2), Some("0.00"));
    assert_eq!(builtin_format_code(3), Some("#,##0"));
    assert_eq!(builtin_format_code(4), Some("#,##0.00"));
    assert_eq!(builtin_format_code(9), Some("0%"));
    assert_eq!(builtin_format_code(10), Some("0.00%"));
    assert_eq!(builtin_format_code(11), Some("0.00E+00"));
    assert_eq!(builtin_format_code(12), Some("# ?/?"));
    assert_eq!(builtin_format_code(13), Some("# ??/??"));
    assert_eq!(builtin_format_code(14), Some("m/d/yyyy"));
    assert_eq!(builtin_format_code(15), Some("d-mmm-yy"));
    assert_eq!(builtin_format_code(16), Some("d-mmm"));
    assert_eq!(builtin_format_code(17), Some("mmm-yy"));
    assert_eq!(builtin_format_code(18), Some("h:mm AM/PM"));
    assert_eq!(builtin_format_code(19), Some("h:mm:ss AM/PM"));
    assert_eq!(builtin_format_code(20), Some("h:mm"));
    assert_eq!(builtin_format_code(21), Some("h:mm:ss"));
    assert_eq!(builtin_format_code(22), Some("m/d/yyyy h:mm"));
    assert_eq!(builtin_format_code(37), Some("#,##0 ;(#,##0)"));
    assert_eq!(builtin_format_code(38), Some("#,##0 ;[Red](#,##0)"));
    assert_eq!(builtin_format_code(39), Some("#,##0.00;(#,##0.00)"));
    assert_eq!(builtin_format_code(40), Some("#,##0.00;[Red](#,##0.00)"));
    assert_eq!(builtin_format_code(45), Some("mm:ss"));
    assert_eq!(builtin_format_code(46), Some("[h]:mm:ss"));
    assert_eq!(builtin_format_code(47), Some("mmss.0"));
    assert_eq!(builtin_format_code(48), Some("##0.0E+0"));
    assert_eq!(builtin_format_code(49), Some("@"));
    assert_eq!(builtin_format_code(5), None);
    assert_eq!(builtin_format_code(8), None);
    assert_eq!(builtin_format_code(23), None);
    assert_eq!(builtin_format_code(44), None);
    assert_eq!(builtin_format_code(163), None);
    assert_eq!(builtin_format_code(164), None);
}

#[test]
fn builtin_number_renders() {
    assert_eq!(fv(1234.5, "General"), "1234.5");
    assert_eq!(fv(1234.0, "0"), "1234");
    assert_eq!(fv(1234.56, "0"), "1235");
    assert_eq!(fv(1234.5, "0.00"), "1234.50");
    assert_eq!(fv(1234567.0, "#,##0"), "1,234,567");
    assert_eq!(fv(1234.5, "#,##0.00"), "1,234.50");
    assert_eq!(fv(0.5, "0%"), "50%");
    assert_eq!(fv(0.5, "0.00%"), "50.00%");
    assert_eq!(fv(12345.0, "0.00E+00"), "1.23E+04");
    assert_eq!(fv(12345.0, "##0.0E+0"), "12.3E+3");
    assert_eq!(fv(1234.0, "@"), "1234");
}

#[test]
fn builtin_date_renders() {
    assert_eq!(fv(43831.0, "m/d/yyyy"), "1/1/2020");
    assert_eq!(fv(43831.0, "d-mmm-yy"), "1-Jan-20");
    assert_eq!(fv(43831.0, "d-mmm"), "1-Jan");
    assert_eq!(fv(43831.0, "mmm-yy"), "Jan-20");
    assert_eq!(fv(43831.5, "m/d/yyyy h:mm"), "1/1/2020 12:00");
}

#[test]
fn builtin_time_renders() {
    assert_eq!(fv(0.5, "h:mm AM/PM"), "12:00 PM");
    assert_eq!(fv(0.5, "h:mm:ss AM/PM"), "12:00:00 PM");
    assert_eq!(fv(0.5, "h:mm"), "12:00");
    assert_eq!(fv(0.5, "h:mm:ss"), "12:00:00");
    assert_eq!(fv(time_serial(90.0), "mm:ss"), "01:30");
    assert_eq!(fv(1.5, "[h]:mm:ss"), "36:00:00");
}

#[test]
fn sections_by_sign() {
    assert_eq!(fv(5.0, "0.00"), "5.00");
    assert_eq!(fv(-5.0, "0.00"), "-5.00");
    assert_eq!(fv(0.0, "0.00"), "0.00");
    assert_eq!(fv(5.0, "0.00;(0.00)"), "5.00");
    assert_eq!(fv(-5.0, "0.00;(0.00)"), "(5.00)");
    assert_eq!(fv(0.0, "0.00;(0.00)"), "0.00");
    assert_eq!(fv(5.0, "\"pos\";\"neg\";\"zero\""), "pos");
    assert_eq!(fv(-5.0, "\"pos\";\"neg\";\"zero\""), "neg");
    assert_eq!(fv(0.0, "\"pos\";\"neg\";\"zero\""), "zero");
    assert_eq!(fv(5.0, ";;;"), "");
}

#[test]
fn conditions_override_selection() {
    assert_eq!(fv(150.0, "[>=100]0;0"), "150");
    assert_eq!(fv(50.0, "[>=100]0;0"), "50");
    assert_eq!(fv(-5.0, "[<0]\"neg\";[>=0]\"pos\""), "neg");
    assert_eq!(fv(5.0, "[<0]\"neg\";[>=0]\"pos\""), "pos");
    assert_eq!(fv(150.0, "[Red][>50]0;[Blue]0"), "150");
    assert_eq!(
        color(150.0, "[Red][>50]0;[Blue]0"),
        Some("#FF0000".to_string())
    );
    assert_eq!(fv(30.0, "[Red][>50]0;[Blue]0"), "30");
    assert_eq!(
        color(30.0, "[Red][>50]0;[Blue]0"),
        Some("#0000FF".to_string())
    );
}

#[test]
fn thousands_and_scaling() {
    assert_eq!(fv(1234567.0, "#,##0"), "1,234,567");
    assert_eq!(fv(12.0, "#,##0"), "12");
    assert_eq!(fv(0.0, "#,##0"), "0");
    assert_eq!(fv(1234567.0, "#,##0,"), "1,235");
    assert_eq!(fv(1234567890.0, "0,,"), "1235");
    assert_eq!(fv(5.0, "000"), "005");
    assert_eq!(fv(5.0, "00,000"), "00,005");
}

#[test]
fn percent_and_scientific() {
    assert_eq!(fv(0.5, "0%"), "50%");
    assert_eq!(fv(0.1234, "0.0%"), "12.3%");
    assert_eq!(fv(1.0, "0%"), "100%");
    assert_eq!(fv(12345.0, "0.00E+00"), "1.23E+04");
    assert_eq!(fv(-12345.0, "0.00E+00"), "-1.23E+04");
    assert_eq!(fv(0.0, "0.0E+0"), "0.0E+0");
    assert_eq!(fv(0.00012345, "0.00E+00"), "1.23E-04");
    assert_eq!(fv(12345.0, "##0.0E+0"), "12.3E+3");
}

#[test]
fn digit_placeholder_padding() {
    assert_eq!(fv(5.0, "??0"), "  5");
    assert_eq!(fv(5.0, "???0"), "   5");
    assert_eq!(fv(1.5, "0.??"), "1.5 ");
    assert_eq!(fv(1.5, "0.00"), "1.50");
    assert_eq!(fv(3.4, "0.0#"), "3.4");
    assert_eq!(fv(3.46, "0.0#"), "3.46");
    assert_eq!(fv(3.5, "0.##"), "3.5");
    assert_eq!(fv(0.5, "#.00"), ".50");
    assert_eq!(fv(0.5, "0.00"), "0.50");
    assert_eq!(fv(12.3, ".00"), "12.30");
}

#[test]
fn rounding_is_half_away_from_zero() {
    assert_eq!(fv(2.5, "0"), "3");
    assert_eq!(fv(3.5, "0"), "4");
    assert_eq!(fv(-2.5, "0"), "-3");
    assert_eq!(fv(0.125, "0.00"), "0.13");
    assert_eq!(fv(2.674, "0.00"), "2.67");
}

#[test]
fn date_anchors_and_phantom_day() {
    // the deliberate 1900 leap-year bug: serial 60 is the phantom 1900-02-29.
    assert_eq!(fv(60.0, "m/d/yyyy"), "2/29/1900");
    assert_eq!(fv(59.0, "m/d/yyyy"), "2/28/1900");
    assert_eq!(fv(61.0, "m/d/yyyy"), "3/1/1900");
    assert_eq!(fv(1.0, "m/d/yyyy"), "1/1/1900");
    assert_eq!(fv(43831.0, "yyyy"), "2020");
    assert_eq!(fv(43831.0, "yy"), "20");
    assert_eq!(fv(-5.0, "m/d/yyyy"), "#######");
}

#[test]
fn date_systems_agree_on_calendar() {
    assert_eq!(fv(43831.0, "m/d/yyyy"), "1/1/2020");
    assert_eq!(fv04(42369.0, "m/d/yyyy"), "1/1/2020");
    assert_eq!(fv04(0.0, "m/d/yyyy"), "1/1/1904");
}

#[test]
fn month_and_day_name_widths() {
    assert_eq!(fv(43831.0, "mmmm d, yyyy"), "January 1, 2020");
    assert_eq!(fv(43831.0, "mmm"), "Jan");
    assert_eq!(fv(43831.0, "mmmmm"), "J");
    assert_eq!(fv(43831.0, "dddd"), "Wednesday");
    assert_eq!(fv(43831.0, "ddd"), "Wed");
    assert_eq!(fv(43831.0, "dd/mm/yyyy"), "01/01/2020");
}

#[test]
fn minute_versus_month_disambiguation() {
    assert_eq!(fv(43831.0, "m/d"), "1/1");
    assert_eq!(fv(0.5, "h:mm"), "12:00");
    assert_eq!(fv(time_serial(90.0), "mm:ss"), "01:30");
    assert_eq!(fv(43831.5, "m/d/yyyy h:mm"), "1/1/2020 12:00");
}

#[test]
fn am_pm_switch() {
    assert_eq!(fv(0.25, "h:mm AM/PM"), "6:00 AM");
    assert_eq!(fv(0.75, "h:mm AM/PM"), "6:00 PM");
    assert_eq!(fv(0.0, "h AM/PM"), "12 AM");
    assert_eq!(fv(0.5, "h AM/PM"), "12 PM");
    assert_eq!(fv(0.25, "h:mm A/P"), "6:00 A");
    assert_eq!(fv(0.75, "h:mm A/P"), "6:00 P");
}

#[test]
fn elapsed_time_beyond_a_day() {
    assert_eq!(fv(1.5, "[h]:mm"), "36:00");
    assert_eq!(fv(1.5, "[h]:mm:ss"), "36:00:00");
    assert_eq!(fv(1.0, "[h]"), "24");
    assert_eq!(fv(1.0, "[mm]"), "1440");
    assert_eq!(fv(time_serial(90.0), "[ss]"), "90");
}

#[test]
fn fractional_seconds() {
    assert_eq!(fv(time_serial(1.5), "ss.0"), "01.5");
    assert_eq!(fv(time_serial(1.25), "ss.00"), "01.25");
}

#[test]
fn colors_named_and_indexed() {
    assert_eq!(color(5.0, "[Red]0"), Some("#FF0000".to_string()));
    assert_eq!(color(5.0, "[Blue]0"), Some("#0000FF".to_string()));
    assert_eq!(color(5.0, "[Green]0"), Some("#00FF00".to_string()));
    assert_eq!(color(5.0, "[Magenta]0"), Some("#FF00FF".to_string()));
    assert_eq!(color(5.0, "[Cyan]0"), Some("#00FFFF".to_string()));
    assert_eq!(color(5.0, "[Yellow]0"), Some("#FFFF00".to_string()));
    assert_eq!(color(5.0, "[White]0"), Some("#FFFFFF".to_string()));
    assert_eq!(color(5.0, "[Black]0"), Some("#000000".to_string()));
    assert_eq!(color(5.0, "[Color 3]0"), Some("#FF0000".to_string()));
    assert_eq!(color(5.0, "[Color 99]0"), None);
    assert_eq!(fv(5.0, "[Red]0.00"), "5.00");
    assert_eq!(
        color(-1234.0, "#,##0 ;[Red](#,##0)"),
        Some("#FF0000".to_string())
    );
    assert_eq!(color(1234.0, "#,##0 ;[Red](#,##0)"), None);
    assert_eq!(fv(-1234.0, "#,##0 ;[Red](#,##0)"), "(1,234)");
    assert_eq!(fv(1234.0, "#,##0 ;[Red](#,##0)"), "1,234 ");
}

#[test]
fn text_values_and_placeholder() {
    let hello = CellValue::Text {
        value: "hello".to_string(),
    };
    let out = format_value(&hello, "0.00;-0.00;0.00;\"txt: \"@", DateSystem::V1900);
    assert_eq!(out.text, "txt: hello");
    assert_eq!(format_value(&hello, "@", DateSystem::V1900).text, "hello");
    assert_eq!(
        format_value(&hello, "0.00", DateSystem::V1900).text,
        "hello"
    );
    let colored = format_value(&hello, "[Red]@", DateSystem::V1900);
    assert_eq!(colored.text, "hello");
    assert_eq!(colored.color, Some("#FF0000".to_string()));
}

#[test]
fn bools_errors_empty() {
    assert_eq!(
        format_value(&CellValue::Bool { value: true }, "0", DateSystem::V1900).text,
        "TRUE"
    );
    assert_eq!(
        format_value(
            &CellValue::Bool { value: false },
            "m/d/yyyy",
            DateSystem::V1900
        )
        .text,
        "FALSE"
    );
    assert_eq!(
        format_value(
            &CellValue::Error {
                value: ErrorValue::Div0
            },
            "0.00",
            DateSystem::V1900
        )
        .text,
        "#DIV/0!"
    );
    assert_eq!(
        format_value(&CellValue::Empty, "0.00", DateSystem::V1900).text,
        ""
    );
}

#[test]
fn general_edge_cases() {
    assert_eq!(fv(0.0, "General"), "0");
    assert_eq!(fv(-0.0, "General"), "0");
    assert_eq!(fv(42.0, "General"), "42");
    assert_eq!(fv(-42.0, "General"), "-42");
    assert_eq!(fv(0.5, "General"), "0.5");
    assert_eq!(fv(12.375, "General"), "12.375");
    assert_eq!(fv(1e11, "General"), "1E+11");
    assert_eq!(fv(123456789012.0, "General"), "1.23457E+11");
    assert_eq!(fv(0.0000001, "General"), "1E-07");
    assert_eq!(fv(-1234.5, "General"), "-1234.5");
    assert_eq!(fv(1000000000.0, "General"), "1000000000");
    assert_eq!(fv(42.0, ""), "42");
}

#[test]
fn unsupported_constructs_degrade() {
    assert_eq!(fv(0.5, "# ?/?"), "0.5");
    assert_eq!(fv(2.75, "# ??/??"), "2.75");
    assert_eq!(fv(5.0, "[$$-409]#,##0.00"), "$5.00");
    assert_eq!(fv(5.0, "[$-409]0"), "5");
    assert_eq!(fv(1234.0, "$#,##0"), "$1,234");
    assert_eq!(fv(5.0, "0_)"), "5 ");
    assert_eq!(fv(5.0, "0*x"), "5");
}

#[test]
fn escaped_and_quoted_literals() {
    assert_eq!(fv(5.0, "0\" kg\""), "5 kg");
    assert_eq!(fv(5.0, "0\\ \\k\\g"), "5 kg");
    assert_eq!(fv(100.0, "\"$\"#,##0"), "$100");
}

#[test]
fn repeated_codes_render_identically_through_cache() {
    let codes = [
        "#,##0.00",
        "[Red]0.00%;\"neg\";\"zero\";\"txt:\"@",
        "m/d/yyyy h:mm",
        "@",
        "\"p\"0.00",
    ];
    for code in codes {
        let expected = fv(1234.5678, code);
        for i in 0..(FORMAT_CACHE_CAP * 2) {
            let _ = fv(i as f64, &format!("\"churn{i}\"0.000"));
            assert_eq!(fv(1234.5678, code), expected);
        }
    }
    let hello = CellValue::Text {
        value: "hello".to_string(),
    };
    let expected = format_value(&hello, "[Red]\"pre \"@", DateSystem::V1900);
    for _ in 0..(FORMAT_CACHE_CAP * 2) {
        assert_eq!(
            format_value(&hello, "[Red]\"pre \"@", DateSystem::V1900),
            expected
        );
    }
}

#[test]
fn eviction_past_capacity_keeps_outputs_identical() {
    let rounds = FORMAT_CACHE_CAP * 3;
    let codes: Vec<String> = (0..rounds).map(|i| format!("\"p{i}\"0.00")).collect();
    let first: Vec<String> = codes.iter().map(|c| fv(7.625, c)).collect();
    for (i, code) in codes.iter().enumerate() {
        assert_eq!(fv(7.625, code), format!("p{i}7.63"));
    }
    for (code, expected) in codes.iter().zip(&first) {
        assert_eq!(fv(-7.625, code), format!("-{expected}"));
    }
}

fn parse_code(code: &str) -> Arc<[Section]> {
    split_sections(code).iter().map(|s| tokenize(s)).collect()
}

#[test]
fn distinct_insertions_never_exceed_two_generations() {
    let mut cache = FormatCache::default();
    for i in 0..=FORMAT_CACHE_CAP * 4 {
        let key = format!("\"s{i}\"0");
        cache.insert(&key, parse_code(&key));
        assert!(cache.hot.len() <= FORMAT_CACHE_CAP);
        assert!(cache.hot.len() + cache.cold.len() <= FORMAT_CACHE_CAP * 2);
    }
}

#[test]
fn repeat_lookup_returns_the_same_arc() {
    let mut cache = FormatCache::default();
    cache.insert("0.000", parse_code("0.000"));
    let first = cache.get("0.000").unwrap();
    for _ in 0..3 {
        let again = cache.get("0.000").unwrap();
        assert!(Arc::ptr_eq(&first, &again));
    }
    assert!(cache.get("0.00").is_none());
}

#[test]
fn a_full_generation_ages_out_but_still_answers_once() {
    let mut cache = FormatCache::default();
    let keys: Vec<String> = (0..FORMAT_CACHE_CAP)
        .map(|i| format!("\"e{i}\"0"))
        .collect();
    for key in &keys {
        cache.insert(key, parse_code(key));
    }
    assert_eq!(cache.hot.len(), FORMAT_CACHE_CAP);
    assert!(cache.cold.is_empty());

    let newcomer = "\"fresh\"0";
    cache.insert(newcomer, parse_code(newcomer));
    assert_eq!(cache.hot.len(), 1);
    assert_eq!(cache.cold.len(), FORMAT_CACHE_CAP);
    assert!(cache.hot.contains_key(newcomer));

    assert!(cache.get(&keys[0]).is_some());
    assert!(cache.hot.contains_key(&keys[0]));
}

#[test]
fn memoized_path_reuses_one_parse_across_lookups() {
    let mut cache = FormatCache::default();
    let a = parsed_sections_in(&mut cache, "[Red]#,##0.00;(#,##0.00)");
    let b = parsed_sections_in(&mut cache, "[Red]#,##0.00;(#,##0.00)");
    assert!(Arc::ptr_eq(&a, &b));
    assert_eq!(cache.hot.len(), 1);
}

#[test]
fn oversize_codes_render_correctly_but_are_never_cached() {
    let literal = "x".repeat(MAX_CACHED_CODE_LEN + 1);
    let code = format!("\"{literal}\"0.00");
    let mut cache = FormatCache::default();
    let a = parsed_sections_in(&mut cache, &code);
    let b = parsed_sections_in(&mut cache, &code);
    assert!(!Arc::ptr_eq(&a, &b));
    assert!(cache.hot.is_empty() && cache.cold.is_empty());
    assert_eq!(fv(7.625, &code), format!("{literal}7.63"));
}

const CACHE_CORPUS: [&str; 14] = [
    "#,##0.00",
    "[Red]0.00%;\"neg\";\"zero\";\"txt:\"@",
    "m/d/yyyy h:mm",
    "[$-409]dddd, mmmm d, yyyy",
    "_($* #,##0.00_);_($* (#,##0.00);_($* \"-\"??_);_(@_)",
    "0.00E+00",
    "[>100]\"big\";[<0][Red]0.00;0.0",
    "# ??/??",
    "0\\ \\k\\g",
    "@",
    // pairs that diverge only late, so a partial key would alias them
    "#,##0.00;[Red](#,##0.00)",
    "#,##0.00;[Blue](#,##0.00)",
    "0.000000000;;\"zero\";@",
    "0.000000000;;\"nil\";@",
];

/// every lookup below is a proven hit, and each must equal a fresh parse.
#[test]
fn a_cache_hit_returns_what_a_fresh_parse_would() {
    for code in CACHE_CORPUS {
        let mut cache = FormatCache::default();
        let seeded = parsed_sections_in(&mut cache, code);
        assert_eq!(*seeded, *tokenize_all(code), "{code}");

        for j in 1..FORMAT_CACHE_CAP {
            let _ = parsed_sections_in(&mut cache, &format!("\"j{j}\"0"));
        }
        assert_eq!(cache.hot.len(), FORMAT_CACHE_CAP);
        assert!(cache.hot.contains_key(code));
        let hot_hit = parsed_sections_in(&mut cache, code);
        assert!(Arc::ptr_eq(&seeded, &hot_hit), "{code}: not a hot hit");

        let _ = parsed_sections_in(&mut cache, "\"flip\"0");
        assert!(cache.cold.contains_key(code) && !cache.hot.contains_key(code));
        let cold_hit = parsed_sections_in(&mut cache, code);
        assert!(Arc::ptr_eq(&seeded, &cold_hit), "{code}: not a cold hit");
        assert_eq!(*cold_hit, *tokenize_all(code), "{code}");
    }
}

#[test]
fn a_cold_hit_promotes_even_when_hot_is_already_full() {
    let code = "[Red]#,##0.00;(#,##0.00);\"-\";@";
    let mut cache = FormatCache::default();
    let seeded = parsed_sections_in(&mut cache, code);
    for j in 1..FORMAT_CACHE_CAP {
        let _ = parsed_sections_in(&mut cache, &format!("\"a{j}\"0"));
    }
    let _ = parsed_sections_in(&mut cache, "\"flip\"0");
    assert!(cache.cold.contains_key(code));
    for j in 1..FORMAT_CACHE_CAP {
        let _ = parsed_sections_in(&mut cache, &format!("\"b{j}\"0"));
    }
    assert_eq!(cache.hot.len(), FORMAT_CACHE_CAP);

    let promoted = parsed_sections_in(&mut cache, code);
    assert!(Arc::ptr_eq(&seeded, &promoted));
    assert!(cache.hot.contains_key(code));
    assert_eq!(*promoted, *tokenize_all(code));
}

#[test]
fn reinserting_a_present_key_at_capacity_loses_nothing() {
    let mut cache = FormatCache::default();
    let keys: Vec<String> = (0..FORMAT_CACHE_CAP)
        .map(|i| format!("\"r{i}\"0"))
        .collect();
    for key in &keys {
        cache.insert(key, parse_code(key));
    }
    let repeat = &keys[FORMAT_CACHE_CAP / 2];
    cache.insert(repeat, parse_code(repeat));
    assert!(cache.hot.contains_key(repeat));
    for key in &keys {
        assert!(
            cache.hot.contains_key(key) || cache.cold.contains_key(key),
            "{key} was dropped"
        );
    }
}

/// only the first four sections are ever selected, so the parse stops there.
#[test]
fn sections_past_the_fourth_are_never_selected() {
    assert_eq!(fv(1.0, "\"a\";\"b\";\"c\";\"d\";\"e\";\"f\""), "a");
    assert_eq!(fv(-1.0, "\"a\";\"b\";\"c\";\"d\";\"e\""), "b");
    assert_eq!(fv(0.0, "\"a\";\"b\";\"c\";\"d\";\"e\""), "c");
    let text = CellValue::Text {
        value: "x".to_string(),
    };
    assert_eq!(
        format_value(&text, "\"a\";\"b\";\"c\";\"d\"@;\"e\"@", DateSystem::V1900).text,
        "dx"
    );
    assert_eq!(tokenize_all("0;0;0;0;0;0;0").len(), USED_SECTIONS);
}

/// the parse is shared but the date system is not: it is applied after lookup.
#[test]
fn one_cached_code_still_honours_both_date_systems() {
    let code = "m/d/yyyy h:mm";
    let serial = 45_000.5;
    let v1900 = fv(serial, code);
    let v1904 = fv04(serial, code);
    assert_ne!(v1900, v1904);
    for _ in 0..8 {
        assert_eq!(fv04(serial, code), v1904);
        assert_eq!(fv(serial, code), v1900);
    }
}

#[test]
fn concurrent_callers_agree_with_a_single_threaded_render() {
    let codes: Vec<String> = (0..FORMAT_CACHE_CAP + 64)
        .map(|i| format!("\"c{i}\"#,##0.00;[Red](#,##0.00)"))
        .collect();
    let expected: Vec<String> = codes.iter().map(|c| fv(-1234.5678, c)).collect();
    std::thread::scope(|s| {
        for _ in 0..4 {
            let (codes, expected) = (&codes, &expected);
            s.spawn(move || {
                for _ in 0..4 {
                    for (code, want) in codes.iter().zip(expected) {
                        assert_eq!(&fv(-1234.5678, code), want);
                    }
                }
            });
        }
    });
}
