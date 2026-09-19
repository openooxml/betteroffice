//! Vertical metrics of the faces Word ships but this package does not bundle,
//! keyed by the requested family. Read off `head` and `hhea` of Word's own
//! copies and cross-checked against the font programs it embeds in its exports.
//! A family with no entry keeps its substitute's own metrics.
//!
//! A family is listed only when its span is the same under either reading of
//! Word's line rule, and when the substitute's advances are already close: the
//! view moves vertical metrics only, so correcting the pitch of a much
//! narrower or wider face moves a document past Word rather than onto it.
//!
//! A family Word has no face for still measures at something: Word substitutes,
//! and the entry carries that substitute's metrics, read off the document's own
//! `w:altName` where Word ships that face and otherwise identified against
//! Word's reference render. A family whose substitute could not be identified
//! is left out.

use crate::font_store::RequestedLineMetrics;

const fn ea(units_per_em: u16, hhea_ascender: i16, hhea_descender: i16) -> RequestedLineMetrics {
    RequestedLineMetrics {
        units_per_em,
        hhea_ascender,
        hhea_descender,
        hhea_line_gap: 0,
        east_asian: true,
    }
}

const fn latin(
    units_per_em: u16,
    hhea_ascender: i16,
    hhea_descender: i16,
    hhea_line_gap: i16,
) -> RequestedLineMetrics {
    RequestedLineMetrics {
        units_per_em,
        hhea_ascender,
        hhea_descender,
        hhea_line_gap,
        east_asian: false,
    }
}

/// The classic 256-unit Japanese and Simplified Chinese bitmap-era faces, all
/// exactly one em from ascender to descender.
const JIS_256: RequestedLineMetrics = ea(256, 220, -36);
/// Batang/Gulim and their fixed-pitch variants — also exactly one em.
const KOREAN_1024: RequestedLineMetrics = ea(1024, 879, -145);
/// MingLiU and its variants — one em at 1024 units.
const MINGLIU_1024: RequestedLineMetrics = ea(1024, 820, -204);
/// Yu Gothic, Yu Gothic Medium/Light and Yu Mincho.
const YU_2048: RequestedLineMetrics = ea(2048, 1802, -455);
/// Malgun Gothic.
const MALGUN_2048: RequestedLineMetrics = ea(2048, 2229, -495);
/// NanumGothic, the Office cloud font Word downloads for `나눔고딕`.
const NANUM_GOTHIC_1000: RequestedLineMetrics = ea(1000, 844, -156);
/// NanumMyeongjo, the Office cloud font Word downloads for `나눔명조`.
const NANUM_MYEONGJO_1024: RequestedLineMetrics = ea(1024, 819, -205);
/// Arial Unicode MS — what Word substitutes for a Hangul-bearing family it has
/// no face for and cannot resolve through `w:altName` either.
const ARIAL_UNICODE_2048: RequestedLineMetrics = ea(2048, 2189, -555);

/// Requested family (lowercased) -> the vertical metrics Word measures it with.
const EAST_ASIAN_FACES: &[(&[&str], RequestedLineMetrics)] = &[
    // Japanese — MS Mincho / MS Gothic and their proportional variants.
    (
        &[
            "ms mincho",
            "ms pmincho",
            "ｍｓ 明朝",
            "ｍｓ ｐ明朝",
            "ms gothic",
            "ms pgothic",
            "ms ui gothic",
            "ｍｓ ゴシック",
            "ｍｓ ｐゴシック",
        ],
        JIS_256,
    ),
    (&["meiryo", "メイリオ"], ea(2048, 2171, -901)),
    (&["meiryo ui"], ea(2048, 2171, -430)),
    (
        &[
            "yu gothic",
            "yu gothic medium",
            "yu gothic light",
            "游ゴシック",
            "yu mincho",
            "游明朝",
        ],
        YU_2048,
    ),
    (
        &["yu gothic ui", "yu gothic ui semilight"],
        ea(2048, 2210, -514),
    ),
    // Simplified Chinese — the SimSun family shares the JIS 256-unit design.
    (
        &[
            "simsun",
            "nsimsun",
            "simhei",
            "kaiti",
            "fangsong",
            "宋体",
            "新宋体",
            "黑体",
            "楷体",
            "仿宋",
        ],
        JIS_256,
    ),
    (&["microsoft yahei", "微软雅黑"], ea(2048, 2167, -536)),
    (&["microsoft yahei ui"], ea(2048, 2080, -521)),
    (&["dengxian", "等线"], ea(2048, 1659, -475)),
    // Traditional Chinese.
    (&["microsoft jhenghei", "微軟正黑體"], ea(2048, 2203, -521)),
    (
        &[
            "mingliu",
            "pmingliu",
            "mingliu_hkscs",
            "mingliu-extb",
            "pmingliu-extb",
            "新細明體",
            "細明體",
            "dfkai-sb",
            "標楷體",
        ],
        MINGLIU_1024,
    ),
    // Korean.
    (&["malgun gothic", "맑은 고딕"], MALGUN_2048),
    (
        &[
            "batang",
            "batangche",
            "gungsuh",
            "gungsuhche",
            "gulim",
            "gulimche",
            "dotum",
            "dotumche",
            "바탕",
            "바탕체",
            "궁서",
            "궁서체",
            "굴림",
            "굴림체",
            "돋움",
            "돋움체",
        ],
        KOREAN_1024,
    ),
    (
        &["나눔고딕", "nanumgothic", "nanum gothic"],
        NANUM_GOTHIC_1000,
    ),
    (
        &["나눔명조", "nanummyeongjo", "nanum myeongjo"],
        NANUM_MYEONGJO_1024,
    ),
    // Hancom's HCR Dotum, which Word never ships: both corpus documents that
    // name it declare `바탕` as its alternate, and Word's reference render
    // draws it in Batang.
    (&["한컴돋움"], KOREAN_1024),
    // The Polaris/Hancom Batang compatibility face. Word has neither it nor
    // the `폴라리스바탕` its `w:altName` names, and falls back to Arial
    // Unicode MS.
    (&["폴라리스새바탕-함초롬바탕호환"], ARIAL_UNICODE_2048),
];

/// The Segoe UI family — UI, Symbol and Emoji ship the same vertical design.
const SEGOE_UI_2048: RequestedLineMetrics = latin(2048, 2210, -514, 0);

/// Requested Latin family (lowercased) -> the vertical metrics Word measures
/// it with. Every entry's span is unambiguous (see the module doc).
const LATIN_FACES: &[(&[&str], RequestedLineMetrics)] = &[
    (&["lato"], latin(2000, 1974, -426, 0)),
    (&["open sans"], latin(2048, 2189, -600, 0)),
    (&["source sans pro"], latin(1000, 984, -273, 0)),
    (&["playfair display"], latin(1000, 1082, -251, 0)),
    (
        &["segoe ui", "segoe ui symbol", "segoe ui emoji"],
        SEGOE_UI_2048,
    ),
    (&["aptos", "aptos display"], latin(2048, 1923, -577, 0)),
    (&["tahoma"], latin(2048, 2049, -423, 0)),
    (&["verdana"], latin(2048, 2059, -430, 0)),
    (&["trebuchet ms"], latin(2048, 1923, -455, 0)),
    (&["symbol"], latin(2048, 2059, -450, 0)),
    (&["wingdings"], latin(2048, 1841, -432, 0)),
    (&["lucida sans unicode"], latin(2048, 2246, -901, 0)),
    (&["georgia"], latin(2048, 1878, -449, 0)),
    (&["comic sans ms"], latin(2048, 2257, -597, 0)),
];

/// Vertical metrics Word measures `family` with, or `None` for a family this
/// table does not cover. Matching is case-insensitive and trimmed, the same
/// normalization hosts apply to a `w:rFonts` name.
pub fn requested_line_metrics(family: &str) -> Option<RequestedLineMetrics> {
    let key = family.trim().to_lowercase();
    EAST_ASIAN_FACES
        .iter()
        .chain(LATIN_FACES)
        .find(|(names, _)| names.contains(&key.as_str()))
        .map(|&(_, metrics)| metrics)
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Every Word family `@betteroffice/fonts` substitutes a Noto CJK face for.
    /// A miss measures that family with Noto's 1.448 em span.
    #[test]
    fn covers_the_families_the_bundled_cjk_faces_stand_in_for() {
        for family in [
            "MS Mincho",
            "MS PMincho",
            "ＭＳ 明朝",
            "ＭＳ Ｐ明朝",
            "MS Gothic",
            "MS PGothic",
            "ＭＳ ゴシック",
            "ＭＳ Ｐゴシック",
            "Meiryo",
            "メイリオ",
            "Meiryo UI",
            "Yu Gothic",
            "游ゴシック",
            "Yu Mincho",
            "游明朝",
            "SimSun",
            "NSimSun",
            "SimHei",
            "KaiTi",
            "FangSong",
            "宋体",
            "黑体",
            "楷体",
            "仿宋",
            "Microsoft YaHei",
            "微软雅黑",
            "DengXian",
            "等线",
            "Microsoft JhengHei",
            "微軟正黑體",
            "MingLiU",
            "PMingLiU",
            "新細明體",
            "細明體",
            "DFKai-SB",
            "標楷體",
            "Malgun Gothic",
            "맑은 고딕",
            "Gulim",
            "Dotum",
            "Batang",
            "Gungsuh",
            "굴림",
            "굴림체",
            "돋움",
            "돋움체",
            "바탕",
            "바탕체",
            "궁서",
            "궁서체",
        ] {
            assert!(
                requested_line_metrics(family).is_some(),
                "{family} has no vertical metrics"
            );
        }
    }

    #[test]
    fn matches_case_insensitively_and_trims() {
        let expected = Some(JIS_256);
        assert_eq!(requested_line_metrics("  ms mincho "), expected);
        assert_eq!(requested_line_metrics("MS MINCHO"), expected);
        assert_eq!(requested_line_metrics("ＭＳ 明朝"), expected);
    }

    /// Bundled families keep the metric-compatible face's own metrics, Word's
    /// own aliases resolve to a bundled face, and a family whose span depends
    /// on which reading of the line rule applies is deliberately absent.
    #[test]
    fn leaves_bundled_aliased_and_ambiguous_families_alone() {
        for family in [
            "Arial",
            "Times New Roman",
            "Calibri",
            "Cambria",
            "Courier New",
            "Helvetica",
            "Times",
            "Lucida Bright",
            "Lucida Sans",
            "Century",
            "Century Gothic",
            "Arial Narrow",
            "Roboto",
            "Cambria Math",
            "Gigi",
            "HGPｺﾞｼｯｸM",
            "제주고딕",
            "폴라리스바탕",
        ] {
            assert_eq!(requested_line_metrics(family), None, "{family}");
        }
    }

    /// Korean families Word either downloads from the cloud font catalog or
    /// substitutes for, none of which `@betteroffice/fonts` maps to a bundled
    /// face: a miss measures them with the Latin last-resort win box.
    #[test]
    fn covers_the_korean_families_that_reach_the_last_resort_face() {
        for (family, span_em) in [
            ("나눔고딕", 1.0),
            ("NanumGothic", 1.0),
            ("나눔명조", 1.0),
            ("NanumMyeongjo", 1.0),
            ("한컴돋움", 1.0),
            ("폴라리스새바탕-함초롬바탕호환", 2744.0 / 2048.0),
        ] {
            let metrics = requested_line_metrics(family).expect(family);
            assert!(metrics.east_asian, "{family}");
            let measured = (f32::from(metrics.hhea_ascender) - f32::from(metrics.hhea_descender))
                / f32::from(metrics.units_per_em);
            assert!(
                (measured - span_em).abs() < 1e-4,
                "{family}: {measured} vs {span_em}"
            );
        }
    }

    #[test]
    fn latin_entries_do_not_claim_the_east_asian_pitch() {
        for family in ["Lato", "Open Sans", "Verdana", "Wingdings"] {
            let metrics = requested_line_metrics(family).expect(family);
            assert!(!metrics.east_asian, "{family}");
        }
        assert!(
            requested_line_metrics("MS Mincho")
                .expect("mincho")
                .east_asian
        );
    }

    #[test]
    fn latin_spans_match_the_faces_word_ships() {
        for (family, span_em) in [
            ("Lato", 2400.0 / 2000.0),
            ("Open Sans", 2789.0 / 2048.0),
            ("Source Sans Pro", 1257.0 / 1000.0),
            ("Playfair Display", 1333.0 / 1000.0),
            ("Segoe UI", 2724.0 / 2048.0),
            ("Segoe UI Symbol", 2724.0 / 2048.0),
            ("Aptos", 2500.0 / 2048.0),
            ("Tahoma", 2472.0 / 2048.0),
            ("Verdana", 2489.0 / 2048.0),
            ("Trebuchet MS", 2378.0 / 2048.0),
            ("Symbol", 2509.0 / 2048.0),
            ("Wingdings", 2273.0 / 2048.0),
            ("Lucida Sans Unicode", 3147.0 / 2048.0),
            ("Georgia", 2327.0 / 2048.0),
            ("Comic Sans MS", 2854.0 / 2048.0),
        ] {
            let metrics = requested_line_metrics(family).expect(family);
            let measured = (f32::from(metrics.hhea_ascender) - f32::from(metrics.hhea_descender)
                + f32::from(metrics.hhea_line_gap))
                / f32::from(metrics.units_per_em);
            assert!(
                (measured - span_em).abs() < 1e-4,
                "{family}: {measured} vs {span_em}"
            );
        }
    }

    #[test]
    fn spans_match_the_faces_word_ships() {
        for (family, span_em) in [
            ("MS Mincho", 1.0),
            ("SimSun", 1.0),
            ("MingLiU", 1.0),
            ("Batang", 1.0),
            ("Yu Gothic", 2257.0 / 2048.0),
            ("Meiryo", 1.5),
            ("Meiryo UI", 1.27),
            ("Malgun Gothic", 2724.0 / 2048.0),
            ("Microsoft YaHei", 2703.0 / 2048.0),
        ] {
            let metrics = requested_line_metrics(family).expect(family);
            let measured = (metrics.hhea_ascender as f32 - metrics.hhea_descender as f32)
                / metrics.units_per_em as f32;
            assert!(
                (measured - span_em).abs() < 1e-4,
                "{family}: {measured} vs {span_em}"
            );
        }
    }
}
