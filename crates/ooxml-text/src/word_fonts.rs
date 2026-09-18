//! Vertical metrics of the East Asian faces Word ships, keyed by the requested
//! family. Read off `head` and `hhea` of Word's own `DFonts` copies (16.113);
//! FangSong and DFKai-SB are Windows-only and take their family's span.
//! A family with no entry keeps its substitute's own metrics.

use crate::font_store::RequestedLineMetrics;

const fn m(units_per_em: u16, hhea_ascender: i16, hhea_descender: i16) -> RequestedLineMetrics {
    RequestedLineMetrics {
        units_per_em,
        hhea_ascender,
        hhea_descender,
    }
}

/// The classic 256-unit Japanese and Simplified Chinese bitmap-era faces, all
/// exactly one em from ascender to descender.
const JIS_256: RequestedLineMetrics = m(256, 220, -36);
/// Batang/Gulim and their fixed-pitch variants — also exactly one em.
const KOREAN_1024: RequestedLineMetrics = m(1024, 879, -145);
/// MingLiU and its variants — one em at 1024 units.
const MINGLIU_1024: RequestedLineMetrics = m(1024, 820, -204);
/// Yu Gothic, Yu Gothic Medium/Light and Yu Mincho.
const YU_2048: RequestedLineMetrics = m(2048, 1802, -455);
/// Malgun Gothic.
const MALGUN_2048: RequestedLineMetrics = m(2048, 2229, -495);

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
    (&["meiryo", "メイリオ"], m(2048, 2171, -901)),
    (&["meiryo ui"], m(2048, 2171, -430)),
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
        m(2048, 2210, -514),
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
    (&["microsoft yahei", "微软雅黑"], m(2048, 2167, -536)),
    (&["microsoft yahei ui"], m(2048, 2080, -521)),
    (&["dengxian", "等线"], m(2048, 1659, -475)),
    // Traditional Chinese.
    (&["microsoft jhenghei", "微軟正黑體"], m(2048, 2203, -521)),
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
];

/// Vertical metrics Word measures `family` with, or `None` for a family this
/// table does not cover. Matching is case-insensitive and trimmed, the same
/// normalization hosts apply to a `w:rFonts` name.
pub fn requested_line_metrics(family: &str) -> Option<RequestedLineMetrics> {
    let key = family.trim().to_lowercase();
    EAST_ASIAN_FACES
        .iter()
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

    #[test]
    fn leaves_latin_and_unlisted_east_asian_families_alone() {
        for family in ["Arial", "Times New Roman", "Calibri", "Century", "HGPｺﾞｼｯｸM"] {
            assert_eq!(requested_line_metrics(family), None, "{family}");
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
