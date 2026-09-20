//! Metrics of the faces Word ships but this package does not bundle, keyed by
//! the requested family. Read off `head` and `hhea` of Word's own copies and
//! cross-checked against the font programs it embeds in its exports. A family
//! with no entry keeps its substitute's own metrics.
//!
//! A family is listed with vertical metrics only when its span is the same
//! under either reading of Word's line rule. Gigi is the measured
//! counter-example — its span is 1.382 em against the last-resort face's
//! 1.150, and correcting it alone lands one corpus document on an extra page.
//!
//! `advance_scale` is the horizontal twin, and stays `1.0` unless the ratio
//! has been measured against the substitute the host actually supplies: a face
//! much wider or narrower than its substitute paginates past Word rather than
//! onto it, and a scale that is right on average is still wrong per glyph, so
//! it buys line and page counts, not line breaks.
//!
//! A family Word has no face for carries its substitute's metrics, taken from
//! the document's `w:altName` or identified against Word's reference render.

use crate::font_store::RequestedLineMetrics;

const fn ea(units_per_em: u16, hhea_ascender: i16, hhea_descender: i16) -> RequestedLineMetrics {
    RequestedLineMetrics {
        units_per_em,
        hhea_ascender,
        hhea_descender,
        hhea_line_gap: 0,
        east_asian: true,
        advance_scale: 1.0,
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
        advance_scale: 1.0,
    }
}

/// [`latin`] for a family whose advance ratio against its substitute has been
/// measured off the font program Word embeds in its own export.
const fn latin_scaled(
    units_per_em: u16,
    hhea_ascender: i16,
    hhea_descender: i16,
    hhea_line_gap: i16,
    advance_scale: f32,
) -> RequestedLineMetrics {
    RequestedLineMetrics {
        units_per_em,
        hhea_ascender,
        hhea_descender,
        hhea_line_gap,
        east_asian: false,
        advance_scale,
    }
}

/// The classic 256-unit Japanese and Simplified Chinese bitmap-era faces, all
/// exactly one em from ascender to descender.
const JIS_256: RequestedLineMetrics = ea(256, 220, -36);
/// Batang/Gulim and their fixed-pitch variants — also exactly one em.
const KOREAN_1024: RequestedLineMetrics = ea(1024, 879, -145);
/// MingLiU and its variants — one em at 1024 units.
const MINGLIU_1024: RequestedLineMetrics = ea(1024, 820, -204);
/// Yu Gothic and Yu Gothic Medium/Light. Yu Mincho is a separate design.
const YU_2048: RequestedLineMetrics = ea(2048, 1802, -455);
/// Yu Mincho — a sixth taller than Yu Gothic, not the same span.
const YU_MINCHO_2048: RequestedLineMetrics = ea(2048, 2038, -598);
/// Malgun Gothic.
const MALGUN_2048: RequestedLineMetrics = ea(2048, 2229, -495);
/// NanumGothic, the Office cloud font Word downloads for `나눔고딕`.
const NANUM_GOTHIC_1000: RequestedLineMetrics = ea(1000, 844, -156);
/// NanumMyeongjo, the Office cloud font Word downloads for `나눔명조`.
const NANUM_MYEONGJO_1024: RequestedLineMetrics = ea(1024, 819, -205);
/// Arial Unicode MS, Word's fallback for a Hangul family it has no face for.
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
        ],
        YU_2048,
    ),
    (&["yu mincho", "游明朝"], YU_MINCHO_2048),
    (
        &["yu gothic ui", "yu gothic ui semilight"],
        ea(2048, 2210, -514),
    ),
    // BIZ UD and Ryumin: Office ships no Mac face, and Word lays every one of
    // them out with Yu Gothic (measured off its own exports).
    (
        &[
            "biz udゴシック",
            "biz udpゴシック",
            "biz udgothic",
            "biz udpgothic",
            "biz ud明朝 medium",
            "biz udp明朝 medium",
            "biz udmincho medium",
            "biz udpmincho medium",
            "ryuminpr5-regular",
        ],
        YU_2048,
    ),
    (
        &["ud デジタル 教科書体 np-b", "ud digi kyokasho np-b"],
        ea(2048, 1802, -567),
    ),
    // The HG family shares the 256-unit design of MS Gothic/Mincho.
    (
        &[
            "hg丸ｺﾞｼｯｸm-pro",
            "hgmarugothicmpro",
            "hgpｺﾞｼｯｸm",
            "hgpgothicm",
            "hgp創英角ｺﾞｼｯｸub",
            "hgpsoeikakugothicub",
            "hg創英角ﾎﾟｯﾌﾟ体",
            "hgp創英角ﾎﾟｯﾌﾟ体",
            "hgs創英角ﾎﾟｯﾌﾟ体",
        ],
        JIS_256,
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
    // Jeju Gothic: Word ships none and falls back to Malgun Gothic.
    (&["제주고딕"], MALGUN_2048),
    // HCR Dotum: Word ships none; both corpus documents alias it to Batang.
    (&["한컴돋움"], KOREAN_1024),
    // Polaris/Hancom Batang compatibility face; Word falls back to Arial Unicode MS.
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
    (&["century"], latin(2048, 2019, -442, 0)),
    (&["century gothic"], latin(2048, 1989, -451, 0)),
    (&["century schoolbook"], latin(2048, 2018, -443, 0)),
    (&["arial narrow"], latin(2048, 1888, -431, 0)),
    (&["agency fb"], latin(2048, 1889, -410, 0)),
    (&["cambria math"], latin(2048, 1595, -455, 353)),
    (&["calibri light"], latin(2048, 1536, -512, 452)),
    (&["roboto"], latin(2048, 1900, -500, 0)),
    // Lucida Bright runs 1.113x the last-resort Liberation Sans it falls back
    // to: its advances over a-z, weighted by English letter frequency and one
    // space per 5.1 letters, against Liberation Sans's. Measured off the
    // subset Word embeds in its own export of `oxi-en-creative-01`. The ratio
    // is the regular face's; Demi measures 1.068 and Italic 1.093, and running
    // text is overwhelmingly regular. Lucida Sans keeps the same span but is a
    // different design, so it stays unscaled until it is measured too.
    (&["lucida bright"], latin_scaled(2048, 1900, -432, 0, 1.113)),
    (&["lucida sans"], latin(2048, 1900, -432, 0)),
    (&["lucida calligraphy"], latin(2048, 1900, -666, 0)),
    (&["wingdings 3"], latin(2048, 1900, -432, 0)),
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
    /// own aliases resolve to a bundled face, and a family this table has no
    /// usable measurement for — Gigi's span is measured but withheld, see the
    /// module doc — is absent.
    #[test]
    fn leaves_bundled_aliased_and_unmeasured_families_alone() {
        for family in [
            "Arial",
            "Times New Roman",
            "Calibri",
            "Cambria",
            "Courier New",
            "Helvetica",
            "Times",
            "Gigi",
            "폴라리스바탕",
        ] {
            assert_eq!(requested_line_metrics(family), None, "{family}");
        }
    }

    /// Lucida Bright is the one family whose advances are corrected, and the
    /// correction is measured against the Liberation Sans the last-resort
    /// chain supplies: per-glyph the two faces run 0.813x (`S`) to 1.558x
    /// (`j`) apart, and 1.113x is where that lands over a-z weighted by
    /// English letter frequency plus one space per 5.1 letters.
    #[test]
    fn lucida_bright_carries_the_advance_ratio_word_embeds() {
        let bright = requested_line_metrics("Lucida Bright").expect("Lucida Bright");
        assert!(
            (bright.advance_scale - 1.113).abs() < 1e-6,
            "{}",
            bright.advance_scale
        );
        let scaled: Vec<&str> = EAST_ASIAN_FACES
            .iter()
            .chain(LATIN_FACES)
            .filter(|(_, metrics)| (metrics.advance_scale - 1.0).abs() >= 1e-6)
            .flat_map(|(names, _)| names.iter().copied())
            .collect();
        assert_eq!(
            scaled,
            ["lucida bright"],
            "an unmeasured family must keep its substitute's advances"
        );
    }

    /// Spans read off the font programs Word embeds in its own exports of the
    /// visual-fidelity corpus. Each of these families measured with the
    /// last-resort face's 1.1499 em span before, which moved every line below
    /// it; Yu Mincho measured with Yu Gothic's 1.102 em.
    #[test]
    fn covers_the_families_word_embeds_in_its_own_exports() {
        for (family, upem, ascender, descender, line_gap, east_asian) in [
            ("Lucida Bright", 2048u16, 1900i16, -432i16, 0i16, false),
            ("Lucida Sans", 2048, 1900, -432, 0, false),
            ("Lucida Calligraphy", 2048, 1900, -666, 0, false),
            ("Century", 2048, 2019, -442, 0, false),
            ("Century Gothic", 2048, 1989, -451, 0, false),
            ("Century Schoolbook", 2048, 2018, -443, 0, false),
            ("Arial Narrow", 2048, 1888, -431, 0, false),
            ("Agency FB", 2048, 1889, -410, 0, false),
            ("Cambria Math", 2048, 1595, -455, 353, false),
            ("Calibri Light", 2048, 1536, -512, 452, false),
            ("Roboto", 2048, 1900, -500, 0, false),
            ("Wingdings 3", 2048, 1900, -432, 0, false),
            ("Yu Mincho", 2048, 2038, -598, 0, true),
            ("游明朝", 2048, 2038, -598, 0, true),
            ("UD デジタル 教科書体 NP-B", 2048, 1802, -567, 0, true),
            ("HGPｺﾞｼｯｸM", 256, 220, -36, 0, true),
            ("HG丸ｺﾞｼｯｸM-PRO", 256, 220, -36, 0, true),
            ("HGP創英角ｺﾞｼｯｸUB", 256, 220, -36, 0, true),
        ] {
            let metrics = requested_line_metrics(family).expect(family);
            assert_eq!(metrics.units_per_em, upem, "{family} upem");
            assert_eq!(metrics.hhea_ascender, ascender, "{family} ascender");
            assert_eq!(metrics.hhea_descender, descender, "{family} descender");
            assert_eq!(metrics.hhea_line_gap, line_gap, "{family} line gap");
            assert_eq!(metrics.east_asian, east_asian, "{family} east asian");
        }
    }

    /// Families Word has no face for, measured against the face its own export
    /// substituted: Jeju Gothic becomes Malgun Gothic, and the BIZ UD and
    /// Ryumin families all become Yu Gothic.
    #[test]
    fn covers_the_japanese_and_korean_families_word_substitutes() {
        for family in ["BIZ UDPゴシック", "BIZ UDゴシック", "BIZ UDP明朝 Medium"] {
            assert_eq!(requested_line_metrics(family), Some(YU_2048), "{family}");
        }
        assert_eq!(requested_line_metrics("제주고딕"), Some(MALGUN_2048));
    }

    /// Korean families Word either downloads from the cloud font catalog or
    /// substitutes for, none of which `@betteroffice/fonts` maps to a bundled
    /// face. Ascender and descender are pinned separately: the East Asian line
    /// box reads them individually and painting takes its baseline from the
    /// ascent, so an equal span is not enough to hold the text in place.
    #[test]
    fn covers_the_korean_families_that_reach_the_last_resort_face() {
        for (family, upem, ascender, descender) in [
            ("나눔고딕", 1000u16, 844i16, -156i16),
            ("NanumGothic", 1000, 844, -156),
            ("나눔명조", 1024, 819, -205),
            ("NanumMyeongjo", 1024, 819, -205),
            ("한컴돋움", 1024, 879, -145),
            ("폴라리스새바탕-함초롬바탕호환", 2048, 2189, -555),
        ] {
            let metrics = requested_line_metrics(family).expect(family);
            assert!(metrics.east_asian, "{family}");
            assert_eq!(metrics.units_per_em, upem, "{family} upem");
            assert_eq!(metrics.hhea_ascender, ascender, "{family} ascender");
            assert_eq!(metrics.hhea_descender, descender, "{family} descender");
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
            ("Yu Mincho", 2636.0 / 2048.0),
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
