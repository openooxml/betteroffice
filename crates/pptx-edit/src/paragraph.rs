//! Font-independent paragraph semantics shared by rendering and export: the layout and master a
//! slide inherits from, the paragraph-property cascade, and list markers.

use ooxml_drawingml::ColorValue;
use pptx_parse::{
    Bullet, ParagraphProperties, Placeholder, PptxPackage, RunProperties, ShapeNode, Slide,
    SlideLayout, SlideMaster, TextBody,
};

use crate::deck::shape_base;

/// The largest number a list marker shows: the largest `startAt` plus one per paragraph a story
/// renders.
pub const MAX_LIST_NUMBER: u32 = 32_767 + 20_000;

/// The source slide, layout and master a slide snapshot inherits from.
#[derive(Clone, Copy, Debug, Default)]
pub struct SlideParents<'a> {
    pub slide: Option<&'a Slide>,
    pub layout: Option<&'a SlideLayout>,
    pub master: Option<&'a SlideMaster>,
}

impl<'a> SlideParents<'a> {
    /// Resolves a slide's parts from its snapshot's source and layout part paths, falling back to
    /// the package's first layout and master.
    pub fn resolve(
        package: &'a PptxPackage,
        source_part_path: Option<&str>,
        layout_part_path: Option<&str>,
    ) -> Self {
        let slide = source_part_path
            .and_then(|path| package.slides.iter().find(|slide| slide.part_path == path));
        let layout_path =
            layout_part_path.or_else(|| slide.and_then(|slide| slide.layout_part_path.as_deref()));
        let layout = layout_path
            .and_then(|path| {
                package
                    .layouts
                    .iter()
                    .find(|layout| layout.part_path == path)
            })
            .or_else(|| package.layouts.first());
        let master = layout
            .and_then(|layout| layout.master_part_path.as_deref())
            .and_then(|path| {
                package
                    .masters
                    .iter()
                    .find(|master| master.part_path == path)
            })
            .or_else(|| {
                layout.and_then(|layout| {
                    package.masters.iter().find(|master| {
                        master
                            .layout_part_paths
                            .iter()
                            .any(|path| path == &layout.part_path)
                    })
                })
            })
            .or_else(|| package.masters.first());
        Self {
            slide,
            layout,
            master,
        }
    }
}

/// The first shape in `nodes`, groups included, whose placeholder matches `target`.
pub fn find_placeholder<'a>(nodes: &'a [ShapeNode], target: &Placeholder) -> Option<&'a ShapeNode> {
    for node in nodes {
        if shape_base(node)
            .placeholder
            .as_ref()
            .is_some_and(|value| placeholders_match(value, target))
        {
            return Some(node);
        }
        if let ShapeNode::Group(group) = node
            && let Some(found) = find_placeholder(&group.children, target)
        {
            return Some(found);
        }
    }
    None
}

/// A slide holds one of each of these, so they inherit by type: PowerPoint writes a slide number
/// as `idx="12"` over a master's `idx="4"` and still draws it where the master put it.
const SINGLETON_PLACEHOLDERS: [&str; 5] = ["title", "sldNum", "dt", "ftr", "hdr"];

/// Placeholders of a singleton type match by normalized type; others match by index when both
/// have one, and by normalized type otherwise.
pub fn placeholders_match(left: &Placeholder, right: &Placeholder) -> bool {
    let left_type = normalize_placeholder_type(left.placeholder_type.as_deref());
    let right_type = normalize_placeholder_type(right.placeholder_type.as_deref());
    if SINGLETON_PLACEHOLDERS.contains(&left_type) || SINGLETON_PLACEHOLDERS.contains(&right_type) {
        return left_type == right_type;
    }
    match (left.index, right.index) {
        (Some(left), Some(right)) => left == right,
        _ => left_type == right_type,
    }
}

/// A placeholder type with its synonyms folded: `ctrTitle` is `title`, `obj` and an absent type
/// are `body`.
pub fn normalize_placeholder_type(value: Option<&str>) -> &str {
    match value.unwrap_or("body") {
        "ctrTitle" => "title",
        "obj" => "body",
        value => value,
    }
}

/// The text bodies a shape's paragraphs inherit from, most specific first.
#[derive(Clone, Copy, Debug, Default)]
pub struct ParagraphCascade<'a> {
    /// The shape's own source text body.
    pub primary: Option<&'a TextBody>,
    /// The matching layout placeholder's text body.
    pub layout: Option<&'a TextBody>,
    /// The matching master placeholder's text body.
    pub master: Option<&'a TextBody>,
    /// The master whose text styles the placeholder type selects.
    pub master_slide: Option<&'a SlideMaster>,
    /// `p:defaultTextStyle`, which outranks the master's `otherStyle` for text outside
    /// placeholders.
    pub default_style: &'a [ParagraphProperties],
    /// `p:defaultTextStyle/a:defPPr`, under every level of `default_style`.
    pub default_paragraph: Option<&'a ParagraphProperties>,
    pub placeholder: Option<&'a Placeholder>,
    /// `p:style/a:fontRef` colour.
    pub style_color: Option<&'a ColorValue>,
}

impl ParagraphCascade<'_> {
    /// The effective properties of paragraph `index` at `level`. An `authored` numbered bullet
    /// replaces the inherited bullet.
    pub fn properties(
        &self,
        index: usize,
        level: u32,
        authored: Option<&Bullet>,
    ) -> ParagraphProperties {
        let mut properties = self
            .master_slide
            .and_then(|master| master_style(master, self.placeholder, level))
            .cloned()
            .unwrap_or_default();
        if self.placeholder.is_none() {
            let level_style = self
                .default_style
                .get(level as usize)
                .or_else(|| self.default_style.first());
            for source in self.default_paragraph.into_iter().chain(level_style) {
                merge_paragraph_properties(&mut properties, source);
            }
        }
        if let Some(color) = self.style_color {
            properties
                .default_run
                .get_or_insert_with(RunProperties::default)
                .color = Some(color.clone());
        }
        for body in [self.master, self.layout, self.primary]
            .into_iter()
            .flatten()
        {
            if let Some(source) = &body.default_list_style {
                merge_paragraph_properties(&mut properties, source);
            }
            if let Some(source) = body.list_style.get(level as usize) {
                merge_paragraph_properties(&mut properties, source);
            }
            if let Some(source) = body
                .paragraphs
                .get(index)
                .or_else(|| body.paragraphs.get(level as usize))
                .map(|paragraph| &paragraph.properties)
            {
                merge_paragraph_properties(&mut properties, source);
            }
        }
        if let Some(Bullet::AutoNumber { restart, .. }) = &mut properties.bullet {
            *restart = self
                .primary
                .and_then(|body| body.paragraphs.get(index))
                .is_some_and(|paragraph| {
                    matches!(
                        paragraph.properties.bullet,
                        Some(Bullet::AutoNumber { restart: true, .. })
                            | Some(Bullet::AutoNumber { start_at: 2.., .. })
                    )
                });
        }
        if let Some(authored @ Bullet::AutoNumber { .. }) = authored {
            properties.bullet = Some(authored.clone());
            if let Some(Bullet::AutoNumber {
                restart, start_at, ..
            }) = &mut properties.bullet
            {
                *restart |= *start_at != 1;
            }
        }
        properties
    }
}

fn master_style<'a>(
    master: &'a SlideMaster,
    placeholder: Option<&Placeholder>,
    level: u32,
) -> Option<&'a ParagraphProperties> {
    let styles = match placeholder {
        Some(placeholder) => {
            match normalize_placeholder_type(placeholder.placeholder_type.as_deref()) {
                "title" => &master.text_styles.title,
                "body" | "subTitle" => &master.text_styles.body,
                _ => &master.text_styles.other,
            }
        }
        None => &master.text_styles.other,
    };
    styles.get(level as usize).or_else(|| styles.first())
}

/// Overlays every property `source` declares onto `target`.
pub fn merge_paragraph_properties(target: &mut ParagraphProperties, source: &ParagraphProperties) {
    if source.alignment.is_some() {
        target.alignment.clone_from(&source.alignment);
    }
    if source.margin_left.is_some() {
        target.margin_left = source.margin_left;
    }
    if source.margin_right.is_some() {
        target.margin_right = source.margin_right;
    }
    if source.indent.is_some() {
        target.indent = source.indent;
    }
    if source.bullet.is_some() {
        target.bullet.clone_from(&source.bullet);
    }
    if source.line_spacing.is_some() {
        target.line_spacing = source.line_spacing;
    }
    if source.space_before.is_some() {
        target.space_before = source.space_before;
    }
    if source.space_after.is_some() {
        target.space_after = source.space_after;
    }
    if source.bullet_font.is_some() {
        target.bullet_font.clone_from(&source.bullet_font);
    }
    if source.bullet_color.is_some() {
        target.bullet_color.clone_from(&source.bullet_color);
    }
    if source.bullet_size.is_some() {
        target.bullet_size.clone_from(&source.bullet_size);
    }
    if source.default_tab_size.is_some() {
        target.default_tab_size = source.default_tab_size;
    }
    if source.tab_stops.is_some() {
        target.tab_stops.clone_from(&source.tab_stops);
    }
    if let Some(source) = &source.default_run {
        let target = target
            .default_run
            .get_or_insert_with(RunProperties::default);
        merge_run_properties(target, source);
    }
}

fn merge_run_properties(target: &mut RunProperties, source: &RunProperties) {
    if source.font_size_pt.is_some() {
        target.font_size_pt = source.font_size_pt;
    }
    if source.bold.is_some() {
        target.bold = source.bold;
    }
    if source.italic.is_some() {
        target.italic = source.italic;
    }
    if source.underline.is_some() {
        target.underline.clone_from(&source.underline);
    }
    if source.font_family.is_some() {
        target.font_family.clone_from(&source.font_family);
    }
    if source.color.is_some() {
        target.color.clone_from(&source.color);
    }
    if source.language.is_some() {
        target.language.clone_from(&source.language);
    }
    if source.spacing_pt.is_some() {
        target.spacing_pt = source.spacing_pt;
    }
    if source.baseline_pct.is_some() {
        target.baseline_pct = source.baseline_pct;
    }
    if source.caps.is_some() {
        target.caps = source.caps;
    }
}

/// A paragraph's list marker, before any symbol-font substitution.
#[derive(Clone, Debug, Eq, PartialEq)]
pub enum ListMarker {
    /// A `buChar` bullet, as authored.
    Character(String),
    /// A `buAutoNum` number; `marker` is `None` for a scheme [`format_autonum`] cannot format.
    Number { value: u32, marker: Option<String> },
}

impl ListMarker {
    /// The marker to draw: an unformattable scheme falls back to `"{value}."`.
    pub fn text(&self) -> String {
        match self {
            Self::Character(value) => value.clone(),
            Self::Number { value, marker } => marker.clone().unwrap_or_else(|| format!("{value}.")),
        }
    }
}

/// Per-level `a:buAutoNum` state of one story: the number last shown and the `startAt` the run
/// was seeded from.
#[derive(Clone, Debug, Default)]
pub struct ListCounters {
    numbers: [u32; 9],
    starts: [u32; 9],
}

impl ListCounters {
    /// The marker of the story's next paragraph. A paragraph without text is spacing, not an
    /// item: PowerPoint neither marks it nor counts it towards the next number.
    pub fn next(
        &mut self,
        bullet: Option<&Bullet>,
        level: u32,
        has_text: bool,
    ) -> Option<ListMarker> {
        if !has_text {
            return None;
        }
        let level = (level as usize).min(self.numbers.len() - 1);
        self.numbers[level + 1..].fill(0);
        self.starts[level + 1..].fill(0);
        match bullet {
            Some(Bullet::AutoNumber {
                scheme,
                start_at,
                restart,
            }) => {
                let start = (*start_at).clamp(1, 32_767);
                // PowerPoint writes the list's `startAt` on every one of its paragraphs, so
                // repeating the seed continues the run; only a different declared start opens a
                // new list.
                self.numbers[level] = match self.numbers[level] {
                    0 => start,
                    _ if *restart && self.starts[level] != start => start,
                    current => current.saturating_add(1),
                };
                if self.numbers[level] == start {
                    self.starts[level] = start;
                }
                let value = self.numbers[level].clamp(1, MAX_LIST_NUMBER);
                Some(ListMarker::Number {
                    value,
                    marker: format_autonum(value, scheme),
                })
            }
            _ => {
                self.numbers[level] = 0;
                self.starts[level] = 0;
                match bullet {
                    Some(Bullet::Character { value }) if !value.trim().is_empty() => {
                        Some(ListMarker::Character(value.clone()))
                    }
                    _ => None,
                }
            }
        }
    }
}

/// Formats a Latin, Roman or decimal `ST_TextAutonumberScheme` marker; `None` for any other
/// scheme.
pub fn format_autonum(value: u32, scheme: &str) -> Option<String> {
    let value = value.clamp(1, MAX_LIST_NUMBER);
    let (numeral, suffix) = ["ParenBoth", "ParenR", "Period", "Plain"]
        .into_iter()
        .find_map(|suffix| Some((scheme.strip_suffix(suffix)?, suffix)))?;
    let body = match numeral {
        "alphaLc" => format_alpha(value, false),
        "alphaUc" => format_alpha(value, true),
        "romanLc" => format_roman(value, false),
        "romanUc" => format_roman(value, true),
        "arabic" => value.to_string(),
        _ => return None,
    };
    Some(match suffix {
        "ParenBoth" => format!("({body})"),
        "ParenR" => format!("{body})"),
        "Plain" => body,
        _ => format!("{body}."),
    })
}

fn format_alpha(value: u32, upper: bool) -> String {
    let base = if upper { b'A' } else { b'a' };
    let mut value = value.max(1);
    let mut out = Vec::new();
    while value > 0 {
        let index = (value - 1) % 26;
        out.push(base + index as u8);
        value = (value - 1) / 26;
    }
    out.reverse();
    String::from_utf8(out).unwrap_or_default()
}

fn format_roman(value: u32, upper: bool) -> String {
    const NUMERALS: [(u32, &str); 13] = [
        (1000, "m"),
        (900, "cm"),
        (500, "d"),
        (400, "cd"),
        (100, "c"),
        (90, "xc"),
        (50, "l"),
        (40, "xl"),
        (10, "x"),
        (9, "ix"),
        (5, "v"),
        (4, "iv"),
        (1, "i"),
    ];
    let mut value = value.max(1);
    let mut out = String::new();
    for (amount, numeral) in NUMERALS {
        while value >= amount {
            out.push_str(numeral);
            value -= amount;
        }
    }
    if upper { out.to_uppercase() } else { out }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn marker(counters: &mut ListCounters, bullet: Option<&Bullet>, level: u32) -> Option<String> {
        counters
            .next(bullet, level, true)
            .map(|marker| marker.text())
    }

    fn number(scheme: &str, start_at: u32, restart: bool) -> Bullet {
        Bullet::AutoNumber {
            scheme: scheme.to_owned(),
            start_at,
            restart,
        }
    }

    fn character(value: &str) -> Bullet {
        Bullet::Character {
            value: value.to_owned(),
        }
    }

    #[test]
    fn autonumbering_counts_per_level_and_resumes_across_other_levels() {
        let mut counters = ListCounters::default();
        let arabic = number("arabicPeriod", 1, false);
        let dash = character("-");
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
        assert_eq!(marker(&mut counters, Some(&dash), 1).as_deref(), Some("-"));
        assert_eq!(marker(&mut counters, Some(&dash), 1).as_deref(), Some("-"));
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("2.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("3.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("2.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("4.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("1.")
        );
        assert_eq!(marker(&mut counters, Some(&Bullet::None), 0), None);
    }

    #[test]
    fn inherited_autonumber_start_at_applies_only_to_the_first_item() {
        let mut counters = ListCounters::default();
        let seven = number("arabicPeriod", 7, false);
        assert_eq!(
            marker(&mut counters, Some(&seven), 0).as_deref(),
            Some("7.")
        );
        assert_eq!(
            marker(&mut counters, Some(&seven), 0).as_deref(),
            Some("8.")
        );
        let mut counters = ListCounters::default();
        let last_start = number("arabicPeriod", 32_767, false);
        assert_eq!(
            marker(&mut counters, Some(&last_start), 0).as_deref(),
            Some("32767.")
        );
        assert_eq!(
            marker(&mut counters, Some(&last_start), 0).as_deref(),
            Some("32768.")
        );
    }

    #[test]
    fn autonumber_schemes_format_their_numeral_and_suffix() {
        assert_eq!(format_autonum(4, "arabicPeriod").as_deref(), Some("4."));
        assert_eq!(format_autonum(4, "arabicParenR").as_deref(), Some("4)"));
        assert_eq!(format_autonum(4, "arabicParenBoth").as_deref(), Some("(4)"));
        assert_eq!(format_autonum(4, "arabicPlain").as_deref(), Some("4"));
        assert_eq!(format_autonum(1, "alphaLcParenR").as_deref(), Some("a)"));
        assert_eq!(format_autonum(27, "alphaUcPeriod").as_deref(), Some("AA."));
        assert_eq!(format_autonum(9, "romanLcPeriod").as_deref(), Some("ix."));
        assert_eq!(
            format_autonum(2024, "romanUcPeriod").as_deref(),
            Some("MMXXIV.")
        );
        for scheme in ["somethingElse", "somethingElsePlain", "arabicDbPeriod"] {
            assert_eq!(format_autonum(3, scheme), None);
        }
    }

    #[test]
    fn unformattable_schemes_fall_back_to_a_decimal_marker() {
        let mut counters = ListCounters::default();
        let circled = number("circleNumDbPlain", 3, false);
        assert_eq!(
            counters.next(Some(&circled), 0, true),
            Some(ListMarker::Number {
                value: 3,
                marker: None
            })
        );
        assert_eq!(
            marker(&mut counters, Some(&circled), 0).as_deref(),
            Some("4.")
        );
    }

    #[test]
    fn paragraphs_without_text_neither_mark_nor_count() {
        let mut counters = ListCounters::default();
        let arabic = number("arabicPeriod", 1, false);
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
        assert_eq!(counters.next(Some(&arabic), 0, false), None);
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("2.")
        );
    }

    #[test]
    fn autonumber_sequences_restart_after_plain_paragraphs_and_explicit_starts() {
        let mut counters = ListCounters::default();
        let arabic = number("arabicPeriod", 1, false);
        let restart = number("arabicPeriod", 7, true);
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&restart), 0).as_deref(),
            Some("7.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("8.")
        );
        assert_eq!(marker(&mut counters, None, 0), None);
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&character("•")), 0).as_deref(),
            Some("•")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 1).as_deref(),
            Some("1.")
        );
        assert_eq!(
            marker(&mut counters, Some(&arabic), 0).as_deref(),
            Some("1.")
        );
    }

    /// `pptarena-018-original` slide 11 declares `startAt="4"` on all four of its paragraphs and
    /// PowerPoint renders 4, 5, 6, 7; `pptarena-034-original` slide 11 declares `startAt="1"` on
    /// all five and PowerPoint renders a) through e).
    #[test]
    fn a_repeated_declared_start_continues_the_list() {
        let mut counters = ListCounters::default();
        let arabic = number("arabicPeriod", 4, true);
        let alpha = number("alphaLcParenR", 1, true);
        assert_eq!(
            (0..4)
                .filter_map(|_| marker(&mut counters, Some(&arabic), 0))
                .collect::<Vec<_>>(),
            ["4.", "5.", "6.", "7."]
        );
        let mut counters = ListCounters::default();
        assert_eq!(
            (0..5)
                .filter_map(|_| marker(&mut counters, Some(&alpha), 0))
                .collect::<Vec<_>>(),
            ["a)", "b)", "c)", "d)", "e)"]
        );
    }

    #[test]
    fn autonumber_roman_markers_bound_untrusted_start_values() {
        assert_eq!(format_autonum(0, "arabicPeriod").as_deref(), Some("1."));
        let largest = format_autonum(u32::MAX, "romanUcPeriod").unwrap();
        assert_eq!(largest.len(), 61);
        assert_eq!(largest, format!("{}DCCLXVII.", "M".repeat(52)));
    }

    #[test]
    fn text_outside_placeholders_takes_the_presentation_default_style() {
        let sized = |size| ParagraphProperties {
            default_run: Some(RunProperties {
                font_size_pt: Some(size),
                ..RunProperties::default()
            }),
            ..ParagraphProperties::default()
        };
        let levels = [sized(27.0), sized(20.0)];
        let under = ParagraphProperties {
            bullet: Some(character("-")),
            ..sized(12.0)
        };
        let body = Placeholder {
            placeholder_type: Some("body".to_owned()),
            index: Some(1),
            orientation: None,
            size: None,
        };
        let cascade = |placeholder| ParagraphCascade {
            default_style: &levels,
            default_paragraph: Some(&under),
            placeholder,
            ..ParagraphCascade::default()
        };
        let size = |properties: ParagraphProperties| properties.default_run.unwrap().font_size_pt;
        let plain = cascade(None).properties(0, 1, None);
        assert_eq!(plain.bullet, Some(character("-")));
        assert_eq!(size(plain), Some(20.0));
        assert_eq!(size(cascade(None).properties(0, 4, None)), Some(27.0));
        let placed = cascade(Some(&body)).properties(0, 1, None);
        assert_eq!((placed.bullet, placed.default_run), (None, None));
    }

    #[test]
    fn singleton_placeholders_match_by_type_and_others_by_index() {
        let placeholder = |kind: &str, index| Placeholder {
            placeholder_type: Some(kind.to_owned()),
            index,
            orientation: None,
            size: None,
        };
        assert!(placeholders_match(
            &placeholder("sldNum", Some(12)),
            &placeholder("sldNum", Some(4))
        ));
        assert!(!placeholders_match(
            &placeholder("title", Some(4)),
            &placeholder("body", Some(4))
        ));
        assert!(placeholders_match(
            &placeholder("ctrTitle", None),
            &placeholder("title", Some(0))
        ));
        assert!(placeholders_match(
            &placeholder("body", Some(4)),
            &placeholder("obj", Some(4))
        ));
    }
}
