//! PPTX packaging around the shared DrawingML chart part.

use ooxml_drawingml::chart::{ChartSpace, ChartXml, parse_chart_space};
use ooxml_drawingml::{Theme, resolve_color_value_to_hex_with_theme};

use crate::drawing::parse_color_container;
use crate::xml::XmlElement;

/// Deeper elements are unreachable for the shared parser, which descends at
/// most 64 levels from any node it searches.
const MAX_CHART_DEPTH: usize = 256;
/// Elements one chart part may present to the shared parser. Well past what
/// the parser itself will read, and it keeps a hostile part from turning
/// package memory into adapter memory.
const MAX_CHART_ELEMENTS: usize = 500_000;

/// A chart element paired with the deck theme, so `a:solidFill` resolves
/// through the presentation's colours instead of Office defaults.
pub(crate) struct ChartElement<'a> {
    element: &'a XmlElement,
    theme: &'a Theme,
    children: Vec<ChartElement<'a>>,
}

impl<'a> ChartElement<'a> {
    fn new(element: &'a XmlElement, theme: &'a Theme, budget: &mut usize, depth: usize) -> Self {
        let mut children = Vec::new();
        if depth < MAX_CHART_DEPTH {
            for child in element.child_elements() {
                let Some(remaining) = budget.checked_sub(1) else {
                    break;
                };
                *budget = remaining;
                children.push(Self::new(child, theme, budget, depth + 1));
            }
        }
        Self {
            element,
            theme,
            children,
        }
    }
}

impl ChartXml for ChartElement<'_> {
    fn local_name(&self) -> &str {
        self.element.local_name()
    }

    fn attribute(&self, prefix: Option<&str>, name: &str) -> Option<&str> {
        prefix
            .and_then(|prefix| self.element.attribute(&format!("{prefix}:{name}")))
            .or_else(|| self.element.attribute(name))
    }

    fn child_elements(&self) -> impl Iterator<Item = &Self> {
        self.children.iter()
    }

    fn descendant_text(&self) -> String {
        self.element.text_content()
    }

    fn solid_fill_hex(&self) -> Option<String> {
        resolve_color_value_to_hex_with_theme(
            parse_color_container(self.element).as_ref(),
            Some(self.theme),
        )
    }
}

/// Parse a `c:chartSpace` root, resolving its colours through `theme`.
pub(crate) fn parse_chart_part(root: &XmlElement, theme: &Theme) -> Option<ChartSpace> {
    let mut budget = MAX_CHART_ELEMENTS;
    parse_chart_space(&ChartElement::new(root, theme, &mut budget, 0))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xml::{ParseBudget, ParseLimits, parse_xml};

    fn parse(xml: &str, theme: &Theme) -> Option<ChartSpace> {
        let limits = ParseLimits::default();
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(xml.as_bytes(), "ppt/charts/chart1.xml", &mut budget).unwrap();
        parse_chart_part(&root, theme)
    }

    /// The same chart with every element name qualified by `ns`.
    fn fixture(ns: &str) -> String {
        format!(
            r#"<{ns}chartSpace xmlns:c="c" xmlns:a="a"><{ns}chart>
                 <{ns}title><{ns}tx><{ns}rich><a:p><a:r><a:t> Sales </a:t></a:r></a:p></{ns}rich></{ns}tx></{ns}title>
                 <{ns}legend><{ns}legendPos val="l"/></{ns}legend>
                 <{ns}plotArea>
                   <{ns}barChart><{ns}barDir val="bar"/><{ns}grouping val="stacked"/>
                     <{ns}ser>
                       <{ns}tx><{ns}v>North</{ns}v></{ns}tx>
                       <{ns}spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></{ns}spPr>
                       <{ns}cat><{ns}strRef><{ns}strCache><{ns}pt><{ns}v>Q1</{ns}v></{ns}pt></{ns}strCache></{ns}strRef></{ns}cat>
                       <{ns}val><{ns}numRef><{ns}numCache><{ns}pt><{ns}v>7</{ns}v></{ns}pt></{ns}numCache></{ns}numRef></{ns}val>
                     </{ns}ser>
                     <{ns}axId val="1"/>
                   </{ns}barChart>
                   <{ns}catAx><{ns}axId val="1"/><{ns}axPos val="b"/></{ns}catAx>
                   <{ns}valAx><{ns}axId val="2"/><{ns}scaling><{ns}min val="0"/><{ns}max val="9"/></{ns}scaling></{ns}valAx>
                 </{ns}plotArea>
               </{ns}chart></{ns}chartSpace>"#
        )
    }

    fn themed() -> Theme {
        let mut theme = Theme::default();
        theme.color_scheme.accent1 = "6254E7".to_owned();
        theme
    }

    #[test]
    fn a_namespace_prefixed_part_parses_the_same_as_a_bare_one() {
        let theme = themed();
        let prefixed = parse(&fixture("c:"), &theme).expect("prefixed parses");
        let bare = parse(&fixture(""), &theme).expect("bare parses");
        assert_eq!(prefixed, bare);
        assert_eq!(prefixed.chart_type, "bar");
        assert_eq!(prefixed.title.as_deref(), Some("Sales"));
        assert_eq!(prefixed.legend.unwrap().position.as_deref(), Some("left"));
        assert_eq!(prefixed.series[0].name.as_deref(), Some("North"));
        assert_eq!(prefixed.series[0].categories, ["Q1"]);
        assert_eq!(prefixed.series[0].values, [7.0]);
        assert_eq!(prefixed.plot_groups[0].grouping.as_deref(), Some("stacked"));
        assert_eq!(prefixed.axis_list.as_ref().unwrap().len(), 2);
    }

    #[test]
    fn series_colours_resolve_through_the_deck_theme() {
        assert_eq!(
            parse(&fixture("c:"), &themed()).unwrap().series[0].color,
            "#6254E7"
        );
        assert_eq!(
            parse(&fixture("c:"), &Theme::default()).unwrap().series[0].color,
            "#4472C4"
        );
    }

    #[test]
    fn theme_tints_shades_and_direct_colours_resolve_per_point() {
        let chart = parse(
            r#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:pieChart><c:ser>
                 <c:val><c:numCache><c:pt><c:v>3</c:v></c:pt><c:pt><c:v>1</c:v></c:pt></c:numCache></c:val>
                 <c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:solidFill></c:spPr></c:dPt>
                 <c:dPt><c:idx val="1"/><c:spPr><a:solidFill><a:srgbClr val="112233"/></a:solidFill></c:spPr></c:dPt>
               </c:ser></c:pieChart></c:plotArea></c:chart></c:chartSpace>"#,
            &themed(),
        )
        .expect("chart parses");
        let points = chart.plot_groups[0].series[0].points.as_ref().unwrap();
        assert_eq!(points[0].color, "#B1AAF3");
        assert_eq!(points[1].color, "#112233");
    }

    #[test]
    fn a_qualified_val_attribute_resolves_through_the_prefix_fallback() {
        let chart = parse(
            r#"<c:chartSpace xmlns:c="c"><c:chart><c:legend><c:legendPos c:val="b"/></c:legend><c:plotArea><c:pieChart/></c:plotArea></c:chart></c:chartSpace>"#,
            &Theme::default(),
        )
        .expect("chart parses");
        assert_eq!(chart.legend.unwrap().position.as_deref(), Some("bottom"));
    }

    #[test]
    fn a_part_without_a_recognized_plot_yields_nothing() {
        assert!(parse(r#"<c:chartSpace xmlns:c="c"/>"#, &Theme::default()).is_none());
        assert!(
            parse(
                r#"<c:chartSpace xmlns:c="c"><c:chart><c:plotArea/></c:chart></c:chartSpace>"#,
                &Theme::default()
            )
            .is_none()
        );
    }

    #[test]
    fn a_deeply_nested_part_is_truncated_instead_of_overflowing() {
        let depth = MAX_CHART_DEPTH + 8;
        let nested = format!(
            r#"<c:chartSpace xmlns:c="c">{}<c:plotArea><c:lineChart/></c:plotArea>{}</c:chartSpace>"#,
            "<c:x>".repeat(depth),
            "</c:x>".repeat(depth)
        );
        let limits = ParseLimits {
            max_xml_depth: depth + 8,
            ..ParseLimits::default()
        };
        let mut budget = ParseBudget::new(&limits);
        let root = parse_xml(nested.as_bytes(), "ppt/charts/chart1.xml", &mut budget).unwrap();
        assert!(parse_chart_part(&root, &Theme::default()).is_none());
    }
}
