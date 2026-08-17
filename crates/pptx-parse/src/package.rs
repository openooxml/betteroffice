use std::collections::{BTreeMap, HashMap, HashSet};

use ooxml_drawingml::Theme;

use crate::chart::parse_chart_part;
use crate::drawing::{common_slide_data, parse_text_styles};
use crate::model::*;
use crate::relationships::{Relationship, parse_relationships, relationship_types};
use crate::theme::parse_theme;
use crate::xml::{ParseBudget, XmlElement, parse_xml};
use crate::{ParseLimits, PptxError};

pub fn parse_pptx(data: &[u8]) -> Result<PptxPackage, PptxError> {
    parse_pptx_with_limits(data, &ParseLimits::default())
}

pub fn parse_pptx_with_limits(data: &[u8], limits: &ParseLimits) -> Result<PptxPackage, PptxError> {
    let source_parts = ooxml_opc::unzip_parts(data).map_err(PptxError::Container)?;
    let parts: HashMap<&str, &[u8]> = source_parts
        .iter()
        .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
        .collect();
    let mut budget = ParseBudget::new(limits);
    let relationships = parse_package_relationships(&source_parts, &mut budget)?;
    let presentation_path = relationships
        .get("")
        .and_then(|entries| {
            entries
                .iter()
                .find(|relationship| relationship.has_type("/officeDocument"))
        })
        .and_then(|relationship| relationship.resolved_target.clone())
        .unwrap_or_else(|| "ppt/presentation.xml".to_owned());
    let presentation_root = parse_part(&parts, &presentation_path, &mut budget)?;
    let presentation_relationships = relationships
        .get(&presentation_path)
        .map(Vec::as_slice)
        .unwrap_or_default();
    let presentation = parse_presentation(
        &presentation_root,
        &presentation_path,
        presentation_relationships,
    )?;

    let mut slides = Vec::with_capacity(presentation.slides.len());
    for reference in &presentation.slides {
        let root = parse_part(&parts, &reference.part_path, &mut budget)?;
        let slide_relationships = relationships
            .get(&reference.part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let data = common_slide_data(
            &root,
            slide_relationships,
            &reference.part_path,
            &mut budget,
        )?;
        slides.push(Slide {
            part_path: reference.part_path.clone(),
            name: data.name,
            layout_part_path: relationship_by_type(
                slide_relationships,
                relationship_types::SLIDE_LAYOUT,
            ),
            show_master_shapes: bool_attribute(&root, "showMasterSp", true),
            background: data.background,
            shapes: data.shapes,
        });
    }

    let master_paths = ordered_part_paths(
        presentation.master_part_paths.clone(),
        &source_parts,
        "ppt/slideMasters/",
    );
    let mut masters = Vec::with_capacity(master_paths.len());
    for part_path in &master_paths {
        let root = parse_part(&parts, part_path, &mut budget)?;
        let master_relationships = relationships
            .get(part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let data = common_slide_data(&root, master_relationships, part_path, &mut budget)?;
        masters.push(SlideMaster {
            part_path: part_path.clone(),
            name: data.name,
            theme_part_path: relationship_by_type(master_relationships, relationship_types::THEME),
            layout_part_paths: master_relationships
                .iter()
                .filter(|relationship| relationship.has_type(relationship_types::SLIDE_LAYOUT))
                .filter_map(|relationship| relationship.resolved_target.clone())
                .collect(),
            background: data.background,
            shapes: data.shapes,
            text_styles: parse_text_styles(&root),
        });
    }

    let layout_paths = ordered_part_paths(
        masters
            .iter()
            .flat_map(|master| master.layout_part_paths.iter().cloned())
            .collect(),
        &source_parts,
        "ppt/slideLayouts/",
    );
    let mut layouts = Vec::with_capacity(layout_paths.len());
    for part_path in &layout_paths {
        let root = parse_part(&parts, part_path, &mut budget)?;
        let layout_relationships = relationships
            .get(part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let data = common_slide_data(&root, layout_relationships, part_path, &mut budget)?;
        layouts.push(SlideLayout {
            part_path: part_path.clone(),
            name: root
                .attribute("matchingName")
                .map(str::to_owned)
                .or(data.name),
            layout_type: root.attribute("type").map(str::to_owned),
            master_part_path: relationship_by_type(
                layout_relationships,
                relationship_types::SLIDE_MASTER,
            ),
            show_master_shapes: bool_attribute(&root, "showMasterSp", true),
            background: data.background,
            shapes: data.shapes,
        });
    }

    let theme_paths = ordered_part_paths(
        masters
            .iter()
            .filter_map(|master| master.theme_part_path.clone())
            .collect(),
        &source_parts,
        "ppt/theme/",
    );
    let mut themes = Vec::with_capacity(theme_paths.len());
    for part_path in theme_paths {
        let root = parse_part(&parts, &part_path, &mut budget)?;
        themes.push(ThemePart {
            part_path,
            theme: parse_theme(&root),
        });
    }

    let charts = parse_chart_parts(
        &parts,
        &ChartSources {
            slides: &slides,
            layouts: &layouts,
            masters: &masters,
            themes: &themes,
            relationships: &relationships,
        },
        limits,
    );

    let content_types = parse_content_types(&parts, &mut budget)?;
    let media = source_parts
        .iter()
        .filter(|(path, _)| path.starts_with("ppt/media/"))
        .map(|(path, bytes)| MediaPart {
            part_path: path.clone(),
            content_type: content_type_for(path, &content_types),
            bytes: bytes.clone(),
        })
        .collect();
    let parts = source_parts
        .into_iter()
        .map(|(path, bytes)| PackagePart { path, bytes })
        .collect();
    Ok(PptxPackage {
        presentation,
        slides,
        layouts,
        masters,
        themes,
        charts,
        media,
        relationships,
        parts,
    })
}

pub fn write_pptx(package: &PptxPackage) -> Result<Vec<u8>, PptxError> {
    let parts = package
        .parts
        .iter()
        .map(|part| (part.path.clone(), part.bytes.clone()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts(&parts).map_err(PptxError::Container)
}

fn parse_package_relationships(
    parts: &[(String, Vec<u8>)],
    budget: &mut ParseBudget<'_>,
) -> Result<BTreeMap<String, Vec<Relationship>>, PptxError> {
    let mut output = BTreeMap::new();
    for (path, bytes) in parts {
        let Some(source) = relationship_source(path) else {
            continue;
        };
        let parsed = parse_relationships(bytes, path, &source, budget)?;
        output.insert(source, parsed);
    }
    Ok(output)
}

fn relationship_source(path: &str) -> Option<String> {
    if path == "_rels/.rels" {
        return Some(String::new());
    }
    let (directory, filename) = path.rsplit_once("/_rels/")?;
    let source_filename = filename.strip_suffix(".rels")?;
    Some(format!("{directory}/{source_filename}"))
}

fn parse_part(
    parts: &HashMap<&str, &[u8]>,
    path: &str,
    budget: &mut ParseBudget<'_>,
) -> Result<XmlElement, PptxError> {
    let bytes = parts
        .get(path)
        .ok_or_else(|| PptxError::MissingPart(path.to_owned()))?;
    parse_xml(bytes, path, budget)
}

fn parse_presentation(
    root: &XmlElement,
    part_path: &str,
    relationships: &[Relationship],
) -> Result<Presentation, PptxError> {
    let slide_size = root.child("sldSz");
    let width_emu = positive_integer_attribute(slide_size, "cx").unwrap_or(12_192_000);
    let height_emu = positive_integer_attribute(slide_size, "cy").unwrap_or(6_858_000);
    let mut slides = Vec::new();
    if let Some(list) = root.child("sldIdLst") {
        for slide in list.children_named("sldId") {
            let relationship_id = slide
                .attribute("r:id")
                .or_else(|| slide.attribute_local("id"))
                .unwrap_or_default()
                .to_owned();
            let part_path =
                relationship_target(relationships, &relationship_id).ok_or_else(|| {
                    PptxError::MissingPart(format!("slide relationship {relationship_id}"))
                })?;
            slides.push(SlideReference {
                id: slide
                    .attribute("id")
                    .and_then(|value| value.parse().ok())
                    .unwrap_or_default(),
                relationship_id,
                part_path,
            });
        }
    }
    let master_part_paths = root
        .child("sldMasterIdLst")
        .into_iter()
        .flat_map(|list| list.children_named("sldMasterId"))
        .filter_map(|master| {
            master
                .attribute("r:id")
                .or_else(|| master.attribute_local("id"))
                .and_then(|id| relationship_target(relationships, id))
        })
        .collect();
    Ok(Presentation {
        part_path: part_path.to_owned(),
        width_emu,
        height_emu,
        slides,
        master_part_paths,
    })
}

/// The parts a chart may be referenced from, plus the theme cascade that
/// decides which colours it resolves against.
struct ChartSources<'a> {
    slides: &'a [Slide],
    layouts: &'a [SlideLayout],
    masters: &'a [SlideMaster],
    themes: &'a [ThemePart],
    relationships: &'a BTreeMap<String, Vec<Relationship>>,
}

impl ChartSources<'_> {
    fn chart_target(&self, source_part: &str, relationship_id: &str) -> Option<String> {
        self.relationships
            .get(source_part)?
            .iter()
            .find(|relationship| relationship.id == relationship_id)
            .filter(|relationship| relationship.has_type(relationship_types::CHART))
            .and_then(|relationship| relationship.resolved_target.clone())
    }

    fn theme_for(&self, source_part: &str) -> Option<&ThemePart> {
        if let Some(slide) = self
            .slides
            .iter()
            .find(|slide| slide.part_path == source_part)
        {
            return self.theme_via_layout(slide.layout_part_path.as_deref());
        }
        if self
            .layouts
            .iter()
            .any(|layout| layout.part_path == source_part)
        {
            return self.theme_via_layout(Some(source_part));
        }
        if self
            .masters
            .iter()
            .any(|master| master.part_path == source_part)
        {
            return self.theme_via_master(Some(source_part));
        }
        self.themes.first()
    }

    fn theme_via_layout(&self, layout_part_path: Option<&str>) -> Option<&ThemePart> {
        let layout = layout_part_path
            .and_then(|path| self.layouts.iter().find(|layout| layout.part_path == path))
            .or_else(|| self.layouts.first());
        let master_part_path = layout
            .and_then(|layout| layout.master_part_path.as_deref())
            .or_else(|| {
                let layout = layout?;
                self.masters
                    .iter()
                    .find(|master| {
                        master
                            .layout_part_paths
                            .iter()
                            .any(|path| path == &layout.part_path)
                    })
                    .map(|master| master.part_path.as_str())
            });
        self.theme_via_master(master_part_path)
    }

    fn theme_via_master(&self, master_part_path: Option<&str>) -> Option<&ThemePart> {
        let master = master_part_path
            .and_then(|path| self.masters.iter().find(|master| master.part_path == path))
            .or_else(|| self.masters.first());
        master
            .and_then(|master| master.theme_part_path.as_deref())
            .and_then(|path| self.themes.iter().find(|theme| theme.part_path == path))
            .or_else(|| self.themes.first())
    }
}

/// Every chart the deck references, read against the theme its source part
/// resolves to.
///
/// A chart part is a decorative leaf: nothing structural is read from it and
/// nothing structural is read after it. So a part `parse_xml` refuses is
/// skipped, leaving the deck to open with that chart missing exactly as it opens
/// when the part is absent. `parse_xml` reports only `MalformedXml`, `UnsafeXml`
/// and `ResourceLimit`, each naming this part; a hostile package is refused
/// earlier, by `ooxml_opc::unzip_parts`.
///
/// What bounds the reading: each part gets its own [`ParseBudget`], so no cached
/// series can starve a slide or spend another part's depth, byte and text
/// limits; all of them draw their XML events from one pool the size of
/// `max_xml_events` ([`read_chart_root`]); and each part is read once however
/// many themes reference it, the tree being the same for all of them and only
/// the colours it resolves against differing.
fn parse_chart_parts(
    parts: &HashMap<&str, &[u8]>,
    sources: &ChartSources<'_>,
    limits: &ParseLimits,
) -> Vec<ChartPart> {
    let mut references = Vec::new();
    for slide in sources.slides {
        collect_chart_references(&slide.shapes, &slide.part_path, &mut references);
    }
    for layout in sources.layouts {
        collect_chart_references(&layout.shapes, &layout.part_path, &mut references);
    }
    for master in sources.masters {
        collect_chart_references(&master.shapes, &master.part_path, &mut references);
    }
    let default_theme = Theme::default();
    let mut reads = Vec::new();
    let mut loaded = HashSet::new();
    let mut pending_reads: HashMap<String, usize> = HashMap::new();
    for (source_part, relationship_id) in references {
        let Some(part_path) = sources.chart_target(&source_part, &relationship_id) else {
            continue;
        };
        let theme_part = sources.theme_for(&source_part);
        let theme_part_path = theme_part.map(|part| part.part_path.clone());
        if !loaded.insert((part_path.clone(), theme_part_path.clone())) {
            continue;
        }
        *pending_reads.entry(part_path.clone()).or_default() += 1;
        let theme = theme_part.map(|part| &part.theme).unwrap_or(&default_theme);
        reads.push((part_path, theme_part_path, theme));
    }

    let mut charts = Vec::new();
    let mut roots: HashMap<String, Option<XmlElement>> = HashMap::new();
    let mut remaining_events = limits.max_xml_events;
    for (part_path, theme_part_path, theme) in reads {
        if !roots.contains_key(&part_path) {
            let root = parts.get(part_path.as_str()).and_then(|bytes| {
                read_chart_root(bytes, &part_path, limits, &mut remaining_events)
            });
            roots.insert(part_path.clone(), root);
        }
        let chart = roots
            .get(&part_path)
            .and_then(Option::as_ref)
            .and_then(|root| parse_chart_part(root, theme));
        if let Some(chart) = chart {
            charts.push(ChartPart {
                part_path: part_path.clone(),
                theme_part_path,
                chart,
            });
        }
        if let Some(pending) = pending_reads.get_mut(&part_path) {
            *pending -= 1;
            if *pending == 0 {
                roots.remove(&part_path);
            }
        }
    }
    charts
}

/// Reads one chart part against its own budget, drawing its XML events from
/// `remaining_events` — a pool the size of the package's `max_xml_events` that
/// every chart part shares. A part that empties the pool is declined like a
/// malformed one, so the deck's charts together cost no more than the one budget
/// they used to be read against.
fn read_chart_root(
    bytes: &[u8],
    part_path: &str,
    limits: &ParseLimits,
    remaining_events: &mut usize,
) -> Option<XmlElement> {
    let part_limits = ParseLimits {
        max_xml_events: *remaining_events,
        ..limits.clone()
    };
    let mut budget = ParseBudget::new(&part_limits);
    let root = parse_xml(bytes, part_path, &mut budget).ok();
    *remaining_events = remaining_events.saturating_sub(budget.xml_events_spent());
    root
}

/// `(referencing part, relationship id)` for every chart in `shapes`, groups
/// included, in document order.
fn collect_chart_references(
    shapes: &[ShapeNode],
    part_path: &str,
    output: &mut Vec<(String, String)>,
) {
    for shape in shapes {
        match shape {
            ShapeNode::GraphicFrame(frame) => {
                if let GraphicFrameData::Chart {
                    relationship_id, ..
                } = &frame.data
                {
                    output.push((part_path.to_owned(), relationship_id.clone()));
                }
            }
            ShapeNode::Group(group) => collect_chart_references(&group.children, part_path, output),
            ShapeNode::Shape(_) | ShapeNode::Picture(_) => {}
        }
    }
}

#[derive(Default)]
struct ContentTypes {
    defaults: HashMap<String, String>,
    overrides: HashMap<String, String>,
}

fn parse_content_types(
    parts: &HashMap<&str, &[u8]>,
    budget: &mut ParseBudget<'_>,
) -> Result<ContentTypes, PptxError> {
    let root = parse_part(parts, "[Content_Types].xml", budget)?;
    let mut content_types = ContentTypes::default();
    for child in root.child_elements() {
        match child.local_name() {
            "Default" => {
                if let (Some(extension), Some(content_type)) =
                    (child.attribute("Extension"), child.attribute("ContentType"))
                {
                    content_types
                        .defaults
                        .insert(extension.to_ascii_lowercase(), content_type.to_owned());
                }
            }
            "Override" => {
                if let (Some(part_name), Some(content_type)) =
                    (child.attribute("PartName"), child.attribute("ContentType"))
                {
                    content_types.overrides.insert(
                        part_name.trim_start_matches('/').to_owned(),
                        content_type.to_owned(),
                    );
                }
            }
            _ => {}
        }
    }
    Ok(content_types)
}

fn content_type_for(path: &str, content_types: &ContentTypes) -> String {
    content_types
        .overrides
        .get(path)
        .cloned()
        .or_else(|| {
            path.rsplit_once('.').and_then(|(_, extension)| {
                content_types
                    .defaults
                    .get(&extension.to_ascii_lowercase())
                    .cloned()
            })
        })
        .unwrap_or_else(|| "application/octet-stream".to_owned())
}

fn ordered_part_paths(
    preferred: Vec<String>,
    parts: &[(String, Vec<u8>)],
    prefix: &str,
) -> Vec<String> {
    let mut seen = HashSet::new();
    preferred
        .into_iter()
        .chain(
            parts
                .iter()
                .map(|(path, _)| path.clone())
                .filter(|path| path.starts_with(prefix) && path.ends_with(".xml")),
        )
        .filter(|path| seen.insert(path.clone()))
        .collect()
}

fn relationship_by_type(relationships: &[Relationship], suffix: &str) -> Option<String> {
    relationships
        .iter()
        .find(|relationship| relationship.has_type(suffix))
        .and_then(|relationship| relationship.resolved_target.clone())
}

fn relationship_target(relationships: &[Relationship], id: &str) -> Option<String> {
    relationships
        .iter()
        .find(|relationship| relationship.id == id)
        .and_then(|relationship| relationship.resolved_target.clone())
}

fn positive_integer_attribute(element: Option<&XmlElement>, name: &str) -> Option<i64> {
    let value = element?.attribute(name)?.parse::<i64>().ok()?;
    (value > 0 && value <= 1_000_000_000_000_000).then_some(value)
}

fn bool_attribute(element: &XmlElement, name: &str, default: bool) -> bool {
    match element.attribute(name) {
        Some("1" | "true" | "on") => true,
        Some("0" | "false" | "off") => false,
        _ => default,
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
    const CHART_FIXTURE: &[u8] = include_bytes!("../tests/fixtures/chart-deck.pptx");
    const CHART_PART: &str = "ppt/charts/chart1.xml";

    #[test]
    fn parses_betteroffice_demo_deck_surface() {
        let package = parse_pptx(FIXTURE).unwrap();
        assert_eq!(package.presentation.width_emu, 12_192_000);
        assert_eq!(package.presentation.height_emu, 6_858_000);
        assert_eq!(package.slides.len(), 3);
        assert_eq!(package.layouts.len(), 1);
        assert_eq!(package.masters.len(), 1);
        assert_eq!(package.themes.len(), 1);
        assert_eq!(package.media.len(), 1);
        assert_eq!(package.themes[0].theme.name, "BetterOffice");

        let kinds = package
            .slides
            .iter()
            .flat_map(|slide| slide.shapes.iter())
            .map(|shape| match shape {
                ShapeNode::Shape(_) => "shape",
                ShapeNode::Picture(_) => "picture",
                ShapeNode::GraphicFrame(_) => "graphicFrame",
                ShapeNode::Group(_) => "group",
            })
            .collect::<HashSet<_>>();
        assert_eq!(
            kinds,
            HashSet::from(["shape", "picture", "graphicFrame", "group"])
        );
        assert!(package.slides.iter().flat_map(|slide| &slide.shapes).any(|shape| {
            matches!(
                shape,
                ShapeNode::Shape(Shape {
                    text: Some(TextBody { paragraphs, .. }),
                    ..
                }) if paragraphs.iter().flat_map(|paragraph| &paragraph.runs).any(|run| run.text.contains("Rust"))
            )
        }));
    }

    #[test]
    fn untouched_save_preserves_every_part_byte_and_order() {
        let package = parse_pptx(FIXTURE).unwrap();
        let written = write_pptx(&package).unwrap();
        let before = ooxml_opc::unzip_parts(FIXTURE).unwrap();
        let after = ooxml_opc::unzip_parts(&written).unwrap();
        assert_eq!(after, before);
    }

    #[test]
    fn loads_every_referenced_chart_part_against_the_deck_theme() {
        let package = parse_pptx(CHART_FIXTURE).unwrap();
        assert_eq!(
            package
                .charts
                .iter()
                .map(|part| part.part_path.as_str())
                .collect::<Vec<_>>(),
            ["ppt/charts/chart1.xml", "ppt/charts/chart2.xml"]
        );

        let column = &package.charts[0].chart;
        assert_eq!(column.chart_type, "column");
        assert_eq!(column.title.as_deref(), Some("Revenue"));
        assert_eq!(
            column.legend.as_ref().unwrap().position.as_deref(),
            Some("right")
        );
        assert_eq!(column.series[0].color, "#6254E7");
        assert_eq!(column.series[1].color, "#1FA97A");
        assert_eq!(column.series[0].values, [12.0, 19.0, 7.0]);
        assert_eq!(column.plot_groups[0].gap_width, Some(150.0));
        assert_eq!(
            column.plot_groups[0].series[0]
                .data_labels
                .as_ref()
                .and_then(|labels| labels.show_value),
            Some(true)
        );
        assert!(column.plot_groups[0].data_labels.is_none());
        assert!(column.plot_groups[0].show_data_labels);
        let axes = column.axis_list.as_ref().unwrap();
        assert_eq!(axes[0].title.as_deref(), Some("Quarter"));
        assert_eq!(axes[1].title.as_deref(), Some("Millions"));
        assert_eq!(axes[1].max, Some(25.0));

        // The grouped chart proves the reference walk descends into groups.
        let pie = &package.charts[1].chart;
        assert_eq!(pie.chart_type, "pie");
        let points = pie.plot_groups[0].series[0].points.as_ref().unwrap();
        assert_eq!(points[0].color, "#E7A954");
        assert_eq!(points[1].color, "#112233");
        assert_eq!(points[1].explosion, Some(10.0));
    }

    /// Three slides sharing `ppt/charts/chart1.xml`, resolving to two themes.
    struct SharedChartDeck {
        slides: Vec<Slide>,
        layouts: Vec<SlideLayout>,
        masters: Vec<SlideMaster>,
        themes: Vec<ThemePart>,
        relationships: BTreeMap<String, Vec<Relationship>>,
    }

    impl SharedChartDeck {
        fn sources(&self) -> ChartSources<'_> {
            ChartSources {
                slides: &self.slides,
                layouts: &self.layouts,
                masters: &self.masters,
                themes: &self.themes,
                relationships: &self.relationships,
            }
        }
    }

    fn shared_chart_deck() -> SharedChartDeck {
        fn chart_shape(id: u32) -> ShapeNode {
            ShapeNode::GraphicFrame(GraphicFrame {
                base: ShapeBase {
                    id,
                    name: format!("Chart {id}"),
                    description: None,
                    hidden: false,
                    placeholder: None,
                    transform: ShapeTransform::default(),
                },
                data: GraphicFrameData::Chart {
                    relationship_id: "rIdChart".to_owned(),
                    part_path: Some("ppt/charts/chart1.xml".to_owned()),
                },
            })
        }

        fn slide(path: &str, layout: &str, id: u32) -> Slide {
            Slide {
                part_path: path.to_owned(),
                name: None,
                layout_part_path: Some(layout.to_owned()),
                show_master_shapes: true,
                background: None,
                shapes: vec![chart_shape(id)],
            }
        }

        fn chart_relationships() -> Vec<Relationship> {
            vec![Relationship {
                id: "rIdChart".to_owned(),
                relationship_type: relationship_types::CHART.to_owned(),
                target: "../charts/chart1.xml".to_owned(),
                target_mode: Default::default(),
                resolved_target: Some("ppt/charts/chart1.xml".to_owned()),
            }]
        }

        let slides = vec![
            slide("ppt/slides/slide1.xml", "ppt/slideLayouts/layout1.xml", 1),
            slide("ppt/slides/slide2.xml", "ppt/slideLayouts/layout2.xml", 2),
            slide("ppt/slides/slide3.xml", "ppt/slideLayouts/layout1.xml", 3),
        ];
        let layouts = vec![
            SlideLayout {
                part_path: "ppt/slideLayouts/layout1.xml".to_owned(),
                name: None,
                layout_type: None,
                master_part_path: Some("ppt/slideMasters/master1.xml".to_owned()),
                show_master_shapes: true,
                background: None,
                shapes: Vec::new(),
            },
            SlideLayout {
                part_path: "ppt/slideLayouts/layout2.xml".to_owned(),
                name: None,
                layout_type: None,
                master_part_path: Some("ppt/slideMasters/master2.xml".to_owned()),
                show_master_shapes: true,
                background: None,
                shapes: Vec::new(),
            },
        ];
        let masters = vec![
            SlideMaster {
                part_path: "ppt/slideMasters/master1.xml".to_owned(),
                name: None,
                theme_part_path: Some("ppt/theme/theme1.xml".to_owned()),
                layout_part_paths: vec!["ppt/slideLayouts/layout1.xml".to_owned()],
                background: None,
                shapes: Vec::new(),
                text_styles: TextStyleSet::default(),
            },
            SlideMaster {
                part_path: "ppt/slideMasters/master2.xml".to_owned(),
                name: None,
                theme_part_path: Some("ppt/theme/theme2.xml".to_owned()),
                layout_part_paths: vec!["ppt/slideLayouts/layout2.xml".to_owned()],
                background: None,
                shapes: Vec::new(),
                text_styles: TextStyleSet::default(),
            },
        ];
        let mut first_theme = Theme::default();
        first_theme.color_scheme.accent1 = "112233".to_owned();
        let mut second_theme = Theme::default();
        second_theme.color_scheme.accent1 = "AABBCC".to_owned();
        let themes = vec![
            ThemePart {
                part_path: "ppt/theme/theme1.xml".to_owned(),
                theme: first_theme,
            },
            ThemePart {
                part_path: "ppt/theme/theme2.xml".to_owned(),
                theme: second_theme,
            },
        ];
        let relationships = BTreeMap::from([
            ("ppt/slides/slide1.xml".to_owned(), chart_relationships()),
            ("ppt/slides/slide2.xml".to_owned(), chart_relationships()),
            ("ppt/slides/slide3.xml".to_owned(), chart_relationships()),
        ]);
        SharedChartDeck {
            slides,
            layouts,
            masters,
            themes,
            relationships,
        }
    }

    #[test]
    fn a_shared_chart_is_resolved_once_per_referencing_theme() {
        let deck = shared_chart_deck();
        let chart_xml = br#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#;
        let parts: HashMap<&str, &[u8]> =
            HashMap::from([("ppt/charts/chart1.xml", chart_xml.as_slice())]);

        let charts = parse_chart_parts(&parts, &deck.sources(), &ParseLimits::default());

        assert_eq!(charts.len(), 2);
        assert_eq!(
            charts
                .iter()
                .map(|chart| chart.theme_part_path.as_deref())
                .collect::<Vec<_>>(),
            [Some("ppt/theme/theme1.xml"), Some("ppt/theme/theme2.xml")]
        );
        assert_eq!(charts[0].chart.series[0].color, "#112233");
        assert_eq!(charts[1].chart.series[0].color, "#AABBCC");
    }

    /// A chart part parses to the same tree whatever theme it is coloured
    /// against, so it is read once: an event pool only one read fits still
    /// resolves the chart against both themes.
    #[test]
    fn a_shared_chart_part_is_read_once_however_many_themes_reference_it() {
        let deck = shared_chart_deck();
        let chart_xml = cached_chart(1_000);
        let parts: HashMap<&str, &[u8]> =
            HashMap::from([("ppt/charts/chart1.xml", chart_xml.as_slice())]);
        let limits = ParseLimits {
            max_xml_events: 6_000,
            ..ParseLimits::default()
        };

        let charts = parse_chart_parts(&parts, &deck.sources(), &limits);

        assert_eq!(charts.len(), 2);
        assert!(
            charts
                .iter()
                .all(|part| part.chart.series[0].values.len() == 1_000)
        );
    }

    #[test]
    fn a_deck_without_charts_loads_no_chart_parts() {
        assert!(parse_pptx(FIXTURE).unwrap().charts.is_empty());
    }

    /// `chart-deck.pptx` with `part`'s bytes replaced.
    fn chart_deck_with(part: &str, bytes: &[u8]) -> Vec<u8> {
        let mut parts = ooxml_opc::unzip_parts(CHART_FIXTURE).unwrap();
        let slot = parts
            .iter_mut()
            .find(|(path, _)| path == part)
            .unwrap_or_else(|| panic!("{part} is not in the fixture"));
        slot.1 = bytes.to_vec();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    /// Opens a deck whose first chart part carries `bytes` and asserts the deck
    /// is whole apart from that one chart.
    fn assert_damaged_chart_is_dropped(bytes: &[u8]) {
        let package = parse_pptx(&chart_deck_with(CHART_PART, bytes)).unwrap();
        assert_eq!(package.slides.len(), 2);
        assert_eq!(package.layouts.len(), 1);
        assert_eq!(package.masters.len(), 1);
        assert_eq!(package.themes.len(), 1);
        assert_eq!(
            package
                .slides
                .iter()
                .map(|slide| slide.name.as_deref())
                .collect::<Vec<_>>(),
            [Some("Charts"), Some("Grouped")]
        );
        assert_eq!(
            package.slides[0]
                .shapes
                .iter()
                .filter(|shape| matches!(shape, ShapeNode::GraphicFrame(_)))
                .count(),
            2
        );
        assert_eq!(
            package
                .charts
                .iter()
                .map(|part| part.part_path.as_str())
                .collect::<Vec<_>>(),
            ["ppt/charts/chart2.xml"]
        );
    }

    #[test]
    fn a_malformed_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(b"<c:chartSpace><c:chart></c:chartSpace>");
    }

    #[test]
    fn an_empty_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(b"");
    }

    #[test]
    fn a_doctype_bearing_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(
            br#"<?xml version="1.0"?><!DOCTYPE c:chartSpace [<!ENTITY x "y">]><c:chartSpace xmlns:c="c"/>"#,
        );
    }

    #[test]
    fn a_truncated_chart_part_drops_only_that_chart() {
        let source = ooxml_opc::unzip_parts(CHART_FIXTURE)
            .unwrap()
            .into_iter()
            .find(|(path, _)| path == CHART_PART)
            .map(|(_, bytes)| bytes)
            .unwrap();
        assert_damaged_chart_is_dropped(&source[..source.len() / 2]);
    }

    #[test]
    fn a_binary_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(&[0x00, 0xff, 0xfe, 0x01, 0x80, 0x7f]);
    }

    #[test]
    fn an_undeclared_entity_in_a_chart_part_drops_only_that_chart() {
        assert_damaged_chart_is_dropped(
            br#"<c:chartSpace xmlns:c="c"><c:chart><c:title>&nbsp;</c:title></c:chart></c:chartSpace>"#,
        );
    }

    /// Declining to read a chart part must not stop the package from carrying
    /// it: an untouched save still writes the source bytes back.
    #[test]
    fn an_unreadable_chart_part_survives_a_save_byte_for_byte() {
        const DAMAGED: &[u8] = b"<c:chartSpace><c:chart></c:chartSpace>";
        let deck = chart_deck_with(CHART_PART, DAMAGED);
        let written = write_pptx(&parse_pptx(&deck).unwrap()).unwrap();
        assert_eq!(
            ooxml_opc::unzip_parts(&written).unwrap(),
            ooxml_opc::unzip_parts(&deck).unwrap()
        );
        assert_eq!(
            ooxml_opc::unzip_parts(&written)
                .unwrap()
                .into_iter()
                .find(|(path, _)| path == CHART_PART)
                .map(|(_, bytes)| bytes),
            Some(DAMAGED.to_vec())
        );
    }

    /// A chart part whose only fault is a large cache. Every point costs the
    /// reader five events.
    fn cached_chart(points: usize) -> Vec<u8> {
        let mut xml = String::from(
            r#"<c:chartSpace xmlns:c="c" xmlns:a="a"><c:chart><c:plotArea><c:barChart><c:barDir val="col"/><c:ser><c:val><c:numRef><c:numCache>"#,
        );
        for point in 0..points {
            xml.push_str(&format!("<c:pt><c:v>{point}</c:v></c:pt>"));
        }
        xml.push_str("</c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>");
        xml.into_bytes()
    }

    /// `chart-deck.pptx` with charts of `points` cached values and `shapes`
    /// more shapes on the second slide, so charts and slides both cost real
    /// events.
    fn cached_chart_deck(points: usize, shapes: usize) -> Vec<u8> {
        let crowd = (0..shapes)
            .map(|shape| {
                format!(
                    r#"<p:sp><p:nvSpPr><p:cNvPr id="{}" name=""/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/></p:sp>"#,
                    shape + 100
                )
            })
            .collect::<String>();
        let mut parts = ooxml_opc::unzip_parts(CHART_FIXTURE).unwrap();
        for (path, bytes) in &mut parts {
            if path.starts_with("ppt/charts/") {
                *bytes = cached_chart(points);
            }
            if path == "ppt/slides/slide2.xml" {
                *bytes = String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace("</p:spTree>", &format!("{crowd}</p:spTree>"))
                    .into_bytes();
            }
        }
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    /// The undamaged deck that used to be refused: spec-valid charts whose
    /// caches, added to what the slides cost, outgrow one budget. The charts
    /// draw from a pool of their own, so slides and charts both come through.
    #[test]
    fn a_deck_whose_charts_outgrow_the_slides_budget_opens_with_every_chart() {
        let deck = cached_chart_deck(2_000, 1_500);
        let limits = ParseLimits {
            max_xml_events: 30_000,
            ..ParseLimits::default()
        };

        let package = parse_pptx_with_limits(&deck, &limits).unwrap();

        assert_eq!(package.slides.len(), 2);
        assert_eq!(
            package
                .slides
                .iter()
                .map(|slide| slide.name.as_deref())
                .collect::<Vec<_>>(),
            [Some("Charts"), Some("Grouped")]
        );
        assert_eq!(package.slides[1].shapes.len(), 1_501);
        assert_eq!(package.charts.len(), 2);
        assert!(
            package
                .charts
                .iter()
                .all(|part| part.chart.series[0].values.len() == 2_000)
        );
    }

    /// The pool the charts share is one budget for the whole deck, so a chart
    /// that empties it leaves the next one declined like a damaged part — and
    /// the deck still opens.
    #[test]
    fn a_deck_whose_charts_outgrow_the_chart_budget_opens_with_the_charts_that_fit() {
        let deck = cached_chart_deck(4_000, 1_500);
        let limits = ParseLimits {
            max_xml_events: 30_000,
            ..ParseLimits::default()
        };

        let package = parse_pptx_with_limits(&deck, &limits).unwrap();

        assert_eq!(package.slides.len(), 2);
        assert_eq!(package.slides[1].shapes.len(), 1_501);
        assert_eq!(
            package.slides[0]
                .shapes
                .iter()
                .filter(|shape| matches!(shape, ShapeNode::GraphicFrame(_)))
                .count(),
            2
        );
        assert_eq!(
            package
                .charts
                .iter()
                .map(|part| part.part_path.as_str())
                .collect::<Vec<_>>(),
            ["ppt/charts/chart1.xml"]
        );
        assert_eq!(package.charts[0].chart.series[0].values.len(), 4_000);
    }

    #[test]
    fn package_limits_apply_across_xml_parts() {
        let limits = ParseLimits {
            max_xml_bytes: 100,
            ..ParseLimits::default()
        };
        assert!(matches!(
            parse_pptx_with_limits(FIXTURE, &limits),
            Err(PptxError::ResourceLimit {
                kind: "xmlBytes",
                ..
            })
        ));
    }
}
