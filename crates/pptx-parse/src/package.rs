use std::collections::{BTreeMap, HashMap, HashSet};

use ooxml_drawingml::{ColorMap, TableStyleList, Theme};

use crate::chart::parse_chart_part;
use crate::comments::{
    Comment, CommentAuthor, CommentFlavor, authors_part, parse_comment_authors, parse_comments,
    slide_comment_parts,
};
use crate::drawing::{
    BasicLayout, DIAGRAM_EXTENT_EMU, classify_layout, common_slide_data,
    common_slide_data_with_vml, layout_basic_diagram, parse_diagram_drawing,
    parse_paragraph_properties, parse_style_levels, parse_text_styles,
};
use crate::model::*;
use crate::relationships::{Relationship, parse_relationships, relationship_types};
use crate::table_style::parse_table_styles;
use crate::theme::{parse_background_pictures, parse_format_scheme, parse_theme};
use crate::xml::{ParseBudget, XmlElement, parse_xml};
use crate::{ParseLimits, PptxError};

pub fn parse_pptx(data: &[u8]) -> Result<PptxPackage, PptxError> {
    parse_pptx_with_limits(data, &ParseLimits::default())
}

pub fn parse_pptx_with_limits(data: &[u8], limits: &ParseLimits) -> Result<PptxPackage, PptxError> {
    parse_package(data, limits, ShapeElements::WithConnectors)
}

/// Preserves pre-connector source ordinals.
pub fn parse_pptx_without_connectors(data: &[u8]) -> Result<PptxPackage, PptxError> {
    parse_package(
        data,
        &ParseLimits::default(),
        ShapeElements::WithoutConnectors,
    )
}

fn parse_package(
    data: &[u8],
    limits: &ParseLimits,
    shape_elements: ShapeElements,
) -> Result<PptxPackage, PptxError> {
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

    let mut has_connectors = false;
    let mut vml_budget = ParseBudget::new(limits);
    let mut slides = Vec::with_capacity(presentation.slides.len());
    let mut sources = BTreeMap::new();
    let mut notes_sources = BTreeMap::new();
    for reference in &presentation.slides {
        let root = parse_part(&parts, &reference.part_path, &mut budget)?;
        let slide_relationships = relationships
            .get(&reference.part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        has_connectors |= !root.descendants_named("cxnSp").is_empty();
        let previews = parse_vml_previews(
            &root,
            slide_relationships,
            &parts,
            &relationships,
            &mut vml_budget,
        );
        let data = common_slide_data_with_vml(
            &root,
            slide_relationships,
            &reference.part_path,
            &mut budget,
            shape_elements,
            &previews,
        )?;
        let mut notes = String::new();
        if let Some(notes_path) = crate::notes::slide_notes_part(slide_relationships)
            && let Some(bytes) = parts.get(notes_path.as_str())
        {
            let notes_root = parse_xml(bytes, &notes_path, &mut budget)?;
            notes = crate::notes::notes_text(&notes_root);
            notes_sources.insert(
                reference.part_path.clone(),
                crate::inventory::notes_source(&notes_root, &notes_path),
            );
        }
        sources.insert(
            reference.part_path.clone(),
            crate::inventory::slide_source(&root, slide_relationships, shape_elements),
        );
        slides.push(Slide {
            part_path: reference.part_path.clone(),
            name: data.name,
            layout_part_path: relationship_by_type(
                slide_relationships,
                relationship_types::SLIDE_LAYOUT,
            ),
            show_master_shapes: bool_attribute(&root, "showMasterSp", true),
            background: data.background,
            background_picture: data.background_picture,
            background_reference: data.background_reference,
            declares_background: root
                .child("cSld")
                .is_some_and(|common| common.child("bg").is_some()),
            shapes: data.shapes,
            color_map_override: parse_color_map_override(&root),
            notes,
            hidden: Some(!bool_attribute(&root, "show", true)),
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
        has_connectors |= !root.descendants_named("cxnSp").is_empty();
        let data = common_slide_data(
            &root,
            master_relationships,
            part_path,
            &mut budget,
            shape_elements,
        )?;
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
            background_picture: data.background_picture,
            background_reference: data.background_reference,
            shapes: data.shapes,
            text_styles: parse_text_styles(&root),
            color_map: parse_color_map(&root),
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
        has_connectors |= !root.descendants_named("cxnSp").is_empty();
        let data = common_slide_data(
            &root,
            layout_relationships,
            part_path,
            &mut budget,
            shape_elements,
        )?;
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
            background_picture: data.background_picture,
            background_reference: data.background_reference,
            shapes: data.shapes,
            color_map_override: parse_color_map_override(&root),
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
        let theme_relationships = relationships
            .get(&part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        themes.push(ThemePart {
            theme: parse_theme(&root),
            format_scheme: parse_format_scheme(&root),
            background_pictures: parse_background_pictures(&root, theme_relationships),
            part_path,
        });
    }

    let table_styles = parse_table_style_part(&parts, presentation_relationships, &mut budget)?;

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

    let deck_comments = parse_package_comments(
        &parts,
        &presentation,
        presentation_relationships,
        &relationships,
        &mut budget,
    )?;

    let content_types = parse_content_types(&parts, &mut budget)?;
    let diagram_drawings =
        parse_diagram_drawings(&parts, &slides, &relationships, &mut budget, shape_elements);
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
        diagram_drawings,
        media,
        table_styles,
        comment_authors: deck_comments.authors,
        comments: deck_comments.comments,
        comment_flavor: deck_comments.flavor,
        relationships,
        parts,
        source_container: ooxml_opc::SourceContainer::new(data.to_vec()),
        shape_elements: if has_connectors {
            shape_elements
        } else {
            ShapeElements::WithoutConnectors
        },
        sources,
        notes_sources,
    })
}

fn parse_package_comments(
    parts: &HashMap<&str, &[u8]>,
    presentation: &Presentation,
    presentation_relationships: &[Relationship],
    relationships: &BTreeMap<String, Vec<Relationship>>,
    budget: &mut ParseBudget<'_>,
) -> Result<PackageComments, PptxError> {
    let mut found: Vec<(CommentFlavor, String, String)> = Vec::new();
    for reference in &presentation.slides {
        let slide_relationships = relationships
            .get(&reference.part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let slide_parts = slide_comment_parts(slide_relationships);
        if let Some(part) = slide_parts.legacy {
            found.push((CommentFlavor::Legacy, reference.part_path.clone(), part));
        }
        if let Some(part) = slide_parts.modern {
            found.push((CommentFlavor::Modern, reference.part_path.clone(), part));
        }
    }
    let Some(flavor) = found
        .iter()
        .any(|(flavor, _, _)| *flavor == CommentFlavor::Legacy)
        .then_some(CommentFlavor::Legacy)
        .or_else(|| (!found.is_empty()).then_some(CommentFlavor::Modern))
    else {
        return Ok(PackageComments::default());
    };

    let mut comments = Vec::new();
    for (_, slide_part_path, part) in found
        .iter()
        .filter(|(candidate, _, _)| *candidate == flavor)
    {
        let Some(bytes) = parts.get(part.as_str()) else {
            continue;
        };
        comments.extend(parse_comments(
            bytes,
            part,
            slide_part_path,
            flavor,
            budget,
        )?);
    }

    let comment_authors = match authors_part(presentation_relationships, flavor)
        .and_then(|part| parts.get(part.as_str()).map(|bytes| (part, *bytes)))
    {
        Some((part, bytes)) => parse_comment_authors(bytes, &part, flavor, budget)?,
        None => Vec::new(),
    };
    Ok(PackageComments {
        authors: comment_authors,
        comments,
        flavor: Some(flavor),
    })
}

#[derive(Default)]
struct PackageComments {
    authors: Vec<CommentAuthor>,
    comments: Vec<Comment>,
    flavor: Option<CommentFlavor>,
}

pub fn write_pptx(package: &PptxPackage) -> Result<Vec<u8>, PptxError> {
    let parts = package
        .parts
        .iter()
        .map(|part| (part.path.clone(), part.bytes.clone()))
        .collect::<Vec<_>>();
    ooxml_opc::rezip_parts_preserving(&parts, package.source_container.as_bytes())
        .map_err(PptxError::Container)
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

fn parse_vml_previews(
    slide: &XmlElement,
    slide_relationships: &[Relationship],
    parts: &HashMap<&str, &[u8]>,
    relationships: &BTreeMap<String, Vec<Relationship>>,
    budget: &mut ParseBudget<'_>,
) -> BTreeMap<String, PictureFill> {
    let wanted: HashSet<_> = slide
        .descendants_named("oleObj")
        .into_iter()
        .filter(|object| object.child("pic").is_none())
        .filter_map(|object| object.attribute("spid"))
        .filter(|spid| !spid.is_empty())
        .collect();
    let mut previews = BTreeMap::new();
    if wanted.is_empty() {
        return previews;
    }
    for relationship in slide_relationships
        .iter()
        .filter(|relationship| relationship.has_type("/vmlDrawing"))
    {
        let Some(path) = relationship.resolved_target.as_deref() else {
            continue;
        };
        let Some(bytes) = parts.get(path) else {
            continue;
        };
        let Ok(root) = parse_xml(bytes, path, budget) else {
            continue;
        };
        let vml_relationships = relationships
            .get(path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        for shape in root.descendants_named("shape") {
            let spids = [shape.attribute("id"), shape.attribute("o:spid")];
            if !spids.iter().flatten().any(|spid| wanted.contains(spid)) {
                continue;
            }
            let Some(image) = shape.child("imagedata") else {
                continue;
            };
            let Some(id) = image
                .attribute("o:relid")
                .or_else(|| image.attribute("r:id"))
            else {
                continue;
            };
            let Some(media_path) = vml_relationships
                .iter()
                .find(|relationship| relationship.id == id)
                .and_then(|relationship| relationship.resolved_target.as_deref())
                .filter(|path| path.starts_with("ppt/media/") && parts.contains_key(path))
            else {
                continue;
            };
            for spid in spids
                .into_iter()
                .flatten()
                .filter(|spid| wanted.contains(spid))
            {
                previews
                    .entry(spid.to_owned())
                    .or_insert_with(|| PictureFill {
                        media_part_path: Some(media_path.to_owned()),
                        ..PictureFill::default()
                    });
            }
        }
    }
    previews
}

fn parse_table_style_part(
    parts: &HashMap<&str, &[u8]>,
    presentation_relationships: &[Relationship],
    budget: &mut ParseBudget<'_>,
) -> Result<TableStyleList, PptxError> {
    let path = relationship_by_type(presentation_relationships, relationship_types::TABLE_STYLES)
        .unwrap_or_else(|| "ppt/tableStyles.xml".to_owned());
    let Some(bytes) = parts.get(path.as_str()) else {
        return Ok(TableStyleList::default());
    };
    Ok(parse_table_styles(&parse_xml(bytes, &path, budget)?))
}

fn parse_presentation(
    root: &XmlElement,
    part_path: &str,
    relationships: &[Relationship],
) -> Result<Presentation, PptxError> {
    let slide_size = root.child("sldSz");
    let width_emu = positive_integer_attribute(slide_size, "cx").unwrap_or(12_192_000);
    let height_emu = positive_integer_attribute(slide_size, "cy").unwrap_or(6_858_000);
    let first_slide_num = root
        .attribute("firstSlideNum")
        .and_then(|value| value.parse().ok())
        .unwrap_or(1);
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
        first_slide_num,
        slides,
        master_part_paths,
        default_text_style: parse_style_levels(root.child("defaultTextStyle")),
        default_text_paragraph: root
            .child("defaultTextStyle")
            .and_then(|style| style.child("defPPr"))
            .map(|properties| Box::new(parse_paragraph_properties(Some(properties)))),
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

struct DiagramReference {
    part_path: String,
    slide_part_path: String,
    relationship_ids: Vec<String>,
}

/// Parses every `ppt/diagrams/drawing#.xml` a slide points at, once each,
/// against what the package budget has left, or synthesizes a basic layout
/// from `ppt/diagrams/data#.xml` when the cached drawing is missing or empty.
fn parse_diagram_drawings(
    parts: &HashMap<&str, &[u8]>,
    slides: &[Slide],
    relationships: &BTreeMap<String, Vec<Relationship>>,
    budget: &mut ParseBudget<'_>,
    elements: ShapeElements,
) -> Vec<DiagramDrawing> {
    let mut wanted: Vec<DiagramReference> = Vec::new();
    for slide in slides {
        collect_diagram_drawings(&slide.shapes, &slide.part_path, &mut wanted);
    }
    wanted.sort_by(|left, right| left.part_path.cmp(&right.part_path));
    wanted.dedup_by(|left, right| left.part_path == right.part_path);
    let mut drawings = Vec::new();
    let mut layouts = HashMap::new();
    'references: for reference in wanted {
        let slide_relationships = relationships
            .get(&reference.slide_part_path)
            .map(Vec::as_slice)
            .unwrap_or_default();
        let data_path = diagram_relationship_target(
            slide_relationships,
            &reference.relationship_ids,
            "/diagramData",
        )
        .or_else(|| sibling_diagram_part(&reference.part_path, "drawing", "data"));
        let mut data_budget = budget.scratch();
        let data_root = if let Some(path) = data_path.as_deref()
            && let Some(bytes) = parts.get(path)
        {
            parse_xml(bytes, path, &mut data_budget).ok()
        } else {
            None
        };
        let mut drawing_relationships = slide_relationships
            .iter()
            .filter(|relationship| relationship.has_type("/diagramDrawing"));
        let related_drawing = drawing_relationships.next().and_then(|first| {
            if drawing_relationships.next().is_none() {
                return first.resolved_target.clone();
            }
            slide_relationships
                .iter()
                .find(|relationship| {
                    relationship.has_type("/diagramDrawing")
                        && relationship.resolved_target.as_deref()
                            == Some(reference.part_path.as_str())
                })
                .and_then(|relationship| relationship.resolved_target.clone())
        });
        let drawing_paths = [
            data_root
                .as_ref()
                .and_then(|root| data_model_drawing_target(root, slide_relationships)),
            related_drawing,
            if data_path.as_deref() == Some(reference.part_path.as_str()) {
                sibling_diagram_part(&reference.part_path, "data", "drawing")
            } else {
                Some(reference.part_path.clone())
            },
        ];
        let mut seen = HashSet::new();
        for path in drawing_paths.into_iter().flatten() {
            if !seen.insert(path.clone()) {
                continue;
            }
            let Some(bytes) = parts.get(path.as_str()) else {
                continue;
            };
            let Ok(root) = parse_xml(bytes, &path, budget) else {
                continue;
            };
            let part_relationships = relationships
                .get(&path)
                .map(Vec::as_slice)
                .unwrap_or_default();
            let Ok(shapes) =
                parse_diagram_drawing(&root, part_relationships, &path, budget, elements)
            else {
                continue;
            };
            if !shapes.is_empty() {
                drawings.push(DiagramDrawing {
                    part_path: reference.part_path,
                    synthesis_extent: None,
                    shapes,
                });
                continue 'references;
            }
        }
        let (Some(data_path), Some(data_root)) = (data_path, data_root) else {
            continue;
        };
        if budget.charge_xml_from(&data_budget, &data_path).is_err() {
            continue;
        }
        let layout_path = diagram_relationship_target(
            slide_relationships,
            &reference.relationship_ids,
            "/diagramLayout",
        )
        .or_else(|| sibling_diagram_part(&data_path, "data", "layout"));
        let kind = if let Some(path) = layout_path.as_deref()
            && let Some(bytes) = parts.get(path)
        {
            let Some(kind) = *layouts.entry(path.to_owned()).or_insert_with(|| {
                parse_xml(bytes, path, budget)
                    .ok()
                    .map(|root| classify_layout(&root))
            }) else {
                continue;
            };
            kind
        } else {
            BasicLayout::List
        };
        if let Ok(shapes) = layout_basic_diagram(&data_root, kind, &data_path, budget) {
            drawings.push(DiagramDrawing {
                part_path: reference.part_path,
                synthesis_extent: Some([DIAGRAM_EXTENT_EMU; 2]),
                shapes,
            });
        }
    }
    drawings
}

fn collect_diagram_drawings(
    shapes: &[ShapeNode],
    slide_part_path: &str,
    output: &mut Vec<DiagramReference>,
) {
    for shape in shapes {
        match shape {
            ShapeNode::GraphicFrame(frame) => {
                if let GraphicFrameData::Diagram {
                    drawing_part_path: Some(path),
                    relationship_ids,
                } = &frame.data
                {
                    output.push(DiagramReference {
                        part_path: path.clone(),
                        slide_part_path: slide_part_path.to_owned(),
                        relationship_ids: relationship_ids.clone(),
                    });
                }
            }
            ShapeNode::Group(group) => {
                collect_diagram_drawings(&group.children, slide_part_path, output);
            }
            ShapeNode::Shape(_) | ShapeNode::Picture(_) => {}
        }
    }
}

fn diagram_relationship_target(
    relationships: &[Relationship],
    relationship_ids: &[String],
    type_suffix: &str,
) -> Option<String> {
    relationship_ids.iter().find_map(|id| {
        relationships
            .iter()
            .find(|relationship| relationship.id == *id && relationship.has_type(type_suffix))
            .and_then(|relationship| relationship.resolved_target.clone())
    })
}

fn sibling_diagram_part(path: &str, from_prefix: &str, to_prefix: &str) -> Option<String> {
    let (directory, name) = path.rsplit_once('/')?;
    let stem = name.strip_prefix(from_prefix)?;
    Some(format!("{directory}/{to_prefix}{stem}"))
}

fn data_model_drawing_target(
    data_root: &XmlElement,
    relationships: &[Relationship],
) -> Option<String> {
    let rel_id = data_root
        .descendants_named("dataModelExt")
        .first()
        .and_then(|element| {
            element
                .attribute("relId")
                .or_else(|| element.attribute_local("relId"))
        })?;
    relationships
        .iter()
        .find(|relationship| relationship.id == rel_id && relationship.has_type("/diagramDrawing"))
        .and_then(|relationship| relationship.resolved_target.clone())
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

/// The `p:clrMap` in effect for a slide: its own `p:clrMapOvr`, else its
/// layout's, else the master's. Every projection resolves scheme colours
/// through this one predicate.
pub fn effective_color_map(
    slide: Option<&Slide>,
    layout: Option<&SlideLayout>,
    master: Option<&SlideMaster>,
) -> ColorMap {
    slide
        .and_then(|slide| slide.color_map_override.clone())
        .or_else(|| layout.and_then(|layout| layout.color_map_override.clone()))
        .or_else(|| master.map(|master| master.color_map.clone()))
        .unwrap_or_default()
}

/// The theme a slide's colours resolve against: its master's theme part
/// carrying [`effective_color_map`].
pub fn slide_theme(
    package: &PptxPackage,
    slide_part_path: Option<&str>,
    layout_part_path: Option<&str>,
) -> Theme {
    let slide = slide_part_path
        .and_then(|path| package.slides.iter().find(|slide| slide.part_path == path));
    let layout = layout_part_path
        .or_else(|| slide.and_then(|slide| slide.layout_part_path.as_deref()))
        .and_then(|path| {
            package
                .layouts
                .iter()
                .find(|layout| layout.part_path == path)
        })
        .or_else(|| package.layouts.first());
    let master = master_for_layout(package, layout);
    let base = master
        .and_then(|master| master.theme_part_path.as_deref())
        .and_then(|path| package.themes.iter().find(|theme| theme.part_path == path))
        .or_else(|| package.themes.first());
    Theme {
        color_map: effective_color_map(slide, layout, master),
        ..base.map(|part| part.theme.clone()).unwrap_or_default()
    }
}

/// The master a layout belongs to, by relationship then by back-reference.
pub fn master_for_layout<'a>(
    package: &'a PptxPackage,
    layout: Option<&SlideLayout>,
) -> Option<&'a SlideMaster> {
    layout
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
        .or_else(|| package.masters.first())
}

fn parse_color_map(root: &XmlElement) -> ColorMap {
    root.child("clrMap").map(color_map).unwrap_or_default()
}

fn parse_color_map_override(root: &XmlElement) -> Option<ColorMap> {
    root.child("clrMapOvr")
        .and_then(|element| element.child("overrideClrMapping"))
        .map(color_map)
}

fn color_map(element: &XmlElement) -> ColorMap {
    let mut map = ColorMap::default();
    for (name, slot) in &element.attributes {
        map.set(name, slot);
    }
    map
}

#[cfg(test)]
mod tests {
    use super::*;

    const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
    const CHART_FIXTURE: &[u8] = include_bytes!("../tests/fixtures/chart-deck.pptx");
    const NUMBERED_FIXTURE: &[u8] = include_bytes!("../tests/fixtures/slide-number-fields.pptx");
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
    fn first_slide_num_is_read_signed_and_defaults_to_one() {
        assert_eq!(parse_pptx(FIXTURE).unwrap().presentation.first_slide_num, 1);
        assert_eq!(
            parse_pptx(NUMBERED_FIXTURE)
                .unwrap()
                .presentation
                .first_slide_num,
            10
        );
        for (value, expected) in [
            ("0", 0),
            ("-3", -3),
            ("-2147483648", i32::MIN),
            ("2147483647", i32::MAX),
            ("2147483648", 1),
            ("-2147483649", 1),
            ("ten", 1),
            ("", 1),
        ] {
            let package = parse_pptx(&demo_deck_numbered_from(value)).unwrap();
            assert_eq!(
                package.presentation.first_slide_num, expected,
                "firstSlideNum={value:?}"
            );
        }
    }

    fn demo_deck_numbered_from(value: &str) -> Vec<u8> {
        let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
        let slot = parts
            .iter_mut()
            .find(|(path, _)| path == "ppt/presentation.xml")
            .unwrap();
        let xml = String::from_utf8(slot.1.clone()).unwrap();
        let numbered = xml.replacen(
            "<p:presentation ",
            &format!("<p:presentation firstSlideNum=\"{value}\" "),
            1,
        );
        assert_ne!(numbered, xml);
        slot.1 = numbered.into_bytes();
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    fn assert_untouched_save_round_trip(deck: &[u8]) {
        let package = parse_pptx(deck).unwrap();
        let written = write_pptx(&package).unwrap();
        let before = ooxml_opc::unzip_parts(deck).unwrap();
        let after = ooxml_opc::unzip_parts(&written).unwrap();
        assert_eq!(after, before);
    }

    #[test]
    fn untouched_save_preserves_every_part_byte_and_order() {
        assert_untouched_save_round_trip(FIXTURE);
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
                background_picture: None,
                background_reference: None,
                declares_background: false,
                shapes: vec![chart_shape(id)],
                color_map_override: None,
                notes: String::new(),
                hidden: Some(false),
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
                background_picture: None,
                background_reference: None,
                shapes: Vec::new(),
                color_map_override: None,
            },
            SlideLayout {
                part_path: "ppt/slideLayouts/layout2.xml".to_owned(),
                name: None,
                layout_type: None,
                master_part_path: Some("ppt/slideMasters/master2.xml".to_owned()),
                show_master_shapes: true,
                background: None,
                background_picture: None,
                background_reference: None,
                shapes: Vec::new(),
                color_map_override: None,
            },
        ];
        let masters = vec![
            SlideMaster {
                part_path: "ppt/slideMasters/master1.xml".to_owned(),
                name: None,
                theme_part_path: Some("ppt/theme/theme1.xml".to_owned()),
                layout_part_paths: vec!["ppt/slideLayouts/layout1.xml".to_owned()],
                background: None,
                background_picture: None,
                background_reference: None,
                shapes: Vec::new(),
                text_styles: TextStyleSet::default(),
                color_map: ColorMap::default(),
            },
            SlideMaster {
                part_path: "ppt/slideMasters/master2.xml".to_owned(),
                name: None,
                theme_part_path: Some("ppt/theme/theme2.xml".to_owned()),
                layout_part_paths: vec!["ppt/slideLayouts/layout2.xml".to_owned()],
                background: None,
                background_picture: None,
                background_reference: None,
                shapes: Vec::new(),
                text_styles: TextStyleSet::default(),
                color_map: ColorMap::default(),
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
                format_scheme: Default::default(),
                background_pictures: Vec::new(),
            },
            ThemePart {
                part_path: "ppt/theme/theme2.xml".to_owned(),
                theme: second_theme,
                format_scheme: Default::default(),
                background_pictures: Vec::new(),
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

    /// A deck of SmartArt whose drawings each hold `shapes` shapes, `drawings` of them.
    fn smartart_deck(drawings: usize, shapes: usize) -> Vec<u8> {
        let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
        let frames = (0..drawings)
            .map(|index| {
                format!(
                    r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="{}" name="SmartArt {index}"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdData{index}" r:lo="" r:qs="" r:cs=""/></a:graphicData></a:graphic></p:graphicFrame>"#,
                    index + 900
                )
            })
            .collect::<String>();
        let links = (0..drawings)
            .map(|index| {
                format!(
                    r#"<Relationship Id="rIdData{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data{index}.xml"/><Relationship Id="rIdDrawing{index}" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing{index}.xml"/>"#
                )
            })
            .collect::<String>();
        let crowd = (0..shapes)
            .map(|_| r#"<dsp:sp modelId="{0}"><dsp:nvSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvSpPr/></dsp:nvSpPr><dsp:spPr/></dsp:sp>"#)
            .collect::<String>();
        let drawing = format!(
            r#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/>{crowd}</dsp:spTree></dsp:drawing>"#
        );
        for (path, bytes) in &mut parts {
            let edited = match path.as_str() {
                "ppt/slides/slide1.xml" => String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace("</p:spTree>", &format!("{frames}</p:spTree>")),
                "ppt/slides/_rels/slide1.xml.rels" => String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace("</Relationships>", &format!("{links}</Relationships>")),
                _ => continue,
            };
            *bytes = edited.into_bytes();
        }
        for index in 0..drawings {
            parts.push((
                format!("ppt/diagrams/drawing{index}.xml"),
                drawing.clone().into_bytes(),
            ));
        }
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    /// Every SmartArt drawing draws from the one package budget, so a deck of
    /// many cannot parse past it; the drawings that would are declined and the
    /// deck still opens.
    #[test]
    fn smartart_drawings_share_the_package_budget() {
        let whole = parse_pptx(&smartart_deck(3, 40)).unwrap();
        assert_eq!(whole.diagram_drawings.len(), 3);
        let spent = whole
            .slides
            .iter()
            .map(|slide| count_shapes(&slide.shapes))
            .sum::<usize>()
            + whole
                .layouts
                .iter()
                .map(|layout| count_shapes(&layout.shapes))
                .sum::<usize>()
            + whole
                .masters
                .iter()
                .map(|master| count_shapes(&master.shapes))
                .sum::<usize>();
        let limits = ParseLimits {
            max_shapes: spent + 100,
            ..ParseLimits::default()
        };

        let package = parse_pptx_with_limits(&smartart_deck(3, 40), &limits).unwrap();

        assert_eq!(package.slides.len(), whole.slides.len());
        assert_eq!(
            package
                .diagram_drawings
                .iter()
                .map(|drawing| drawing.shapes.len())
                .collect::<Vec<_>>(),
            [40, 40]
        );
    }

    fn count_shapes(shapes: &[ShapeNode]) -> usize {
        shapes
            .iter()
            .map(|shape| match shape {
                ShapeNode::Group(group) => 1 + count_shapes(&group.children),
                _ => 1,
            })
            .sum()
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

    fn diagram_deck_without_drawing(
        data_xml: &[u8],
        layout_xml: Option<&[u8]>,
        include_drawing_rel: bool,
    ) -> Vec<u8> {
        let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
        let frame = r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="901" name="Diagram 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdData1" r:lo="rIdLayout1" r:qs="" r:cs=""/></a:graphicData></a:graphic></p:graphicFrame>"#;
        let mut rels = String::from(
            r#"<Relationship Id="rIdData1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/data1.xml"/>"#,
        );
        if layout_xml.is_some() {
            rels.push_str(
                r#"<Relationship Id="rIdLayout1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/layout1.xml"/>"#,
            );
        }
        if include_drawing_rel {
            rels.push_str(
                r#"<Relationship Id="rIdDrawing1" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/>"#,
            );
        }
        for (path, bytes) in &mut parts {
            let edited = match path.as_str() {
                "ppt/slides/slide1.xml" => String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace("</p:spTree>", &format!("{frame}</p:spTree>")),
                "ppt/slides/_rels/slide1.xml.rels" => String::from_utf8(bytes.clone())
                    .unwrap()
                    .replace("</Relationships>", &format!("{rels}</Relationships>")),
                _ => continue,
            };
            *bytes = edited.into_bytes();
        }
        parts.push(("ppt/diagrams/data1.xml".to_owned(), data_xml.to_vec()));
        parts.push((
            "ppt/diagrams/drawing1.xml".to_owned(),
            br#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id="0" name=""/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/></dsp:spTree></dsp:drawing>"#.to_vec(),
        ));
        if let Some(layout) = layout_xml {
            parts.push(("ppt/diagrams/layout1.xml".to_owned(), layout.to_vec()));
        }
        ooxml_opc::rezip_parts(&parts).unwrap()
    }

    const BASIC_DIAGRAM_DATA: &[u8] = br#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="p1"><dgm:t><a:p><a:r><a:t>First</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="p2"><dgm:t><a:p><a:r><a:t>Second</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst></dgm:dataModel>"#;

    #[test]
    fn cached_diagram_ignores_malformed_data() {
        let mut parts = ooxml_opc::unzip_parts(&smartart_deck(1, 1)).unwrap();
        parts.push(("ppt/diagrams/data0.xml".to_owned(), b"<broken>".to_vec()));
        let package = parse_pptx(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
        assert_eq!(package.diagram_drawings.len(), 1);
        assert_eq!(package.diagram_drawings[0].shapes.len(), 1);
        assert_eq!(package.diagram_drawings[0].synthesis_extent, None);
    }

    #[test]
    fn cached_diagram_does_not_charge_data_against_remaining_budget() {
        let deck = smartart_deck(1, 1);
        let package = parse_pptx(&deck).unwrap();
        let mut source = ooxml_opc::unzip_parts(&deck).unwrap();
        let drawing = source
            .iter()
            .find(|(path, _)| path == "ppt/diagrams/drawing0.xml")
            .unwrap()
            .1
            .clone();
        for data in [
            b"<dataModel/>".to_vec(),
            format!("<dataModel>{}</dataModel>", "x".repeat(4_096)).into_bytes(),
        ] {
            source.retain(|(path, _)| path != "ppt/diagrams/data0.xml");
            source.push(("ppt/diagrams/data0.xml".to_owned(), data));
            let parts: HashMap<_, _> = source
                .iter()
                .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
                .collect();
            let spent = b"<spent/>";
            let limits = ParseLimits {
                max_xml_bytes: spent.len() + drawing.len(),
                ..ParseLimits::default()
            };
            let mut budget = ParseBudget::new(&limits);
            parse_xml(spent, "spent.xml", &mut budget).unwrap();
            let drawings = parse_diagram_drawings(
                &parts,
                &package.slides,
                &package.relationships,
                &mut budget,
                ShapeElements::WithConnectors,
            );
            assert_eq!(drawings.len(), 1);
            assert_eq!(drawings[0].shapes.len(), 1);
            assert_eq!(drawings[0].synthesis_extent, None);
            assert!(parse_xml(b"<extra/>", "extra.xml", &mut budget).is_err());
        }
    }

    #[test]
    fn cached_diagram_falls_back_from_unusable_explicit_drawing() {
        let data = String::from_utf8(BASIC_DIAGRAM_DATA.to_vec()).unwrap().replace(
            "</dgm:dataModel>",
            r#"<dgm:extLst><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rIdDrawing1"/></dgm:extLst></dgm:dataModel>"#,
        );
        for candidate in [
            None,
            Some(b"<drawing/>".as_slice()),
            Some(b"<broken>".as_slice()),
        ] {
            let mut parts = ooxml_opc::unzip_parts(&smartart_deck(1, 1)).unwrap();
            parts.push((
                "ppt/diagrams/data0.xml".to_owned(),
                data.as_bytes().to_vec(),
            ));
            let rels = &mut parts
                .iter_mut()
                .find(|(path, _)| path == "ppt/slides/_rels/slide1.xml.rels")
                .unwrap()
                .1;
            *rels = String::from_utf8(rels.clone()).unwrap().replace(
                "</Relationships>",
                r#"<Relationship Id="rIdDrawing1" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing1.xml"/></Relationships>"#,
            ).into_bytes();
            if let Some(bytes) = candidate {
                parts.push(("ppt/diagrams/drawing1.xml".to_owned(), bytes.to_vec()));
            }
            let package = parse_pptx(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
            assert_eq!(package.diagram_drawings.len(), 1);
            assert_eq!(package.diagram_drawings[0].shapes.len(), 1);
            assert_eq!(package.diagram_drawings[0].synthesis_extent, None);
        }
    }

    #[test]
    fn synthesized_diagram_layout_obeys_xml_limits() {
        let layout = format!(
            r#"<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="process">{}<dgm:title>{}</dgm:title>{}</dgm:layoutDef>"#,
            "<dgm:layoutNode>".repeat(7),
            "x".repeat(4_096),
            "</dgm:layoutNode>".repeat(7)
        );
        let deck = diagram_deck_without_drawing(BASIC_DIAGRAM_DATA, Some(layout.as_bytes()), false);
        let package = parse_pptx(&deck).unwrap();
        let source = ooxml_opc::unzip_parts(&deck).unwrap();
        let parts: HashMap<_, _> = source
            .iter()
            .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
            .collect();
        for limits in [
            ParseLimits {
                max_xml_bytes: 2_048,
                ..ParseLimits::default()
            },
            ParseLimits {
                max_xml_text_bytes: 100,
                ..ParseLimits::default()
            },
            ParseLimits {
                max_xml_depth: 7,
                ..ParseLimits::default()
            },
        ] {
            let mut budget = ParseBudget::new(&limits);
            let drawings = parse_diagram_drawings(
                &parts,
                &package.slides,
                &package.relationships,
                &mut budget,
                ShapeElements::WithConnectors,
            );
            assert!(
                drawings.is_empty(),
                "layout exceeding XML limits must be declined"
            );
        }
    }

    #[test]
    fn synthesized_diagrams_parse_a_shared_layout_once() {
        let layout = br#"<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="process"/>"#;
        let deck = diagram_deck_without_drawing(BASIC_DIAGRAM_DATA, Some(layout), false);
        let mut package = parse_pptx(&deck).unwrap();
        let mut frame = package.slides[0]
            .shapes
            .iter()
            .find(|shape| matches!(shape, ShapeNode::GraphicFrame(_)))
            .unwrap()
            .clone();
        let ShapeNode::GraphicFrame(value) = &mut frame else {
            unreachable!()
        };
        let GraphicFrameData::Diagram {
            drawing_part_path,
            relationship_ids,
            ..
        } = &mut value.data
        else {
            unreachable!()
        };
        *drawing_part_path = Some("ppt/diagrams/data2.xml".to_owned());
        relationship_ids[0] = "rIdData2".to_owned();
        package.slides[0].shapes.push(frame);
        let rels = package
            .relationships
            .get_mut("ppt/slides/slide1.xml")
            .unwrap();
        let mut data_rel = rels
            .iter()
            .find(|rel| rel.id == "rIdData1")
            .unwrap()
            .clone();
        data_rel.id = "rIdData2".to_owned();
        data_rel.resolved_target = Some("ppt/diagrams/data2.xml".to_owned());
        rels.push(data_rel);
        let mut source = ooxml_opc::unzip_parts(&deck).unwrap();
        source.push((
            "ppt/diagrams/data2.xml".to_owned(),
            BASIC_DIAGRAM_DATA.to_vec(),
        ));
        let parts: HashMap<_, _> = source
            .iter()
            .map(|(path, bytes)| (path.as_str(), bytes.as_slice()))
            .collect();
        let limits = ParseLimits::default();
        let mut expected = ParseBudget::new(&limits);
        for path in [
            "ppt/diagrams/data1.xml",
            "ppt/diagrams/drawing1.xml",
            "ppt/diagrams/layout1.xml",
            "ppt/diagrams/data2.xml",
        ] {
            parse_xml(parts[path], path, &mut expected).unwrap();
        }
        let limits = ParseLimits {
            max_xml_events: expected.xml_events_spent(),
            ..limits
        };
        let mut budget = ParseBudget::new(&limits);
        let drawings = parse_diagram_drawings(
            &parts,
            &package.slides,
            &package.relationships,
            &mut budget,
            ShapeElements::WithConnectors,
        );
        assert_eq!(drawings.len(), 2);
        assert!(drawings.iter().all(|drawing| drawing.shapes.len() == 2));
        assert_eq!(budget.xml_events_spent(), expected.xml_events_spent());
    }

    #[test]
    fn explicit_diagram_drawing_wins_over_filename_guess() {
        for extension in [false, true] {
            let data = String::from_utf8(BASIC_DIAGRAM_DATA.to_vec()).unwrap().replace(
                "</dgm:dataModel>",
                if extension { r#"<dgm:extLst><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="rIdDrawing2"/></dgm:extLst></dgm:dataModel>"# } else { "</dgm:dataModel>" },
            );
            let deck = diagram_deck_without_drawing(data.as_bytes(), None, extension);
            let mut parts = ooxml_opc::unzip_parts(&deck).unwrap();
            let drawing = |name: &str| {
                format!(r#"<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:spTree><dsp:sp><dsp:nvSpPr><dsp:cNvPr id="1" name="{name}"/></dsp:nvSpPr><dsp:spPr/></dsp:sp></dsp:spTree></dsp:drawing>"#).into_bytes()
            };
            parts
                .iter_mut()
                .find(|(path, _)| path == "ppt/diagrams/drawing1.xml")
                .unwrap()
                .1 = drawing("Unrelated");
            parts.push(("ppt/diagrams/drawing2.xml".to_owned(), drawing("Explicit")));
            let rels = &mut parts
                .iter_mut()
                .find(|(path, _)| path == "ppt/slides/_rels/slide1.xml.rels")
                .unwrap()
                .1;
            *rels = String::from_utf8(rels.clone()).unwrap().replace("</Relationships>", r#"<Relationship Id="rIdDrawing2" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="../diagrams/drawing2.xml"/></Relationships>"#).into_bytes();
            let package = parse_pptx(&ooxml_opc::rezip_parts(&parts).unwrap()).unwrap();
            let ShapeNode::Shape(shape) = &package.diagram_drawings[0].shapes[0] else {
                panic!("expected drawing shape")
            };
            assert_eq!(shape.base.name, "Explicit");
        }
    }

    #[test]
    fn text_free_diagram_with_empty_drawing_records_an_empty_diagram_drawing() {
        let data = br#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="{0}" type="doc"/><dgm:pt modelId="{1}" type="pres"/><dgm:pt modelId="{2}" type="parTrans"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>ignored</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/></dgm:dataModel>"#;
        for include_drawing_rel in [true, false] {
            let deck = diagram_deck_without_drawing(data, None, include_drawing_rel);
            let package = parse_pptx(&deck).unwrap();
            assert_eq!(package.diagram_drawings.len(), 1);
            assert!(package.diagram_drawings[0].shapes.is_empty());
        }
    }

    #[test]
    fn diagram_without_cached_drawing_synthesizes_boxes_from_data_model_and_round_trips() {
        let data = br#"<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="doc" type="doc"/><dgm:pt modelId="p1"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First header</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="c1"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>First child</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="p2"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second header</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="c2"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second child</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst><dgm:cxn modelId="x1" srcId="doc" destId="p1" srcOrd="0"/><dgm:cxn modelId="x2" srcId="doc" destId="p2" srcOrd="1"/><dgm:cxn modelId="x3" srcId="p1" destId="c1" srcOrd="0"/><dgm:cxn modelId="x4" srcId="p2" destId="c2" srcOrd="0"/></dgm:cxnLst></dgm:dataModel>"#;
        let list_layout = br#"<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/hList1"/>"#;
        let deck = diagram_deck_without_drawing(data, Some(list_layout), false);
        let package = parse_pptx(&deck).unwrap();
        assert_eq!(package.diagram_drawings.len(), 1);
        let shapes = &package.diagram_drawings[0].shapes;
        assert_eq!(shapes.len(), 2);
        let mut texts = Vec::new();
        for node in shapes {
            let ShapeNode::Shape(shape) = node else {
                panic!("expected a synthesized shape");
            };
            let t = &shape.base.transform;
            assert!(t.x >= 0 && t.y >= 0);
            assert!(t.width > 0 && t.height > 0);
            assert!(t.x + t.width <= 3_657_600);
            assert!(t.y + t.height <= 1_828_800);
            let body = shape.text.as_ref().expect("shape has text body");
            for paragraph in &body.paragraphs {
                for run in &paragraph.runs {
                    texts.push(run.text.clone());
                }
            }
        }
        assert_eq!(
            texts,
            [
                "First header",
                "First child",
                "Second header",
                "Second child"
            ]
        );
        assert_untouched_save_round_trip(&deck);

        let process_layout = br#"<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:microsoft.com/office/officeart/2005/8/layout/process5"/>"#;
        let process_deck = diagram_deck_without_drawing(data, Some(process_layout), true);
        let process_package = parse_pptx(&process_deck).unwrap();
        assert_eq!(process_package.diagram_drawings.len(), 1);
        assert_eq!(process_package.diagram_drawings[0].shapes.len(), 4);
        assert_untouched_save_round_trip(&process_deck);
    }
}
