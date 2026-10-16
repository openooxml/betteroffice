use quick_xml::Reader;
use quick_xml::events::{BytesStart, Event};
use xlsx_model::SheetId;

use crate::ParseError;
use crate::package::{PreservedPackage, XmlTemplate};
use crate::xml::{attr, xml_err};

pub(super) struct ViewPlan {
    pub(super) book_views: Vec<u8>,
    active_tabs: Vec<usize>,
    /// Per workbook view: the active tab is still the source's active sheet,
    /// so its other selected source sheets stay grouped.
    keeps_group: Vec<bool>,
    has_source: Vec<bool>,
}

impl ViewPlan {
    /// `origins` are the claimed origins, naming each source sheet at most once.
    pub(super) fn new(
        package: &PreservedPackage,
        origins: &[Option<usize>],
        active_sheet: SheetId,
    ) -> Result<Option<Self>, ParseError> {
        let structural = origins.len() != package.sheets.len()
            || origins
                .iter()
                .enumerate()
                .any(|(index, origin)| *origin != Some(index));
        if !structural && active_sheet == package.active_sheet {
            return Ok(None);
        }
        let mapped = |index| {
            origins
                .iter()
                .position(|origin| *origin == Some(index))
                .unwrap_or(index.min(origins.len().saturating_sub(1)))
        };
        let source = package.workbook_template.child("bookViews");
        let mut template = XmlTemplate::capture(
            source
                .map(|child| child.bytes.as_slice())
                .unwrap_or(b"<bookViews></bookViews>"),
        )?;
        let keeps = |active: usize, source: usize| origins.get(active) == Some(&Some(source));
        let mut active_tabs = Vec::new();
        let mut keeps_group = Vec::new();
        for index in 0..template.children.len() {
            if !template.is_root_child_named(&template.children[index], "workbookView") {
                continue;
            }
            let bytes = &template.children[index].bytes;
            let element = root_element(bytes)?;
            let (active, source_active) = if active_tabs.is_empty() {
                (active_sheet.0 as usize, package.active_sheet.0 as usize)
            } else {
                let source_active = index_attribute(&element, b"activeTab")?;
                (mapped(source_active), source_active)
            };
            let first = mapped(index_attribute(&element, b"firstSheet")?);
            let mut changes = vec![("activeTab", active.to_string())];
            if first != 0 || attr(&element, b"firstSheet")?.is_some() {
                changes.push(("firstSheet", first.to_string()));
            }
            template.children[index].bytes = patch_attributes(bytes, &changes)?;
            active_tabs.push(active);
            keeps_group.push(keeps(active, source_active));
        }
        let book_views = if active_tabs.is_empty() {
            let active = active_sheet.0 as usize;
            active_tabs.push(active);
            keeps_group.push(keeps(active, package.active_sheet.0 as usize));
            let first = mapped(0);
            let first_attribute = if first == 0 {
                String::new()
            } else {
                format!(" firstSheet=\"{first}\"")
            };
            let view = format!(
                "<workbookView activeTab=\"{}\"{first_attribute}/>",
                active_sheet.0
            );
            template.render(vec![("workbookView", Some(view.into_bytes()))], view_rank)?
        } else {
            template.render(vec![], view_rank)?
        };
        Ok(Some(Self {
            book_views,
            active_tabs,
            keeps_group,
            has_source: origins.iter().map(Option::is_some).collect(),
        }))
    }

    pub(super) fn patch_sheet(
        &self,
        source: &[u8],
        sheet: usize,
    ) -> Result<Option<Vec<u8>>, ParseError> {
        let mut template = XmlTemplate::capture(source)?;
        let views = template
            .children
            .iter()
            .position(|child| template.is_root_child_named(child, "sheetViews"));
        let mut view_template = XmlTemplate::capture(
            views
                .map(|index| template.children[index].bytes.as_slice())
                .unwrap_or(b"<sheetViews></sheetViews>"),
        )?;
        let mut present = vec![false; self.active_tabs.len()];
        let has_source = self.has_source.get(sheet).copied().unwrap_or(false);
        let mut changed = false;
        for index in 0..view_template.children.len() {
            if !view_template.is_root_child_named(&view_template.children[index], "sheetView") {
                continue;
            }
            let bytes = &view_template.children[index].bytes;
            let element = root_element(bytes)?;
            let view = index_attribute(&element, b"workbookViewId")?;
            let Some(active) = self.active_tabs.get(view) else {
                continue;
            };
            present[view] = true;
            let was_selected = attr(&element, b"tabSelected")?
                .is_some_and(|value| value == "1" || value == "true");
            let selected =
                *active == sheet || (self.keeps_group[view] && has_source && was_selected);
            if selected != was_selected {
                changed = true;
                view_template.children[index].bytes = patch_attributes(
                    bytes,
                    &[("tabSelected", if selected { "1" } else { "0" }.to_owned())],
                )?;
            }
        }
        let missing = self
            .active_tabs
            .iter()
            .enumerate()
            .filter(|(view, active)| !present[*view] && **active == sheet)
            .map(|(view, _)| {
                view_template.qualify_fragment(
                    format!("<sheetView workbookViewId=\"{view}\" tabSelected=\"1\"/>").as_bytes(),
                )
            })
            .collect::<Result<Vec<_>, _>>()?;
        if !changed && missing.is_empty() {
            return Ok(None);
        }
        let patched_views = if missing.is_empty() {
            view_template.render(vec![], view_rank)?
        } else {
            let retained = view_template
                .children_named("sheetView")
                .map(|child| child.bytes.clone())
                .chain(missing)
                .collect::<Vec<_>>();
            view_template.render_repeated("sheetView", &retained, &[])?
        };
        let output = if let Some(index) = views {
            template.children[index].bytes = patched_views;
            template.render(vec![], super::worksheet_child_rank)?
        } else if self.active_tabs.contains(&sheet) {
            template.render(
                vec![("sheetViews", Some(patched_views))],
                super::worksheet_child_rank,
            )?
        } else {
            return Ok(None);
        };
        Ok((output != source).then_some(output))
    }
}

fn index_attribute(element: &BytesStart<'_>, name: &[u8]) -> Result<usize, ParseError> {
    Ok(attr(element, name)?
        .and_then(|value| value.parse().ok())
        .unwrap_or(0))
}

fn root_element(source: &[u8]) -> Result<BytesStart<'_>, ParseError> {
    let mut reader = Reader::from_reader(source);
    loop {
        match reader.read_event().map_err(xml_err)? {
            Event::Start(element) | Event::Empty(element) => return Ok(element),
            Event::Eof => return Err(ParseError::Malformed("missing view element".into())),
            _ => {}
        }
    }
}

fn patch_attributes(source: &[u8], changes: &[(&str, String)]) -> Result<Vec<u8>, ParseError> {
    let mut reader = Reader::from_reader(source);
    loop {
        let before = reader.buffer_position() as usize;
        let event = reader.read_event().map_err(xml_err)?;
        let empty = matches!(event, Event::Empty(_));
        let element = match event {
            Event::Start(element) | Event::Empty(element) => element,
            Event::Eof => return Err(ParseError::Malformed("missing view element".into())),
            _ => continue,
        };
        let mut replacements = Vec::new();
        let mut additions = String::new();
        for (name, value) in changes {
            let mut found = false;
            for attribute in element.attributes() {
                let attribute = attribute.map_err(xml_err)?;
                if attribute.key.as_ref() == name.as_bytes() {
                    found = true;
                    if attribute.value.as_ref() != value.as_bytes() {
                        let offset = attribute
                            .value
                            .as_ptr()
                            .addr()
                            .checked_sub(element.as_ref().as_ptr().addr())
                            .ok_or_else(|| {
                                ParseError::Malformed("invalid view attribute span".into())
                            })?;
                        let start = before + 1 + offset;
                        replacements.push((
                            start,
                            start + attribute.value.len(),
                            value.as_bytes().to_vec(),
                        ));
                    }
                    break;
                }
            }
            if !found {
                additions.push_str(&format!(" {name}=\"{value}\""));
            }
        }
        if !additions.is_empty() {
            let position = reader.buffer_position() as usize - if empty { 2 } else { 1 };
            replacements.push((position, position, additions.into_bytes()));
        }
        replacements.sort_by_key(|(start, _, _)| *start);
        let mut output = Vec::with_capacity(source.len());
        let mut cursor = 0;
        for (start, end, value) in replacements {
            output.extend_from_slice(&source[cursor..start]);
            output.extend_from_slice(&value);
            cursor = end;
        }
        output.extend_from_slice(&source[cursor..]);
        return Ok(output);
    }
}

fn view_rank(name: &str) -> usize {
    if name == "workbookView" || name == "sheetView" {
        0
    } else {
        usize::MAX
    }
}
