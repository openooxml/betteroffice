use super::*;

#[derive(Default)]
pub(super) struct StyleOverrides {
    pub fonts: HashMap<usize, Vec<u8>>,
    pub fills: HashMap<usize, Vec<u8>>,
    pub borders: HashMap<usize, Vec<u8>>,
    pub xfs: HashMap<usize, Vec<u8>>,
}

pub(super) struct StylePlan {
    pub styles: Stylesheet,
    pub pairs: HashMap<(u32, u32), u32>,
    pub overrides: StyleOverrides,
}

impl StylePlan {
    pub fn new(
        wb: &Workbook,
        package: &PreservedPackage,
        origins: &[Option<usize>],
        axes: &[Option<SheetAxes>],
    ) -> Result<Self, ParseError> {
        let mut plan = Self {
            styles: wb.styles.clone(),
            pairs: HashMap::new(),
            overrides: StyleOverrides::default(),
        };
        let Some(template) = &package.stylesheet_template else {
            return Ok(plan);
        };
        let original = &package.original_workbook.styles;
        let source_fonts = pool_items(template, "fonts", "font")?;
        let source_fills = pool_items(template, "fills", "fill")?;
        let source_borders = pool_items(template, "borders", "border")?;
        let source_xfs = pool_items(template, "cellXfs", "xf")?;
        let mut fonts = pool_xml(&source_fonts, &wb.styles.fonts, &original.fonts, write_font)?;
        let mut fills = pool_xml(&source_fills, &wb.styles.fills, &original.fills, write_fill)?;
        let mut borders = pool_xml(
            &source_borders,
            &wb.styles.borders,
            &original.borders,
            write_border,
        )?;
        let mut xfs = pool_xml(
            &source_xfs,
            &wb.styles.cell_xfs,
            &original.cell_xfs,
            write_xf,
        )?;
        for (index, sheet) in wb.sheets.iter().enumerate() {
            let Some((origin, axes)) = origins[index].zip(axes.get(index).and_then(Option::as_ref))
            else {
                continue;
            };
            let source = &package.original_workbook.sheets[origin];
            for (at, cell) in sheet.iter_cells() {
                let Some((row, col)) = axes.rows.source(at.row).zip(axes.cols.source(at.col))
                else {
                    continue;
                };
                let old = source
                    .cell(CellRef::new(row, col))
                    .and_then(|cell| cell.style)
                    .unwrap_or(0);
                let Some(current) = cell.style else {
                    continue;
                };
                if plan.pairs.contains_key(&(old, current))
                    || original.resolved_format(Some(old))
                        == wb.styles.resolved_format(Some(current))
                {
                    continue;
                }
                let (Some(base), Some(target)) = (original.xf(old), wb.styles.xf(current)) else {
                    continue;
                };
                let Some(source_xml) = source_xfs.get(old as usize) else {
                    continue;
                };
                let mut xf = target.clone();
                let before = original.cell_format(Some(old));
                let after = wb.styles.cell_format(Some(current));
                if before.number_format == after.number_format {
                    xf.num_fmt_id = base.num_fmt_id;
                }
                if before.alignment == after.alignment {
                    xf.alignment = base.alignment.clone();
                }
                xf.font = if before.font == after.font {
                    base.font
                } else {
                    let bytes = source_fonts.get(base.font.unwrap_or(0) as usize);
                    let bytes = match bytes {
                        Some(bytes) => patch_font(bytes, &before.font, &after.font)?,
                        None => fragment(|writer| write_font(writer, &after.font))?,
                    };
                    Some(intern_xml(
                        &mut plan.styles.fonts,
                        &mut fonts,
                        &mut plan.overrides.fonts,
                        after.font.clone(),
                        bytes,
                    ))
                };
                xf.fill = if before.fill == after.fill {
                    base.fill
                } else {
                    let bytes = source_fills.get(base.fill.unwrap_or(0) as usize);
                    let bytes = match bytes {
                        Some(bytes) => patch_fill(bytes, &after.fill)?,
                        None => fragment(|writer| write_fill(writer, &after.fill))?,
                    };
                    Some(intern_xml(
                        &mut plan.styles.fills,
                        &mut fills,
                        &mut plan.overrides.fills,
                        after.fill.clone(),
                        bytes,
                    ))
                };
                xf.border = if before.border == after.border {
                    base.border
                } else {
                    let bytes = source_borders.get(base.border.unwrap_or(0) as usize);
                    let bytes = match bytes {
                        Some(bytes) => patch_border(bytes, &before.border, &after.border)?,
                        None => fragment(|writer| write_border(writer, &after.border))?,
                    };
                    Some(intern_xml(
                        &mut plan.styles.borders,
                        &mut borders,
                        &mut plan.overrides.borders,
                        after.border.clone(),
                        bytes,
                    ))
                };
                let bytes = patch_xf(source_xml, base, &xf)?;
                let index = intern_xml(
                    &mut plan.styles.cell_xfs,
                    &mut xfs,
                    &mut plan.overrides.xfs,
                    xf,
                    bytes,
                );
                plan.pairs.insert((old, current), index);
            }
        }
        Ok(plan)
    }
}

fn pool_items(template: &XmlTemplate, pool: &str, item: &str) -> Result<Vec<Vec<u8>>, ParseError> {
    let Some(pool) = template.child(pool) else {
        return Ok(Vec::new());
    };
    Ok(XmlTemplate::capture(&pool.bytes)?
        .children_named(item)
        .map(|child| child.bytes.clone())
        .collect())
}

fn pool_xml<T: PartialEq>(
    sources: &[Vec<u8>],
    current: &[T],
    original: &[T],
    write: fn(&mut Writer<Vec<u8>>, &T) -> io::Result<()>,
) -> Result<Vec<Vec<u8>>, ParseError> {
    current
        .iter()
        .enumerate()
        .map(|(index, value)| {
            if original.get(index) == Some(value)
                && let Some(bytes) = sources.get(index)
            {
                return Ok(bytes.clone());
            }
            fragment(|writer| write(writer, value))
        })
        .collect()
}

fn intern_xml<T>(
    values: &mut Vec<T>,
    xml: &mut Vec<Vec<u8>>,
    overrides: &mut HashMap<usize, Vec<u8>>,
    value: T,
    bytes: Vec<u8>,
) -> u32 {
    let key = xml_key(&bytes);
    if let Some(index) = xml
        .iter()
        .position(|candidate| candidate == &bytes || xml_key(candidate) == key)
    {
        return index as u32;
    }
    let index = values.len();
    values.push(value);
    xml.push(bytes.clone());
    overrides.insert(index, bytes);
    index as u32
}

fn xml_key(bytes: &[u8]) -> Vec<u8> {
    let Ok(template) = XmlTemplate::capture(bytes) else {
        return bytes.to_vec();
    };
    let mut reader = Reader::from_reader(bytes);
    let name = match reader.read_event() {
        Ok(Event::Start(element) | Event::Empty(element)) => {
            String::from_utf8_lossy(element.local_name().as_ref()).into_owned()
        }
        _ => return bytes.to_vec(),
    };
    if ![
        "font",
        "xf",
        "alignment",
        "fill",
        "patternFill",
        "border",
        "left",
        "right",
        "top",
        "bottom",
        "diagonal",
        "b",
        "i",
        "u",
        "strike",
        "sz",
        "color",
        "fgColor",
        "bgColor",
        "name",
        "family",
        "charset",
        "scheme",
        "protection",
    ]
    .contains(&name.as_str())
    {
        return bytes.to_vec();
    }
    let Ok(attributes) = attributes_from_fragment(bytes) else {
        return bytes.to_vec();
    };
    let mut attributes: Vec<_> = attributes
        .into_iter()
        .filter(|attribute| !attribute.name.starts_with("xmlns"))
        .map(|attribute| (attribute.name, attribute.value))
        .collect();
    attributes.sort();
    let mut children: Vec<_> = template
        .children
        .iter()
        .map(|child| xml_key(&child.bytes))
        .collect();
    if name == "font" {
        children.sort();
    }
    format!("{name:?}{attributes:?}{children:?}").into_bytes()
}

fn patch_font(source: &[u8], base: &Font, font: &Font) -> Result<Vec<u8>, ParseError> {
    let template = XmlTemplate::capture(source)?;
    let generated = XmlTemplate::capture(&fragment(|writer| write_font(writer, font))?)?;
    let changes = [
        ("b", base.bold != font.bold),
        ("i", base.italic != font.italic),
        ("u", base.underline != font.underline),
        ("strike", base.strike != font.strike),
        ("sz", base.size_pt != font.size_pt),
        ("color", base.color != font.color),
        ("name", base.name != font.name),
    ];
    template.render(
        changes
            .into_iter()
            .filter(|(_, changed)| *changed)
            .map(|(name, _)| (name, generated.child(name).map(|child| child.bytes.clone())))
            .collect(),
        |_| 0,
    )
}

fn patch_fill(source: &[u8], fill: &Fill) -> Result<Vec<u8>, ParseError> {
    let template = XmlTemplate::capture(source)?;
    let generated = XmlTemplate::capture(&fragment(|writer| write_fill(writer, fill))?)?;
    let pattern = match (template.child("patternFill"), fill) {
        (Some(source), Fill::Solid(color)) => {
            Some(XmlTemplate::capture(&source.bytes)?.render_with_attributes(
                vec![(
                    "fgColor",
                    Some(fragment(|writer| write_color(writer, "fgColor", color))?),
                )],
                |_| 0,
                &[("patternType", Some("solid".into()))],
            )?)
        }
        _ => generated
            .child("patternFill")
            .map(|child| child.bytes.clone()),
    };
    template.render(
        vec![("patternFill", pattern), ("gradientFill", None)],
        |_| 0,
    )
}

fn patch_border(source: &[u8], base: &Border, border: &Border) -> Result<Vec<u8>, ParseError> {
    let template = XmlTemplate::capture(source)?;
    let generated = XmlTemplate::capture(&fragment(|writer| write_border(writer, border))?)?;
    let mut replacements = Vec::new();
    for (name, base, edge) in [
        ("left", &base.left, &border.left),
        ("right", &base.right, &border.right),
        ("top", &base.top, &border.top),
        ("bottom", &base.bottom, &border.bottom),
    ] {
        if base == edge {
            continue;
        }
        let replacement = match (template.child(name), base, edge) {
            (Some(source), Some(base), Some(edge)) => {
                let source = XmlTemplate::capture(&source.bytes)?;
                let mut children = Vec::new();
                if base.color != edge.color {
                    children.push((
                        "color",
                        edge.color
                            .as_ref()
                            .map(|color| fragment(|writer| write_color(writer, "color", color)))
                            .transpose()?,
                    ));
                }
                let attributes = if base.style != edge.style {
                    vec![("style", Some(edge.style.as_sml().into()))]
                } else {
                    Vec::new()
                };
                Some(source.render_with_attributes(children, |_| 0, &attributes)?)
            }
            _ => generated.child(name).map(|child| child.bytes.clone()),
        };
        replacements.push((name, replacement));
    }
    template.render(replacements, |name| match name {
        "left" => 0,
        "right" => 1,
        "top" => 2,
        "bottom" => 3,
        "diagonal" => 4,
        _ => 5,
    })
}

fn patch_xf(source: &[u8], base: &Xf, xf: &Xf) -> Result<Vec<u8>, ParseError> {
    let template = XmlTemplate::capture(source)?;
    let mut attributes = Vec::new();
    for (name, flag, base, value) in [
        ("fontId", "applyFont", base.font, xf.font),
        ("fillId", "applyFill", base.fill, xf.fill),
        ("borderId", "applyBorder", base.border, xf.border),
        (
            "numFmtId",
            "applyNumberFormat",
            base.num_fmt_id.map(u32::from),
            xf.num_fmt_id.map(u32::from),
        ),
    ] {
        if base != value {
            attributes.push((name, Some(value.unwrap_or(0).to_string())));
            attributes.push((flag, Some("1".into())));
        }
    }
    let mut replacements = Vec::new();
    if base.alignment != xf.alignment {
        let alignment = xf.alignment.clone().unwrap_or_default();
        let before = base.alignment.clone().unwrap_or_default();
        let changes = [
            (
                "horizontal",
                before.h != alignment.h,
                alignment.h.map(|value| value.as_sml().into()),
            ),
            (
                "vertical",
                before.v != alignment.v,
                alignment.v.map(|value| value.as_sml().into()),
            ),
            (
                "wrapText",
                before.wrap_text != alignment.wrap_text,
                alignment.wrap_text.then(|| "1".into()),
            ),
            (
                "shrinkToFit",
                before.shrink_to_fit != alignment.shrink_to_fit,
                alignment.shrink_to_fit.then(|| "1".into()),
            ),
        ]
        .into_iter()
        .filter(|(_, changed, _)| *changed)
        .map(|(name, _, value)| (name, value))
        .collect::<Vec<_>>();
        let value = match template.child("alignment") {
            Some(source) => XmlTemplate::capture(&source.bytes)?.render_with_attributes(
                Vec::new(),
                |_| 0,
                &changes,
            )?,
            None => fragment(|writer| write_alignment(writer, &alignment))?,
        };
        replacements.push(("alignment", Some(value)));
        attributes.push(("applyAlignment", Some("1".into())));
    }
    template.render_with_attributes(replacements, |_| 0, &attributes)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::tests::{package, parse_workbook_with_package};

    #[test]
    fn existing_style_reuse_requires_matching_preserved_properties() {
        let styles = concat!(
            r#"<styleSheet><fonts count="2"><font><sz val="11"/></font><font><b/><sz val="11"/></font></fonts>"#,
            r#"<fills count="1"><fill><patternFill patternType="none"/></fill></fills><borders count="1"><border/></borders>"#,
            r#"<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>"#,
            r#"<xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyProtection="1"><protection locked="0"/></xf>"#,
            r#"<xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1"/></cellXfs></styleSheet>"#,
        );
        let sheet = r#"<sheetData><row r="1"><c r="A1" s="1"><v>1</v></c><c r="B1"><v>2</v></c></row></sheetData>"#;
        let mut parts = package(sheet, &[], false);
        parts.push(("xl/styles.xml".into(), styles.as_bytes().to_vec()));
        let parsed = parse_workbook_with_package(&parts).unwrap();
        let mut workbook = parsed.workbook;
        for col in 0..2 {
            workbook.sheets[0]
                .cell_mut(CellRef::new(0, col))
                .unwrap()
                .style = Some(2);
        }
        let plan = StylePlan::new(
            &workbook,
            &parsed.package,
            &[Some(0)],
            &[Some(SheetAxes::default())],
        )
        .unwrap();
        assert_eq!(plan.pairs[&(0, 2)], 2);
        assert_eq!(plan.pairs[&(1, 2)], 3);
        assert_eq!(plan.styles.fonts.len(), 2);
        assert_eq!(plan.styles.cell_xfs.len(), 4);
        let derived = std::str::from_utf8(&plan.overrides.xfs[&3]).unwrap();
        assert!(derived.contains(r#"fontId="1""#));
        assert!(derived.contains(r#"applyProtection="1""#));
        assert!(derived.contains(r#"<protection locked="0"/>"#));
    }
}
