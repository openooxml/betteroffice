use xlsx_model::{ChartEdit, ChartRef, ChartSeriesName, SheetChart};
pub use xlsx_model::{ChartMetadata, ChartSeriesInfo};

use super::*;

const NS_TEXT: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";

fn child<'a>(element: &'a Element, local: &str) -> Option<&'a Element> {
    element
        .child_elements()
        .find(|item| item.is(NS_CHART, local))
}

fn descendant<'a>(element: &'a Element, namespace: &str, local: &str) -> Option<&'a Element> {
    if element.is(namespace, local) {
        return Some(element);
    }
    element
        .child_elements()
        .find_map(|item| descendant(item, namespace, local))
}

fn chart_root(root: &Element) -> Result<&Element, ParseError> {
    if !root.is(NS_CHART, "chartSpace") || !names_are_resolvable(root) {
        return Err(ParseError::UnsupportedEdit(
            "chart must use the classic DrawingML chart vocabulary".into(),
        ));
    }
    child(root, "chart")
        .ok_or_else(|| ParseError::UnsupportedEdit("chart part has no chart element".into()))
}

fn plots(chart: &Element) -> Vec<&Element> {
    child(chart, "plotArea")
        .map(|area| {
            area.child_elements()
                .filter(|item| {
                    item.namespace() == Some(NS_CHART) && item.local_name().ends_with("Chart")
                })
                .collect()
        })
        .unwrap_or_default()
}

fn series(chart: &Element) -> Vec<(&Element, &Element)> {
    plots(chart)
        .into_iter()
        .flat_map(|plot| {
            plot.child_elements()
                .filter(|item| item.is(NS_CHART, "ser"))
                .map(move |item| (item, plot))
        })
        .collect()
}

fn formula(element: Option<&Element>) -> Option<String> {
    element
        .and_then(|item| descendant(item, NS_CHART, "f"))
        .map(|item| item.text_content().trim().to_owned())
}

fn text_runs(element: &Element, out: &mut String, count: &mut usize) {
    if element.is(NS_TEXT, "r") || element.is(NS_TEXT, "fld") {
        *count += 1;
    }
    if element.is(NS_TEXT, "t") {
        out.push_str(&element.text_content());
    } else if element.is(NS_TEXT, "br") {
        out.push('\n');
    } else {
        let mut paragraph = false;
        for item in element.child_elements() {
            if item.is(NS_TEXT, "p") {
                if paragraph {
                    out.push('\n');
                }
                paragraph = true;
            }
            text_runs(item, out, count);
        }
    }
}

fn referenced_text(
    element: &Element,
    workbook: &xlsx_model::Workbook,
    owner: &str,
) -> Option<String> {
    if let Some(reference) = formula(Some(element))
        && let Ok(cells) = reference_inputs(
            &reference,
            workbook,
            owner,
            1,
            ReferenceShape::OneDimensional,
        )
        && let Some((value, format)) = cells.first()
    {
        return cache_point(
            value,
            false,
            CacheFidelity::Displayed,
            format,
            workbook.date_system,
        );
    }
    descendant(element, NS_CHART, "v").map(Element::text_content)
}

pub fn chart_metadata(
    part: &[u8],
    workbook: &xlsx_model::Workbook,
    owner: &str,
) -> Result<ChartMetadata, ParseError> {
    let root = parse_tree(part)?;
    let chart = chart_root(&root)?;
    let mut rich_title = false;
    let title = child(chart, "title").map(|title| {
        if let Some(rich) = descendant(title, NS_CHART, "rich") {
            let mut text = String::new();
            let mut runs = 0;
            text_runs(rich, &mut text, &mut runs);
            rich_title = runs > 1;
            text
        } else {
            referenced_text(title, workbook, owner).unwrap_or_default()
        }
    });
    Ok(ChartMetadata {
        chart_type: plots(chart)
            .iter()
            .map(|item| item.local_name())
            .collect::<Vec<_>>()
            .join(","),
        title,
        rich_title,
        series: series(chart)
            .iter()
            .enumerate()
            .map(|(index, (item, _))| {
                let tx = child(item, "tx");
                ChartSeriesInfo {
                    index,
                    name: tx.and_then(|tx| referenced_text(tx, workbook, owner)),
                    name_ref: formula(tx),
                    category_ref: formula(child(item, "cat").or_else(|| child(item, "xVal"))),
                    value_ref: formula(child(item, "val").or_else(|| child(item, "yVal"))),
                }
            })
            .collect(),
    })
}

pub fn chart_bytes(
    part: &[u8],
    chart: &SheetChart,
    workbook: &xlsx_model::Workbook,
    owner: &str,
) -> Result<Vec<u8>, ParseError> {
    let base = match &chart.raw_xml {
        Some(xml) => Part::decode(part)?.encode_text(xml),
        None => part.to_vec(),
    };
    patch_chart_refs(
        &base,
        &chart.part,
        &chart.refs,
        workbook,
        owner,
        (workbook, owner),
    )
}

fn prefix(element: &Element) -> String {
    element
        .name
        .rsplit_once(':')
        .map(|(prefix, _)| format!("{prefix}:"))
        .unwrap_or_default()
}

fn replace_child(
    source: &Part,
    parent: &Element,
    old: Option<&Element>,
    markup: String,
) -> Result<Vec<u8>, ParseError> {
    let range = match old {
        Some(item) => source.element_range(item)?,
        None => parent.content.start..parent.content.start,
    };
    if old.is_none() && parent.splice_target().is_none() {
        return source.splice(&[(
            source.element_range(parent)?,
            Replacement::Markup(source.element_with_content(parent, &markup)?),
        )]);
    }
    source.splice(&[(range, Replacement::Markup(markup))])
}

fn plain_run(source: &Part, rich: &Element, text: &str) -> Result<String, ParseError> {
    let paragraph = rich.child_elements().find(|item| item.is(NS_TEXT, "p"));
    let run = paragraph.and_then(|p| p.child_elements().find(|item| item.is(NS_TEXT, "r")));
    let properties = run
        .and_then(|run| run.child_elements().find(|item| item.is(NS_TEXT, "rPr")))
        .map(|item| source.element_markup(item))
        .transpose()?
        .unwrap_or_default();
    let text = format!(
        "<a:t xmlns:a=\"{NS_TEXT}\" xml:space=\"preserve\">{}</a:t>",
        escape_text(text)?
    );
    let run = if let Some(run) = run.filter(|item| item.splice_target().is_some()) {
        source.element_with_content(run, &format!("{properties}{text}"))?
    } else {
        format!("<a:r xmlns:a=\"{NS_TEXT}\">{text}</a:r>")
    };
    let paragraph = if let Some(p) = paragraph.filter(|item| item.splice_target().is_some()) {
        let properties = p
            .child_elements()
            .find(|item| item.is(NS_TEXT, "pPr"))
            .map(|item| source.element_markup(item))
            .transpose()?
            .unwrap_or_default();
        source.element_with_content(p, &format!("{properties}{run}"))?
    } else {
        format!("<a:p xmlns:a=\"{NS_TEXT}\">{run}</a:p>")
    };
    let mut markup = String::new();
    let mut replaced = false;
    for item in rich.child_elements() {
        if item.is(NS_TEXT, "p") {
            if !replaced {
                markup.push_str(&paragraph);
                replaced = true;
            }
        } else {
            markup.push_str(&source.element_markup(item)?);
        }
    }
    if !replaced {
        markup.push_str(&paragraph);
    }
    Ok(markup)
}

fn set_title(part: &[u8], title: &Option<String>) -> Result<Vec<u8>, ParseError> {
    let source = Part::decode(part)?;
    let root = source.tree()?;
    let chart = chart_root(&root)?;
    let old_title = child(chart, "title");
    let bytes = if let Some(text) = title {
        validate_text(text, "title")?;
        let p = old_title.map(prefix).unwrap_or_else(|| prefix(chart));
        if let Some(rich) = old_title
            .and_then(|title| child(title, "tx"))
            .and_then(|tx| child(tx, "rich"))
        {
            let markup = plain_run(&source, rich, text)?;
            if let Some(range) = rich.splice_target() {
                source.splice(&[(range, Replacement::Markup(markup))])?
            } else {
                replace_child(
                    &source,
                    old_title.unwrap(),
                    Some(child(old_title.unwrap(), "tx").unwrap()),
                    format!(
                        "<{p}tx><{p}rich><a:bodyPr xmlns:a=\"{NS_TEXT}\"/><a:lstStyle xmlns:a=\"{NS_TEXT}\"/>{markup}</{p}rich></{p}tx>"
                    ),
                )?
            }
        } else {
            let rich = format!(
                "<{p}tx><{p}rich><a:bodyPr xmlns:a=\"{NS_TEXT}\"/><a:lstStyle xmlns:a=\"{NS_TEXT}\"/><a:p xmlns:a=\"{NS_TEXT}\"><a:r><a:t xml:space=\"preserve\">{}</a:t></a:r></a:p></{p}rich></{p}tx>",
                escape_text(text)?
            );
            if let Some(old) = old_title {
                replace_child(&source, old, child(old, "tx"), rich)?
            } else {
                replace_child(&source, chart, None, format!("<{p}title>{rich}</{p}title>"))?
            }
        }
    } else if let Some(old) = old_title {
        source.splice(&[(
            source.element_range(old)?,
            Replacement::Markup(String::new()),
        )])?
    } else {
        part.to_vec()
    };
    let source = Part::decode(&bytes)?;
    let root = source.tree()?;
    let chart = chart_root(&root)?;
    let p = prefix(chart);
    let markup = format!(
        "<{p}autoTitleDeleted val=\"{}\"/>",
        u8::from(title.is_none())
    );
    if let Some(auto) = child(chart, "autoTitleDeleted") {
        let value = if title.is_none() { "1" } else { "0" };
        if auto.attribute(None, "val") == Some(value) {
            return Ok(bytes);
        }
        let markup = source.element_with_attribute(auto, "val", value)?;
        replace_child(&source, chart, Some(auto), markup)
    } else {
        let at = child(chart, "title")
            .map(|item| source.element_range(item).map(|range| range.end))
            .transpose()?
            .unwrap_or(chart.content.start);
        source.splice(&[(at..at, Replacement::Markup(markup))])
    }
}

fn validate_text(text: &str, field: &str) -> Result<(), ParseError> {
    if text.chars().count() > 32767 || escape_text(text).is_err() {
        return Err(ParseError::UnsupportedEdit(format!(
            "{field} must contain at most 32767 XML text characters"
        )));
    }
    Ok(())
}

fn clean_reference(reference: &str, field: &str) -> Result<String, ParseError> {
    let reference = reference
        .trim()
        .strip_prefix('=')
        .unwrap_or(reference.trim())
        .trim();
    if reference.is_empty() || reference.len() > 8192 || escape_text(reference).is_err() {
        return Err(ParseError::UnsupportedEdit(format!(
            "{field} must be a nonempty formula reference of at most 8192 characters, such as Sheet1!$A$1:$A$5 or a defined name"
        )));
    }
    Ok(reference.to_owned())
}

fn cache_markup(
    reference: &str,
    numeric: bool,
    p: &str,
    workbook: &xlsx_model::Workbook,
    owner: &str,
    existing: Option<&Element>,
) -> Result<String, ParseError> {
    let local = if numeric { "numCache" } else { "strCache" };
    let cache = CacheSite {
        local: local.into(),
        prefix: p.into(),
        span: None,
        format_code: existing
            .and_then(|item| child(item, "formatCode"))
            .map(Element::text_content),
        unmodelled_content: false,
        authored_points: 0,
    };
    match regenerated_cache(
        &cache,
        reference,
        workbook,
        owner,
        &mut CacheRebuild::persisted(),
    ) {
        Ok(content) => Ok(format!("<{p}{local}>{content}</{p}{local}>")),
        Err(ParseError::UnsupportedEdit(_)) => Ok(String::new()),
        Err(error) => Err(error),
    }
}

fn set_series(
    part: &[u8],
    index: usize,
    name: &Option<ChartSeriesName>,
    categories: &Option<String>,
    values: &Option<String>,
    workbook: &xlsx_model::Workbook,
    owner: &str,
) -> Result<Vec<u8>, ParseError> {
    if name.is_none() && categories.is_none() && values.is_none() {
        return Err(ParseError::UnsupportedEdit(
            "set_series needs at least one field: name, categories, values".into(),
        ));
    }
    let mut bytes = part.to_vec();
    for (field, reference) in [
        ("name", None),
        ("categories", categories.as_deref()),
        ("values", values.as_deref()),
    ] {
        if (field == "name" && name.is_none()) || (field != "name" && reference.is_none()) {
            continue;
        }
        let source = Part::decode(&bytes)?;
        let root = source.tree()?;
        let chart = chart_root(&root)?;
        let series = series(chart);
        let (item, plot) = series.get(index).ok_or_else(|| {
            ParseError::UnsupportedEdit(format!(
                "series {index} is out of range; valid series indices: {}",
                (0..series.len())
                    .map(|index| index.to_string())
                    .collect::<Vec<_>>()
                    .join(", ")
            ))
        })?;
        let xy = matches!(plot.local_name(), "scatterChart" | "bubbleChart");
        let (local, numeric) = if field == "name" {
            ("tx", false)
        } else if field == "values" {
            (
                if xy || child(item, "yVal").is_some() {
                    "yVal"
                } else {
                    "val"
                },
                true,
            )
        } else {
            (
                if xy || child(item, "xVal").is_some() {
                    "xVal"
                } else {
                    "cat"
                },
                xy || child(item, "xVal").is_some()
                    || child(item, "cat").is_some_and(|cat| {
                        child(cat, "numRef").is_some() || child(cat, "numLit").is_some()
                    }),
            )
        };
        let old = child(item, local);
        let p = old.map(prefix).unwrap_or_else(|| prefix(item));
        let content = if let Some(ChartSeriesName::Literal { text }) =
            name.as_ref().filter(|_| field == "name")
        {
            validate_text(text, "name.text")?;
            format!("<{p}v>{}</{p}v>", escape_text(text)?)
        } else {
            let reference = if field == "name" {
                match name.as_ref().unwrap() {
                    ChartSeriesName::Reference { reference } => reference,
                    ChartSeriesName::Literal { .. } => unreachable!(),
                }
            } else {
                reference.unwrap()
            };
            let reference =
                clean_reference(reference, if field == "name" { "name.ref" } else { field })?;
            let reference_tag = if numeric { "numRef" } else { "strRef" };
            let container = old.and_then(|old| child(old, reference_tag));
            let rp = container.map(prefix).unwrap_or_else(|| p.clone());
            let existing = container
                .and_then(|item| child(item, if numeric { "numCache" } else { "strCache" }));
            let cache = cache_markup(&reference, numeric, &rp, workbook, owner, existing)?;
            let mut content = format!("<{rp}f>{}</{rp}f>{cache}", escape_text(&reference)?);
            if let Some(container) = container {
                for item in container.child_elements().filter(|item| {
                    !(item.namespace() == Some(NS_CHART)
                        && matches!(
                            item.local_name(),
                            "f" | "numCache" | "strCache" | "multiLvlStrCache"
                        ))
                }) {
                    content.push_str(&source.element_markup(item)?);
                }
            }
            if let Some(container) = container {
                source.element_with_content(container, &content)?
            } else {
                format!("<{p}{reference_tag}>{content}</{p}{reference_tag}>")
            }
        };
        let content = if field != "name" {
            let mut extras = String::new();
            if let Some(old) = old {
                for item in old.child_elements().filter(|item| {
                    !(item.namespace() == Some(NS_CHART)
                        && matches!(
                            item.local_name(),
                            "strRef" | "numRef" | "strLit" | "numLit" | "multiLvlStrRef"
                        ))
                }) {
                    extras.push_str(&source.element_markup(item)?);
                }
            }
            format!("{content}{extras}")
        } else {
            content
        };
        let markup = if let Some(old) = old {
            source.element_with_content(old, &content)?
        } else {
            format!("<{p}{local}>{content}</{p}{local}>")
        };
        if let Some(old) = old {
            bytes = source.splice(&[(source.element_range(old)?, Replacement::Markup(markup))])?;
        } else {
            let before = item.child_elements().find(|item| match field {
                "name" => !matches!(item.local_name(), "idx" | "order"),
                "categories" => matches!(
                    item.local_name(),
                    "val" | "yVal" | "bubbleSize" | "smooth" | "extLst"
                ),
                _ => matches!(item.local_name(), "bubbleSize" | "smooth" | "extLst"),
            });
            let at = before
                .map(|item| source.element_range(item).map(|range| range.start))
                .transpose()?
                .unwrap_or(item.content.end);
            bytes = source.splice(&[(at..at, Replacement::Markup(markup))])?;
        }
    }
    Ok(bytes)
}

pub fn chart_references(part: &[u8]) -> Result<Vec<ChartRef>, ParseError> {
    chart_refs(&parse_tree(part)?)
}

pub fn edit_chart(
    part: &[u8],
    edit: &ChartEdit,
    workbook: &xlsx_model::Workbook,
    owner: &str,
) -> Result<(String, Vec<ChartRef>), ParseError> {
    let root = parse_tree(part)?;
    chart_root(&root)?;
    if matches!(edit, ChartEdit::SetSeries { .. }) && unsupported_reference_form(&root, 0) {
        return Err(ParseError::UnsupportedEdit("chart edits require a classic chart without pivot, external data or filtered series references".into()));
    }
    let bytes = match edit {
        ChartEdit::SetChartTitle { title } => set_title(part, title)?,
        ChartEdit::SetSeries {
            series,
            name,
            categories,
            values,
        } => set_series(part, *series, name, categories, values, workbook, owner)?,
    };
    let part = Part::decode(&bytes)?;
    let refs = chart_refs(&part.tree()?)?;
    Ok((part.text().to_owned(), refs))
}
