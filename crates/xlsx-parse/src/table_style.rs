//! `tableStyleInfo` plus the workbook's `tableStyles` or Excel's presets ->
//! the [`TableStyleElement`]s a table draws.

use quick_xml::events::{BytesStart, Event};
use xlsx_model::{
    Color, DifferentialFormat, Stylesheet, Table, TableStyleElement, TableStyleElementType,
};

use crate::read::is_truthy;
use crate::table_style_presets::{PRESETS, Preset};
use crate::xml::{attr, local_name, next_event, reader};
use crate::{MAX_STYLE_ENTRIES, ParseError};

/// A table's `tableStyleInfo`.
pub(crate) struct TableStyleInfo {
    name: String,
    first_column: bool,
    last_column: bool,
    row_stripes: bool,
    column_stripes: bool,
}

impl TableStyleInfo {
    pub(crate) fn parse(e: &BytesStart) -> Result<Option<Self>, ParseError> {
        let Some(name) = attr(e, b"name")? else {
            return Ok(None);
        };
        let flag = |key: &[u8]| -> Result<bool, ParseError> {
            Ok(attr(e, key)?.is_some_and(|value| is_truthy(&value)))
        };
        Ok(Some(Self {
            name,
            first_column: flag(b"showFirstColumn")?,
            last_column: flag(b"showLastColumn")?,
            row_stripes: flag(b"showRowStripes")?,
            column_stripes: flag(b"showColumnStripes")?,
        }))
    }

    fn shows(&self, kind: TableStyleElementType) -> bool {
        use TableStyleElementType::*;
        match kind {
            FirstRowStripe | SecondRowStripe => self.row_stripes,
            FirstColumnStripe | SecondColumnStripe => self.column_stripes,
            FirstColumn | FirstHeaderCell | FirstTotalCell => self.first_column,
            LastColumn | LastHeaderCell | LastTotalCell => self.last_column,
            WholeTable | HeaderRow | TotalRow => true,
        }
    }
}

/// A `tableStyle` in `styles.xml`: its name and `(type, dxfId, size)` elements.
type WorkbookTableStyle = (String, Vec<(TableStyleElementType, Option<u32>, u32)>);

/// Fill each table's style from its `tableStyleInfo`, preferring a style the
/// workbook defines with elements over a preset of the same name.
pub(crate) fn resolve_tables(
    tables: Vec<(Table, Option<TableStyleInfo>)>,
    styles_xml: Option<&[u8]>,
    styles: &Stylesheet,
) -> Result<Vec<Table>, ParseError> {
    let defined = match styles_xml {
        Some(xml) if tables.iter().any(|(_, info)| info.is_some()) => workbook_table_styles(xml)?,
        _ => Vec::new(),
    };
    Ok(tables
        .into_iter()
        .map(|(mut table, info)| {
            if let Some(info) = info {
                table.style = resolve(&info, &defined, styles);
            }
            table
        })
        .collect())
}

fn resolve(
    info: &TableStyleInfo,
    defined: &[WorkbookTableStyle],
    styles: &Stylesheet,
) -> Vec<TableStyleElement> {
    let mut elements = match defined
        .iter()
        .find(|(name, elements)| *name == info.name && !elements.is_empty())
    {
        Some((_, elements)) => elements
            .iter()
            .map(|&(kind, dxf, size)| {
                #[cfg(test)]
                tests::VISITS.with(|visits| visits.set(visits.get() + 1));
                TableStyleElement {
                    kind,
                    size,
                    format: dxf
                        .and_then(|dxf| styles.dxfs.get(dxf as usize))
                        .cloned()
                        .unwrap_or_default(),
                }
            })
            .collect(),
        None => preset(&info.name, styles),
    };
    elements.retain(|element| {
        info.shows(element.kind) && (element.format != Default::default() || spans(element))
    });
    if elements.iter().all(|element| element.format == Default::default()) {
        return Vec::new();
    }
    elements.sort_by_key(|element| element.kind);
    elements.dedup_by_key(|element| element.kind);
    elements
}

/// Whether an element that paints nothing still sets a stripe band wider than one.
fn spans(element: &TableStyleElement) -> bool {
    use TableStyleElementType::*;
    element.size > 1
        && matches!(
            element.kind,
            FirstRowStripe | SecondRowStripe | FirstColumnStripe | SecondColumnStripe
        )
}

fn preset(name: &str, styles: &Stylesheet) -> Vec<TableStyleElement> {
    let Some((_, presets)) = PRESETS.iter().find(|(preset, _)| *preset == name) else {
        return Vec::new();
    };
    let color = |color: Option<(u8, f64)>| {
        color.and_then(|(idx, tint)| styles.resolve_color(&Color::Theme { idx, tint }))
    };
    presets
        .iter()
        .map(|&Preset(kind, bold, font, fill)| TableStyleElement {
            kind,
            size: 1,
            format: DifferentialFormat {
                font_color: color(font),
                bold: bold.then_some(true),
                italic: None,
                fill: color(fill),
            },
        })
        .collect()
}

fn workbook_table_styles(data: &[u8]) -> Result<Vec<WorkbookTableStyle>, ParseError> {
    let mut reader = reader(data);
    let mut buf = Vec::new();
    let mut depth = 0;
    let mut styles: Vec<WorkbookTableStyle> = Vec::new();
    let mut open = false;
    loop {
        match next_event(&mut reader, &mut buf, &mut depth)? {
            Event::Start(e) if local_name(&e) == b"tableStyle" => {
                open = styles.len() < MAX_STYLE_ENTRIES
                    && attr(&e, b"table")?.is_none_or(|value| is_truthy(&value));
                if open && let Some(name) = attr(&e, b"name")? {
                    styles.push((name, Vec::new()));
                } else {
                    open = false;
                }
            }
            Event::Start(e) if open && local_name(&e) == b"tableStyleElement" => {
                let Some(kind) = attr(&e, b"type")?.and_then(|value| element_type(&value)) else {
                    continue;
                };
                let dxf = attr(&e, b"dxfId")?.and_then(|value| value.parse::<u32>().ok());
                let size = attr(&e, b"size")?
                    .and_then(|value| value.parse::<u32>().ok())
                    .unwrap_or(1)
                    .max(1);
                if let Some((_, elements)) = styles.last_mut()
                    && elements.len() < MAX_STYLE_ENTRIES
                {
                    elements.push((kind, dxf, size));
                }
            }
            Event::End(e) if e.local_name().as_ref() == b"tableStyle" => open = false,
            Event::Eof => break,
            _ => {}
        }
    }
    Ok(styles)
}

fn element_type(token: &str) -> Option<TableStyleElementType> {
    use TableStyleElementType::*;
    Some(match token {
        "wholeTable" => WholeTable,
        "firstColumnStripe" => FirstColumnStripe,
        "secondColumnStripe" => SecondColumnStripe,
        "firstRowStripe" => FirstRowStripe,
        "secondRowStripe" => SecondRowStripe,
        "lastColumn" => LastColumn,
        "firstColumn" => FirstColumn,
        "headerRow" => HeaderRow,
        "totalRow" => TotalRow,
        "firstHeaderCell" => FirstHeaderCell,
        "lastHeaderCell" => LastHeaderCell,
        "firstTotalCell" => FirstTotalCell,
        "lastTotalCell" => LastTotalCell,
        _ => return None,
    })
}

#[cfg(test)]
mod tests {
    use std::cell::Cell;

    use xlsx_model::{CellRange, CellRef, SheetId};

    use super::*;

    thread_local! {
        pub(super) static VISITS: Cell<usize> = const { Cell::new(0) };
    }

    const STYLES: &[u8] = br#"<styleSheet><tableStyles count="1"><tableStyle name="Budget" pivot="0" count="2"><tableStyleElement type="headerRow" dxfId="0"/><tableStyleElement type="firstRowStripe" dxfId="1"/></tableStyle></tableStyles></styleSheet>"#;

    fn table(name: &str) -> (Table, Option<TableStyleInfo>) {
        let table = Table {
            name: name.into(),
            sheet: SheetId(0),
            range: CellRange::new(CellRef::new(0, 0), CellRef::new(4, 2)),
            header_rows: 1,
            totals_rows: 0,
            columns: Vec::new(),
            style: Vec::new(),
        };
        let info = TableStyleInfo {
            name: "Budget".into(),
            first_column: false,
            last_column: false,
            row_stripes: true,
            column_stripes: false,
        };
        (table, Some(info))
    }

    #[test]
    fn tables_sharing_a_style_visit_its_elements_once() {
        let mut styles = Stylesheet::default();
        styles.dxfs = vec![
            DifferentialFormat {
                bold: Some(true),
                ..Default::default()
            },
            DifferentialFormat {
                fill: Some("#fff78c".into()),
                ..Default::default()
            },
        ];
        let tables = (0..3).map(|n| table(&format!("T{n}"))).collect();
        VISITS.with(|visits| visits.set(0));
        let resolved = resolve_tables(tables, Some(STYLES), &styles).unwrap();
        assert!(resolved.iter().all(|table| table.style.len() == 2));
        assert_eq!(VISITS.with(Cell::get), 2);
    }
}
