//! S10 border, conditional-format, and table-grid serializers.

use crate::borders::BorderSpec;
use crate::formatting::ConditionalFormatStyle;
use crate::table::{Table, TableRow};

use super::xml_writer::{XmlWriter, int_attr};

const DEFAULT_TABLE_WIDTH_DXA: f64 = 9_360.0;

/// Trusted border element names used by the shared `CT_Border` serializer.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum BorderSide {
    Top,
    Bottom,
    Left,
    Right,
    InsideH,
    InsideV,
    Between,
    Bar,
    Start,
    End,
    TopLeftToBottomRight,
    TopRightToBottomLeft,
}

impl BorderSide {
    fn qname(self) -> &'static str {
        match self {
            Self::Top => "w:top",
            Self::Bottom => "w:bottom",
            Self::Left => "w:left",
            Self::Right => "w:right",
            Self::InsideH => "w:insideH",
            Self::InsideV => "w:insideV",
            Self::Between => "w:between",
            Self::Bar => "w:bar",
            Self::Start => "w:start",
            Self::End => "w:end",
            Self::TopLeftToBottomRight => "w:tl2br",
            Self::TopRightToBottomLeft => "w:tr2bl",
        }
    }
}

/// Serialize one border element with incumbent attribute ordering.
pub fn serialize_border(border: Option<&BorderSpec>, side: BorderSide) -> String {
    let Some(border) = border else {
        return String::new();
    };
    let mut writer = XmlWriter::with_capacity(96);
    write_border(&mut writer, border, side);
    writer.finish()
}

pub(super) fn write_border(writer: &mut XmlWriter, border: &BorderSpec, side: BorderSide) {
    writer.start_element(side.qname());
    if matches!(border.style.as_str(), "none" | "nil") {
        writer.attribute("w:val", &border.style).end_element();
        return;
    }

    writer.attribute("w:val", &border.style);
    if let Some(value) = border.size {
        writer.attribute("w:sz", &int_attr(Some(value)));
    }
    if let Some(value) = border.space {
        writer.attribute("w:space", &int_attr(Some(value)));
    }
    if let Some(color) = border.color.as_ref() {
        if color.auto == Some(true) {
            writer.attribute("w:color", "auto");
        } else if let Some(rgb) = color.rgb.as_deref().filter(|value| !value.is_empty()) {
            writer.attribute("w:color", rgb);
        }
        if let Some(value) = color
            .theme_color
            .as_deref()
            .filter(|value| !value.is_empty())
        {
            writer.attribute("w:themeColor", value);
        }
        if let Some(value) = color
            .theme_tint
            .as_deref()
            .filter(|value| !value.is_empty())
        {
            writer.attribute("w:themeTint", value);
        }
        if let Some(value) = color
            .theme_shade
            .as_deref()
            .filter(|value| !value.is_empty())
        {
            writer.attribute("w:themeShade", value);
        }
    }
    if border.shadow == Some(true) {
        writer.attribute("w:shadow", "true");
    }
    if border.frame == Some(true) {
        writer.attribute("w:frame", "true");
    }
    writer.end_element();
}

/// Serialize the 12-bit `w:cnfStyle` mask shared by table rows and cells.
pub fn serialize_conditional_format_style(style: Option<&ConditionalFormatStyle>) -> String {
    let Some(style) = style else {
        return String::new();
    };
    let flags = [
        style.first_row,
        style.last_row,
        style.first_column,
        style.last_column,
        style.odd_v_band,
        style.even_v_band,
        style.odd_h_band,
        style.even_h_band,
        style.nw_cell,
        style.ne_cell,
        style.sw_cell,
        style.se_cell,
    ];
    let value: String = flags
        .into_iter()
        .map(|flag| if flag == Some(true) { '1' } else { '0' })
        .collect();
    if value == "000000000000" {
        return String::new();
    }

    let mut writer = XmlWriter::with_capacity(42);
    writer
        .start_element("w:cnfStyle")
        .attribute("w:val", &value)
        .end_element();
    writer.finish()
}

/// Serialize the required `w:tblGrid`, inferring widths exactly like the
/// incumbent serializer when explicit column widths are absent.
pub fn serialize_table_grid(table: &Table) -> String {
    let column_widths = infer_table_grid_widths(table);
    if column_widths.is_empty() {
        return String::new();
    }

    let mut writer = XmlWriter::with_capacity(32 + column_widths.len() * 28);
    writer.start_element("w:tblGrid");
    for width in column_widths {
        writer
            .start_element("w:gridCol")
            .attribute("w:w", &int_attr(Some(width)))
            .end_element();
    }
    writer.end_element();
    writer.finish()
}

fn positive_integer(value: Option<f64>) -> Option<f64> {
    let value = value.filter(|value| value.is_finite() && *value > 0.0)?;
    Some((value + 0.5).floor().max(1.0))
}

fn positive_count(value: Option<f64>) -> usize {
    positive_integer(value).unwrap_or(1.0) as usize
}

fn grid_column_count(table: &Table) -> usize {
    if let Some(widths) = table
        .column_widths
        .as_ref()
        .filter(|widths| !widths.is_empty())
    {
        return widths.len();
    }
    table
        .rows
        .iter()
        .map(|row| {
            row.cells
                .iter()
                .map(|cell| {
                    positive_count(
                        cell.formatting
                            .as_ref()
                            .and_then(|formatting| formatting.grid_span),
                    )
                })
                .sum()
        })
        .max()
        .unwrap_or(0)
}

fn row_column_widths(row: &TableRow) -> Option<Vec<f64>> {
    let mut widths = Vec::new();
    for cell in &row.cells {
        let formatting = cell.formatting.as_ref()?;
        let width = positive_integer(formatting.width.as_ref().map(|width| width.value))?;
        let grid_span = positive_count(formatting.grid_span);
        widths.extend(distribute_column_widths(width, grid_span));
    }
    Some(widths)
}

fn infer_table_grid_widths(table: &Table) -> Vec<f64> {
    let explicit_widths: Vec<_> = table
        .column_widths
        .as_deref()
        .unwrap_or_default()
        .iter()
        .filter_map(|width| positive_integer(Some(*width)))
        .collect();
    if !explicit_widths.is_empty() {
        return explicit_widths;
    }

    let column_count = grid_column_count(table);
    if column_count == 0 {
        return Vec::new();
    }

    let mut best: Option<(&TableRow, Vec<f64>)> = None;
    for row in &table.rows {
        let Some(widths) = row_column_widths(row).filter(|widths| widths.len() == column_count)
        else {
            continue;
        };
        if best
            .as_ref()
            .is_none_or(|(best_row, _)| row.cells.len() > best_row.cells.len())
        {
            best = Some((row, widths));
        }
    }
    if let Some((_, widths)) = best {
        return widths;
    }

    let table_width = table
        .formatting
        .as_ref()
        .and_then(|formatting| formatting.width.as_ref())
        .filter(|width| width.kind == "dxa")
        .and_then(|width| positive_integer(Some(width.value)))
        .unwrap_or(DEFAULT_TABLE_WIDTH_DXA);
    distribute_column_widths(table_width, column_count)
}

fn distribute_column_widths(total_width: f64, column_count: usize) -> Vec<f64> {
    if column_count == 0 {
        return Vec::new();
    }
    let base_width = (total_width / column_count as f64).floor();
    let mut remainder = total_width - base_width * column_count as f64;
    (0..column_count)
        .map(|_| {
            let width = base_width + f64::from(remainder > 0.0);
            remainder = (remainder - 1.0).max(0.0);
            width.max(1.0)
        })
        .collect()
}

#[cfg(test)]
mod tests {
    use crate::formatting::{TableCellFormatting, TableMeasurement};
    use crate::scalars::ColorValue;
    use crate::table::{TableCell, TableRow};

    use super::*;

    fn cell(width: Option<f64>, grid_span: Option<f64>) -> TableCell {
        TableCell {
            node_type: "tableCell".to_owned(),
            formatting: (width.is_some() || grid_span.is_some()).then(|| TableCellFormatting {
                width: width.map(|value| TableMeasurement {
                    value,
                    kind: "dxa".to_owned(),
                }),
                grid_span,
                ..TableCellFormatting::default()
            }),
            property_changes: None,
            structural_change: None,
            content: Vec::new(),
        }
    }

    fn row(cells: Vec<TableCell>) -> TableRow {
        TableRow {
            node_type: "tableRow".to_owned(),
            formatting: None,
            property_changes: None,
            structural_change: None,
            cells,
        }
    }

    #[test]
    fn border_bytes_match_typescript_and_escape_attacker_values() {
        let border = BorderSpec {
            style: "single\" onmouseover=\"x<&".to_owned(),
            size: Some(7.6),
            space: Some(1.5),
            color: Some(ColorValue {
                rgb: Some("FF00\"/><evil".to_owned()),
                theme_color: Some("accent&1".to_owned()),
                theme_tint: Some("80'".to_owned()),
                ..ColorValue::default()
            }),
            shadow: Some(true),
            frame: Some(true),
        };
        assert_eq!(
            serialize_border(Some(&border), BorderSide::Top),
            "<w:top w:val=\"single&quot; onmouseover=&quot;x&lt;&amp;\" w:sz=\"8\" w:space=\"2\" w:color=\"FF00&quot;/&gt;&lt;evil\" w:themeColor=\"accent&amp;1\" w:themeTint=\"80&apos;\" w:shadow=\"true\" w:frame=\"true\"/>"
        );
        let nil = BorderSpec {
            style: "nil".to_owned(),
            ..border
        };
        assert_eq!(
            serialize_border(Some(&nil), BorderSide::Right),
            "<w:right w:val=\"nil\"/>"
        );
    }

    #[test]
    fn conditional_format_uses_incumbent_bit_order_and_omits_zero() {
        assert_eq!(
            serialize_conditional_format_style(Some(&ConditionalFormatStyle::default())),
            ""
        );
        let style = ConditionalFormatStyle {
            first_row: Some(true),
            odd_v_band: Some(true),
            even_h_band: Some(true),
            se_cell: Some(true),
            ..ConditionalFormatStyle::default()
        };
        assert_eq!(
            serialize_conditional_format_style(Some(&style)),
            "<w:cnfStyle w:val=\"100010010001\"/>"
        );
    }

    #[test]
    fn table_grid_prefers_explicit_positive_widths() {
        let mut table = Table::empty();
        table.column_widths = Some(vec![1008.000_000_000_000_1, -1.0, 2000.4]);
        assert_eq!(
            serialize_table_grid(&table),
            "<w:tblGrid><w:gridCol w:w=\"1008\"/><w:gridCol w:w=\"2000\"/></w:tblGrid>"
        );
    }

    #[test]
    fn table_grid_uses_most_granular_complete_row() {
        let mut table = Table::empty();
        table.rows = vec![
            row(vec![cell(Some(6000.0), Some(3.0))]),
            row(vec![
                cell(Some(1000.0), None),
                cell(Some(2000.0), None),
                cell(Some(3000.0), None),
            ]),
        ];
        assert_eq!(
            serialize_table_grid(&table),
            "<w:tblGrid><w:gridCol w:w=\"1000\"/><w:gridCol w:w=\"2000\"/><w:gridCol w:w=\"3000\"/></w:tblGrid>"
        );
    }

    #[test]
    fn table_grid_evenly_splits_default_width_with_leading_remainder() {
        let mut table = Table::empty();
        table.rows = vec![row(vec![cell(None, None); 7])];
        assert_eq!(
            serialize_table_grid(&table),
            "<w:tblGrid><w:gridCol w:w=\"1338\"/><w:gridCol w:w=\"1337\"/><w:gridCol w:w=\"1337\"/><w:gridCol w:w=\"1337\"/><w:gridCol w:w=\"1337\"/><w:gridCol w:w=\"1337\"/><w:gridCol w:w=\"1337\"/></w:tblGrid>"
        );
    }
}
