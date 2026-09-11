//! Table styles from `a:tblStyleLst` and the per-cell cascade they drive.

use serde::{Deserialize, Serialize};

use crate::{ColorValue, ShapeFill, ShapeOutline};

/// An `a:tblStyleLst` part.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStyleList {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub default_style_id: Option<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub styles: Vec<TableStyle>,
}

impl TableStyleList {
    pub fn is_empty(&self) -> bool {
        self.default_style_id.is_none() && self.styles.is_empty()
    }

    /// The style `style_id` names, or `def` when a table names none. An id the
    /// package does not define resolves to nothing, not to `def`.
    pub fn style(&self, style_id: Option<&str>) -> Option<&TableStyle> {
        match style_id {
            Some(id) => self.by_id(id),
            None => self.by_id(self.default_style_id.as_deref()?),
        }
    }

    /// Resolves one cell against the named style, or to defaults when none resolves.
    pub fn resolve_cell(
        &self,
        style_id: Option<&str>,
        flags: TableStyleFlags,
        position: TableCellPosition,
    ) -> ResolvedCellStyle {
        self.style(style_id)
            .map(|style| style.resolve_cell(flags, position))
            .unwrap_or_default()
    }

    fn by_id(&self, style_id: &str) -> Option<&TableStyle> {
        self.styles.iter().find(|style| style.style_id == style_id)
    }
}

/// One `a:tblStyle`.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStyle {
    pub style_id: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub style_name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub whole_table: Option<TableStylePart>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band1_row: Option<TableStylePart>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub band2_row: Option<TableStylePart>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_row: Option<TableStylePart>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_row: Option<TableStylePart>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub first_column: Option<TableStylePart>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub last_column: Option<TableStylePart>,
}

impl TableStyle {
    /// Applies `wholeTbl`, the row band, `firstCol`/`lastCol` then `firstRow`/`lastRow`.
    pub fn resolve_cell(
        &self,
        flags: TableStyleFlags,
        position: TableCellPosition,
    ) -> ResolvedCellStyle {
        let mut resolved = ResolvedCellStyle::default();
        let last_row = position.row_count.saturating_sub(1);
        let last_column = position.column_count.saturating_sub(1);

        resolved.apply_part(
            self.whole_table.as_ref(),
            RegionEdges {
                left: position.column == 0,
                right: position.column >= last_column,
                top: position.row == 0,
                bottom: position.row >= last_row,
            },
        );
        resolved.apply_part(self.band_part(flags, position), RegionEdges::row(position));
        if flags.first_column && position.column == 0 {
            resolved.apply_part(self.first_column.as_ref(), RegionEdges::column(position));
        }
        if flags.last_column && position.column_count > 0 && position.column == last_column {
            resolved.apply_part(self.last_column.as_ref(), RegionEdges::column(position));
        }
        if flags.first_row && position.row == 0 {
            resolved.apply_part(self.first_row.as_ref(), RegionEdges::row(position));
        }
        if flags.last_row && position.row_count > 0 && position.row == last_row {
            resolved.apply_part(self.last_row.as_ref(), RegionEdges::row(position));
        }
        resolved
    }

    fn band_part(
        &self,
        flags: TableStyleFlags,
        position: TableCellPosition,
    ) -> Option<&TableStylePart> {
        if !flags.band_row {
            return None;
        }
        let first = usize::from(flags.first_row);
        let last = position
            .row_count
            .checked_sub(1 + usize::from(flags.last_row))?;
        if position.row < first || position.row > last {
            return None;
        }
        if (position.row - first).is_multiple_of(2) {
            self.band1_row.as_ref()
        } else {
            self.band2_row.as_ref()
        }
    }
}

/// An `a:wholeTbl`, band, row or column style part.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStylePart {
    #[serde(default)]
    pub text: TableTextStyle,
    #[serde(default)]
    pub cell: TableCellStyle,
}

/// An `a:tcTxStyle`.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableTextStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub color: Option<ColorValue>,
}

/// An `a:tcStyle`, or the fill and borders of a cell's own `a:tcPr`.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCellStyle {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<ShapeFill>,
    #[serde(default)]
    pub borders: TableCellBorders,
}

/// The `a:tcBdr` edges.
#[derive(Clone, Debug, Default, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableCellBorders {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<TableCellBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<TableCellBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<TableCellBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<TableCellBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_horizontal: Option<TableCellBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub inside_vertical: Option<TableCellBorder>,
}

/// One edge of an `a:tcBdr`.
#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "camelCase")]
pub enum TableCellBorder {
    /// An `a:noFill` line, which clears the edge a lower style part set.
    None,
    Line(Box<ShapeOutline>),
}

/// The `a:tblPr` flags that select which style parts apply.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TableStyleFlags {
    pub first_row: bool,
    pub last_row: bool,
    pub first_column: bool,
    pub last_column: bool,
    pub band_row: bool,
}

/// Where a cell sits in its table.
#[derive(Clone, Copy, Debug, Default, PartialEq, Eq)]
pub struct TableCellPosition {
    pub row: usize,
    pub column: usize,
    pub row_count: usize,
    pub column_count: usize,
}

/// A cell's formatting after the style cascade.
#[derive(Clone, Debug, Default, PartialEq)]
pub struct ResolvedCellStyle {
    pub fill: Option<ShapeFill>,
    pub left: Option<ShapeOutline>,
    pub right: Option<ShapeOutline>,
    pub top: Option<ShapeOutline>,
    pub bottom: Option<ShapeOutline>,
    pub bold: Option<bool>,
    pub italic: Option<bool>,
    pub color: Option<ColorValue>,
}

impl ResolvedCellStyle {
    /// Applies a cell's own `a:tcPr`, the top of the cascade.
    pub fn apply_cell_style(&mut self, style: &TableCellStyle) {
        self.apply_cell(style, RegionEdges::CELL);
    }

    fn apply_part(&mut self, part: Option<&TableStylePart>, edges: RegionEdges) {
        let Some(part) = part else {
            return;
        };
        self.bold = part.text.bold.or(self.bold);
        self.italic = part.text.italic.or(self.italic);
        self.color = part.text.color.clone().or(self.color.take());
        self.apply_cell(&part.cell, edges);
    }

    fn apply_cell(&mut self, cell: &TableCellStyle, edges: RegionEdges) {
        if let Some(fill) = &cell.fill {
            self.fill = Some(fill.clone());
        }
        let borders = &cell.borders;
        let horizontal = borders.inside_horizontal.as_ref();
        let vertical = borders.inside_vertical.as_ref();
        apply_edge(
            &mut self.left,
            edge(edges.left, borders.left.as_ref(), vertical),
        );
        apply_edge(
            &mut self.right,
            edge(edges.right, borders.right.as_ref(), vertical),
        );
        apply_edge(
            &mut self.top,
            edge(edges.top, borders.top.as_ref(), horizontal),
        );
        apply_edge(
            &mut self.bottom,
            edge(edges.bottom, borders.bottom.as_ref(), horizontal),
        );
    }
}

/// Which sides of a style part's region a cell lies on.
#[derive(Clone, Copy, Debug)]
struct RegionEdges {
    left: bool,
    right: bool,
    top: bool,
    bottom: bool,
}

impl RegionEdges {
    const CELL: Self = Self {
        left: true,
        right: true,
        top: true,
        bottom: true,
    };

    fn row(position: TableCellPosition) -> Self {
        Self {
            left: position.column == 0,
            right: position.column + 1 >= position.column_count,
            top: true,
            bottom: true,
        }
    }

    fn column(position: TableCellPosition) -> Self {
        Self {
            left: true,
            right: true,
            top: position.row == 0,
            bottom: position.row + 1 >= position.row_count,
        }
    }
}

fn edge<'a>(
    outer: bool,
    boundary: Option<&'a TableCellBorder>,
    inside: Option<&'a TableCellBorder>,
) -> Option<&'a TableCellBorder> {
    if outer { boundary } else { inside }
}

fn apply_edge(target: &mut Option<ShapeOutline>, border: Option<&TableCellBorder>) {
    match border {
        Some(TableCellBorder::None) => *target = None,
        Some(TableCellBorder::Line(outline)) => *target = Some(outline.as_ref().clone()),
        None => {}
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn rgb(value: &str) -> ColorValue {
        ColorValue {
            rgb: Some(value.to_owned()),
            ..Default::default()
        }
    }

    fn solid(value: &str) -> ShapeFill {
        ShapeFill {
            fill_type: "solid".to_owned(),
            color: Some(rgb(value)),
            gradient: None,
        }
    }

    fn line(width: f64, value: &str) -> TableCellBorder {
        TableCellBorder::Line(Box::new(ShapeOutline {
            width: Some(width),
            color: Some(rgb(value)),
            ..Default::default()
        }))
    }

    fn fixture_style() -> TableStyle {
        TableStyle {
            style_id: "{STYLE}".to_owned(),
            style_name: Some("Fixture".to_owned()),
            whole_table: Some(TableStylePart {
                text: TableTextStyle {
                    color: Some(rgb("000000")),
                    ..Default::default()
                },
                cell: TableCellStyle {
                    fill: Some(solid("FFFFFF")),
                    borders: TableCellBorders {
                        left: Some(line(12700.0, "7F7F7F")),
                        right: Some(line(12700.0, "7F7F7F")),
                        top: Some(line(12700.0, "7F7F7F")),
                        bottom: Some(line(12700.0, "7F7F7F")),
                        inside_horizontal: Some(line(6350.0, "BFBFBF")),
                        inside_vertical: Some(line(6350.0, "BFBFBF")),
                    },
                },
            }),
            band1_row: Some(TableStylePart {
                cell: TableCellStyle {
                    fill: Some(solid("F2F2F2")),
                    ..Default::default()
                },
                ..Default::default()
            }),
            first_row: Some(TableStylePart {
                text: TableTextStyle {
                    bold: Some(true),
                    color: Some(rgb("FFFFFF")),
                    italic: None,
                },
                cell: TableCellStyle {
                    fill: Some(solid("4472C4")),
                    borders: TableCellBorders {
                        bottom: Some(line(38100.0, "FFFFFF")),
                        ..Default::default()
                    },
                },
            }),
            ..Default::default()
        }
    }

    fn flags() -> TableStyleFlags {
        TableStyleFlags {
            first_row: true,
            band_row: true,
            ..Default::default()
        }
    }

    fn position(row: usize, column: usize) -> TableCellPosition {
        TableCellPosition {
            row,
            column,
            row_count: 4,
            column_count: 3,
        }
    }

    #[test]
    fn header_row_overrides_whole_table_fill_and_text() {
        let resolved = fixture_style().resolve_cell(flags(), position(0, 1));

        assert_eq!(resolved.fill, Some(solid("4472C4")));
        assert_eq!(resolved.bold, Some(true));
        assert_eq!(resolved.color, Some(rgb("FFFFFF")));
        assert_eq!(resolved.bottom, Some(line_outline(38100.0, "FFFFFF")));
        assert_eq!(resolved.top, Some(line_outline(12700.0, "7F7F7F")));
        assert_eq!(resolved.left, Some(line_outline(6350.0, "BFBFBF")));
    }

    #[test]
    fn banding_alternates_over_data_rows_and_skips_the_header() {
        let style = fixture_style();

        assert_eq!(
            style.resolve_cell(flags(), position(1, 0)).fill,
            Some(solid("F2F2F2"))
        );
        assert_eq!(
            style.resolve_cell(flags(), position(2, 0)).fill,
            Some(solid("FFFFFF"))
        );
        assert_eq!(
            style.resolve_cell(flags(), position(3, 0)).fill,
            Some(solid("F2F2F2"))
        );
    }

    #[test]
    fn banding_is_ignored_when_band_row_is_off() {
        let style = fixture_style();
        let flags = TableStyleFlags {
            first_row: true,
            ..Default::default()
        };

        assert_eq!(
            style.resolve_cell(flags, position(1, 0)).fill,
            Some(solid("FFFFFF"))
        );
    }

    #[test]
    fn interior_cells_take_the_inside_edges() {
        let resolved = fixture_style().resolve_cell(flags(), position(2, 1));

        assert_eq!(resolved.left, Some(line_outline(6350.0, "BFBFBF")));
        assert_eq!(resolved.right, Some(line_outline(6350.0, "BFBFBF")));
        assert_eq!(resolved.top, Some(line_outline(6350.0, "BFBFBF")));
        assert_eq!(resolved.bottom, Some(line_outline(6350.0, "BFBFBF")));
    }

    #[test]
    fn a_no_fill_edge_clears_the_inherited_one() {
        let mut style = fixture_style();
        style.first_row.as_mut().unwrap().cell.borders.top = Some(TableCellBorder::None);

        let resolved = style.resolve_cell(flags(), position(0, 0));

        assert_eq!(resolved.top, None);
        assert_eq!(resolved.left, Some(line_outline(12700.0, "7F7F7F")));
    }

    #[test]
    fn first_and_last_column_apply_only_when_enabled() {
        let mut style = fixture_style();
        style.first_column = Some(TableStylePart {
            cell: TableCellStyle {
                fill: Some(solid("00FF00")),
                ..Default::default()
            },
            ..Default::default()
        });
        let mut flags = flags();

        assert_eq!(
            style.resolve_cell(flags, position(2, 0)).fill,
            Some(solid("FFFFFF"))
        );
        flags.first_column = true;
        assert_eq!(
            style.resolve_cell(flags, position(2, 0)).fill,
            Some(solid("00FF00"))
        );
    }

    #[test]
    fn a_cells_own_properties_win_over_the_style() {
        let mut resolved = fixture_style().resolve_cell(flags(), position(0, 0));
        resolved.apply_cell_style(&TableCellStyle {
            fill: Some(solid("DDEBF7")),
            borders: TableCellBorders {
                bottom: Some(line(19050.0, "C00000")),
                ..Default::default()
            },
        });

        assert_eq!(resolved.fill, Some(solid("DDEBF7")));
        assert_eq!(resolved.bottom, Some(line_outline(19050.0, "C00000")));
        assert_eq!(resolved.bold, Some(true));
    }

    #[test]
    fn an_unknown_style_id_leaves_the_cell_unstyled_rather_than_taking_the_default() {
        let list = TableStyleList {
            default_style_id: Some("{STYLE}".to_owned()),
            styles: vec![fixture_style()],
        };

        assert_eq!(
            list.resolve_cell(Some("{MISSING}"), flags(), position(0, 0))
                .fill,
            None
        );
    }

    #[test]
    fn a_table_naming_no_style_takes_the_default() {
        let list = TableStyleList {
            default_style_id: Some("{STYLE}".to_owned()),
            styles: vec![fixture_style()],
        };

        assert_eq!(
            list.resolve_cell(None, flags(), position(0, 0)).fill,
            Some(solid("4472C4"))
        );
    }

    #[test]
    fn an_unresolvable_default_leaves_the_cell_unstyled() {
        let list = TableStyleList {
            default_style_id: Some("{ALSO-MISSING}".to_owned()),
            styles: Vec::new(),
        };

        assert_eq!(
            list.resolve_cell(Some("{MISSING}"), flags(), position(0, 0)),
            ResolvedCellStyle::default()
        );
    }

    fn line_outline(width: f64, value: &str) -> ShapeOutline {
        match line(width, value) {
            TableCellBorder::Line(outline) => *outline,
            TableCellBorder::None => unreachable!(),
        }
    }
}
