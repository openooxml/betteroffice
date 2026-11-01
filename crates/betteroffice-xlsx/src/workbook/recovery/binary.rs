use crate::authority::snapshot::{Codec, EditedChart, edited_charts};
use crate::snapshot::SnapshotError;
use xlsx_model::{CellFormat, Color, ErrorValue, Font, NumberFormat};
use xlsx_ops::{BorderPatch, KeptFacets, PreservedCellFormat, StyleProperty};

use super::*;

#[cfg(test)]
mod properties;
#[cfg(test)]
mod tests;

macro_rules! record_codec {
    ($ty:ident = $tag:literal { $($field:ident),+ $(,)? } $(; $default:ident = $value:expr)* $(=> $validate:ident)?) => {
        impl Codec for $ty {
            fn write(&self, w: &mut Writer) {
                let Self { $($field),+ $(, $default: _)* } = self;
                w.u8($tag);
                $($field.write(w);)+
            }

            fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
                if r.u8()? != $tag {
                    return Err(SnapshotError::new(concat!("invalid recovery ", stringify!($ty))));
                }
                let value = Self { $($field: Codec::read(r)?),+ $(, $default: $value)* };
                $($validate(&value)?;)?
                Ok(value)
            }
        }
    };
}

macro_rules! enum_codec {
    ($ty:ident { $($variant:ident = $tag:literal),+ $(,)? }) => {
        impl Codec for $ty {
            fn write(&self, w: &mut Writer) {
                w.u8(match self { $(Self::$variant => $tag),+ });
            }

            fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
                match r.u8()? {
                    $($tag => Ok(Self::$variant),)+
                    _ => Err(SnapshotError::new(concat!("invalid recovery ", stringify!($ty)))),
                }
            }
        }
    };
}

macro_rules! tagged_codec {
    ($ty:ident { $($variant:ident = $tag:literal $( { $($field:ident),+ $(,)? } )?),+ $(,)? } $(=> $validate:ident)?) => {
        impl Codec for $ty {
            fn write(&self, w: &mut Writer) {
                match self {
                    $(Self::$variant $( { $($field),+ } )? => {
                        w.u8($tag);
                        $($($field.write(w);)+)?
                    }),+
                }
            }

            fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
                let value = match r.u8()? {
                    $($tag => Self::$variant $( { $($field: Codec::read(r)?),+ } )?,)+
                    _ => return Err(SnapshotError::new(concat!("invalid recovery ", stringify!($ty)))),
                };
                $($validate(&value)?;)?
                Ok(value)
            }
        }
    };
}

record_codec!(RecoveryCheckpoint = 0 {
    engine, through_seq, header, calculation, proposals, proposal_id_counter, axes, history
}; collaborative = false);
record_codec!(RecoveryHeader = 0 {
    edited_since_open, recalculated_since_open, moved_references_since_open, active_sheet,
    rand_seed, version_nonce, committed_changes, last_calculation, calculation_context,
    client_id, next_sheet_id
});
record_codec!(RecoveryDiagnostics = 0 { cycle_cells, limited_cells, budget_exhausted });
record_codec!(CalculationOptions = 0 { now_serial } => validate_calculation_options);
record_codec!(CellAddress = 0 { sheet, cell });
record_codec!(RecoveryCalculation = 0 { cells, deleted_cells, arrays, styles, hidden_sizes });
record_codec!(RecoveryHiddenSizes = 0 { cols, rows } => validate_hidden_sizes);
record_codec!(RecoveryAxes = 0 { rows, cols });
record_codec!(RecoveryHistory = 0 { undo, sheet_order, preserved_undo, preserved_redo });
record_codec!(RecoveryPreservedHistory = 0 { before, after });
record_codec!(RecoveryPreserved = 0 { state, axes });
record_codec!(Transaction = 0 { ops, author, proposed });
record_codec!(CellState = 0 { value, formula, style });
record_codec!(Proposal = 0 { id, agent_id, note, edits, ghosts });
record_codec!(ProposedEdit = 0 {
    sheet, row, col, input, old_state, number_format, a1, old_text, new_text
});
record_codec!(ProposalGhost = 0 { sheet, row, col, old_text, new_text, alignment_value });
record_codec!(StylePatch = 0 {
    bold, italic, strikethrough, font_family, font_size, text_color, fill_color, border,
    horizontal_alignment, vertical_alignment, text_wrapping, clear
} => validate_style_patch);
record_codec!(BorderPatch = 0 { preset, style, color });
record_codec!(CapturedFormat = 0 {
    rows, columns, formats, kept, source_styles, source_formats
} => validate_captured_format);
record_codec!(KeptFacets = 0 {
    protection, xf_id, quote_prefix, pivot_button, text_rotation, indent, relative_indent,
    justify_last_line, reading_order
});
record_codec!(PreservedCellFormat = 0 { format, source_xf, derived_from });
record_codec!(CellFormat = 0 {
    font, fill, border, number_format, alignment, protection, xf_id, quote_prefix, pivot_button
} => validate_cell_format);

enum_codec!(StyleProperty {
    Bold = 0, Italic = 1, Strikethrough = 2, FontFamily = 3, FontSize = 4, TextColor = 5,
    FillColor = 6, Borders = 7, HorizontalAlignment = 8, VerticalAlignment = 9, TextWrapping = 10,
});
enum_codec!(HorizontalAlignment { Left = 0, Center = 1, Right = 2 });
enum_codec!(VerticalAlignment { Top = 0, Middle = 1, Bottom = 2 });
enum_codec!(TextWrapping { Overflow = 0, Wrap = 1, Clip = 2 });
enum_codec!(BorderPreset {
    All = 0, Inner = 1, Horizontal = 2, Vertical = 3, Outer = 4, Left = 5, Top = 6,
    Right = 7, Bottom = 8, None = 9,
});
enum_codec!(BorderLineStyle { Solid = 0, Dashed = 1, Dotted = 2, Double = 3 });
enum_codec!(ErrorValue {
    Div0 = 0, NA = 1, Name = 2, Null = 3, Num = 4, Ref = 5, Value = 6, Spill = 7, Calc = 8,
});
tagged_codec!(CellValue {
    Empty = 0, Number = 1 { value }, Text = 2 { value }, Bool = 3 { value }, Error = 4 { value },
} => validate_cell_value);
tagged_codec!(NumberFormat { Builtin = 0 { id }, Custom = 1 { pattern } });
tagged_codec!(NumberFormatMutation {
    Automatic = 0, PlainText = 1, Number = 2, Percent = 3, Scientific = 4, Currency = 5,
    Date = 6, Time = 7, Custom = 8 { pattern }, IncreaseDecimal = 9, DecreaseDecimal = 10,
});
tagged_codec!(Provenance { User = 0, Agent = 1 { id }, Remote = 2 { actor } });

fn finite(value: f64) -> SnapshotResult<()> {
    if value.is_finite() {
        Ok(())
    } else {
        Err(SnapshotError::new("non-finite recovery number"))
    }
}

fn finite_option(value: Option<f64>) -> SnapshotResult<()> {
    value.map_or(Ok(()), finite)
}

fn read_finite_option(r: &mut Reader<'_>) -> SnapshotResult<Option<f64>> {
    let value = Codec::read(r)?;
    finite_option(value)?;
    Ok(value)
}

fn validate_cell_value(value: &CellValue) -> SnapshotResult<()> {
    match value {
        CellValue::Number { value } => finite(*value),
        CellValue::Empty
        | CellValue::Text { .. }
        | CellValue::Bool { .. }
        | CellValue::Error { .. } => Ok(()),
    }
}

fn validate_captured_format(format: &CapturedFormat) -> SnapshotResult<()> {
    if format.kept.is_empty() || format.kept.len() == format.formats.len() {
        Ok(())
    } else {
        Err(SnapshotError::new(
            "recovery captured format facets do not match its formats",
        ))
    }
}

fn validate_calculation_options(options: &CalculationOptions) -> SnapshotResult<()> {
    finite_option(options.now_serial)
}

fn validate_hidden_sizes(sizes: &RecoveryHiddenSizes) -> SnapshotResult<()> {
    sizes
        .cols
        .values()
        .chain(sizes.rows.values())
        .try_for_each(|value| finite(*value))
}

fn validate_style_patch(patch: &StylePatch) -> SnapshotResult<()> {
    finite_option(patch.font_size)
}

fn validate_color(color: &Color) -> SnapshotResult<()> {
    match color {
        Color::Theme { tint, .. } => finite(*tint),
        Color::Rgb(_) | Color::Indexed(_) | Color::Auto => Ok(()),
    }
}

fn validate_font(font: &Font) -> SnapshotResult<()> {
    finite_option(font.size_pt)?;
    font.color.as_ref().map_or(Ok(()), validate_color)
}

fn validate_fill(fill: &Fill) -> SnapshotResult<()> {
    match fill {
        Fill::None => Ok(()),
        Fill::Solid(color) => validate_color(color),
        Fill::Pattern { fg, bg, .. } => fg.iter().chain(bg.iter()).try_for_each(validate_color),
    }
}

fn validate_border(border: &Border) -> SnapshotResult<()> {
    [&border.left, &border.right, &border.top, &border.bottom]
        .into_iter()
        .flatten()
        .filter_map(|edge| edge.color.as_ref())
        .try_for_each(validate_color)
}

fn validate_cell_format(format: &CellFormat) -> SnapshotResult<()> {
    validate_font(&format.font)?;
    validate_fill(&format.fill)?;
    validate_border(&format.border)
}

fn read_filter_range(r: &mut Reader<'_>) -> SnapshotResult<Option<CellRange>> {
    let range = Option::<CellRange>::read(r)?;
    range
        .map(|range| {
            validate_cell_ref(range.start)
                .map_err(|error| SnapshotError::new(error.to_string()))?;
            validate_cell_ref(range.end).map_err(|error| SnapshotError::new(error.to_string()))?;
            Ok(CellRange::new(range.start, range.end))
        })
        .transpose()
}

impl<A: Codec, B: Codec, C: Codec> Codec for (A, B, C) {
    fn write(&self, w: &mut Writer) {
        self.0.write(w);
        self.1.write(w);
        self.2.write(w);
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        Ok((Codec::read(r)?, Codec::read(r)?, Codec::read(r)?))
    }
}

fn write_slice<T: Codec>(values: &[T], w: &mut Writer) {
    w.var_usize(values.len());
    for value in values {
        value.write(w);
    }
}

impl Codec for UndoStack {
    fn write(&self, w: &mut Writer) {
        let (undo, redo) = self.recovery_stacks();
        w.u8(0);
        write_slice(undo, w);
        write_slice(redo, w);
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        if r.u8()? != 0 {
            return Err(SnapshotError::new("invalid recovery undo stack"));
        }
        Ok(Self::from_recovery_stacks(Codec::read(r)?, Codec::read(r)?))
    }
}

impl Codec for Stylesheet {
    fn write(&self, w: &mut Writer) {
        w.u8(0);
        self.fonts.write(w);
        self.fills.write(w);
        self.borders.write(w);
        self.cell_xfs.write(w);
        self.num_fmts.write(w);
        self.theme.write(w);
        self.indexed_colors.write(w);
        self.dxfs.write(w);
        let (sources, exact, released) = self.recovery_provenance();
        sources.write(w);
        exact.write(w);
        write_slice(released, w);
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        if r.u8()? != 0 {
            return Err(SnapshotError::new("invalid recovery stylesheet"));
        }
        let mut styles = Self::default();
        styles.fonts = Codec::read(r)?;
        styles.fills = Codec::read(r)?;
        styles.borders = Codec::read(r)?;
        styles.cell_xfs = Codec::read(r)?;
        styles.num_fmts = Codec::read(r)?;
        styles.theme = Codec::read(r)?;
        styles.indexed_colors = Codec::read(r)?;
        styles.dxfs = Codec::read(r)?;
        styles.restore_recovery_provenance(Codec::read(r)?, Codec::read(r)?, Codec::read(r)?);
        styles.fonts.iter().try_for_each(validate_font)?;
        styles.fills.iter().try_for_each(validate_fill)?;
        styles.borders.iter().try_for_each(validate_border)?;
        Ok(styles)
    }
}

impl Codec for Op {
    fn write(&self, w: &mut Writer) {
        match self {
            Self::SetCell { sheet, at, cell } => {
                w.u8(0);
                sheet.write(w);
                at.write(w);
                cell.write(w);
            }
            Self::InsertRows { sheet, at, count } => {
                w.u8(1);
                sheet.write(w);
                at.write(w);
                count.write(w);
            }
            Self::DeleteRows { sheet, at, count } => {
                w.u8(2);
                sheet.write(w);
                at.write(w);
                count.write(w);
            }
            Self::InsertCols { sheet, at, count } => {
                w.u8(3);
                sheet.write(w);
                at.write(w);
                count.write(w);
            }
            Self::DeleteCols { sheet, at, count } => {
                w.u8(4);
                sheet.write(w);
                at.write(w);
                count.write(w);
            }
            Self::SetColWidth { sheet, col, width } => {
                w.u8(5);
                sheet.write(w);
                col.write(w);
                width.write(w);
            }
            Self::SetColVisibility { sheet, col, hidden } => {
                w.u8(6);
                sheet.write(w);
                col.write(w);
                hidden.write(w);
            }
            Self::RestoreColDimension {
                sheet,
                col,
                width,
                hidden_width,
            } => {
                w.u8(7);
                sheet.write(w);
                col.write(w);
                width.write(w);
                hidden_width.write(w);
            }
            Self::SetRowHeight { sheet, row, height } => {
                w.u8(8);
                sheet.write(w);
                row.write(w);
                height.write(w);
            }
            Self::SetRowVisibility { sheet, row, hidden } => {
                w.u8(9);
                sheet.write(w);
                row.write(w);
                hidden.write(w);
            }
            Self::RestoreRowDimension {
                sheet,
                row,
                height,
                hidden_height,
            } => {
                w.u8(10);
                sheet.write(w);
                row.write(w);
                height.write(w);
                hidden_height.write(w);
            }
            Self::SetFreezePane { sheet, pane } => {
                w.u8(11);
                sheet.write(w);
                pane.write(w);
            }
            Self::SetHyperlinks { sheet, hyperlinks } => {
                w.u8(12);
                sheet.write(w);
                hyperlinks.write(w);
            }
            Self::SetHyperlink { sheet, hyperlink } => {
                w.u8(13);
                sheet.write(w);
                hyperlink.write(w);
            }
            Self::RemoveHyperlink { sheet, range } => {
                w.u8(14);
                sheet.write(w);
                range.write(w);
            }
            Self::SetAutoFilter { sheet, range } => {
                w.u8(15);
                sheet.write(w);
                range.write(w);
            }
            Self::RestoreAutoFilter { sheet, filter } => {
                w.u8(16);
                sheet.write(w);
                filter.write(w);
            }
            Self::AddConditionalFormat {
                sheet,
                range,
                kind,
                operator,
                formulas,
                format,
            } => {
                w.u8(17);
                sheet.write(w);
                range.write(w);
                kind.write(w);
                operator.write(w);
                formulas.write(w);
                format.write(w);
            }
            Self::RemoveConditionalFormat { sheet, id } => {
                w.u8(18);
                sheet.write(w);
                id.write(w);
            }
            Self::SetConditionalFormats { sheet, formats } => {
                w.u8(19);
                sheet.write(w);
                formats.write(w);
            }
            Self::RestoreColStyles { sheet, styles } => {
                w.u8(20);
                sheet.write(w);
                styles.write(w);
            }
            Self::RestoreRowStyles { sheet, styles } => {
                w.u8(21);
                sheet.write(w);
                styles.write(w);
            }
            Self::SetCharts { sheet, charts } => {
                w.u8(22);
                sheet.write(w);
                w.var_usize(charts.len());
                for chart in charts {
                    EditedChart::from(chart).write(w);
                }
            }
            Self::SetChartAnchor {
                sheet,
                frame,
                part,
                from,
                to,
            } => {
                w.u8(23);
                sheet.write(w);
                frame.write(w);
                part.write(w);
                from.write(w);
                to.write(w);
            }
            Self::MergeCells { sheet, range } => {
                w.u8(24);
                sheet.write(w);
                range.write(w);
            }
            Self::UnmergeCells { sheet, range } => {
                w.u8(25);
                sheet.write(w);
                range.write(w);
            }
            Self::PatchRangeStyle {
                sheet,
                range,
                patch,
            } => {
                w.u8(26);
                sheet.write(w);
                range.write(w);
                patch.write(w);
            }
            Self::SetRangeNumberFormat {
                sheet,
                range,
                format,
            } => {
                w.u8(27);
                sheet.write(w);
                range.write(w);
                format.write(w);
            }
            Self::ApplyRangeFormat {
                sheet,
                range,
                format,
            } => {
                w.u8(28);
                sheet.write(w);
                range.write(w);
                format.write(w);
            }
            Self::AddSheet { index, name } => {
                w.u8(29);
                index.write(w);
                name.write(w);
            }
            Self::RemoveSheet { index } => {
                w.u8(30);
                index.write(w);
            }
            Self::MoveSheet { from, to } => {
                w.u8(31);
                from.write(w);
                to.write(w);
            }
            Self::RenameSheet { sheet, name } => {
                w.u8(32);
                sheet.write(w);
                name.write(w);
            }
            Self::SetDefinedName {
                name,
                refers_to,
                scope,
                hidden,
            } => {
                w.u8(33);
                name.write(w);
                refers_to.write(w);
                scope.write(w);
                hidden.write(w);
            }
            Self::DeleteDefinedName { name, scope } => {
                w.u8(34);
                name.write(w);
                scope.write(w);
            }
            Self::SetDefinedNames { defined_names } => {
                w.u8(35);
                defined_names.write(w);
            }
            Self::RestoreSheet {
                sheet,
                name,
                formulas,
            } => {
                w.u8(36);
                sheet.write(w);
                name.write(w);
                formulas.write(w);
            }
        }
    }

    fn read(r: &mut Reader<'_>) -> SnapshotResult<Self> {
        Ok(match r.u8()? {
            0 => Self::SetCell {
                sheet: Codec::read(r)?,
                at: Codec::read(r)?,
                cell: Codec::read(r)?,
            },
            1 => Self::InsertRows {
                sheet: Codec::read(r)?,
                at: Codec::read(r)?,
                count: Codec::read(r)?,
            },
            2 => Self::DeleteRows {
                sheet: Codec::read(r)?,
                at: Codec::read(r)?,
                count: Codec::read(r)?,
            },
            3 => Self::InsertCols {
                sheet: Codec::read(r)?,
                at: Codec::read(r)?,
                count: Codec::read(r)?,
            },
            4 => Self::DeleteCols {
                sheet: Codec::read(r)?,
                at: Codec::read(r)?,
                count: Codec::read(r)?,
            },
            5 => Self::SetColWidth {
                sheet: Codec::read(r)?,
                col: Codec::read(r)?,
                width: read_finite_option(r)?,
            },
            6 => Self::SetColVisibility {
                sheet: Codec::read(r)?,
                col: Codec::read(r)?,
                hidden: Codec::read(r)?,
            },
            7 => Self::RestoreColDimension {
                sheet: Codec::read(r)?,
                col: Codec::read(r)?,
                width: read_finite_option(r)?,
                hidden_width: read_finite_option(r)?,
            },
            8 => Self::SetRowHeight {
                sheet: Codec::read(r)?,
                row: Codec::read(r)?,
                height: read_finite_option(r)?,
            },
            9 => Self::SetRowVisibility {
                sheet: Codec::read(r)?,
                row: Codec::read(r)?,
                hidden: Codec::read(r)?,
            },
            10 => Self::RestoreRowDimension {
                sheet: Codec::read(r)?,
                row: Codec::read(r)?,
                height: read_finite_option(r)?,
                hidden_height: read_finite_option(r)?,
            },
            11 => Self::SetFreezePane {
                sheet: Codec::read(r)?,
                pane: Codec::read(r)?,
            },
            12 => Self::SetHyperlinks {
                sheet: Codec::read(r)?,
                hyperlinks: Codec::read(r)?,
            },
            13 => Self::SetHyperlink {
                sheet: Codec::read(r)?,
                hyperlink: Codec::read(r)?,
            },
            14 => Self::RemoveHyperlink {
                sheet: Codec::read(r)?,
                range: Codec::read(r)?,
            },
            15 => Self::SetAutoFilter {
                sheet: Codec::read(r)?,
                range: read_filter_range(r)?,
            },
            16 => Self::RestoreAutoFilter {
                sheet: Codec::read(r)?,
                filter: Codec::read(r)?,
            },
            17 => Self::AddConditionalFormat {
                sheet: Codec::read(r)?,
                range: Codec::read(r)?,
                kind: Codec::read(r)?,
                operator: Codec::read(r)?,
                formulas: Codec::read(r)?,
                format: Codec::read(r)?,
            },
            18 => Self::RemoveConditionalFormat {
                sheet: Codec::read(r)?,
                id: Codec::read(r)?,
            },
            19 => Self::SetConditionalFormats {
                sheet: Codec::read(r)?,
                formats: Codec::read(r)?,
            },
            20 => Self::RestoreColStyles {
                sheet: Codec::read(r)?,
                styles: Codec::read(r)?,
            },
            21 => Self::RestoreRowStyles {
                sheet: Codec::read(r)?,
                styles: Codec::read(r)?,
            },
            22 => Self::SetCharts {
                sheet: Codec::read(r)?,
                charts: edited_charts(Vec::<EditedChart>::read(r)?)?,
            },
            23 => Self::SetChartAnchor {
                sheet: Codec::read(r)?,
                frame: Codec::read(r)?,
                part: Codec::read(r)?,
                from: Codec::read(r)?,
                to: Codec::read(r)?,
            },
            24 => Self::MergeCells {
                sheet: Codec::read(r)?,
                range: Codec::read(r)?,
            },
            25 => Self::UnmergeCells {
                sheet: Codec::read(r)?,
                range: Codec::read(r)?,
            },
            26 => Self::PatchRangeStyle {
                sheet: Codec::read(r)?,
                range: Codec::read(r)?,
                patch: Codec::read(r)?,
            },
            27 => Self::SetRangeNumberFormat {
                sheet: Codec::read(r)?,
                range: Codec::read(r)?,
                format: Codec::read(r)?,
            },
            28 => Self::ApplyRangeFormat {
                sheet: Codec::read(r)?,
                range: Codec::read(r)?,
                format: Codec::read(r)?,
            },
            29 => Self::AddSheet {
                index: Codec::read(r)?,
                name: Codec::read(r)?,
            },
            30 => Self::RemoveSheet {
                index: Codec::read(r)?,
            },
            31 => Self::MoveSheet {
                from: Codec::read(r)?,
                to: Codec::read(r)?,
            },
            32 => Self::RenameSheet {
                sheet: Codec::read(r)?,
                name: Codec::read(r)?,
            },
            33 => Self::SetDefinedName {
                name: Codec::read(r)?,
                refers_to: Codec::read(r)?,
                scope: Codec::read(r)?,
                hidden: Codec::read(r)?,
            },
            34 => Self::DeleteDefinedName {
                name: Codec::read(r)?,
                scope: Codec::read(r)?,
            },
            35 => Self::SetDefinedNames {
                defined_names: Codec::read(r)?,
            },
            36 => Self::RestoreSheet {
                sheet: Codec::read(r)?,
                name: Codec::read(r)?,
                formulas: Codec::read(r)?,
            },
            _ => return Err(SnapshotError::new("invalid recovery Op")),
        })
    }
}
