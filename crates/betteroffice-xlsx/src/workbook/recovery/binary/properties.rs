use std::fmt::Debug;

use proptest::prelude::*;
use xlsx_model::{
    Alignment, AnchorCell, AnchorEditAs, AnchorExtent, AnchorPos, AutoFilter, Border, BorderEdge,
    BorderStyle, ChartRef, ChartRefKind, ColStyle, Color, ConditionalFormatRule,
    ConditionalFormatting, DefinedName, DifferentialFormat, Fill, Font, HAlign, PatternType,
    Protection, VAlign,
};

use super::*;

type S<T> = BoxedStrategy<T>;

macro_rules! record {
    ($ty:ident { $($field:ident: $strategy:expr),+ $(,)? }) => {
        ($($strategy,)+).prop_map(|($($field,)+)| $ty { $($field),+ }).boxed()
    };
}

macro_rules! operation {
    ($variant:ident { $($field:ident: $strategy:expr),+ $(,)? }) => {
        ($($strategy,)+).prop_map(|($($field,)+)| Op::$variant { $($field),+ }).boxed()
    };
}

fn text() -> S<String> {
    prop::collection::vec(any::<char>(), 0..24)
        .prop_map(|chars| chars.into_iter().collect())
        .boxed()
}

fn finite() -> S<f64> {
    prop_oneof![
        (-1.0e6..1.0e6),
        Just(-0.0),
        Just(f64::MAX),
        Just(f64::MIN_POSITIVE)
    ]
    .boxed()
}

fn opt<T: Strategy + 'static>(strategy: T) -> S<Option<T::Value>> {
    prop::option::of(strategy).boxed()
}

fn select<T: Clone + Debug + 'static>(values: Vec<T>) -> S<T> {
    prop::sample::select(values).boxed()
}

fn sheet() -> S<SheetId> {
    any::<u32>().prop_map(SheetId).boxed()
}

fn at() -> S<CellRef> {
    record!(CellRef {
        row: any::<u32>(),
        col: any::<u32>(),
        abs_row: any::<bool>(),
        abs_col: any::<bool>()
    })
}

fn range() -> S<CellRange> {
    record!(CellRange {
        start: at(),
        end: at()
    })
}

fn value() -> S<CellValue> {
    prop_oneof![
        Just(CellValue::Empty),
        finite().prop_map(|value| CellValue::Number { value }),
        text().prop_map(|value| CellValue::Text { value }),
        any::<bool>().prop_map(|value| CellValue::Bool { value }),
        select(vec![
            ErrorValue::Div0,
            ErrorValue::NA,
            ErrorValue::Name,
            ErrorValue::Null,
            ErrorValue::Num,
            ErrorValue::Ref,
            ErrorValue::Value,
            ErrorValue::Spill,
            ErrorValue::Calc
        ])
        .prop_map(|value| CellValue::Error { value }),
    ]
    .boxed()
}

fn cell() -> S<CellState> {
    record!(CellState {
        value: value(),
        formula: opt(text()),
        style: opt(any::<u32>())
    })
}

fn color() -> S<Color> {
    prop_oneof![
        text().prop_map(Color::Rgb),
        (any::<u8>(), finite()).prop_map(|(idx, tint)| Color::Theme { idx, tint }),
        any::<u8>().prop_map(Color::Indexed),
        Just(Color::Auto),
    ]
    .boxed()
}

fn font() -> S<Font> {
    record!(Font {
        name: opt(text()),
        size_pt: opt(finite()),
        bold: any::<bool>(),
        italic: any::<bool>(),
        underline: any::<bool>(),
        strike: any::<bool>(),
        color: opt(color()),
    })
}

fn fill() -> S<Fill> {
    prop_oneof![
        Just(Fill::None),
        color().prop_map(Fill::Solid),
        (
            select(PatternType::ALL.to_vec()),
            opt(color()),
            opt(color())
        )
            .prop_map(|(pattern, fg, bg)| Fill::Pattern { pattern, fg, bg }),
    ]
    .boxed()
}

fn edge() -> S<BorderEdge> {
    record!(BorderEdge {
        style: select(vec![
            BorderStyle::Thin,
            BorderStyle::Medium,
            BorderStyle::Thick,
            BorderStyle::Dashed,
            BorderStyle::Dotted,
            BorderStyle::Double,
            BorderStyle::Hair,
            BorderStyle::MediumDashed,
            BorderStyle::DashDot,
            BorderStyle::MediumDashDot,
            BorderStyle::DashDotDot,
            BorderStyle::MediumDashDotDot,
            BorderStyle::SlantDashDot
        ]),
        color: opt(color()),
    })
}

fn border() -> S<Border> {
    record!(Border {
        left: opt(edge()),
        right: opt(edge()),
        top: opt(edge()),
        bottom: opt(edge())
    })
}

fn alignment() -> S<Alignment> {
    record!(Alignment {
        h: opt(select(vec![
            HAlign::General,
            HAlign::Left,
            HAlign::Center,
            HAlign::Right,
            HAlign::Fill,
            HAlign::Justify,
            HAlign::CenterContinuous,
            HAlign::Distributed
        ])),
        v: opt(select(vec![
            VAlign::Top,
            VAlign::Center,
            VAlign::Bottom,
            VAlign::Justify,
            VAlign::Distributed
        ])),
        wrap_text: any::<bool>(),
        shrink_to_fit: any::<bool>(),
        text_rotation: opt(any::<u16>()),
        indent: opt(any::<u32>()),
        relative_indent: opt(any::<i32>()),
        justify_last_line: opt(any::<bool>()),
        reading_order: opt(any::<u32>()),
    })
}

fn protection() -> S<Protection> {
    record!(Protection {
        locked: any::<bool>(),
        hidden: any::<bool>()
    })
}

fn number_format() -> S<NumberFormat> {
    prop_oneof![
        any::<u16>().prop_map(|id| NumberFormat::Builtin { id }),
        text().prop_map(|pattern| NumberFormat::Custom { pattern }),
    ]
    .boxed()
}

fn cell_format() -> S<CellFormat> {
    record!(CellFormat {
        font: font(),
        fill: fill(),
        border: border(),
        number_format: number_format(),
        alignment: alignment(),
        protection: opt(protection()),
        xf_id: opt(any::<u32>()),
        quote_prefix: any::<bool>(),
        pivot_button: any::<bool>(),
    })
}

fn kept() -> S<KeptFacets> {
    record!(KeptFacets {
        protection: any::<bool>(),
        xf_id: any::<bool>(),
        quote_prefix: any::<bool>(),
        pivot_button: any::<bool>(),
        text_rotation: any::<bool>(),
        indent: any::<bool>(),
        relative_indent: any::<bool>(),
        justify_last_line: any::<bool>(),
        reading_order: any::<bool>(),
    })
}

fn captured_format() -> S<CapturedFormat> {
    record!(CapturedFormat {
        rows: 1u32..4,
        columns: 1u32..4,
        formats: prop::collection::vec(cell_format(), 0..4),
        kept: prop::collection::vec(kept(), 0..4),
        source_styles: prop::collection::vec(opt(any::<u32>()), 0..4),
        source_formats: prop::collection::vec(
            record!(PreservedCellFormat {
                format: cell_format(),
                source_xf: opt(any::<u32>()),
                derived_from: opt(any::<u32>()),
            }),
            0..4
        ),
    })
}

fn number_mutation() -> S<NumberFormatMutation> {
    prop_oneof![
        select(vec![
            NumberFormatMutation::Automatic,
            NumberFormatMutation::PlainText,
            NumberFormatMutation::Number,
            NumberFormatMutation::Percent,
            NumberFormatMutation::Scientific,
            NumberFormatMutation::Currency,
            NumberFormatMutation::Date,
            NumberFormatMutation::Time,
            NumberFormatMutation::IncreaseDecimal,
            NumberFormatMutation::DecreaseDecimal
        ]),
        text().prop_map(|pattern| NumberFormatMutation::Custom { pattern }),
    ]
    .boxed()
}

fn style_patch() -> S<StylePatch> {
    record!(StylePatch {
        bold: opt(any::<bool>()),
        italic: opt(any::<bool>()),
        strikethrough: opt(any::<bool>()),
        font_family: opt(text()),
        font_size: opt(finite()),
        text_color: opt(text()),
        fill_color: opt(text()),
        border: opt(record!(BorderPatch {
            preset: opt(select(vec![
                BorderPreset::All,
                BorderPreset::Inner,
                BorderPreset::Horizontal,
                BorderPreset::Vertical,
                BorderPreset::Outer,
                BorderPreset::Left,
                BorderPreset::Top,
                BorderPreset::Right,
                BorderPreset::Bottom,
                BorderPreset::None
            ])),
            style: opt(select(vec![
                BorderLineStyle::Solid,
                BorderLineStyle::Dashed,
                BorderLineStyle::Dotted,
                BorderLineStyle::Double
            ])),
            color: opt(text()),
        })),
        horizontal_alignment: opt(select(vec![
            HorizontalAlignment::Left,
            HorizontalAlignment::Center,
            HorizontalAlignment::Right
        ])),
        vertical_alignment: opt(select(vec![
            VerticalAlignment::Top,
            VerticalAlignment::Middle,
            VerticalAlignment::Bottom
        ])),
        text_wrapping: opt(select(vec![
            TextWrapping::Overflow,
            TextWrapping::Wrap,
            TextWrapping::Clip
        ])),
        clear: prop::collection::vec(
            select(vec![
                StyleProperty::Bold,
                StyleProperty::Italic,
                StyleProperty::Strikethrough,
                StyleProperty::FontFamily,
                StyleProperty::FontSize,
                StyleProperty::TextColor,
                StyleProperty::FillColor,
                StyleProperty::Borders,
                StyleProperty::HorizontalAlignment,
                StyleProperty::VerticalAlignment,
                StyleProperty::TextWrapping
            ]),
            0..12
        ),
    })
}

fn anchor_cell() -> S<AnchorCell> {
    record!(AnchorCell {
        col: any::<u32>(),
        col_off: any::<i64>(),
        row: any::<u32>(),
        row_off: any::<i64>()
    })
}

fn extent() -> S<AnchorExtent> {
    record!(AnchorExtent {
        cx: any::<i64>(),
        cy: any::<i64>()
    })
}

fn anchor() -> S<ChartAnchor> {
    prop_oneof![
        (
            anchor_cell(),
            anchor_cell(),
            select(vec![
                AnchorEditAs::TwoCell,
                AnchorEditAs::OneCell,
                AnchorEditAs::Absolute
            ])
        )
            .prop_map(|(from, to, edit_as)| ChartAnchor::TwoCell { from, to, edit_as }),
        (anchor_cell(), extent()).prop_map(|(from, extent)| ChartAnchor::OneCell { from, extent }),
        (
            record!(AnchorPos {
                x: any::<i64>(),
                y: any::<i64>()
            }),
            extent()
        )
            .prop_map(|(pos, extent)| ChartAnchor::Absolute { pos, extent }),
    ]
    .boxed()
}

fn chart() -> S<SheetChart> {
    record!(SheetChart {
        part: text(), drawing: text(), anchor_index: any::<usize>(), anchor: anchor(),
        refs: prop::collection::vec(record!(ChartRef {
            kind: select(vec![ChartRefKind::SeriesName, ChartRefKind::Categories, ChartRefKind::Values,
                ChartRefKind::BubbleSize, ChartRefKind::Title, ChartRefKind::DataLabels, ChartRefKind::Other]),
            formula: text(),
        }), 0..4),
        raw_xml: opt(text().prop_map(|text| format!("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart><!--{}--></c:chart></c:chartSpace>", text.replace('-', "x")))),
    })
}

fn hyperlink() -> S<Hyperlink> {
    record!(Hyperlink {
        range: range(),
        external_target: opt(text()),
        location: opt(text()),
        tooltip: opt(text()),
        display: opt(text())
    })
}

fn differential() -> S<DifferentialFormat> {
    record!(DifferentialFormat {
        font_color: opt(text()),
        bold: opt(any::<bool>()),
        italic: opt(any::<bool>()),
        fill: opt(text())
    })
}

fn conditional() -> S<ConditionalFormatting> {
    record!(ConditionalFormatting {
        sqref: text(),
        rules: prop::collection::vec(
            record!(ConditionalFormatRule {
                id: text(),
                rule_type: text(),
                operator: opt(text()),
                formulas: prop::collection::vec(text(), 0..4),
                priority: any::<u32>(),
                dxf_id: opt(any::<u32>()),
                format: differential(),
                raw_xml: opt(text()),
            }),
            0..4
        ),
    })
}

fn defined_name() -> S<DefinedName> {
    record!(DefinedName {
        name: text(),
        formula: text(),
        local_sheet: opt(sheet()),
        hidden: any::<bool>()
    })
}

fn op(tag: u8) -> S<Op> {
    match tag {
        0 => operation!(SetCell {
            sheet: sheet(),
            at: at(),
            cell: cell()
        }),
        1 => operation!(InsertRows {
            sheet: sheet(),
            at: any::<u32>(),
            count: any::<u32>()
        }),
        2 => operation!(DeleteRows {
            sheet: sheet(),
            at: any::<u32>(),
            count: any::<u32>()
        }),
        3 => operation!(InsertCols {
            sheet: sheet(),
            at: any::<u32>(),
            count: any::<u32>()
        }),
        4 => operation!(DeleteCols {
            sheet: sheet(),
            at: any::<u32>(),
            count: any::<u32>()
        }),
        5 => operation!(SetColWidth {
            sheet: sheet(),
            col: any::<u32>(),
            width: opt(finite())
        }),
        6 => operation!(SetColVisibility {
            sheet: sheet(),
            col: any::<u32>(),
            hidden: any::<bool>()
        }),
        7 => operation!(RestoreColDimension {
            sheet: sheet(),
            col: any::<u32>(),
            width: opt(finite()),
            hidden_width: opt(finite())
        }),
        8 => operation!(SetRowHeight {
            sheet: sheet(),
            row: any::<u32>(),
            height: opt(finite())
        }),
        9 => operation!(SetRowVisibility {
            sheet: sheet(),
            row: any::<u32>(),
            hidden: any::<bool>()
        }),
        10 => operation!(RestoreRowDimension {
            sheet: sheet(),
            row: any::<u32>(),
            height: opt(finite()),
            hidden_height: opt(finite())
        }),
        11 => operation!(SetFreezePane {
            sheet: sheet(),
            pane: opt(record!(FreezePane {
                rows: any::<u32>(),
                cols: any::<u32>(),
                top_left: at()
            }))
        }),
        12 => operation!(SetHyperlinks {
            sheet: sheet(),
            hyperlinks: prop::collection::vec(hyperlink(), 0..4)
        }),
        13 => operation!(SetHyperlink {
            sheet: sheet(),
            hyperlink: hyperlink()
        }),
        14 => operation!(RemoveHyperlink {
            sheet: sheet(),
            range: range()
        }),
        15 => operation!(SetAutoFilter {
            sheet: sheet(),
            range: opt(range())
        }),
        16 => operation!(RestoreAutoFilter {
            sheet: sheet(),
            filter: opt(record!(AutoFilter {
                range: range(),
                raw_xml: opt(text()),
                source_count: any::<u32>(),
                filter_columns: prop::collection::vec(any::<u32>(), 0..4),
                has_sort_state: any::<bool>(),
            }))
        }),
        17 => operation!(AddConditionalFormat {
            sheet: sheet(),
            range: range(),
            kind: text(),
            operator: opt(text()),
            formulas: prop::collection::vec(text(), 0..4),
            format: differential()
        }),
        18 => operation!(RemoveConditionalFormat {
            sheet: sheet(),
            id: text()
        }),
        19 => operation!(SetConditionalFormats {
            sheet: sheet(),
            formats: prop::collection::vec(conditional(), 0..4)
        }),
        20 => operation!(RestoreColStyles {
            sheet: sheet(),
            styles: prop::collection::vec(
                record!(ColStyle {
                    first: any::<u32>(),
                    last: any::<u32>(),
                    xf: any::<u32>()
                }),
                0..4
            )
        }),
        21 => operation!(RestoreRowStyles {
            sheet: sheet(),
            styles: prop::collection::btree_map(any::<u32>(), any::<u32>(), 0..4)
        }),
        22 => operation!(SetCharts {
            sheet: sheet(),
            charts: prop::collection::vec(chart(), 0..4)
        }),
        23 => operation!(SetChartAnchor {
            sheet: sheet(),
            frame: text(),
            part: text(),
            from: anchor(),
            to: anchor()
        }),
        24 => operation!(MergeCells {
            sheet: sheet(),
            range: range()
        }),
        25 => operation!(UnmergeCells {
            sheet: sheet(),
            range: range()
        }),
        26 => operation!(PatchRangeStyle {
            sheet: sheet(),
            range: range(),
            patch: style_patch()
        }),
        27 => operation!(SetRangeNumberFormat {
            sheet: sheet(),
            range: range(),
            format: number_mutation()
        }),
        28 => operation!(ApplyRangeFormat {
            sheet: sheet(),
            range: range(),
            format: captured_format()
        }),
        29 => operation!(AddSheet {
            index: any::<usize>(),
            name: text()
        }),
        30 => operation!(RemoveSheet {
            index: any::<usize>()
        }),
        31 => operation!(MoveSheet {
            from: any::<usize>(),
            to: any::<usize>()
        }),
        32 => operation!(RenameSheet {
            sheet: sheet(),
            name: text()
        }),
        33 => operation!(SetDefinedName {
            name: text(),
            refers_to: text(),
            scope: opt(sheet()),
            hidden: opt(any::<bool>())
        }),
        34 => operation!(DeleteDefinedName {
            name: text(),
            scope: opt(sheet())
        }),
        35 => operation!(SetDefinedNames {
            defined_names: prop::collection::vec(defined_name(), 0..4)
        }),
        36 => operation!(RestoreSheet {
            sheet: sheet(),
            name: text(),
            formulas: prop::collection::vec((sheet(), at(), cell()), 0..4)
        }),
        _ => unreachable!(),
    }
}

fn all_ops() -> S<Vec<Op>> {
    (0..37).fold(Just(Vec::new()).boxed(), |ops, tag| {
        (ops, op(tag))
            .prop_map(|(mut ops, op)| {
                ops.push(op);
                ops
            })
            .boxed()
    })
}

fn any_op() -> S<Op> {
    (0u8..37).prop_flat_map(op).boxed()
}

fn history() -> S<Vec<Vec<Op>>> {
    prop::collection::vec(prop::collection::vec(any_op(), 0..4), 0..4).boxed()
}

fn encoded(value: &impl Codec) -> Vec<u8> {
    let mut writer = Writer::new();
    value.write(&mut writer);
    writer.into_bytes()
}

fn preserved_history(count: usize) -> Vec<RecoveryPreservedHistory> {
    (0..count)
        .map(|_| RecoveryPreservedHistory {
            before: RecoveryPreserved {
                state: 0,
                axes: vec![None],
            },
            after: RecoveryPreserved {
                state: 0,
                axes: vec![None],
            },
        })
        .collect()
}

proptest! {
    #![proptest_config(ProptestConfig::with_cases(64))]

    #[test]
    fn every_generated_op_round_trips(ops in all_ops()) {
        for (tag, op) in ops.into_iter().enumerate() {
            let bytes = encoded(&op);
            prop_assert_eq!(bytes[0], tag as u8);
            let mut reader = Reader::new(&bytes);
            prop_assert_eq!(Op::read(&mut reader).unwrap(), op);
            prop_assert!(reader.finish().is_ok());
        }
    }

    #[test]
    fn checkpoint_round_trips_generated_history(
        undo in history(), redo in history(), through_seq in any::<u64>(),
        nonce in text(), collaborative in any::<bool>(), seed in opt(any::<u32>()),
        now in opt(finite()),
    ) {
        let workbook = Workbook::from_model(WorkbookModel {
            sheets: vec![Sheet::new("Data")], ..WorkbookModel::default()
        }).unwrap();
        let bytes = workbook.recovery_checkpoint(through_seq, Some(CalculationOptions { now_serial: now })).unwrap();
        let mut outer = Reader::new(&bytes[42..]);
        let mut metadata = Reader::new(outer.bytes().unwrap());
        let mut checkpoint = RecoveryCheckpoint::read(&mut metadata).unwrap();
        metadata.finish().unwrap();
        checkpoint.header.version_nonce = nonce;
        checkpoint.header.rand_seed = seed;
        checkpoint.collaborative = collaborative;
        checkpoint.history = Some(RecoveryHistory {
            preserved_undo: preserved_history(undo.len()), preserved_redo: preserved_history(redo.len()),
            undo: UndoStack::from_recovery_stacks(undo, redo), sheet_order: SheetOrderHistory::default(),
        });
        let original = encoded(&checkpoint);
        let mut payload = Writer::new();
        payload.bool(checkpoint.collaborative);
        payload.bytes(&original);
        let payload = payload.into_bytes();
        let mut outer = Reader::new(&payload);
        let restored_collaborative = outer.bool().unwrap();
        let mut reader = Reader::new(outer.bytes().unwrap());
        let mut restored = RecoveryCheckpoint::read(&mut reader).unwrap();
        prop_assert!(reader.finish().is_ok());
        prop_assert!(outer.finish().is_ok());
        restored.collaborative = restored_collaborative;
        prop_assert_eq!(encoded(&restored), original);
        prop_assert_eq!(restored.collaborative, collaborative);
        prop_assert_eq!(restored.through_seq, through_seq);
        prop_assert_eq!(restored.header.rand_seed, seed);
        prop_assert_eq!(restored.header.calculation_context, Some(CalculationOptions { now_serial: now }));
    }
}
