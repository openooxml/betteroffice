use std::fmt::Debug;

use xlsx_model::{
    Alignment, AnchorCell, AnchorEditAs, AnchorExtent, AutoFilter, ColStyle, DefinedName,
    Protection,
};

use super::*;

fn encoded(value: &impl Codec) -> Vec<u8> {
    let mut writer = Writer::new();
    value.write(&mut writer);
    writer.into_bytes()
}

fn round_trip<T: Codec + PartialEq + Debug>(value: T) {
    let bytes = encoded(&value);
    let mut reader = Reader::new(&bytes);
    assert_eq!(T::read(&mut reader).unwrap(), value);
    reader.finish().unwrap();
}

fn captured() -> CapturedFormat {
    let format = CellFormat {
        alignment: Alignment {
            text_rotation: Some(45),
            indent: Some(2),
            relative_indent: Some(-1),
            justify_last_line: Some(true),
            reading_order: Some(2),
            ..Alignment::default()
        },
        protection: Some(Protection {
            locked: false,
            hidden: true,
        }),
        xf_id: Some(3),
        quote_prefix: true,
        pivot_button: true,
        number_format: NumberFormat::Custom {
            pattern: "0.000".into(),
        },
        ..CellFormat::default()
    };
    CapturedFormat {
        rows: 1,
        columns: 2,
        formats: vec![format.clone(), CellFormat::default()],
        kept: vec![
            KeptFacets {
                protection: true,
                xf_id: true,
                quote_prefix: true,
                pivot_button: true,
                text_rotation: true,
                indent: true,
                relative_indent: true,
                justify_last_line: true,
                reading_order: true,
            },
            KeptFacets::default(),
        ],
        source_styles: vec![Some(7), None],
        source_formats: vec![
            PreservedCellFormat {
                format,
                source_xf: Some(7),
                derived_from: None,
            },
            PreservedCellFormat {
                format: CellFormat::default(),
                source_xf: None,
                derived_from: Some(9),
            },
        ],
    }
}

fn all_ops() -> Vec<Op> {
    let sheet = SheetId(0);
    let at = CellRef::new(2, 3);
    let range = CellRange::new(at, CellRef::new(4, 5));
    let cell = CellState {
        value: CellValue::Number { value: -0.5 },
        formula: Some("SUM(A1:A2)".into()),
        style: Some(3),
    };
    let hyperlink = Hyperlink {
        range,
        external_target: Some("https://example.com".into()),
        location: Some("Data!A1".into()),
        tooltip: Some("tip".into()),
        display: Some("link".into()),
    };
    let anchor = ChartAnchor::OneCell {
        from: AnchorCell {
            row: 2,
            col: 3,
            row_off: -4,
            col_off: 5,
        },
        extent: AnchorExtent { cx: 100, cy: 200 },
    };
    let chart = SheetChart {
        part: "xl/charts/chart1.xml".into(),
        drawing: "xl/drawings/drawing1.xml".into(),
        anchor_index: 4,
        anchor,
        refs: vec![xlsx_model::ChartRef { kind: xlsx_model::ChartRefKind::Values, formula: "Data!A1:A2".into() }],
        raw_xml: Some("<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart/></c:chartSpace>".into()),
    };
    vec![
        Op::SetCell {
            sheet,
            at,
            cell: cell.clone(),
        },
        Op::InsertRows {
            sheet,
            at: 2,
            count: 3,
        },
        Op::DeleteRows {
            sheet,
            at: 3,
            count: 4,
        },
        Op::InsertCols {
            sheet,
            at: 4,
            count: 5,
        },
        Op::DeleteCols {
            sheet,
            at: 5,
            count: 6,
        },
        Op::SetColWidth {
            sheet,
            col: 2,
            width: Some(12.5),
        },
        Op::SetColVisibility {
            sheet,
            col: 3,
            hidden: true,
        },
        Op::RestoreColDimension {
            sheet,
            col: 4,
            width: Some(0.0),
            hidden_width: Some(13.5),
        },
        Op::SetRowHeight {
            sheet,
            row: 5,
            height: None,
        },
        Op::SetRowVisibility {
            sheet,
            row: 6,
            hidden: false,
        },
        Op::RestoreRowDimension {
            sheet,
            row: 7,
            height: Some(0.0),
            hidden_height: Some(24.5),
        },
        Op::SetFreezePane {
            sheet,
            pane: Some(FreezePane::new(1, 2, at)),
        },
        Op::SetHyperlinks {
            sheet,
            hyperlinks: vec![hyperlink.clone()],
        },
        Op::SetHyperlink { sheet, hyperlink },
        Op::RemoveHyperlink { sheet, range },
        Op::SetAutoFilter {
            sheet,
            range: Some(range),
        },
        Op::RestoreAutoFilter {
            sheet,
            filter: Some(AutoFilter {
                range,
                raw_xml: Some("<autoFilter/>".into()),
                source_count: 2,
                filter_columns: vec![1, 2],
                has_sort_state: true,
            }),
        },
        Op::AddConditionalFormat {
            sheet,
            range,
            kind: "cellIs".into(),
            operator: Some("greaterThan".into()),
            formulas: vec!["1".into()],
            format: xlsx_model::DifferentialFormat::default(),
        },
        Op::RemoveConditionalFormat {
            sheet,
            id: "cf:7".into(),
        },
        Op::SetConditionalFormats {
            sheet,
            formats: vec![xlsx_model::ConditionalFormatting {
                sqref: "A1:B2".into(),
                rules: Vec::new(),
            }],
        },
        Op::RestoreColStyles {
            sheet,
            styles: vec![ColStyle {
                first: 1,
                last: 3,
                xf: 2,
            }],
        },
        Op::RestoreRowStyles {
            sheet,
            styles: BTreeMap::from([(3, 1), (70_000, 2)]),
        },
        Op::SetCharts {
            sheet,
            charts: vec![chart],
        },
        Op::SetChartAnchor {
            sheet,
            frame: "drawing#4".into(),
            part: "chart1.xml".into(),
            from: anchor,
            to: ChartAnchor::TwoCell {
                from: AnchorCell::default(),
                to: AnchorCell {
                    row: 10,
                    col: 8,
                    ..AnchorCell::default()
                },
                edit_as: AnchorEditAs::TwoCell,
            },
        },
        Op::MergeCells { sheet, range },
        Op::UnmergeCells { sheet, range },
        Op::PatchRangeStyle {
            sheet,
            range,
            patch: StylePatch {
                bold: Some(true),
                italic: Some(false),
                strikethrough: Some(true),
                font_family: Some("Example".into()),
                font_size: Some(13.5),
                text_color: Some("#123456".into()),
                fill_color: Some("#654321".into()),
                border: Some(BorderPatch {
                    preset: Some(BorderPreset::Outer),
                    style: Some(BorderLineStyle::Double),
                    color: Some("#112233".into()),
                }),
                horizontal_alignment: Some(HorizontalAlignment::Center),
                vertical_alignment: Some(VerticalAlignment::Middle),
                text_wrapping: Some(TextWrapping::Clip),
                clear: vec![StyleProperty::TextColor],
            },
        },
        Op::SetRangeNumberFormat {
            sheet,
            range,
            format: NumberFormatMutation::Custom {
                pattern: "0.000".into(),
            },
        },
        Op::ApplyRangeFormat {
            sheet,
            range,
            format: captured(),
        },
        Op::AddSheet {
            index: 2,
            name: "Added".into(),
        },
        Op::RemoveSheet { index: 1 },
        Op::MoveSheet { from: 2, to: 0 },
        Op::RenameSheet {
            sheet,
            name: "Renamed".into(),
        },
        Op::SetDefinedName {
            name: "Name".into(),
            refers_to: "Data!A1".into(),
            scope: Some(sheet),
            hidden: Some(true),
        },
        Op::DeleteDefinedName {
            name: "Name".into(),
            scope: None,
        },
        Op::SetDefinedNames {
            defined_names: vec![DefinedName {
                name: "Name".into(),
                formula: "Data!A1".into(),
                local_sheet: Some(sheet),
                hidden: true,
            }],
        },
        Op::RestoreSheet {
            sheet,
            name: "Restored".into(),
            formulas: vec![(sheet, at, cell)],
        },
    ]
}

#[test]
fn every_op_has_a_fixed_tag_and_round_trips() {
    let ops = all_ops();
    assert_eq!(ops.len(), 37);
    for (tag, op) in ops.into_iter().enumerate() {
        assert_eq!(encoded(&op)[0], tag as u8);
        round_trip(op);
    }
}

#[test]
fn checkpoint_round_trips_every_op_and_calculation_and_proposal_fields() {
    let mut model = WorkbookModel {
        sheets: vec![Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    model.styles.cell_xfs.push(xlsx_model::Xf::default());
    model.sheets[0].col_widths.insert(2, 0.0);
    model.sheets[0].row_heights.insert(3, 0.0);
    model.sheets[0].set_cell(
        CellRef::new(8, 0),
        xlsx_model::Cell {
            value: CellValue::Number { value: 3.0 },
            ..xlsx_model::Cell::default()
        },
    );
    let mut workbook = Workbook::from_model(model).unwrap();
    let options = CalculationOptions::default();
    for row in 0..2 {
        workbook
            .edit_cell(SheetId(0), CellRef::new(row, 0), "1", options)
            .unwrap();
    }
    workbook.undo(options).unwrap();
    let ops = all_ops();
    workbook.undo = UndoStack::from_recovery_stacks(vec![ops.clone()], vec![ops]);
    let sheet = &mut workbook.model.sheets[0];
    sheet.set_cell(CellRef::new(8, 0), xlsx_model::Cell::default());
    sheet.set_cell(
        CellRef::new(0, 0),
        xlsx_model::Cell {
            value: CellValue::Number { value: 9.0 },
            formula: Some("1+8".into()),
            style: None,
        },
    );
    sheet.set_cell(
        CellRef::new(4, 0),
        xlsx_model::Cell {
            value: CellValue::Text {
                value: "calculated".into(),
            },
            ..xlsx_model::Cell::default()
        },
    );
    sheet.set_array_formula(
        CellRef::new(4, 0),
        CellRange::new(CellRef::new(4, 0), CellRef::new(5, 0)),
    );
    sheet.hidden_col_widths.insert(2, 25.0);
    sheet.hidden_row_heights.insert(3, 45.0);
    workbook
        .model
        .styles
        .fonts
        .push(xlsx_model::Font::default());
    workbook.model.styles.mark_generated_format(0);
    workbook.model.styles.preserve_style_identity();
    workbook.last_calculation.budget_exhausted = true;
    workbook.proposals = ProposalSet::with_id_counter(41);
    workbook.proposals.add(Proposal {
        id: "p1".into(),
        agent_id: "agent".into(),
        note: Some("review".into()),
        edits: vec![ProposedEdit {
            sheet: 0,
            row: 0,
            col: 0,
            input: "0.25".into(),
            old_state: CellState {
                value: CellValue::Number { value: 9.0 },
                formula: Some("1+8".into()),
                style: None,
            },
            number_format: Some(NumberFormatMutation::Percent),
            a1: "A1".into(),
            old_text: "9".into(),
            new_text: "25%".into(),
        }],
        ghosts: vec![ProposalGhost {
            sheet: 0,
            row: 1,
            col: 1,
            old_text: "old".into(),
            new_text: "new".into(),
            alignment_value: CellValue::Bool { value: true },
        }],
    });
    let context = Some(CalculationOptions {
        now_serial: Some(45_000.25),
    });
    let bytes = workbook
        .recovery_checkpoint_with_history(77, context, true)
        .unwrap();
    let (restored, seq, restored_context, history_reset) =
        Workbook::restore_recovery_checkpoint(&bytes).unwrap();
    assert_eq!(seq, 77);
    assert_eq!(restored_context, context);
    assert!(!history_reset);
    assert_eq!(restored.model, workbook.model);
    assert_eq!(restored.proposals(), workbook.proposals());
    assert_eq!(
        restored.proposals.id_counter(),
        workbook.proposals.id_counter()
    );
    assert_eq!(restored.history_state(), workbook.history_state());
    assert_eq!(encoded(&restored.undo), encoded(&workbook.undo));
    assert_eq!(restored.last_calculation, workbook.last_calculation);
    assert_eq!(
        restored.model.styles.has_style_identity(),
        workbook.model.styles.has_style_identity()
    );
    assert_eq!(restored.save().unwrap(), workbook.save().unwrap());
}

#[test]
fn enum_fields_round_trip_every_tag() {
    for value in [
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
        StyleProperty::TextWrapping,
    ] {
        round_trip(value);
    }
    for value in [
        BorderPreset::All,
        BorderPreset::Inner,
        BorderPreset::Horizontal,
        BorderPreset::Vertical,
        BorderPreset::Outer,
        BorderPreset::Left,
        BorderPreset::Top,
        BorderPreset::Right,
        BorderPreset::Bottom,
        BorderPreset::None,
    ] {
        round_trip(value);
    }
    for value in [
        BorderLineStyle::Solid,
        BorderLineStyle::Dashed,
        BorderLineStyle::Dotted,
        BorderLineStyle::Double,
    ] {
        round_trip(value);
    }
    for value in [
        HorizontalAlignment::Left,
        HorizontalAlignment::Center,
        HorizontalAlignment::Right,
    ] {
        round_trip(value);
    }
    for value in [
        VerticalAlignment::Top,
        VerticalAlignment::Middle,
        VerticalAlignment::Bottom,
    ] {
        round_trip(value);
    }
    for value in [
        TextWrapping::Overflow,
        TextWrapping::Wrap,
        TextWrapping::Clip,
    ] {
        round_trip(value);
    }
    for value in [
        NumberFormatMutation::Automatic,
        NumberFormatMutation::PlainText,
        NumberFormatMutation::Number,
        NumberFormatMutation::Percent,
        NumberFormatMutation::Scientific,
        NumberFormatMutation::Currency,
        NumberFormatMutation::Date,
        NumberFormatMutation::Time,
        NumberFormatMutation::Custom {
            pattern: "0.0".into(),
        },
        NumberFormatMutation::IncreaseDecimal,
        NumberFormatMutation::DecreaseDecimal,
    ] {
        round_trip(value);
    }
    for value in [
        CellValue::Empty,
        CellValue::Number { value: -0.0 },
        CellValue::Text { value: "é".into() },
        CellValue::Bool { value: true },
    ] {
        round_trip(value);
    }
    for value in [
        ErrorValue::Div0,
        ErrorValue::NA,
        ErrorValue::Name,
        ErrorValue::Null,
        ErrorValue::Num,
        ErrorValue::Ref,
        ErrorValue::Value,
        ErrorValue::Spill,
        ErrorValue::Calc,
    ] {
        round_trip(CellValue::Error { value });
    }
    for author in [
        Provenance::User,
        Provenance::Agent { id: "agent".into() },
        Provenance::Remote {
            actor: "peer".into(),
        },
    ] {
        round_trip(Transaction {
            ops: all_ops(),
            author,
            proposed: true,
        });
    }
    round_trip(captured());
    round_trip(NumberFormat::Builtin { id: 14 });
}

#[test]
fn malformed_metadata_tags_counts_and_truncations_are_refused() {
    fn invalid_tag<T: Codec>() {
        assert!(T::read(&mut Reader::new(&[255])).is_err());
    }

    invalid_tag::<RecoveryCheckpoint>();
    invalid_tag::<RecoveryHistory>();
    invalid_tag::<RecoveryPreservedHistory>();
    invalid_tag::<SheetOrderHistory>();
    invalid_tag::<UndoStack>();
    invalid_tag::<Proposal>();
    invalid_tag::<ProposedEdit>();
    invalid_tag::<ProposalGhost>();
    invalid_tag::<StylePatch>();
    invalid_tag::<BorderPatch>();
    invalid_tag::<StyleProperty>();
    invalid_tag::<BorderPreset>();
    invalid_tag::<BorderLineStyle>();
    invalid_tag::<HorizontalAlignment>();
    invalid_tag::<VerticalAlignment>();
    invalid_tag::<TextWrapping>();
    invalid_tag::<KeptFacets>();
    invalid_tag::<Provenance>();
    invalid_tag::<NumberFormat>();
    invalid_tag::<Stylesheet>();
    assert!(Op::read(&mut Reader::new(&[37])).is_err());
    assert!(CellState::read(&mut Reader::new(&[1])).is_err());
    assert!(CellValue::read(&mut Reader::new(&[5])).is_err());
    assert!(NumberFormatMutation::read(&mut Reader::new(&[11])).is_err());
    assert!(CapturedFormat::read(&mut Reader::new(&[1])).is_err());
    let mut count = Writer::new();
    count.var_u64(u64::MAX);
    let count = count.into_bytes();
    assert!(Vec::<Op>::read(&mut Reader::new(&count)).is_err());
    assert!(BTreeMap::<u32, f64>::read(&mut Reader::new(&count)).is_err());
    for op in all_ops() {
        let bytes = encoded(&op);
        for end in 0..bytes.len() {
            assert!(Op::read(&mut Reader::new(&bytes[..end])).is_err());
        }
    }
    let mut bounded = Writer::bounded(16);
    captured().write(&mut bounded);
    assert!(bounded.len() > 16);
    assert_eq!(bounded.into_bytes().len(), 16);
}

#[test]
fn stylesheet_recovery_keeps_provenance_and_legacy_facet_markers() {
    let mut styles = Stylesheet::default();
    styles.cell_xfs.push(xlsx_model::Xf::default());
    styles.push_released_xf(xlsx_model::Xf::default());
    styles.restore_recovery_provenance(
        BTreeMap::from([(0, None), (1, Some(7))]),
        true,
        vec![false, true],
    );
    let bytes = encoded(&styles);
    let mut reader = Reader::new(&bytes);
    let restored = Stylesheet::read(&mut reader).unwrap();
    reader.finish().unwrap();
    assert_eq!(restored, styles);
    assert_eq!(restored.recovery_provenance(), styles.recovery_provenance());
}
