use super::*;
use yrs::updates::decoder::Decode;
use yrs::updates::encoder::Encode;
use yrs::{Any, Array, Doc, Map, MapPrelim, ReadTxn, StateVector, Transact, Update, WriteTxn};

fn empty_worker() -> Workbook {
    Workbook::from_model(WorkbookModel {
        sheets: vec![crate::Sheet::new("Data")],
        ..WorkbookModel::default()
    })
    .unwrap()
}

fn change_header(
    records: &mut [(ChunkKind, u64, Vec<u8>)],
    change: impl FnOnce(&mut SnapshotHeader),
) {
    let mut reader = Reader::new(&records[0].2);
    let mut header = SnapshotHeader::decode(reader.bytes().unwrap()).unwrap();
    change(&mut header);
    let mut writer = Writer::new();
    writer.bytes(&header.encode());
    writer.raw(reader.rest());
    records[0].2 = writer.into_bytes();
}

fn transport_records(
    mut records: Vec<(ChunkKind, u64, Vec<u8>)>,
    budget: SnapshotBudget,
) -> Vec<Vec<u8>> {
    let mut counts = [0u64; 9];
    for (kind, _, _) in &records {
        counts[index(*kind)] += 1;
    }
    counts[index(ChunkKind::End)] = 1;
    change_header(&mut records, |header| {
        header.snapshot_id = 17;
        header.chunk_counts = counts;
    });
    let mut digest = Sha256::new();
    let mut logical_ordinals = [0; 9];
    let mut transport_ordinals = [0; 9];
    let mut chunks = Vec::new();
    for (kind, _, payload) in records {
        let ordinal = if kind == ChunkKind::Cells {
            logical_ordinals[index(ChunkKind::Model)] + logical_ordinals[index(kind)]
        } else {
            logical_ordinals[index(kind)]
        };
        logical_ordinals[index(kind)] += 1;
        let logical = frame(kind, ordinal, &payload);
        digest.update(&logical);
        append_transport(
            &mut chunks,
            &mut transport_ordinals,
            kind,
            ordinal,
            &logical,
            budget,
        );
    }
    let logical = frame(ChunkKind::End, 0, &digest.finalize());
    append_transport(
        &mut chunks,
        &mut transport_ordinals,
        ChunkKind::End,
        0,
        &logical,
        budget,
    );
    chunks
}

fn append_transport(
    chunks: &mut Vec<Vec<u8>>,
    ordinals: &mut [u64; 9],
    kind: ChunkKind,
    ordinal: u64,
    logical: &[u8],
    budget: SnapshotBudget,
) {
    let part_bytes = budget.max_bytes() - 64;
    for (part, bytes) in logical.chunks(part_bytes).enumerate() {
        let mut writer = Writer::new();
        writer.var_u64(17);
        writer.var_u64(ordinal);
        writer.var_usize(logical.len());
        writer.var_usize(part * part_bytes);
        writer.raw(bytes);
        chunks.push(frame(kind, ordinals[index(kind)], &writer.into_bytes()));
        ordinals[index(kind)] += 1;
    }
}

fn replace_yrs(
    worker: &Workbook,
    updates: Vec<Vec<u8>>,
    vector: Vec<u8>,
    budget: SnapshotBudget,
) -> Vec<Vec<u8>> {
    let mut records = decoded_snapshot(&encode(worker, budget));
    records.retain(|(kind, _, _)| *kind != ChunkKind::Yrs);
    for update in updates {
        records.push((ChunkKind::Yrs, 0, update));
    }
    records.sort_by_key(|(kind, _, _)| index(*kind));
    change_header(&mut records, |header| header.state_vector = vector);
    transport_records(records, budget)
}

fn split(update: &[u8], budget: SnapshotBudget) -> Vec<Vec<u8>> {
    crate::snapshot::yrs_split::split_fallback_v1_bounded(
        update,
        budget.max_records(),
        budget.max_bytes() - 64,
    )
    .unwrap()
}

fn assert_refuses_before_ready(chunks: &[Vec<u8>], budget: SnapshotBudget) -> String {
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in chunks {
        if let Err(failure) = builder.push(chunk) {
            assert_step_budget(budget);
            assert!(builder.finish().is_err());
            return failure.to_string();
        }
        assert_step_budget(budget);
        if let Err(failure) = builder.advance(budget) {
            assert_step_budget(budget);
            assert!(builder.finish().is_err());
            return failure.to_string();
        }
    }
    assert!(builder.ended);
    for _ in 0..200_000 {
        match builder.advance(budget) {
            Ok(progress) => assert!(!progress.is_ready()),
            Err(failure) => {
                assert_step_budget(budget);
                assert!(builder.finish().is_err());
                return failure.to_string();
            }
        }
        assert_step_budget(budget);
    }
    panic!("snapshot did not refuse");
}

fn apply_matching_history_step(worker: &mut Workbook, peer: &mut Workbook, operation: Op) {
    assert_eq!(
        worker
            .apply_ops(vec![operation.clone()], context())
            .unwrap(),
        peer.apply_ops(vec![operation], context()).unwrap(),
    );
    assert_edited_identity(worker, peer);
    assert!(worker.can_undo());
    assert!(peer.can_undo());
    assert_eq!(
        worker.undo(context()).unwrap(),
        peer.undo(context()).unwrap()
    );
    assert_edited_identity(worker, peer);
    assert!(worker.can_redo());
    assert!(peer.can_redo());
    assert_eq!(
        worker.redo(context()).unwrap(),
        peer.redo(context()).unwrap()
    );
    assert_edited_identity(worker, peer);
}

fn delta_doc(worker: &Workbook) -> (Doc, StateVector) {
    let baseline = StateVector::decode_v1(&worker.encode_state_vector_v1()).unwrap();
    let mut client = 1;
    while baseline.contains_client(&yrs::block::ClientID::new(client)) {
        client += 1;
    }
    let doc = Doc::with_client_id(client);
    doc.transact_mut()
        .apply_update(Update::decode_v1(&worker.encode_state_as_update_v1()).unwrap())
        .unwrap();
    (doc, baseline)
}

fn replace_yrs_with_delta(
    worker: &Workbook,
    doc: &Doc,
    baseline: &StateVector,
    budget: SnapshotBudget,
) -> Vec<Vec<u8>> {
    let mut updates = split(&worker.encode_state_as_update_v1(), budget);
    let txn = doc.transact();
    let delta = txn.encode_diff_v1(baseline);
    let mut cursor = crate::snapshot::yrs_split::UpdateCursor::default();
    while let Some(part) = cursor
        .next(&delta, budget.max_records(), budget.max_bytes() - 64)
        .unwrap()
    {
        updates.push(part.bytes);
    }
    let state = txn.state_vector();
    let mut entries = state
        .iter()
        .map(|(client, clock)| (client.get(), *clock))
        .collect::<Vec<_>>();
    entries.sort_unstable();
    let mut writer = Writer::new();
    writer.var_usize(entries.len());
    for (client, clock) in entries {
        writer.var_u64(client);
        writer.var_u32(clock);
    }
    replace_yrs(worker, updates, writer.into_bytes(), budget)
}

#[test]
fn snapshot_initial_peer_matches_sheet_edit_history_and_live_reorder_refusal() {
    for (records, bytes) in [(1, 1024), (256, 16_384)] {
        let budget = SnapshotBudget::new(records, bytes).unwrap();
        let mut worker = empty_worker();
        let chunks = encode(&worker, budget);
        let (builder, _) = completed_builder_with_step_budget(&chunks, budget);
        assert!(builder.ready.is_some());
        let (mut peer, _) = builder.finish().unwrap().into_parts();
        assert_edited_identity(&worker, &peer);
        for index in 0..32 {
            apply_matching_history_step(
                &mut worker,
                &mut peer,
                Op::AddSheet {
                    index,
                    name: format!("Added {index}"),
                },
            );
        }
        for operation in [
            Op::RenameSheet {
                sheet: SheetId(1),
                name: "Renamed".to_owned(),
            },
            Op::RemoveSheet { index: 0 },
        ] {
            apply_matching_history_step(&mut worker, &mut peer, operation);
        }
        assert_eq!(worker.model().sheets.len(), 32);
        assert_eq!(worker.model().sheets[0].name, "Renamed");
        assert!(deleted_clock_total(&worker) > 1);
        for workbook in [&mut worker, &mut peer] {
            {
                let mut txn = workbook.authority.snapshot_transaction_for_test();
                let order = txn.get_array("xlsx:sheet-order").unwrap();
                let key = order.get(&txn, 0).unwrap().cast::<String>().unwrap();
                order.remove_range(&mut txn, 0, 1);
                order.insert(&mut txn, 31, key);
            }
            workbook.model = workbook.authority.materialize().unwrap();
        }
        assert_edited_identity(&worker, &peer);
        assert_eq!(worker.model().sheets[31].name, "Renamed");
        let initial = Workbook::from_model(WorkbookModel {
            sheets: vec![crate::Sheet::new("First"), crate::Sheet::new("Second")],
            ..WorkbookModel::default()
        })
        .unwrap();
        let (doc, baseline) = delta_doc(&initial);
        {
            let mut txn = doc.transact_mut();
            let order = txn.get_array("xlsx:sheet-order").unwrap();
            order.remove_range(&mut txn, 0, 1);
            order.insert(&mut txn, 1, "sheet:0");
        }
        let chunks = replace_yrs_with_delta(&initial, &doc, &baseline, budget);
        assert!(
            assert_refuses_before_ready(&chunks, budget).contains("authority and model disagree")
        );
    }
}

#[test]
fn snapshot_initial_envelope_refuses_retained_sheet_maps_and_non_maps() {
    let worker = empty_worker();
    for retained_map in [false, true] {
        let (doc, baseline) = delta_doc(&worker);
        {
            let mut txn = doc.transact_mut();
            let sheets = txn.get_map("xlsx:sheets").unwrap();
            if retained_map {
                let retained = sheets.insert(&mut txn, "sheet:retained", MapPrelim::default());
                retained.insert(&mut txn, "name", "Retained");
            } else {
                sheets.insert(&mut txn, "sheet:retained", "not a sheet map");
            }
        }
        for (records, bytes) in [(1, 1024), (256, 16_384)] {
            let budget = SnapshotBudget::new(records, bytes).unwrap();
            let chunks = replace_yrs_with_delta(&worker, &doc, &baseline, budget);
            if retained_map {
                assert_eq!(
                    assert_refuses_before_ready(&chunks, budget),
                    "snapshot retained sheet map is outside the live sheet order",
                );
            } else {
                assert!(
                    assert_refuses_before_ready(&chunks, budget)
                        .contains("authority and model disagree")
                );
            }
        }
    }
}

#[test]
fn snapshot_live_sheet_order_requires_strings_and_existing_sheets() {
    for invalid_order in [
        Any::BigInt(7),
        Any::from("sheet:missing"),
        Any::from("sheet:0"),
    ] {
        let worker = empty_worker();
        let (doc, baseline_state) = delta_doc(&worker);
        let duplicate = invalid_order == Any::from("sheet:0");
        let non_string = !matches!(&invalid_order, Any::String(_));
        {
            let mut txn = doc.transact_mut();
            let order = txn.get_array("xlsx:sheet-order").unwrap();
            order.remove_range(&mut txn, 0, 1);
            order.insert(&mut txn, 0, invalid_order);
            if duplicate {
                order.insert(&mut txn, 1, "sheet:0");
            }
        }
        for (records, bytes) in [(1, 1024), (256, 16_384)] {
            let budget = SnapshotBudget::new(records, bytes).unwrap();
            let chunks = replace_yrs_with_delta(&worker, &doc, &baseline_state, budget);
            let failure = assert_refuses_before_ready(&chunks, budget);
            assert!(failure.contains(if non_string {
                "unsupported_content"
            } else {
                "authority"
            }));
        }
    }
}

fn sheet_order_history_update(client: u64, deleted_items: usize, deleted_len: u32) -> Vec<u8> {
    let mut update = Writer::new();
    update.var_usize(1);
    update.var_usize(deleted_items + 1);
    update.var_u64(client);
    update.var_u32(0);
    let mut clock = 0;
    for index in 0..=deleted_items {
        let kind = if index == deleted_items {
            yrs::block::BLOCK_ITEM_ANY_REF_NUMBER
        } else {
            yrs::block::BLOCK_ITEM_DELETED_REF_NUMBER
        };
        if index == 0 {
            update.u8(kind);
            update.var_u32(1);
            update.str("xlsx:sheet-order");
        } else {
            update.u8(yrs::block::HAS_ORIGIN | kind);
            update.var_u64(client);
            update.var_u32(clock - 1);
        }
        if index == deleted_items {
            update.var_u32(1);
            update.u8(119);
            update.str("sheet:0");
        } else {
            update.var_u32(deleted_len);
            clock += deleted_len;
        }
    }
    if deleted_items == 0 {
        update.var_usize(0);
    } else {
        update.var_usize(1);
        update.var_u64(client);
        update.var_usize(1);
        update.var_u32(0);
        update.var_u32(clock);
    }
    update.into_bytes()
}

#[test]
fn snapshot_sheet_order_cap_refuses_before_integration() {
    use crate::snapshot::yrs_split::{CausalState, SHEET_ORDER_MAX_ITEMS, SplitError};

    let at_cap = sheet_order_history_update(1, SHEET_ORDER_MAX_ITEMS - 1, 1);
    assert!(CausalState::default().admit(&at_cap).is_ok());
    for update in [
        sheet_order_history_update(1, SHEET_ORDER_MAX_ITEMS, 1),
        sheet_order_history_update(1, 1, SHEET_ORDER_MAX_ITEMS as u32),
    ] {
        let mut causal = CausalState::default();
        assert_eq!(causal.admit(&update), Err(SplitError::SheetOrderLimit));
        assert!(causal.is_empty());
        assert!(causal.take_keys().is_empty());
        assert!(causal.admit(&at_cap).is_ok());
        let mut vector = Writer::new();
        vector.var_usize(1);
        vector.var_u64(1);
        vector.var_usize(SHEET_ORDER_MAX_ITEMS + 1);
        let vector = vector.into_bytes();
        for (records, bytes) in [(1, 1024), (256, 16_384)] {
            let budget = SnapshotBudget::new(records, bytes).unwrap();
            let chunks = replace_yrs(
                &empty_worker(),
                vec![update.clone()],
                vector.clone(),
                budget,
            );
            assert!(assert_refuses_before_ready(&chunks, budget).contains("sheet_order"));
        }
    }
}

#[test]
fn snapshot_packed_live_sheet_order_does_not_charge_keys_as_history() {
    use crate::snapshot::yrs_split::{CausalState, SHEET_ORDER_MAX_ITEMS};

    let mut update = Writer::new();
    update.var_usize(1);
    update.var_usize(1);
    update.var_u64(1);
    update.var_u32(0);
    update.u8(yrs::block::BLOCK_ITEM_ANY_REF_NUMBER);
    update.var_u32(1);
    update.str("xlsx:sheet-order");
    update.var_usize(SHEET_ORDER_MAX_ITEMS + 1);
    for index in 0..=SHEET_ORDER_MAX_ITEMS {
        update.u8(119);
        update.str(&format!("sheet:{index}"));
    }
    update.var_usize(0);
    assert!(CausalState::default().admit(&update.into_bytes()).is_ok());
}

#[test]
fn snapshot_refusal_leaves_the_non_snapshot_open_path_usable() {
    for (records, bytes) in [(1, 1024), (256, 16_384)] {
        let mut worker = Workbook::open(&empty_worker().save().unwrap()).unwrap();
        let budget = SnapshotBudget::new(records, bytes).unwrap();
        let mut snapshot = decoded_snapshot(&encode(&worker, budget));
        worker.authority.snapshot_skip_gc_for_test();
        {
            let mut txn = worker.authority.snapshot_transaction_for_test();
            let mut client = 1;
            while txn
                .state_vector()
                .contains_client(&yrs::block::ClientID::new(client))
            {
                client += 1;
            }
            txn.get_array("xlsx:sheet-order")
                .unwrap()
                .remove_range(&mut txn, 0, 1);
            let update = sheet_order_history_update(
                client,
                crate::snapshot::yrs_split::SHEET_ORDER_MAX_ITEMS,
                1,
            );
            txn.apply_update(Update::decode_v1(&update).unwrap())
                .unwrap();
        }
        assert_eq!(worker.authority.materialize().unwrap(), *worker.model());
        let failure = WorkbookSnapshotEncoder::new(&worker, Some(context()), budget)
            .err()
            .unwrap();
        assert!(failure.to_string().contains("sheet_order"));
        snapshot.retain(|(kind, _, _)| *kind != ChunkKind::Yrs);
        snapshot.push((ChunkKind::Yrs, 0, worker.encode_state_as_update_v1()));
        snapshot.sort_by_key(|(kind, _, _)| index(*kind));
        change_header(&mut snapshot, |header| {
            header.state_vector = worker.encode_state_vector_v1();
        });
        let chunks = transport_records(snapshot, budget);
        let mut builder = WorkbookSnapshotBuilder::new();
        for chunk in &chunks {
            builder.push(chunk).unwrap();
            assert_step_budget(budget);
        }
        let mut refused = false;
        for _ in 0..200_000 {
            match builder.advance(budget) {
                Ok(progress) => assert!(!progress.is_ready()),
                Err(failure) => {
                    assert!(failure.to_string().contains("sheet_order"));
                    assert!(builder.failed);
                    assert!(builder.ready.is_none());
                    assert!(builder.restored.is_none());
                    refused = true;
                    break;
                }
            }
            assert_step_budget(budget);
        }
        assert!(refused);
        assert_step_budget(budget);
        assert!(builder.finish().is_err());
        let saved = worker.save().unwrap();
        let mut peer = Workbook::open(&saved).unwrap();
        assert_eq!(worker.model(), peer.model());
        for operation in [
            Op::SetCell {
                sheet: SheetId(0),
                at: CellRef::new(0, 0),
                cell: CellState {
                    value: CellValue::Number { value: 7.0 },
                    ..CellState::default()
                },
            },
            Op::RenameSheet {
                sheet: SheetId(0),
                name: "Fallback edit".to_owned(),
            },
        ] {
            assert_eq!(
                worker
                    .apply_ops(vec![operation.clone()], context())
                    .unwrap(),
                peer.apply_ops(vec![operation], context()).unwrap(),
            );
            assert_eq!(worker.model(), peer.model());
        }
    }
}

#[test]
fn snapshot_initial_peer_matches_applied_edit_histories() {
    for (records, bytes) in [(1, 1024), (256, 16_384)] {
        let budget = SnapshotBudget::new(records, bytes).unwrap();
        let mut parts = ooxml_opc::unzip_parts(&source(false, false)).unwrap();
        let (_, data) = parts
            .iter_mut()
            .find(|(name, _)| name == "xl/worksheets/sheet1.xml")
            .unwrap();
        *data = std::str::from_utf8(data)
            .unwrap()
            .replace("SUM(Items[Qty])+Input", "SUM(A5:A6)+Input")
            .into_bytes();
        let mut worker = worker(&ooxml_opc::rezip_parts(&parts).unwrap());
        let chunks = encode(&worker, budget);
        let (builder, _) = completed_builder_with_step_budget(&chunks, budget);
        assert!(builder.ready.is_some());
        let (mut peer, _) = builder.finish().unwrap().into_parts();
        assert_edited_identity(&worker, &peer);
        let sheet = SheetId(0);
        let at = CellRef::new(20, 4);
        let cell_range = crate::CellRange { start: at, end: at };
        let merge = crate::CellRange {
            start: CellRef::new(22, 4),
            end: CellRef::new(22, 5),
        };
        let chart = worker.model().sheets[0].charts[0].clone();
        let crate::ChartAnchor::OneCell { mut from, extent } = chart.anchor else {
            panic!("expected a movable chart");
        };
        from.row += 1;
        from.col += 1;
        for operation in [
            Op::SetCell {
                sheet,
                at,
                cell: CellState {
                    value: CellValue::Number { value: 7.0 },
                    ..CellState::default()
                },
            },
            Op::PatchRangeStyle {
                sheet,
                range: cell_range,
                patch: crate::StylePatch {
                    bold: Some(true),
                    ..crate::StylePatch::default()
                },
            },
            Op::SetRangeNumberFormat {
                sheet,
                range: cell_range,
                format: crate::NumberFormatMutation::Custom {
                    pattern: "0.0000".to_owned(),
                },
            },
            Op::SetHyperlinks {
                sheet,
                hyperlinks: vec![crate::Hyperlink {
                    range: cell_range,
                    external_target: Some("https://example.com".to_owned()),
                    location: None,
                    tooltip: Some("Edited link".to_owned()),
                    display: Some("Link".to_owned()),
                }],
            },
            Op::MergeCells {
                sheet,
                range: merge,
            },
            Op::UnmergeCells {
                sheet,
                range: merge,
            },
            Op::SetCell {
                sheet,
                at,
                cell: CellState::default(),
            },
            Op::SetChartAnchor {
                sheet,
                frame: chart.frame_id(),
                part: chart.part.clone(),
                from: chart.anchor,
                to: crate::ChartAnchor::OneCell { from, extent },
            },
            Op::InsertRows {
                sheet,
                at: 0,
                count: 1,
            },
            Op::RenameSheet {
                sheet,
                name: "Renamed".to_owned(),
            },
        ] {
            let defined_names = worker.model().defined_names.clone();
            let remaps_names = matches!(&operation, Op::InsertRows { .. } | Op::RenameSheet { .. });
            let catalog_edit = matches!(
                &operation,
                Op::PatchRangeStyle { .. } | Op::SetRangeNumberFormat { .. }
            );
            let styles = worker.model().styles.cell_xfs.len();
            apply_matching_history_step(&mut worker, &mut peer, operation);
            if remaps_names {
                assert_ne!(worker.model().defined_names, defined_names);
            }
            if catalog_edit {
                assert!(worker.model().styles.cell_xfs.len() > styles);
            }
        }
    }
}

fn replace_snapshot_model(
    chunks: &[Vec<u8>],
    model: &WorkbookModel,
    budget: SnapshotBudget,
) -> Vec<Vec<u8>> {
    let mut records = decoded_snapshot(chunks);
    records.retain(|(kind, _, _)| !matches!(kind, ChunkKind::Model | ChunkKind::Cells));
    let mut encoder = ModelSnapshotEncoder::new();
    while let Some(chunk) = encoder.next(model, budget).unwrap() {
        let (kind, ordinal, payload) = unframe(&chunk).unwrap();
        records.push((kind, ordinal, payload.to_vec()));
    }
    records.sort_by_key(|(kind, _, _)| index(*kind));
    let mut reader = Reader::new(&records[0].2);
    let mut writer = Writer::new();
    writer.bytes(reader.bytes().unwrap());
    for _ in 0..3 {
        writer.bool(reader.bool().unwrap());
    }
    writer.option(reader.option(Reader::var_usize).unwrap(), Writer::var_usize);
    for count in model.styles.snapshot_field_counts() {
        reader.var_usize().unwrap();
        writer.var_usize(count);
    }
    writer.raw(reader.rest());
    records[0].2 = writer.into_bytes();
    transport_records(records, budget)
}

#[test]
fn snapshot_initial_envelope_refuses_appended_styles_and_contradictions() {
    let worker = empty_worker();
    let mut model = worker.model().clone();
    let mut format = model.styles.cell_format(None);
    format.font.bold = true;
    let style = model.styles.intern_cell_format(&format).unwrap().unwrap();
    model.sheets[0].set_cell(
        CellRef::new(0, 0),
        Cell {
            style: Some(style),
            ..Cell::default()
        },
    );
    let payload = serde_json::to_string(&format).unwrap();
    let key = format!("{:x}", Sha256::digest(payload.as_bytes()));
    let (doc, baseline) = delta_doc(&worker);
    {
        let mut txn = doc.transact_mut();
        txn.get_map("xlsx:cell-formats")
            .unwrap()
            .insert(&mut txn, key.clone(), payload);
        let sheet = txn
            .get_map("xlsx:sheets")
            .unwrap()
            .get(&txn, "sheet:0")
            .unwrap()
            .cast::<yrs::MapRef>()
            .unwrap();
        sheet
            .get(&txn, "styles")
            .unwrap()
            .cast::<yrs::MapRef>()
            .unwrap()
            .insert(&mut txn, "0:0", key);
    }
    for (records, bytes) in [(1, 1024), (256, 16_384)] {
        let budget = SnapshotBudget::new(records, bytes).unwrap();
        let envelope = replace_yrs_with_delta(&worker, &doc, &baseline, budget);
        let chunks = replace_snapshot_model(&envelope, &model, budget);
        assert_eq!(
            assert_refuses_before_ready(&chunks, budget),
            "snapshot authority and model disagree: style tables extend beyond the authority base",
        );
        for unused in [false, true] {
            let mut contradictory = model.clone();
            if unused {
                contradictory.styles.fonts.push(xlsx_model::styles::Font {
                    name: Some("Unlisted".to_owned()),
                    ..xlsx_model::styles::Font::default()
                });
            } else {
                let font = contradictory.styles.cell_xfs[style as usize].font.unwrap();
                contradictory.styles.fonts[font as usize].bold = false;
            }
            let chunks = replace_snapshot_model(&envelope, &contradictory, budget);
            assert!(
                assert_refuses_before_ready(&chunks, budget)
                    .contains("authority and model disagree")
            );
        }
        let mut default_equivalent = worker.model().clone();
        let style = default_equivalent.styles.cell_xfs.len() as u32;
        default_equivalent
            .styles
            .cell_xfs
            .push(xlsx_model::styles::Xf {
                num_fmt_id: Some(0),
                ..xlsx_model::styles::Xf::default()
            });
        default_equivalent.sheets[0].set_cell(
            CellRef::new(0, 0),
            Cell {
                style: Some(style),
                ..Cell::default()
            },
        );
        let format = default_equivalent.styles.cell_format(Some(style));
        assert_eq!(format, xlsx_model::styles::CellFormat::default());
        let payload = serde_json::to_string(&format).unwrap();
        let key = format!("{:x}", Sha256::digest(payload.as_bytes()));
        let (doc, baseline) = delta_doc(&worker);
        {
            let mut txn = doc.transact_mut();
            let sheet = txn
                .get_map("xlsx:sheets")
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<yrs::MapRef>()
                .unwrap();
            sheet
                .get(&txn, "styles")
                .unwrap()
                .cast::<yrs::MapRef>()
                .unwrap()
                .insert(&mut txn, "0:0", key);
        }
        let envelope = replace_yrs_with_delta(&worker, &doc, &baseline, budget);
        let chunks = replace_snapshot_model(&envelope, &default_equivalent, budget);
        assert_eq!(
            assert_refuses_before_ready(&chunks, budget),
            "snapshot authority and model disagree: style tables extend beyond the authority base",
        );
    }
}

#[test]
fn snapshot_initial_catalog_rejects_changed_and_unlisted_authority_formats() {
    let mut model = WorkbookModel {
        sheets: vec![crate::Sheet::new("Data")],
        ..WorkbookModel::default()
    };
    model.styles.fonts.push(xlsx_model::styles::Font {
        bold: true,
        ..xlsx_model::styles::Font::default()
    });
    model.styles.cell_xfs.push(xlsx_model::styles::Xf {
        font: Some(0),
        ..xlsx_model::styles::Xf::default()
    });
    model.sheets[0].set_cell(
        CellRef::new(0, 0),
        Cell {
            style: Some(0),
            ..Cell::default()
        },
    );
    let worker = Workbook::from_model(model).unwrap();
    for unused in [false, true] {
        let (doc, baseline) = delta_doc(&worker);
        {
            let mut txn = doc.transact_mut();
            let formats = txn.get_map("xlsx:cell-formats").unwrap();
            if unused {
                let mut styles = worker.model().styles.clone();
                styles.fonts[0].name = Some("Unlisted".to_owned());
                let format = styles.cell_format(Some(0));
                let payload = serde_json::to_string(&format).unwrap();
                let key = format!("{:x}", Sha256::digest(payload.as_bytes()));
                formats.insert(&mut txn, key, payload);
            } else {
                let format = worker.model().styles.cell_format(Some(0));
                let payload = serde_json::to_string(&format).unwrap();
                let key = format!("{:x}", Sha256::digest(payload.as_bytes()));
                let mut changed = format;
                changed.font.bold = false;
                formats.insert(&mut txn, key, serde_json::to_string(&changed).unwrap());
            }
        }
        for (records, bytes) in [(1, 1024), (256, 16_384)] {
            let budget = SnapshotBudget::new(records, bytes).unwrap();
            let chunks = replace_yrs_with_delta(&worker, &doc, &baseline, budget);
            assert!(
                assert_refuses_before_ready(&chunks, budget)
                    .contains("authority and model disagree")
            );
        }
    }
}

#[test]
fn snapshot_authority_schema_and_authored_model_agree_before_ready() {
    let budget = budgets()[0];
    let worker = empty_worker();
    let chunks = replace_yrs(&worker, vec![vec![0, 0]], vec![0], budget);
    assert!(assert_refuses_before_ready(&chunks, budget).contains("authority"));
    let mut renamed = worker.encode_state_as_update_v1();
    let mut matches = 0;
    for offset in 0..renamed.len().saturating_sub(6) {
        if renamed.get(offset..offset + 7) == Some(b"sheet:0") {
            renamed[offset + 6] = b'9';
            matches += 1;
        }
    }
    assert_eq!(matches, 2);
    let chunks = replace_yrs(
        &worker,
        vec![renamed],
        worker.encode_state_vector_v1(),
        budget,
    );
    let (builder, _) = completed_builder_with_step_budget(&chunks, budget);
    let (peer, _) = builder.finish().unwrap().into_parts();
    assert_eq!(
        canonical_model(worker.model()),
        canonical_model(peer.model())
    );
    {
        let mut txn = worker.authority.snapshot_transaction_for_test();
        txn.get_map("xlsx")
            .unwrap()
            .insert(&mut txn, "schemaVersion", 999_i64);
    }
    let chunks = encode(&worker, budget);
    assert!(assert_refuses_before_ready(&chunks, budget).contains("schema version"));
    for metadata in [false, true] {
        let mut worker = empty_worker();
        if metadata {
            worker.model.sheets[0].name = "Contradiction".to_owned();
        } else {
            worker.model.sheets[0].set_cell(
                CellRef::new(0, 0),
                Cell {
                    value: CellValue::Number { value: 99.0 },
                    ..Cell::default()
                },
            );
        }
        let chunks = encode(&worker, budget);
        assert!(
            assert_refuses_before_ready(&chunks, budget).contains("authority and model disagree")
        );
    }
    let mut sheet = crate::Sheet::new("Data");
    sheet.set_cell(
        CellRef::new(0, 0),
        Cell {
            value: CellValue::Number { value: 1.0 },
            ..Cell::default()
        },
    );
    let worker = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    {
        let mut txn = worker.authority.snapshot_transaction_for_test();
        let sheets = txn.get_map("xlsx:sheets").unwrap();
        let sheet = sheets
            .get(&txn, "sheet:0")
            .unwrap()
            .cast::<yrs::MapRef>()
            .unwrap();
        let contents = sheet
            .get(&txn, "contents")
            .unwrap()
            .cast::<yrs::MapRef>()
            .unwrap();
        let value = Any::Array(
            vec![
                Any::BigInt(0),
                Any::Array(vec![Any::BigInt(1), Any::Number(99.0)].into()),
            ]
            .into(),
        );
        contents.insert(&mut txn, "0:0", value);
    }
    assert!(
        assert_refuses_before_ready(&encode(&worker, budget), budget)
            .contains("authority and model disagree")
    );
    let mut sheet = crate::Sheet::new("Data");
    for row in 0..512 {
        sheet.set_cell(
            CellRef::new(row, 0),
            Cell {
                value: CellValue::Number {
                    value: f64::from(row),
                },
                ..Cell::default()
            },
        );
    }
    let worker = Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap();
    for records in [1, 7, 256] {
        let budget = SnapshotBudget::new(records, 16_384).unwrap();
        let chunks = encode(&worker, budget);
        let (builder, _) = completed_builder_with_step_budget(&chunks, budget);
        assert!(builder.authority_validated);
        assert!(builder.authority_keys.is_empty());
        let (peer, _) = builder.finish().unwrap().into_parts();
        assert_current_identity(&worker, &peer);
    }
}

#[test]
fn snapshot_shared_text_surrogate_interior_refuses_without_integration() {
    let budget = budgets()[0];
    let mut update = Writer::new();
    update.var_usize(2);
    update.var_usize(1);
    update.var_u64(2);
    update.var_u32(0);
    update.u8(yrs::block::BLOCK_ITEM_STRING_REF_NUMBER);
    update.var_u32(1);
    update.str("hostile-text");
    update.str("😀");
    update.var_usize(1);
    update.var_u64(1);
    update.var_u32(0);
    update.u8(yrs::block::HAS_ORIGIN | yrs::block::BLOCK_ITEM_STRING_REF_NUMBER);
    update.var_u64(2);
    update.var_u32(0);
    update.str("x");
    update.var_usize(0);
    let mut vector = Writer::new();
    vector.var_usize(2);
    vector.var_u64(2);
    vector.var_u32(2);
    vector.var_u64(1);
    vector.var_u32(1);
    let chunks = replace_yrs(
        &empty_worker(),
        vec![update.into_bytes()],
        vector.into_bytes(),
        budget,
    );
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in &chunks {
        builder.push(chunk).unwrap();
    }
    let mut refused = false;
    for _ in 0..10_000 {
        match builder.advance(budget) {
            Ok(progress) => assert!(!progress.is_ready()),
            Err(failure) => {
                assert!(
                    failure
                        .to_string()
                        .contains("shared Yrs text is not supported")
                );
                assert!(
                    builder
                        .authority
                        .as_ref()
                        .unwrap()
                        .snapshot_vector_for_test()
                        .is_empty()
                );
                assert_step_budget(budget);
                refused = true;
                break;
            }
        }
        assert_step_budget(budget);
    }
    assert!(refused);
    assert!(builder.finish().is_err());
}

#[test]
fn snapshot_charts_require_source_package_before_ready() {
    let budget = budgets()[0];
    let mut worker = worker(&source(false, false));
    assert!(!worker.model.sheets[0].charts.is_empty());
    worker.source_package = None;
    worker.source_container = None;
    let chunks = encode(&worker, budget);
    assert_eq!(
        assert_refuses_before_ready(&chunks, budget),
        "snapshot charts require a source package"
    );
}

#[test]
fn snapshot_export_source_helper_retains_base_contract() {
    let worker = empty_worker();
    let source: crate::structured::ExportSource<'_> = worker.export_source();
    assert!(std::ptr::eq(source.model, worker.model()));
    assert!(source.package.is_none());
    assert!(worker.try_export_source().unwrap().package.is_none());
}

#[test]
fn snapshot_transport_frame_limit_refuses_before_payload_work() {
    let budget = SnapshotBudget::new(1, 64).unwrap();
    let mut chunks = encode(&empty_worker(), budget);
    chunks[0].resize(65, 0);
    let mut builder = WorkbookSnapshotBuilder::with_max_frame_bytes(64).unwrap();
    let digest = builder.digest.clone().finalize();
    let allocations = FRAGMENT_ALLOCATED_BYTES.get();
    assert_eq!(
        builder.push(&chunks[0]).unwrap_err().to_string(),
        "snapshot transport frame exceeds accepted byte limit"
    );
    let work = crate::snapshot::step::current();
    assert_eq!(work.records, 0);
    assert_eq!(work.bytes, 0);
    assert_eq!(work.allocated_bytes, 0);
    assert_eq!(FRAGMENT_ALLOCATED_BYTES.get(), allocations);
    assert_eq!(builder.digest.clone().finalize(), digest);
    assert!(builder.fragment.is_none());
    assert!(builder.queue.is_empty());
    assert!(builder.finish().is_err());
    let worker = empty_worker();
    let chunks = encode(&worker, budgets()[0]);
    assert!(completed_builder(&chunks, budgets()[0]).finish().is_ok());
}

#[test]
fn snapshot_identity_fields_are_bounded_at_header_admission() {
    let budget = budgets()[0];
    let worker = empty_worker();
    for field in [0, 1, 2, 3] {
        let mut records = decoded_snapshot(&encode(&worker, budget));
        change_header(&mut records, |header| match field {
            0 => header.version_nonce = "a".repeat(2 * 1024 * 1024),
            1 => header.version_nonce = "z".repeat(32),
            2 => header.guid = "a".repeat(2 * 1024 * 1024),
            _ => header.guid = "z".repeat(36),
        });
        let chunks = transport_records(records, budget);
        let mut builder = WorkbookSnapshotBuilder::new();
        let mut refused = false;
        for chunk in chunks {
            match builder.push(&chunk) {
                Ok(progress) => assert!(!progress.is_ready()),
                Err(failure) => {
                    assert!(failure.to_string().contains("snapshot identity"));
                    assert!(builder.header.is_none());
                    assert!(builder.authority.is_none());
                    refused = true;
                    break;
                }
            }
        }
        assert!(refused);
        assert!(builder.finish().is_err());
    }
    let chunks = encode(&worker, budget);
    let (peer, _) = completed_builder(&chunks, budget)
        .finish()
        .unwrap()
        .into_parts();
    assert_eq!(
        decoded_snapshot(&chunks),
        decoded_snapshot(&encode(&peer, budget))
    );
}

fn facts_record(tag: u64) -> Vec<u8> {
    let mut record = Writer::new();
    record.var_u64(tag);
    match tag {
        1 => {
            record.str("xl/worksheets/sheet1.xml");
            for value in [0, 0, 1, 0] {
                record.var_u64(value);
            }
        }
        2 => {
            record.str("opaque.xml");
            record.bool(false);
        }
        3 => {
            record.str("xl/charts/chart1.xml");
            record.bytes(b"<chart/>");
        }
        _ => unreachable!(),
    }
    let mut framed = Writer::new();
    framed.bytes(&record.into_bytes());
    framed.into_bytes()
}

#[test]
fn snapshot_batched_facts_refuse_before_capacity_boundary_mutation() {
    let budget = SnapshotBudget::new(2, 16_384).unwrap();
    for tag in 1..=3 {
        let worker = worker(&source(false, false));
        let mut records = decoded_snapshot(&encode(&worker, budget));
        records.retain(|(kind, _, _)| *kind != ChunkKind::Facts);
        let mut header = Writer::new();
        header.var_u64(0);
        header.var_u64(1);
        header.bool(false);
        for kind in 1..=3 {
            header.var_usize(if kind == tag { 9 } else { 0 });
        }
        let mut framed = Writer::new();
        framed.bytes(&header.into_bytes());
        records.push((ChunkKind::Facts, 0, framed.into_bytes()));
        for _ in 0..7 {
            records.push((ChunkKind::Facts, 0, facts_record(tag)));
        }
        let mut batch = facts_record(tag);
        batch.extend(facts_record(tag));
        records.push((ChunkKind::Facts, 0, batch));
        records.sort_by_key(|(kind, _, _)| index(*kind));
        let chunks = transport_records(records, budget);
        let mut builder = WorkbookSnapshotBuilder::new();
        for chunk in &chunks {
            builder.push(chunk).unwrap();
        }
        let mut refused = false;
        for _ in 0..200_000 {
            let before = builder.facts_admission;
            match builder.advance(budget) {
                Ok(progress) => assert!(!progress.is_ready()),
                Err(failure) => {
                    assert_eq!(
                        failure.to_string(),
                        "snapshot facts chunk completes multiple records"
                    );
                    assert_eq!(builder.facts_admission.remaining, before.remaining);
                    assert_eq!(builder.facts_admission.length, before.length);
                    let work = assert_step_budget(budget);
                    assert_eq!(work.records, 0);
                    assert_eq!(work.allocated_bytes, 0);
                    refused = true;
                    break;
                }
            }
            assert_step_budget(budget);
        }
        assert!(refused);
        assert!(builder.finish().is_err());
    }
}

#[test]
fn snapshot_oversized_yrs_scan_and_allocations_respect_byte_allowance() {
    let budget = budgets()[0];
    for map in [false, true] {
        let doc = Doc::with_client_id(77);
        {
            let mut txn = doc.transact_mut();
            let root = txn.get_or_insert_map("hostile-record");
            let value = if map {
                Any::Map(
                    (0..100_000)
                        .map(|index| (index.to_string(), Any::Null))
                        .collect::<std::collections::HashMap<_, _>>()
                        .into(),
                )
            } else {
                Any::from("a".repeat(4 * 1024 * 1024))
            };
            root.insert(&mut txn, "value", value);
        }
        let update = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let chunks = replace_yrs(
            &empty_worker(),
            vec![update],
            doc.transact().state_vector().encode_v1(),
            budget,
        );
        let mut builder = WorkbookSnapshotBuilder::new();
        for chunk in &chunks {
            builder.push(chunk).unwrap();
        }
        let mut refused = false;
        for _ in 0..10_000 {
            match builder.advance(budget) {
                Ok(progress) => assert!(!progress.is_ready()),
                Err(failure) => {
                    assert_eq!(
                        failure.to_string(),
                        "Yrs snapshot record exceeds advance byte budget"
                    );
                    let work = assert_step_budget(budget);
                    assert!(work.scanned_bytes > 0);
                    assert!(work.scanned_bytes <= budget.max_bytes(), "{work:?}");
                    assert!(work.allocated_bytes <= budget.max_bytes(), "{work:?}");
                    assert!(
                        builder
                            .authority
                            .as_ref()
                            .unwrap()
                            .snapshot_vector_for_test()
                            .is_empty()
                    );
                    refused = true;
                    break;
                }
            }
            assert_step_budget(budget);
        }
        assert!(refused);
        assert!(builder.finish().is_err());
    }
}

#[test]
fn snapshot_anchor_reconstruction_visits_allocations_and_refusal_are_bounded() {
    let budget = budgets()[0];
    let mut original = worker(&source(false, false));
    let mut model = original.model.clone();
    let template = model.sheets[0].charts[0].clone();
    model.sheets[0].charts = (0..32)
        .map(|index| {
            let mut chart = template.clone();
            chart.anchor_index = index;
            chart.refs.clear();
            chart
        })
        .collect();
    let Some(PackageSlot::Present(package)) = original.source_package.take() else {
        panic!("source package is not materialized");
    };
    let mut worker = Workbook::from_parts(model, Some(package), SheetId(0), true, None).unwrap();
    worker.source_container = original.source_container.take();
    let chunks = encode(&worker, budget);
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in &chunks {
        builder.push(chunk).unwrap();
    }
    let mut anchors = 0;
    loop {
        let before = builder
            .restored
            .as_ref()
            .map_or(0, |workbook| workbook.opened_anchors.len());
        let ready = builder.advance(budget).unwrap().is_ready();
        let work = assert_step_budget(budget);
        assert!(work.allocated_bytes <= budget.max_bytes(), "{work:?}");
        let after = builder
            .restored
            .as_ref()
            .or_else(|| builder.ready.as_ref().map(|ready| &ready.workbook))
            .map_or(0, |workbook| workbook.opened_anchors.len());
        let added = after.saturating_sub(before);
        assert!(added <= budget.max_records());
        anchors += added;
        if ready {
            break;
        }
    }
    assert_eq!(anchors, 32);
    let (peer, _) = builder.finish().unwrap().into_parts();
    for chart in &worker.model.sheets[0].charts {
        assert_eq!(
            peer.opened_anchors.get(&chart.frame_id()),
            Some(&chart.anchor)
        );
    }
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in &chunks {
        builder.push(chunk).unwrap();
    }
    while !builder.authority_validated {
        crate::snapshot::step::reset();
        builder.advance_unit(budget).unwrap();
        assert_step_budget(budget);
    }
    builder.restored.as_mut().unwrap().model.sheets[0].charts[0].drawing =
        "x".repeat(budget.max_bytes() + 1);
    assert_eq!(
        builder.advance(budget).unwrap_err().to_string(),
        "snapshot chart exceeds advance byte budget"
    );
    let work = assert_step_budget(budget);
    assert_eq!(work.allocated_bytes, 0);
    assert!(builder.restored.as_ref().unwrap().opened_anchors.is_empty());
    assert!(builder.finish().is_err());
}

#[path = "measurement_tests.rs"]
mod measurements;

#[test]
fn snapshot_large_metadata_windows_reach_ready_and_refuse_a_late_contradiction() {
    for bytes in [1024, 16_384] {
        let budget = SnapshotBudget::new(1, bytes).unwrap();
        let model = WorkbookModel {
            sheets: vec![crate::Sheet::new("Data")],
            defined_names: (0..4096)
                .map(|index| xlsx_model::DefinedName {
                    name: format!("N_{index}"),
                    formula: "Data!A1".to_owned(),
                    local_sheet: None,
                    hidden: false,
                })
                .collect(),
            shared_strings: vec!["水\"\\\n".repeat(8192)],
            ..WorkbookModel::default()
        };
        let mut worker = Workbook::from_model(model).unwrap();
        let chunks = encode(&worker, budget);
        let mut builder = WorkbookSnapshotBuilder::new();
        for chunk in &chunks {
            builder.push(chunk).unwrap();
            assert_step_budget(budget);
        }
        let mut ready = false;
        for _ in 0..200_000 {
            ready = builder.advance(budget).unwrap().is_ready();
            let work = assert_step_budget(budget);
            assert!(work.allocated_bytes <= budget.max_bytes(), "{work:?}");
            assert!(work.scanned_bytes <= budget.max_bytes(), "{work:?}");
            if ready {
                break;
            }
        }
        assert!(ready);
        let (peer, _) = builder.finish().unwrap().into_parts();
        assert_current_identity(&worker, &peer);
        assert_eq!(worker.model.shared_strings, peer.model.shared_strings);
        let last = worker.model.shared_strings[0].len() - 1;
        worker.model.shared_strings[0].replace_range(last.., "x");
        let chunks = encode(&worker, budget);
        assert!(
            assert_refuses_before_ready(&chunks, budget).contains("authority and model disagree")
        );
    }
}

#[test]
fn snapshot_near_allowance_cell_content_validation_is_resumable() {
    for budget in [SnapshotBudget::new(1, 1024).unwrap(), budgets()[0]] {
        let mut model = WorkbookModel {
            sheets: vec![crate::Sheet::new("Data")],
            ..WorkbookModel::default()
        };
        model.sheets[0].set_cell(
            CellRef::new(0, 0),
            Cell {
                value: CellValue::Text {
                    value: "a".repeat(budget.max_bytes() - 384),
                },
                ..Cell::default()
            },
        );
        let worker = Workbook::from_model(model).unwrap();
        let chunks = encode(&worker, budget);
        let (builder, _) = completed_builder_with_step_budget(&chunks, budget);
        let (peer, _) = builder.finish().unwrap().into_parts();
        assert_current_identity(&worker, &peer);
    }
}
