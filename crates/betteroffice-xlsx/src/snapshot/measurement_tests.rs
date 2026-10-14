use super::*;

fn measure_frames(chunks: &[Vec<u8>], label: &str, budget: SnapshotBudget) -> Workbook {
    let mut builder = WorkbookSnapshotBuilder::new();
    let mut max_advance = std::time::Duration::ZERO;
    let mut advances = 0;
    for chunk in chunks {
        builder.push(chunk).unwrap();
        let start = std::time::Instant::now();
        let progress = builder.advance(budget).unwrap();
        advances += 1;
        max_advance = max_advance.max(start.elapsed());
        assert!(!progress.is_ready());
        assert_step_budget(budget);
    }
    loop {
        let start = std::time::Instant::now();
        let ready = builder.advance(budget).unwrap().is_ready();
        advances += 1;
        max_advance = max_advance.max(start.elapsed());
        assert_step_budget(budget);
        if ready {
            break;
        }
    }
    eprintln!(
        "{label}: max advance ms={:.3}, advances={advances}",
        max_advance.as_secs_f64() * 1_000.0
    );
    let (peer, _) = builder.finish().unwrap().into_parts();
    peer
}

fn frames_for_parts(
    worker: &Workbook,
    parts: Vec<Vec<u8>>,
    budget: SnapshotBudget,
) -> (Vec<Vec<u8>>, Vec<u8>) {
    let merged = yrs::merge_updates_v1(parts.iter()).unwrap();
    let doc = non_gc_doc(u64::from(u32::MAX));
    for part in &parts {
        doc.transact_mut()
            .apply_update(Update::decode_v1(part).unwrap())
            .unwrap();
    }
    let txn = doc.transact();
    assert!(txn.store().pending_update().is_none());
    assert!(txn.store().pending_ds().is_none());
    let state = txn.state_vector();
    assert_eq!(state, Update::decode_v1(&merged).unwrap().state_vector());
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
    let vector = writer.into_bytes();
    let chunks = replace_yrs(worker, parts, vector.clone(), budget);
    (chunks, vector)
}

fn split_delta(update: &[u8], budget: SnapshotBudget) -> Vec<Vec<u8>> {
    let mut cursor = crate::snapshot::yrs_split::UpdateCursor::default();
    let mut parts = Vec::new();
    while let Some(part) = cursor
        .next(update, budget.max_records(), budget.max_bytes() - 64)
        .unwrap()
    {
        parts.push(part.bytes);
    }
    parts
}

fn non_gc_doc(client: u64) -> Doc {
    let mut options = yrs::Options::with_client_id(yrs::block::ClientID::new(client));
    options.skip_gc = true;
    Doc::with_options(options)
}

fn alternating_sheet_order_history_update(client: u64, items: usize) -> Vec<u8> {
    let mut update = Writer::new();
    update.var_usize(2);
    for parity in 0..2 {
        update.var_usize((items - parity).div_ceil(2));
        update.var_u64(client - parity as u64);
        update.var_u32(0);
        for index in (parity..items).step_by(2) {
            let kind = if index + 1 == items {
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
                update.var_u64(client - (1 - parity) as u64);
                update.var_usize((index - 1) / 2);
            }
            update.var_u32(1);
            if index + 1 == items {
                update.u8(119);
                update.str("sheet:0");
            }
        }
    }
    update.var_usize(2);
    for parity in 0..2 {
        update.var_u64(client - parity as u64);
        update.var_usize(1);
        update.var_u32(0);
        update.var_usize((items - 1 - parity).div_ceil(2));
    }
    update.into_bytes()
}

#[test]
#[ignore]
fn snapshot_sheet_order_cap_steps_respect_budget() {
    if cfg!(debug_assertions) {
        panic!("run this measurement in release mode");
    }
    let worker = empty_worker();
    let doc = non_gc_doc(u64::from(u32::MAX));
    {
        let mut txn = doc.transact_mut();
        txn.apply_update(Update::decode_v1(&worker.encode_state_as_update_v1()).unwrap())
            .unwrap();
        txn.get_array("xlsx:sheet-order")
            .unwrap()
            .remove_range(&mut txn, 0, 1);
    }
    let baseline = doc
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let history = alternating_sheet_order_history_update(
        doc.client_id().get(),
        crate::snapshot::yrs_split::SHEET_ORDER_MAX_ITEMS - 1,
    );
    for (records, bytes) in [(1, 1024), (256, 16_384)] {
        let budget = SnapshotBudget::new(records, bytes).unwrap();
        let mut parts = split(&baseline, budget);
        parts.extend(split(&history, budget));
        let (chunks, vector) = frames_for_parts(&worker, parts, budget);
        let peer = measure_frames(&chunks, "sheet order at the item and clock cap", budget);
        assert_eq!(peer.encode_state_vector_v1(), vector);
        assert_eq!(
            canonical_model(worker.model()),
            canonical_model(peer.model())
        );
    }
}

#[test]
#[ignore]
fn snapshot_live_subtree_replacement_steps_respect_budget() {
    if cfg!(debug_assertions) {
        panic!("run this measurement in release mode");
    }
    let budget = budgets()[0];
    let worker = empty_worker();
    let doc = non_gc_doc(1);
    {
        let mut txn = doc.transact_mut();
        let root = txn.get_or_insert_map("snapshot-live-subtree");
        let subtree = root.insert(&mut txn, "value", MapPrelim::default());
        for index in 0..50_000 {
            subtree.insert(&mut txn, index.to_string(), index as i64);
        }
    }
    assert!(doc.skip_gc());
    let initial = yrs::merge_updates_v1([
        worker.encode_state_as_update_v1(),
        doc.transact()
            .encode_state_as_update_v1(&StateVector::default()),
    ])
    .unwrap();
    let writer = non_gc_doc(2);
    {
        let mut txn = writer.transact_mut();
        txn.get_or_insert_map("snapshot-live-subtree")
            .insert(&mut txn, "value", 2_i64);
    }
    let replacement = writer
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let mut parts = split(&initial, budget);
    parts.extend(split(&replacement, budget));
    let (chunks, vector) = frames_for_parts(&worker, parts, budget);
    let peer = measure_frames(
        &chunks,
        "(a) live subtree replacement with 50000 leaves",
        budget,
    );
    assert_eq!(peer.encode_state_vector_v1(), vector);
    let txn = peer.authority.snapshot_transaction_for_test();
    assert_eq!(
        txn.get_map("snapshot-live-subtree")
            .unwrap()
            .get(&txn, "value"),
        Some(yrs::Out::Any(Any::from(2_i64)))
    );
    assert_eq!(
        canonical_model(worker.model()),
        canonical_model(peer.model())
    );
}

#[test]
#[ignore]
fn snapshot_deep_shared_maps_steps_respect_budget() {
    if cfg!(debug_assertions) {
        panic!("run this measurement in release mode");
    }
    std::thread::Builder::new()
        .stack_size(64 * 1024 * 1024)
        .spawn(|| {
            let budget = budgets()[0];
            let mut worker = empty_worker();
            worker.authority.snapshot_skip_gc_for_test();
            {
                let mut txn = worker.authority.snapshot_transaction_for_test();
                let mut parent = txn.get_or_insert_map("snapshot-deep-maps");
                for _ in 0..10_000 {
                    parent = parent.insert(&mut txn, "child", MapPrelim::default());
                }
                parent.insert(&mut txn, "value", 1_i64);
            }
            let chunks = encode(&worker, budget);
            let peer = measure_frames(&chunks, "(b) 10000 nested shared maps", budget);
            assert_current_identity(&worker, &peer);
            let txn = peer.authority.snapshot_transaction_for_test();
            let mut parent = txn.get_map("snapshot-deep-maps").unwrap();
            for _ in 0..10_000 {
                parent = parent
                    .get(&txn, "child")
                    .unwrap()
                    .cast::<yrs::MapRef>()
                    .unwrap();
            }
            assert_eq!(
                parent.get(&txn, "value"),
                Some(yrs::Out::Any(Any::from(1_i64)))
            );
        })
        .unwrap()
        .join()
        .unwrap();
}

#[test]
#[ignore]
fn snapshot_accumulated_any_append_and_interior_delete_steps_respect_budget() {
    if cfg!(debug_assertions) {
        panic!("run this measurement in release mode");
    }
    let budget = budgets()[0];
    let worker = empty_worker();
    let doc = non_gc_doc(1);
    let mut parts = split(&worker.encode_state_as_update_v1(), budget);
    let array = doc.get_or_insert_array("snapshot-accumulated-any");
    let mut vector = StateVector::default();
    for index in 0..20_000 {
        {
            let mut txn = doc.transact_mut();
            array.insert(&mut txn, index, index as i64);
        }
        let txn = doc.transact();
        parts.extend(split_delta(&txn.encode_state_as_update_v1(&vector), budget));
        vector = txn.state_vector();
    }
    let (chunks, expected) = frames_for_parts(&worker, parts.clone(), budget);
    let peer = measure_frames(&chunks, "(e) 20000 accumulated Any appends", budget);
    assert_eq!(peer.encode_state_vector_v1(), expected);
    {
        let mut txn = doc.transact_mut();
        array.remove_range(&mut txn, 10_000, 1);
    }
    parts.extend(split_delta(
        &doc.transact().encode_state_as_update_v1(&vector),
        budget,
    ));
    let (chunks, expected) = frames_for_parts(&worker, parts, budget);
    let peer = measure_frames(
        &chunks,
        "(f) interior deletion in accumulated Any block",
        budget,
    );
    assert_eq!(peer.encode_state_vector_v1(), expected);
    let txn = peer.authority.snapshot_transaction_for_test();
    assert_eq!(
        txn.get_array("snapshot-accumulated-any").unwrap().len(&txn),
        19_999
    );
}

#[test]
#[ignore]
fn snapshot_increasing_client_sequence_conflicts_steps_respect_budget() {
    if cfg!(debug_assertions) {
        panic!("run this measurement in release mode");
    }
    let budget = budgets()[0];
    let worker = empty_worker();
    let doc = non_gc_doc(50_000);
    {
        let mut txn = doc.transact_mut();
        txn.get_or_insert_array("snapshot-sequence-conflicts")
            .insert(&mut txn, 0, 0_i64);
    }
    let baseline = doc
        .transact()
        .encode_state_as_update_v1(&StateVector::default());
    let vector = doc.transact().state_vector();
    let initial =
        yrs::merge_updates_v1([worker.encode_state_as_update_v1(), baseline.clone()]).unwrap();
    let mut parts = split(&initial, budget);
    for client in 1..=4_000 {
        let writer = non_gc_doc(client);
        {
            let mut txn = writer.transact_mut();
            txn.apply_update(Update::decode_v1(&baseline).unwrap())
                .unwrap();
            txn.get_array("snapshot-sequence-conflicts")
                .unwrap()
                .insert(&mut txn, 1, client as i64);
        }
        parts.extend(split_delta(
            &writer.transact().encode_state_as_update_v1(&vector),
            budget,
        ));
    }
    let (chunks, expected) = frames_for_parts(&worker, parts, budget);
    let peer = measure_frames(
        &chunks,
        "4000 increasing-client sequence insertions",
        budget,
    );
    assert_eq!(peer.encode_state_vector_v1(), expected);
    let txn = peer.authority.snapshot_transaction_for_test();
    assert_eq!(
        txn.get_array("snapshot-sequence-conflicts")
            .unwrap()
            .len(&txn),
        4_001
    );
}

fn million_cell_worker() -> Workbook {
    let mut sheet = crate::Sheet::new("Large");
    for row in 0..50_000 {
        for col in 0..20 {
            sheet.set_cell(
                CellRef::new(row, col),
                Cell {
                    value: CellValue::Number {
                        value: f64::from(row * 20 + col),
                    },
                    ..Cell::default()
                },
            );
        }
    }
    Workbook::from_model(WorkbookModel {
        sheets: vec![sheet],
        ..WorkbookModel::default()
    })
    .unwrap()
}

fn mid_hydration(
    chunks: &[Vec<u8>],
    budget: SnapshotBudget,
) -> (WorkbookSnapshotBuilder, std::time::Duration) {
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in chunks {
        builder.push(chunk).unwrap();
    }
    let mut max_advance = std::time::Duration::ZERO;
    loop {
        let start = std::time::Instant::now();
        assert!(!builder.advance(budget).unwrap().is_ready());
        max_advance = max_advance.max(start.elapsed());
        assert_step_budget(budget);
        let clocks = builder.authority.as_ref().map_or(0, |authority| {
            authority
                .snapshot_vector_for_test()
                .iter()
                .map(|(_, &clock)| u64::from(clock))
                .sum::<u64>()
        });
        if clocks >= 500_000 {
            break;
        }
    }
    assert!(builder.restored.is_none());
    (builder, max_advance)
}

#[test]
#[ignore]
fn snapshot_refusal_teardown_steps_respect_budget() {
    if cfg!(debug_assertions) {
        panic!("run this measurement in release mode");
    }
    let budget = SnapshotBudget::new(256, 16_384).unwrap();
    let worker = million_cell_worker();
    assert_eq!(worker.model.sheets[0].iter_cells().count(), 1_000_000);
    let chunks = encode(&worker, budget);
    let mut records = decoded_snapshot(&chunks);
    let mut reader = Reader::new(&records[0].2);
    let header = reader.bytes().unwrap().to_vec();
    let flags = [
        reader.bool().unwrap(),
        reader.bool().unwrap(),
        reader.bool().unwrap(),
    ];
    let source_length = reader.option(Reader::var_usize).unwrap();
    let mut styles = [0; 7];
    for count in &mut styles {
        *count = reader.var_usize().unwrap();
    }
    styles[0] += 1;
    let mut writer = Writer::new();
    writer.bytes(&header);
    for flag in flags {
        writer.bool(flag);
    }
    writer.option(source_length, Writer::var_usize);
    for count in styles {
        writer.var_usize(count);
    }
    writer.raw(reader.rest());
    records[0].2 = writer.into_bytes();
    let refused_chunks = transport_records(records, budget);
    let mut builder = WorkbookSnapshotBuilder::new();
    for chunk in &refused_chunks {
        builder.push(chunk).unwrap();
    }
    let mut max_advance = std::time::Duration::ZERO;
    let refusal;
    loop {
        let start = std::time::Instant::now();
        let result = builder.advance(budget);
        let elapsed = start.elapsed();
        max_advance = max_advance.max(elapsed);
        assert_step_budget(budget);
        match result {
            Ok(progress) => assert!(!progress.is_ready()),
            Err(failure) => {
                assert_eq!(failure.to_string(), "snapshot model header differs");
                refusal = elapsed;
                break;
            }
        }
    }
    let start = std::time::Instant::now();
    drop(builder);
    let refused_drop = start.elapsed();
    eprintln!(
        "B7 late completion refusal: max advance ms={:.3}, refusal ms={:.3}, refused drop ms={:.3}",
        max_advance.max(refused_drop).as_secs_f64() * 1_000.0,
        refusal.as_secs_f64() * 1_000.0,
        refused_drop.as_secs_f64() * 1_000.0
    );
    drop(refused_chunks);
    for finish in [false, true] {
        let (builder, max_advance) = mid_hydration(&chunks, budget);
        let start = std::time::Instant::now();
        if finish {
            assert!(builder.finish().is_err());
        } else {
            drop(builder);
        }
        let abandon = start.elapsed();
        eprintln!(
            "B7 incomplete {}: max advance ms={:.3}, abandonment ms={:.3}",
            if finish { "finish" } else { "drop" },
            max_advance.max(abandon).as_secs_f64() * 1_000.0,
            abandon.as_secs_f64() * 1_000.0
        );
    }
}
