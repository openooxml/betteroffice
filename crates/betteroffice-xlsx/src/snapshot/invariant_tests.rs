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
        builder.advance(budget).unwrap();
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
