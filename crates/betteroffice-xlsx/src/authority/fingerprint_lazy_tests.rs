use super::*;

fn cases() -> Vec<open_test_cases::OpenCase> {
    let mut cases = open_test_cases::matrix();
    let mut plain = cases.remove(0);
    plain.label = "plain".into();
    let mut combined = cases.pop().unwrap();
    combined.label = "combined".into();
    combined.model.styles.fills[1] = xlsx_model::Fill::Pattern {
        pattern: xlsx_model::PatternType::LightUp,
        fg: Some(xlsx_model::Color::Rgb("#123456".into())),
        bg: Some(xlsx_model::Color::Rgb("#fedcba".into())),
    };
    combined.model.styles.cell_xfs[1].quote_prefix = true;
    combined.model.styles.cell_xfs[1].xf_id = Some(2);
    combined.model.sheets[0].row_styles.insert(3, 1);
    combined.model.styles.fonts[0].size_pt = Some(-0.0);
    let mut imprecise = open_test_cases::OpenCase {
        label: "imprecise".into(),
        model: combined.model.clone(),
        legacy_dimensions: combined.legacy_dimensions.clone(),
        legacy_styles: combined.legacy_styles.clone(),
    };
    imprecise.model.full_precision = false;
    vec![plain, combined, imprecise]
}

fn eager_base(case: &open_test_cases::OpenCase, readings: &[Stylesheet]) -> WorkbookBase {
    let mut base = WorkbookBase::from_model_with_legacy_dimensions_oracle(
        &case.model,
        &case.legacy_dimensions,
        case.legacy_styles.as_ref(),
    )
    .unwrap();
    for reading in readings {
        let mut model = case.model.clone();
        model.styles = reading.clone();
        let projected = WorkbookBase::from_model_with_legacy_dimensions_oracle(
            &model,
            &case.legacy_dimensions,
            None,
        )
        .unwrap();
        for (version, accepted) in projected.fingerprints {
            for fingerprint in accepted {
                append_fingerprint(&mut base.fingerprints, version, fingerprint);
            }
        }
    }
    base
}

fn fingerprint_bytes(value: &impl snapshot::Codec) -> Vec<u8> {
    let mut writer = crate::snapshot::wire::Writer::new();
    value.write(&mut writer);
    writer.into_bytes()
}

fn fingerprints_with_budget(
    case: &open_test_cases::OpenCase,
    readings: &[Stylesheet],
    limit: usize,
) -> (String, u64, Fingerprints) {
    let mut builder = FingerprintBuilder::new(&case.model, &case.legacy_dimensions);
    let accepted = builder
        .plan(|builder| {
            let styles = FingerprintStyles::new(&case.model.styles);
            builder.current(if released_styles(&case.model.styles) {
                styles
            } else {
                styles.with_written_fills()
            })?;
            let mut accepted =
                builder.style_fingerprints(&case.model.styles, case.legacy_styles.as_ref())?;
            for reading in readings {
                for (version, values) in builder.style_fingerprints(reading, None)? {
                    for value in values {
                        append_fingerprint(&mut accepted, version, value);
                    }
                }
            }
            Ok(accepted)
        })
        .unwrap();
    builder.finish_with_budget(accepted, limit).unwrap()
}

fn encoded_reading(
    model: &WorkbookModel,
    base: &WorkbookBase,
    version: i64,
    reading: FingerprintReading,
    rows: bool,
) -> Vec<u8> {
    let mut model = model.clone();
    if !reading.has_facets() {
        model.styles = pre_facet_styles(&model.styles);
    }
    if !reading.as_written() {
        model.styles.fills = model
            .styles
            .fills
            .iter()
            .map(|fill| released_fill(fill).into_owned())
            .collect();
        model.styles.borders = model
            .styles
            .borders
            .iter()
            .map(|border| released_border(border).into_owned())
            .collect();
    }
    if !rows {
        for sheet in &mut model.sheets {
            sheet.row_styles.clear();
        }
    }
    let keys = (0..model.sheets.len())
        .map(|index| format!("sheet:{index}"))
        .collect::<Vec<_>>();
    let doc = Doc::with_client_id(base.bootstrap_client_id);
    let entry: FormatEntryWriter = if reading.has_facets() {
        style_entry
    } else {
        |styles, index| cell_format_entry(&styles.cell_format(Some(index)))
    };
    seed_with_format_entries(&doc, base, &model, &keys, entry).unwrap();
    {
        let mut txn = doc.transact_mut();
        let meta = txn.get_map(META).unwrap();
        meta.insert(&mut txn, "schemaVersion", version);
        let sheets = txn.get_map(SHEETS).unwrap();
        for key in &keys {
            let sheet = sheets.get(&txn, key).unwrap().cast::<MapRef>().unwrap();
            for (introduced, name) in [
                (FREEZE_PANE_SCHEMA_VERSION, FREEZE_PANE),
                (HYPERLINK_SCHEMA_VERSION, HYPERLINKS),
                (CHARTS_SCHEMA_VERSION, CHARTS),
                (CHARTS_SCHEMA_VERSION, COL_STYLES),
            ] {
                if version < introduced {
                    sheet.remove(&mut txn, name);
                }
            }
            if !rows {
                sheet.remove(&mut txn, ROW_STYLES);
            }
        }
    }
    doc.transact()
        .encode_state_as_update_v1(&StateVector::default())
}

#[test]
fn historical_state_encodings_match_eager_acceptance() {
    for case in cases() {
        let reading = FingerprintReading::of(true, !released_styles(&case.model.styles));
        let defaults = has_style_defaults(&case.model);
        let mut encodings = (MIN_SUPPORTED_SCHEMA_VERSION..6)
            .map(|version| ("released", version, FingerprintReading::Released, false))
            .collect::<Vec<_>>();
        encodings.extend([
            ("0.4.3-era", 6, FingerprintReading::Released, false),
            ("pre-#1312", 6, reading.legacy().unwrap(), false),
            (
                "current",
                SCHEMA_VERSION,
                reading.with_defaults(defaults),
                defaults,
            ),
        ]);
        for (era, version, reading, rows) in encodings {
            let (actual, _, _) = WorkbookAuthority::from_source_with_projection(
                &case.model,
                Some(71),
                &case.legacy_dimensions,
                case.legacy_styles.as_ref(),
                &[],
            )
            .unwrap();
            let expected = WorkbookAuthority::from_source_oracle(
                &case.model,
                Some(71),
                &case.legacy_dimensions,
                case.legacy_styles.as_ref(),
                &[],
            )
            .unwrap();
            let domain = fingerprint_domain(version, reading);
            let (fingerprint, client) = fingerprint_model_with_row_domain_oracle(
                &case.model,
                version,
                version >= 4,
                true,
                rows,
                domain,
            )
            .unwrap();
            if version == SCHEMA_VERSION {
                let accepted = if era == "current" {
                    WorkbookBase::current_fingerprints_oracle(&case.model, None).unwrap()
                } else {
                    WorkbookBase::reading_fingerprints_oracle(&case.model, None, reading).unwrap()
                };
                assert!(accepted.contains(&fingerprint), "{}: {era}", case.label);
            }
            assert!(expected.base.accepts_fingerprint(version, &fingerprint));
            for schema in (MIN_SUPPORTED_SCHEMA_VERSION - 1)..=(SCHEMA_VERSION + 1) {
                assert_eq!(
                    actual.base.accepts_fingerprint(schema, &fingerprint),
                    expected.base.accepts_fingerprint(schema, &fingerprint),
                    "{}: {era} {schema}",
                    case.label,
                );
            }
            let mut base = expected.base.as_ref().clone();
            base.fingerprint = fingerprint;
            base.bootstrap_client_id = client;
            let update = encoded_reading(&case.model, &base, version, reading, rows);
            let adopted =
                |authority: &WorkbookAuthority| match authority.snapshot_replacement(&update) {
                    SnapshotAdoption::Replacement(adopted) => adopted,
                    SnapshotAdoption::Incompatible(error) => {
                        panic!("{}: {era}: {error}", case.label)
                    }
                    SnapshotAdoption::NotApplicable => {
                        panic!("{}: {era}: not adopted", case.label)
                    }
                };
            let actual_adopted = adopted(&actual);
            let expected_adopted = adopted(&expected);
            assert_eq!(
                actual_adopted.strict_materialize().unwrap(),
                expected_adopted.strict_materialize().unwrap(),
                "{}: {era}",
                case.label,
            );
            let mut changed = case.model.clone();
            changed.sheets[0].name.push('!');
            (base.fingerprint, base.bootstrap_client_id) =
                fingerprint_model_with_row_domain_oracle(
                    &changed,
                    version,
                    version >= 4,
                    true,
                    rows,
                    domain,
                )
                .unwrap();
            let rejected = encoded_reading(&changed, &base, version, reading, rows);
            for authority in [&expected, &actual] {
                assert!(
                    !authority
                        .base
                        .accepts_fingerprint(version, &base.fingerprint)
                );
                assert!(matches!(
                    authority.snapshot_replacement(&rejected),
                    SnapshotAdoption::Incompatible(_)
                ));
            }
            let adopted_fingerprints = actual_adopted.base.fingerprints.clone();
            let expected_map = expected
                .base
                .fingerprints
                .clone()
                .into_iter()
                .collect::<BTreeMap<_, _>>();
            assert_eq!(
                adopted_fingerprints
                    .clone()
                    .into_iter()
                    .collect::<BTreeMap<_, _>>(),
                expected_map,
            );
            assert_eq!(
                fingerprint_bytes(&actual_adopted.base.fingerprints),
                fingerprint_bytes(&expected_map),
                "{}: {era}",
                case.label,
            );
            assert_eq!(&*actual_adopted.base.fingerprints, &expected_map);
            let restored = WorkbookAuthority::restore_recovery_state(
                &actual_adopted.encode_recovery_state().unwrap(),
                72,
                0,
                true,
            )
            .unwrap();
            for mut fingerprints in [adopted_fingerprints, restored.base.fingerprints.clone()] {
                fingerprints.retain(|version, _| *version <= SCHEMA_VERSION);
                append_fingerprint(
                    &mut fingerprints,
                    version,
                    expected_map[&version][0].clone(),
                );
                assert_eq!(
                    fingerprints.into_iter().collect::<BTreeMap<_, _>>(),
                    expected_map,
                );
            }
        }
    }
}

#[test]
fn construction_paths_materialize_before_mutation_and_iteration() {
    use snapshot::Codec;

    let case = cases().remove(1);
    let planned = || {
        WorkbookBase::from_model_with_style_readings(
            &case.model,
            &case.legacy_dimensions,
            case.legacy_styles.as_ref(),
            &[],
        )
        .unwrap()
        .fingerprints
    };
    let expected = eager_base(&case, &[])
        .fingerprints
        .into_iter()
        .collect::<BTreeMap<_, _>>();
    let mut duplicated = expected.clone();
    let accepted = duplicated.get_mut(&SCHEMA_VERSION).unwrap();
    accepted.push(accepted[0].clone());
    let bytes = fingerprint_bytes(&duplicated);
    let mut reader = crate::snapshot::wire::Reader::new(&bytes);
    let restored = Fingerprints::read(&mut reader).unwrap();
    reader.finish().unwrap();
    let lazy = planned();
    let materialized = planned();
    assert_eq!(&*materialized, &expected);
    let ready = Fingerprints::from(duplicated.clone());
    let mut paths = vec![
        ("default", Fingerprints::default(), BTreeMap::new()),
        ("empty-map", BTreeMap::new().into(), BTreeMap::new()),
        ("map", ready.clone(), duplicated.clone()),
        ("persisted-clone", restored.clone(), duplicated.clone()),
        ("persisted", restored, duplicated.clone()),
        ("planned", planned(), expected.clone()),
        ("planned-clone", lazy.clone(), expected.clone()),
        ("materialized-clone", materialized.clone(), expected.clone()),
    ];
    for (label, mut left, right, expected) in [
        (
            "empty-planned",
            Fingerprints::default(),
            lazy.clone(),
            expected.clone(),
        ),
        (
            "planned-planned",
            lazy.clone(),
            lazy.clone(),
            expected.clone(),
        ),
        (
            "planned-ready",
            lazy.clone(),
            ready.clone(),
            expected.clone(),
        ),
        ("ready-planned", ready, lazy.clone(), duplicated),
        ("materialized-planned", materialized, lazy, expected),
    ] {
        left.merge(right);
        paths.push((label, left, expected));
    }
    for (label, mut actual, mut expected) in paths {
        assert_eq!(
            actual.clone().into_iter().collect::<BTreeMap<_, _>>(),
            expected,
            "{label}",
        );
        let mut appended = actual.clone();
        let mut expected_appended = expected.clone();
        append_fingerprint(&mut appended, SCHEMA_VERSION, case.label.clone());
        append_fingerprint(&mut expected_appended, SCHEMA_VERSION, case.label.clone());
        assert_eq!(
            appended.materialized_bytes(),
            Some(appended.retained_bytes()),
            "{label}"
        );
        assert_eq!(
            appended.into_iter().collect::<BTreeMap<_, _>>(),
            expected_appended,
            "{label}",
        );
        let removed = expected.get(&SCHEMA_VERSION).cloned().unwrap_or_default();
        actual.retain(|version, _| *version < SCHEMA_VERSION);
        expected.retain(|version, _| *version < SCHEMA_VERSION);
        assert_eq!(
            actual.materialized_bytes(),
            Some(actual.retained_bytes()),
            "{label}"
        );
        assert!(
            removed
                .iter()
                .all(|value| !actual.accepts(SCHEMA_VERSION, value)),
            "{label}",
        );
        for _ in 0..2 {
            append_fingerprint(&mut actual, SCHEMA_VERSION, case.label.clone());
            append_fingerprint(&mut expected, SCHEMA_VERSION, case.label.clone());
        }
        actual.merge(Fingerprints::default());
        assert_eq!(
            fingerprint_bytes(&actual),
            fingerprint_bytes(&expected),
            "{label}",
        );
        assert_eq!(
            actual.into_iter().collect::<BTreeMap<_, _>>(),
            expected,
            "{label}",
        );
    }
}

#[test]
fn accepted_fingerprint_matrix_matches_released_and_current_states() {
    for case in cases() {
        let mut reading = case.model.styles.clone();
        reading.cell_xfs[1].num_fmt_id = None;
        reading.cell_xfs[1].font = None;
        for readings in [Vec::new(), vec![reading.clone(), reading.clone()]] {
            let expected = eager_base(&case, &readings);
            let actual = WorkbookBase::from_model_with_style_readings(
                &case.model,
                &case.legacy_dimensions,
                case.legacy_styles.as_ref(),
                &readings,
            )
            .unwrap();
            for accepted in expected.fingerprints.values() {
                for fingerprint in accepted {
                    for version in (MIN_SUPPORTED_SCHEMA_VERSION - 1)..=(SCHEMA_VERSION + 1) {
                        assert_eq!(
                            actual.accepts_fingerprint(version, fingerprint),
                            expected.accepts_fingerprint(version, fingerprint),
                            "{}: {version} {fingerprint}",
                            case.label,
                        );
                    }
                }
            }
            let mut changed = case.model.clone();
            changed.sheets[0].name.push('!');
            let rejected = fingerprint_model_current_oracle(&changed).unwrap().0;
            for version in MIN_SUPPORTED_SCHEMA_VERSION..=SCHEMA_VERSION {
                assert!(!expected.accepts_fingerprint(version, &rejected));
                assert!(!actual.accepts_fingerprint(version, &rejected));
                assert!(!actual.accepts_fingerprint(version, "invalid"));
            }
        }
    }
}

#[test]
fn incoming_state_matrix_matches_eager_validation() {
    let case = cases().remove(1);
    let expected = eager_base(&case, &[]);
    let actual = WorkbookBase::from_model_with_style_readings(
        &case.model,
        &case.legacy_dimensions,
        case.legacy_styles.as_ref(),
        &[],
    )
    .unwrap();
    let keys = (0..case.model.sheets.len())
        .map(|index| format!("sheet:{index}"))
        .collect::<Vec<_>>();
    for (&version, accepted) in expected.fingerprints.iter() {
        for fingerprint in accepted {
            let source = Doc::with_client_id(fingerprint_client_id(fingerprint).unwrap());
            seed(&source, &expected, &case.model, &keys).unwrap();
            {
                let mut txn = source.transact_mut();
                let meta = txn.get_map(META).unwrap();
                meta.insert(&mut txn, BASE_FINGERPRINT, fingerprint.as_str());
                meta.insert(&mut txn, "schemaVersion", version);
                let sheets = txn.get_map(SHEETS).unwrap();
                for key in &keys {
                    let sheet = sheets.get(&txn, key).unwrap().cast::<MapRef>().unwrap();
                    for (introduced, name) in [
                        (FREEZE_PANE_SCHEMA_VERSION, FREEZE_PANE),
                        (HYPERLINK_SCHEMA_VERSION, HYPERLINKS),
                        (CHARTS_SCHEMA_VERSION, CHARTS),
                        (CHARTS_SCHEMA_VERSION, COL_STYLES),
                        (CHARTS_SCHEMA_VERSION, ROW_STYLES),
                    ] {
                        if version < introduced {
                            sheet.remove(&mut txn, name);
                        }
                    }
                }
            }
            let update = source
                .transact()
                .encode_state_as_update_v1(&StateVector::default());
            let target = Doc::with_client_id(71);
            hydrate_doc(&target, &update).unwrap();
            let authority = WorkbookAuthority::hydrated(target, Arc::new(actual.clone()), 0);
            let before = WorkbookAuthority::hydrated(source, Arc::new(expected.clone()), 0)
                .strict_materialize()
                .unwrap();
            assert_eq!(
                authority.strict_materialize().unwrap(),
                before,
                "{version} {fingerprint}"
            );
            {
                let mut txn = authority.doc.transact_mut();
                let meta = txn.get_map(META).unwrap();
                meta.insert(&mut txn, BASE_FINGERPRINT, "unrelated-workbook");
            }
            assert!(authority.strict_materialize().is_err());
        }
    }
}

#[test]
fn persisted_base_bytes_match_the_eager_oracle() {
    use snapshot::Codec;

    for case in cases() {
        let mut reading = case.model.styles.clone();
        reading.cell_xfs[1].num_fmt_id = None;
        let readings = [reading];
        let (actual, _, _) = WorkbookAuthority::from_source_with_projection(
            &case.model,
            Some(71),
            &case.legacy_dimensions,
            case.legacy_styles.as_ref(),
            &readings,
        )
        .unwrap();
        let expected = WorkbookAuthority::from_source_oracle(
            &case.model,
            Some(71),
            &case.legacy_dimensions,
            case.legacy_styles.as_ref(),
            &readings,
        )
        .unwrap();
        let mut before = crate::snapshot::wire::Writer::new();
        let mut after = crate::snapshot::wire::Writer::new();
        expected.base.fingerprints.write(&mut before);
        actual.base.fingerprints.write(&mut after);
        assert_eq!(before.into_bytes(), after.into_bytes(), "{}", case.label);
        assert_eq!(
            expected.encode_recovery_state().unwrap(),
            actual.encode_recovery_state().unwrap(),
            "{}",
            case.label,
        );
    }
}

#[test]
fn opening_hashes_only_the_canonical_key_and_memoizes_validation() {
    use xlsx_model::open_timing;

    let case = cases().remove(1);
    let expected = eager_base(&case, &[]);
    let fingerprint = expected.fingerprints[&3].last().unwrap().clone();
    open_timing::begin();
    let (authority, _, _) = WorkbookAuthority::from_source_with_projection(
        &case.model,
        Some(71),
        &case.legacy_dimensions,
        case.legacy_styles.as_ref(),
        &[],
    )
    .unwrap();
    let opened = open_timing::finish();
    assert_eq!(opened.counts["fingerprint_planned_keys"], 1);
    assert!(opened.counts["fingerprint_dimension_cols"] > 0);
    let cloned = authority.base.fingerprints.clone();
    open_timing::begin();
    assert!(
        authority
            .base
            .accepts_fingerprint(SCHEMA_VERSION, &authority.base.fingerprint)
    );
    assert_eq!(
        authority.seed_clients(&authority.doc.transact().state_vector()),
        vec![authority.base.bootstrap_client_id],
    );
    assert!(
        !open_timing::finish()
            .counts
            .contains_key("fingerprint_planned_keys")
    );
    open_timing::begin();
    assert!(authority.base.accepts_fingerprint(3, &fingerprint));
    assert!(open_timing::finish().counts["fingerprint_planned_keys"] > 0);
    open_timing::begin();
    assert!(authority.base.accepts_fingerprint(3, &fingerprint));
    assert!(cloned.accepts(3, &fingerprint));
    assert!(
        !open_timing::finish()
            .counts
            .contains_key("fingerprint_planned_keys")
    );
    open_timing::begin();
    assert_eq!(authority.base.fingerprints, expected.fingerprints);
    let completed = open_timing::finish();
    assert!(completed.counts["fingerprint_planned_keys"] > 0);
    assert!(!completed.counts.contains_key("fingerprint_dimension_cols"));
    open_timing::begin();
    assert_eq!(cloned, expected.fingerprints);
    assert!(
        !open_timing::finish()
            .counts
            .contains_key("fingerprint_planned_keys")
    );
}

#[test]
fn merging_accepted_sets_defers_hashing_and_preserves_order() {
    use xlsx_model::open_timing;

    let mut case = cases().remove(1);
    let mut expected = eager_base(&case, &[]).fingerprints;
    let mut actual = WorkbookBase::from_model_with_style_readings(
        &case.model,
        &case.legacy_dimensions,
        case.legacy_styles.as_ref(),
        &[],
    )
    .unwrap()
    .fingerprints;
    case.model.styles.cell_xfs[1].quote_prefix = false;
    let next_expected = eager_base(&case, &[]);
    for (version, accepted) in next_expected.fingerprints {
        for fingerprint in accepted {
            append_fingerprint(&mut expected, version, fingerprint);
        }
    }
    let next = WorkbookBase::from_model_with_style_readings(
        &case.model,
        &case.legacy_dimensions,
        case.legacy_styles.as_ref(),
        &[],
    )
    .unwrap();
    open_timing::begin();
    actual.merge(next.fingerprints);
    assert!(actual.accepts(SCHEMA_VERSION, &next.fingerprint));
    assert!(
        !open_timing::finish()
            .counts
            .contains_key("fingerprint_planned_keys")
    );
    assert_eq!(actual, expected);
}

#[test]
fn bounded_stream_matches_eager_acceptance_and_persisted_bytes() {
    use xlsx_model::open_timing;

    for case in cases() {
        let mut reading = case.model.styles.clone();
        reading.cell_xfs[1].num_fmt_id = None;
        for readings in [Vec::new(), vec![reading.clone(), reading.clone()]] {
            let expected = eager_base(&case, &readings);
            for limit in [0, 32, 128] {
                open_timing::begin();
                let (fingerprint, client, actual) =
                    fingerprints_with_budget(&case, &readings, limit);
                let timing = open_timing::finish();
                assert_eq!(timing.counts["fingerprint_payload_fallbacks"], 1);
                assert_eq!(fingerprint, expected.fingerprint);
                assert_eq!(client, expected.bootstrap_client_id);
                assert_eq!(actual.materialized_bytes(), Some(actual.retained_bytes()));
                for accepted in expected.fingerprints.values() {
                    for value in accepted {
                        for version in (MIN_SUPPORTED_SCHEMA_VERSION - 1)..=(SCHEMA_VERSION + 1) {
                            assert_eq!(
                                actual.accepts(version, value),
                                expected.accepts_fingerprint(version, value)
                            );
                        }
                        assert!(actual.contains_client(fingerprint_client_id(value).unwrap()));
                    }
                }
                assert!(!actual.accepts(SCHEMA_VERSION, "invalid"));
                assert_eq!(
                    fingerprint_bytes(&actual),
                    fingerprint_bytes(&expected.fingerprints)
                );
            }
        }
    }
}

#[test]
fn oversized_cell_payload_retains_only_materialized_fingerprints() {
    use xlsx_model::open_timing;

    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Payload");
    for row in 0..=(fingerprints::MAX_LAZY_BYTES / 32_767) as u32 {
        sheet.set_cell(
            CellRef::new(row, 0),
            Cell {
                value: CellValue::Text {
                    value: "x".repeat(32_767),
                },
                formula: Some(format!("\"{}\"", "y".repeat(257))),
                ..Cell::default()
            },
        );
    }
    model.sheets.push(sheet);
    model.sheets.push(Sheet::new("Tail"));
    open_timing::begin();
    let actual = WorkbookBase::from_model(&model).unwrap();
    assert_eq!(
        open_timing::finish().counts["fingerprint_payload_fallbacks"],
        1
    );
    let before = actual.fingerprints.retained_bytes();
    assert!(before < 64 * 1024, "{before}");
    assert_eq!(actual.fingerprints.materialized_bytes(), Some(before));
    let expected =
        WorkbookBase::from_model_with_legacy_dimensions_oracle(&model, &[], None).unwrap();
    assert_eq!(actual.fingerprint, expected.fingerprint);
    assert_eq!(actual.bootstrap_client_id, expected.bootstrap_client_id);
    assert_eq!(
        fingerprint_bytes(&actual.fingerprints),
        fingerprint_bytes(&expected.fingerprints)
    );
    let after = actual.fingerprints.retained_bytes();
    assert_eq!(before, after);
    assert_eq!(actual.fingerprints.materialized_bytes(), Some(after));
}

#[test]
fn oversized_shared_strings_do_not_survive_in_a_plan() {
    use xlsx_model::open_timing;

    let model = WorkbookModel {
        shared_strings: (0..=fingerprints::MAX_LAZY_BYTES / 32_767)
            .map(|_| "x".repeat(32_767))
            .collect(),
        sheets: vec![Sheet::new("Payload")],
        ..WorkbookModel::default()
    };
    open_timing::begin();
    let actual = WorkbookBase::from_model(&model).unwrap();
    assert_eq!(
        open_timing::finish().counts["fingerprint_payload_fallbacks"],
        1
    );
    let retained = actual.fingerprints.retained_bytes();
    assert!(retained < 64 * 1024, "{retained}");
    assert_eq!(actual.fingerprints.materialized_bytes(), Some(retained));
    let expected =
        WorkbookBase::from_model_with_legacy_dimensions_oracle(&model, &[], None).unwrap();
    assert_eq!(actual.fingerprint, expected.fingerprint);
    assert_eq!(
        fingerprint_bytes(&actual.fingerprints),
        fingerprint_bytes(&expected.fingerprints)
    );
}

#[test]
fn dimension_and_cell_heavy_workbooks_remain_lazy() {
    use xlsx_model::open_timing;

    for dimensions in [true, false] {
        let mut case = cases().remove(1);
        for sheet in &mut case.model.sheets {
            if dimensions {
                sheet.col_widths = (0..MAX_COLS).map(|col| (col, 12.0)).collect();
            } else {
                for row in 0..65_536 {
                    sheet.set_cell(
                        CellRef::new(row, 0),
                        Cell {
                            value: CellValue::Number {
                                value: f64::from(row),
                            },
                            style: Some(1),
                            ..Cell::default()
                        },
                    );
                }
            }
        }
        open_timing::begin();
        let actual = WorkbookBase::from_model_with_style_readings(
            &case.model,
            &case.legacy_dimensions,
            case.legacy_styles.as_ref(),
            &[],
        )
        .unwrap();
        let timing = open_timing::finish();
        assert_eq!(timing.counts["fingerprint_planned_keys"], 1);
        assert!(!timing.counts.contains_key("fingerprint_payload_fallbacks"));
        assert_eq!(actual.fingerprints.materialized_bytes(), None);
        assert!(actual.fingerprints.retained_bytes() < fingerprints::MAX_LAZY_BYTES + 64 * 1024);
        let expected = fingerprint_model_current_oracle(&case.model).unwrap();
        assert_eq!((actual.fingerprint, actual.bootstrap_client_id), expected);
    }
}

#[test]
fn materialization_releases_payload_across_clones() {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Payload");
    for row in 0..64 {
        sheet.set_cell(
            CellRef::new(row, 0),
            Cell {
                value: CellValue::Text {
                    value: "x".repeat(32_767),
                },
                ..Cell::default()
            },
        );
    }
    model.sheets.push(sheet);
    for path in 0..3 {
        let mut actual = WorkbookBase::from_model(&model).unwrap().fingerprints;
        let cloned = actual.clone();
        let before = actual.retained_bytes();
        assert!(before > 1024 * 1024);
        match path {
            0 => {
                let _ = fingerprint_bytes(&actual);
            }
            1 => actual.retain(|_, _| true),
            _ => {
                let _ = actual.clone().into_iter().collect::<BTreeMap<_, _>>();
            }
        }
        assert!(cloned.retained_bytes() < 64 * 1024);
        assert_eq!(cloned.materialized_bytes(), None);
        if path < 2 {
            assert_eq!(actual.materialized_bytes(), Some(actual.retained_bytes()));
        }
        assert_eq!(fingerprint_bytes(&actual), fingerprint_bytes(&cloned));
        assert_eq!(cloned.materialized_bytes(), Some(cloned.retained_bytes()));
    }
}

#[test]
fn merging_plans_keeps_the_combined_payload_within_budget() {
    let mut model = WorkbookModel::default();
    let mut sheet = Sheet::new("Payload");
    for row in 0..=(fingerprints::MAX_LAZY_BYTES / 2 / 32_767) as u32 {
        sheet.set_cell(
            CellRef::new(row, 0),
            Cell {
                value: CellValue::Text {
                    value: "x".repeat(32_767),
                },
                ..Cell::default()
            },
        );
    }
    model.sheets.push(sheet);
    let mut left = WorkbookBase::from_model(&model).unwrap().fingerprints;
    model.sheets[0].name.push('!');
    let right = WorkbookBase::from_model(&model).unwrap().fingerprints;
    let left_clone = left.clone();
    let right_clone = right.clone();
    assert_eq!(left.materialized_bytes(), None);
    assert_eq!(right.materialized_bytes(), None);
    left.merge(right);
    assert_eq!(left.materialized_bytes(), Some(left.retained_bytes()));
    assert!(left.retained_bytes() < 64 * 1024);
    assert!(left_clone.retained_bytes() < 64 * 1024);
    assert!(right_clone.retained_bytes() < 64 * 1024);
    let mut expected = left_clone.into_iter().collect::<BTreeMap<_, _>>();
    for (version, accepted) in right_clone {
        for fingerprint in accepted {
            append_fingerprint(&mut expected, version, fingerprint);
        }
    }
    assert_eq!(fingerprint_bytes(&left), fingerprint_bytes(&expected));
}
