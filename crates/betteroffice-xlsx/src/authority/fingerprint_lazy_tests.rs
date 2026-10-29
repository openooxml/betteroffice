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
