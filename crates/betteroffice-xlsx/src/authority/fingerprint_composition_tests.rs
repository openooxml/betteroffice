use super::*;

const COMBINED_DOMAIN: &[u8] = b"betteroffice-xlsx-yrs-v7-xf-facets-pattern-fills";
const COMBINED_TAG: &str = "xf-facets-pattern-fills:";

fn patterned_facets() -> WorkbookModel {
    use xlsx_model::{Border, BorderEdge, BorderStyle, Color, Fill, PatternType};
    let mut model = open_test_cases::matrix().remove(0).model;
    model.styles.fills[1] = Fill::Pattern {
        pattern: PatternType::LightUp,
        fg: Some(Color::Rgb("#123456".into())),
        bg: Some(Color::Rgb("#fedcba".into())),
    };
    model.styles.borders[0] = Border {
        top: Some(BorderEdge {
            style: BorderStyle::MediumDashDot,
            color: None,
        }),
        ..Border::default()
    };
    model.styles.cell_xfs[1].xf_id = Some(1);
    model.styles.cell_xfs[1].quote_prefix = true;
    model
}

#[test]
fn composition_fingerprint_hashes_both_readings_in_its_own_domain() {
    let model = patterned_facets();
    let base = WorkbookBase::from_model(&model).unwrap();
    let hash = |domain| {
        fingerprint_model_with_domain_oracle(&model, SCHEMA_VERSION, true, true, domain)
            .unwrap()
            .0
            .rsplit(':')
            .next()
            .unwrap()
            .to_string()
    };
    let digest = hash(COMBINED_DOMAIN);
    assert_eq!(base.fingerprint, format!("{COMBINED_TAG}{digest}"));
    assert_ne!(digest, hash(FACETS_FINGERPRINT_DOMAIN));
    assert_ne!(digest, hash(PATTERN_FILLS_FINGERPRINT_DOMAIN));
    assert_eq!(
        fingerprint_client_id(&base.fingerprint),
        Some(base.bootstrap_client_id)
    );
    let mut changed = model.clone();
    changed.styles.cell_xfs[1].quote_prefix = false;
    assert_ne!(
        base.fingerprint,
        WorkbookBase::from_model(&changed).unwrap().fingerprint
    );
    changed = model.clone();
    if let xlsx_model::Fill::Pattern { bg, .. } = &mut changed.styles.fills[1] {
        *bg = Some(xlsx_model::Color::Indexed(9));
    }
    assert_ne!(
        base.fingerprint,
        WorkbookBase::from_model(&changed).unwrap().fingerprint
    );
}

#[test]
fn composition_predicates_classify_each_reading() {
    let model = patterned_facets();
    let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 71).unwrap();
    Arc::make_mut(&mut authority.base).fingerprint = format!("{COMBINED_TAG}0123456789abcdef");
    for (tag, before_facets, before_patterns) in [
        ("", true, true),
        (FACETS_FINGERPRINT_TAG, false, true),
        (PATTERN_FILLS_FINGERPRINT_TAG, true, false),
        (COMBINED_TAG, false, false),
    ] {
        let fingerprint = format!("{tag}0123456789abcdef");
        {
            let mut txn = authority.doc.transact_mut();
            let meta = txn.get_map(META).unwrap();
            meta.insert(&mut txn, BASE_FINGERPRINT, fingerprint.clone());
        }
        assert_eq!(released_fingerprint(&fingerprint), before_facets, "{tag}");
        assert_eq!(authority.holds_released_reading(), before_patterns, "{tag}");
        assert_eq!(
            fingerprint_client_id(&fingerprint),
            Some(client_id_of_digest(0x0123456789abcdef)),
            "{tag}"
        );
    }
    Arc::make_mut(&mut authority.base).fingerprint =
        format!("{FACETS_FINGERPRINT_TAG}0123456789abcdef");
    assert!(!authority.holds_released_reading());
}

#[test]
fn composition_bootstrap_refuses_another_whole_state_and_accepts_its_own() {
    let model = patterned_facets();
    let own = WorkbookAuthority::from_model_with_client_id(&model, 71).unwrap();
    let mut changed = model.clone();
    changed.styles.cell_xfs[1].pivot_button = true;
    let other = WorkbookAuthority::from_model_with_client_id(&changed, 72).unwrap();
    assert!(own.base.fingerprint.starts_with(COMBINED_TAG));
    assert!(other.base.fingerprint.starts_with(COMBINED_TAG));
    assert!(!own.seeds_another_bootstrap(&own.encode_state_as_update_v1()));
    assert!(own.seeds_another_bootstrap(&other.encode_state_as_update_v1()));
    assert_eq!(
        own.seed_clients(&own.doc.transact().state_vector()),
        vec![own.base.bootstrap_client_id]
    );
}

#[test]
fn composition_main_released_demo_seed_restores() {
    let bytes = include_bytes!("../../../../apps/demo/public/showcase.xlsx");
    let seed = include_bytes!("released_demo_seed.bin");
    let options = crate::CalculationOptions {
        now_serial: Some(45_000.25),
    };
    let mut workbook =
        crate::Workbook::open_collaborative_recalculated(bytes, 71, options).unwrap();
    let before = workbook.model().clone();
    let result = workbook.apply_update_v1(seed, options).unwrap();
    assert!(result.applied);
    assert_eq!(workbook.model(), &before);
    let mut peer = crate::Workbook::open_collaborative(bytes, 72).unwrap();
    peer.apply_update_v1(&workbook.encode_state_as_update_v1(), options)
        .unwrap();
    assert_eq!(peer.model(), workbook.model());
}

#[test]
fn composition_single_reading_states_restore_source_formats() {
    let model = patterned_facets();
    for reading in [
        FingerprintReading::Released,
        FingerprintReading::Facets,
        FingerprintReading::Patterns,
    ] {
        let mut legacy = model.clone();
        if !reading.has_facets() {
            legacy.styles = pre_facet_styles(&legacy.styles);
        }
        if !reading.as_written() {
            legacy.styles.fills = legacy
                .styles
                .fills
                .iter()
                .map(|fill| released_fill(fill).into_owned())
                .collect();
            legacy.styles.borders = legacy
                .styles
                .borders
                .iter()
                .map(|border| released_border(border).into_owned())
                .collect();
        }
        let mut base = WorkbookBase::from_model(&legacy).unwrap();
        (base.fingerprint, base.bootstrap_client_id) = fingerprint_model_with_domain_oracle(
            &model,
            SCHEMA_VERSION,
            true,
            true,
            fingerprint_domain(SCHEMA_VERSION, reading),
        )
        .unwrap();
        let doc = Doc::with_client_id(base.bootstrap_client_id);
        seed(&doc, &base, &legacy, &["sheet:0".to_string()]).unwrap();
        let update = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let current = WorkbookAuthority::from_model_with_client_id(&model, 71).unwrap();
        let adopted = match current.snapshot_replacement(&update) {
            SnapshotAdoption::Replacement(adopted) => adopted,
            SnapshotAdoption::Incompatible(error) => panic!("{reading:?}: {error}"),
            SnapshotAdoption::NotApplicable => panic!("{reading:?}: snapshot not adopted"),
        };
        let restored = adopted.materialize().unwrap();
        for (at, cell) in model.sheets[0].iter_cells() {
            let actual = restored.sheets[0].cell(at).unwrap();
            assert_eq!(
                restored.styles.cell_format(actual.style),
                model.styles.cell_format(cell.style),
                "{reading:?}: {at:?}"
            );
            assert_eq!(actual.value, cell.value);
            assert_eq!(actual.formula, cell.formula);
        }
    }
}
