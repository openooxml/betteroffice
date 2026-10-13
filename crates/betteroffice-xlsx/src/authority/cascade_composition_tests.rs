use super::*;

fn composed_model() -> WorkbookModel {
    let mut model = open_test_cases::matrix().remove(0).model;
    model.styles.fills[1] = xlsx_model::Fill::Pattern {
        pattern: xlsx_model::PatternType::LightUp,
        fg: Some(xlsx_model::Color::Rgb("#123456".into())),
        bg: Some(xlsx_model::Color::Rgb("#fedcba".into())),
    };
    model.styles.cell_xfs[1].xf_id = Some(2);
    model.styles.cell_xfs[1].quote_prefix = true;
    model.sheets[0].row_styles.insert(3, 1);
    model.sheets[0].col_styles = vec![ColStyle {
        first: 1,
        last: 3,
        xf: 1,
    }];
    model
}

#[test]
fn cascade_composition_readings_have_distinct_domains_and_clients() {
    let model = composed_model();
    let base = WorkbookBase::from_model(&model).unwrap();
    let domain = b"betteroffice-xlsx-yrs-v7-xf-facets-pattern-fills-style-defaults";
    let expected =
        fingerprint_model_with_row_domain_oracle(&model, SCHEMA_VERSION, true, true, true, domain)
            .unwrap();
    assert_eq!(base.fingerprint, expected.0);
    assert!(
        base.fingerprint
            .starts_with("xf-facets-pattern-fills-style-defaults:")
    );
    assert_eq!(
        fingerprint_client_id(&base.fingerprint),
        Some(base.bootstrap_client_id)
    );
    let mut builder = FingerprintBuilder::new(&model, &[]);
    let mut digests = BTreeSet::new();
    for (facets, patterns, defaults, tag) in [
        (false, false, false, ""),
        (true, false, false, "xf-facets:"),
        (false, true, false, "pattern-fills:"),
        (true, true, false, "xf-facets-pattern-fills:"),
        (false, false, true, "style-defaults:"),
        (true, false, true, "xf-facets-style-defaults:"),
        (false, true, true, "pattern-fills-style-defaults:"),
        (true, true, true, "xf-facets-pattern-fills-style-defaults:"),
    ] {
        let mut styles = FingerprintStyles::new(&model.styles);
        if !facets {
            styles = styles.without_xf_facets();
        }
        if patterns {
            styles = styles.with_written_fills();
        }
        let (fingerprint, client) = builder
            .hash(styles, SCHEMA_VERSION, true, true, defaults, false)
            .unwrap();
        assert!(fingerprint.starts_with(tag), "{tag}: {fingerprint}");
        assert_eq!(fingerprint_client_id(&fingerprint), Some(client));
        assert!(digests.insert(fingerprint.rsplit(':').next().unwrap().to_string()));
    }
    for change in 0..3 {
        let mut changed = model.clone();
        match change {
            0 => {
                changed.sheets[0].row_styles.insert(3, 0);
            }
            1 => {
                changed.styles.cell_xfs[1].quote_prefix = false;
            }
            _ => {
                if let xlsx_model::Fill::Pattern { bg, .. } = &mut changed.styles.fills[1] {
                    *bg = None;
                }
            }
        }
        assert_ne!(
            base.fingerprint,
            WorkbookBase::from_model(&changed).unwrap().fingerprint
        );
    }
}

#[test]
fn cascade_composition_tags_preserve_predicates_and_bootstrap_guards() {
    let model = composed_model();
    let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 71).unwrap();
    for (tag, facets, patterns) in [
        ("style-defaults:", false, false),
        ("xf-facets-style-defaults:", true, false),
        ("pattern-fills-style-defaults:", false, true),
        ("xf-facets-pattern-fills-style-defaults:", true, true),
    ] {
        let fingerprint = format!("{tag}0123456789abcdef");
        {
            let mut txn = authority.doc.transact_mut();
            txn.get_map(META)
                .unwrap()
                .insert(&mut txn, BASE_FINGERPRINT, fingerprint.clone());
        }
        assert_eq!(released_fingerprint(&fingerprint), !facets, "{tag}");
        assert_eq!(authority.holds_released_reading(), !patterns, "{tag}");
        assert_eq!(
            fingerprint_client_id(&fingerprint),
            Some(client_id_of_digest(0x0123456789abcdef)),
            "{tag}"
        );
    }
    authority = WorkbookAuthority::from_model_with_client_id(&model, 71).unwrap();
    let mut changed = model.clone();
    changed.sheets[0].row_styles.insert(3, 0);
    let other = WorkbookAuthority::from_model_with_client_id(&changed, 72).unwrap();
    let update = authority.encode_state_as_update_v1();
    assert!(!authority.seeds_another_bootstrap(&update));
    assert!(authority.seeds_another_bootstrap(&other.encode_state_as_update_v1()));
    assert_eq!(
        authority.seed_clients(&authority.doc.transact().state_vector()),
        vec![authority.base.bootstrap_client_id]
    );
}

#[test]
fn cascade_composition_released_fingerprint_layouts_round_trip() {
    let mut model = open_test_cases::matrix().remove(0).model;
    model.sheets = vec![Sheet::new("Data")];
    model.sheets[0].set_cell(
        CellRef::new(0, 0),
        Cell {
            value: CellValue::Number { value: 42.0 },
            style: Some(1),
            ..Cell::default()
        },
    );
    for version in 3..=6 {
        let bytes = tests::legacy_update(&model, version, version >= 4);
        let authority = tests::authority_from_update(&model, &bytes, 71);
        assert_eq!(authority.materialize().unwrap(), model);
        let encoded = authority.encode_state_as_update_v1();
        let peer = tests::authority_from_update(&model, &encoded, 72);
        assert_eq!(peer.materialize().unwrap(), model);
    }
}

#[test]
fn cascade_composition_v6_feature_layouts_round_trip() {
    let model = composed_model();
    for (domain, facets, patterns, defaults) in [
        (
            b"betteroffice-xlsx-yrs-v6-xf-facets".as_slice(),
            true,
            false,
            false,
        ),
        (
            b"betteroffice-xlsx-yrs-v6-pattern-fills".as_slice(),
            false,
            true,
            false,
        ),
        (
            b"betteroffice-xlsx-yrs-v6-xf-facets-pattern-fills".as_slice(),
            true,
            true,
            false,
        ),
        (
            b"betteroffice-xlsx-yrs-v6-style-defaults".as_slice(),
            false,
            false,
            true,
        ),
    ] {
        let mut legacy = model.clone();
        if !facets {
            legacy.styles = pre_facet_styles(&legacy.styles);
        }
        if !patterns {
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
        if !defaults {
            legacy.sheets[0].row_styles.clear();
        }
        let mut base = WorkbookBase::from_model(&legacy).unwrap();
        (base.fingerprint, base.bootstrap_client_id) = fingerprint_model_with_row_domain_oracle(
            &model,
            SCHEMA_VERSION,
            true,
            true,
            defaults,
            domain,
        )
        .unwrap();
        let doc = Doc::with_client_id(base.bootstrap_client_id);
        seed(&doc, &base, &legacy, &["sheet:0".to_string()]).unwrap();
        if !defaults {
            let mut txn = doc.transact_mut();
            let sheet = txn
                .get_map(SHEETS)
                .unwrap()
                .get(&txn, "sheet:0")
                .unwrap()
                .cast::<MapRef>()
                .unwrap();
            sheet.remove(&mut txn, COL_STYLES);
        }
        let update = doc
            .transact()
            .encode_state_as_update_v1(&StateVector::default());
        let current = WorkbookAuthority::from_model_with_client_id(&model, 71).unwrap();
        let adopted = match current.snapshot_replacement(&update) {
            SnapshotAdoption::Replacement(adopted) => adopted,
            SnapshotAdoption::Incompatible(error) => panic!("{domain:?}: {error}"),
            SnapshotAdoption::NotApplicable => panic!("{domain:?}: not adopted"),
        };
        let restored = adopted.materialize().unwrap();
        assert_eq!(restored.sheets[0].row_styles, model.sheets[0].row_styles);
        for (at, cell) in model.sheets[0].iter_cells() {
            let actual = restored.sheets[0].cell(at).unwrap();
            assert_eq!(
                restored.styles.cell_format(actual.style),
                model.styles.cell_format(cell.style),
                "{domain:?}: {at:?}"
            );
            assert_eq!(actual.value, cell.value);
        }
        let peer = tests::authority_from_update(&model, &adopted.encode_state_as_update_v1(), 72);
        assert_eq!(peer.materialize().unwrap(), restored);
    }
}
