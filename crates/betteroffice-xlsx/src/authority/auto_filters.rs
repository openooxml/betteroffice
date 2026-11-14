//! Auto filters in a collaborative document. The `autoFilter` key stays one
//! whole value on the sheet's map, read as the source's filter when absent. The
//! filter's `_xlnm._FilterDatabase` name is worked out from that value on every
//! read and never stored, so the two cannot pick different winners.

use xlsx_model::AutoFilter;

use super::*;

pub(super) fn stored_auto_filter<T: ReadTxn>(
    sheet: &MapRef,
    txn: &T,
    fallback: Option<&AutoFilter>,
) -> Result<Option<AutoFilter>, String> {
    match sheet.get(txn, AUTO_FILTER) {
        Some(Out::Any(Any::String(json))) => serde_json::from_str(&json)
            .map_err(|error| format!("sheet auto filter is invalid: {error}")),
        Some(_) => Err("sheet auto filter is not a string".to_string()),
        None => Ok(fallback.cloned()),
    }
}

/// Writes `sheet`'s filter unless the stored one already reads as it. The range
/// is stored unanchored, the one spelling a peer accepts.
pub(super) fn sync_auto_filter(
    txn: &mut TransactionMut<'_>,
    sheet_map: &MapRef,
    sheet: &Sheet,
    fallback: Option<&AutoFilter>,
) -> Result<(), String> {
    let filter = sheet.auto_filter.clone().map(|mut filter| {
        for at in [&mut filter.range.start, &mut filter.range.end] {
            at.abs_row = false;
            at.abs_col = false;
        }
        filter
    });
    if stored_auto_filter(sheet_map, txn, fallback)? != filter {
        let json = serde_json::to_string(&filter)
            .map_err(|error| format!("cannot encode sheet auto filter: {error}"))?;
        sheet_map.try_update(txn, AUTO_FILTER, json);
    }
    Ok(())
}

/// Whether `filter` is one a peer's operation can have written over `source`,
/// the filter this sheet was opened with. That value itself is exempt however
/// odd, because every replica opened it. Anything else holds what the source
/// held or nothing beyond a range, so no XML the source lacked reaches the file.
pub(super) fn validate_stored_auto_filter(
    filter: Option<&AutoFilter>,
    source: Option<&AutoFilter>,
) -> Result<(), String> {
    if filter == source {
        return Ok(());
    }
    if source.is_some_and(|source| source.source_count > 1) {
        return Err("sheet auto filter replaces a repeated source filter".to_string());
    }
    let Some(filter) = filter else {
        return Ok(());
    };
    let from_source = source.filter(|source| {
        source.raw_xml == filter.raw_xml
            && source.source_count == filter.source_count
            && source.filter_columns == filter.filter_columns
            && source.has_sort_state == filter.has_sort_state
    });
    let bare = filter.raw_xml.is_none()
        && filter.source_count == 1
        && filter.filter_columns.is_empty()
        && !filter.has_sort_state;
    if from_source.is_none() && !bare {
        return Err("sheet auto filter holds what its source does not".to_string());
    }
    let anchored = [filter.range.start, filter.range.end]
        .iter()
        .any(|at| at.abs_row || at.abs_col);
    if anchored {
        return Err("sheet auto filter range is anchored".to_string());
    }
    xlsx_ops::validate_auto_filter_range(&filter.range, from_source)
        .map_err(|error| error.to_string())
}

/// Gives every filter that differs from its source the hidden name it carries.
/// A sheet still showing its source's filter keeps whatever the source table
/// holds, so undoing every filter edit restores `definedNames` byte for byte.
pub(super) fn project_filter_names(
    model: &mut WorkbookModel,
    keys: &[String],
    base: &WorkbookBase,
) -> Result<(), String> {
    let WorkbookModel {
        sheets,
        defined_names,
        ..
    } = model;
    for (index, (key, sheet)) in keys.iter().zip(sheets.iter()).enumerate() {
        let source = base_sheet_index(key)
            .and_then(|base_sheet| base.auto_filters.get(base_sheet))
            .and_then(Option::as_ref);
        if sheet.auto_filter.as_ref() == source {
            continue;
        }
        let scope =
            SheetId(u32::try_from(index).map_err(|_| "sheet index is out of range".to_string())?);
        let entry = defined_names.iter().position(|defined| {
            defined.local_sheet == Some(scope)
                && defined.name.eq_ignore_ascii_case(FILTER_DATABASE_NAME)
        });
        match (&sheet.auto_filter, entry) {
            (Some(filter), Some(entry)) => {
                let defined = &mut defined_names[entry];
                defined.formula = xlsx_ops::filter_database_formula(&sheet.name, filter.range);
                defined.hidden = true;
            }
            (Some(filter), None) => defined_names.push(DefinedName {
                name: FILTER_DATABASE_NAME.to_string(),
                formula: xlsx_ops::filter_database_formula(&sheet.name, filter.range),
                local_sheet: Some(scope),
                hidden: true,
            }),
            (None, Some(entry)) => {
                defined_names.remove(entry);
            }
            (None, None) => {}
        }
    }
    Ok(())
}

#[cfg(test)]
mod tests {
    use super::*;

    fn bare(range: &str) -> AutoFilter {
        AutoFilter {
            range: CellRange::parse_a1(range).unwrap(),
            raw_xml: None,
            source_count: 1,
            filter_columns: Vec::new(),
            has_sort_state: false,
        }
    }

    fn sourced(range: &str) -> AutoFilter {
        AutoFilter {
            raw_xml: Some(format!(
                r#"<autoFilter ref="{range}"><filterColumn colId="1"/></autoFilter>"#
            )),
            filter_columns: vec![1],
            ..bare(range)
        }
    }

    #[test]
    fn a_filter_alone_leaves_the_names_root_unmade() {
        let mut model = WorkbookModel::default();
        model.sheets.push(Sheet::new("Data"));
        let mut authority = WorkbookAuthority::from_model_with_client_id(&model, 106).unwrap();
        let op = Op::SetAutoFilter {
            sheet: SheetId(0),
            range: Some(CellRange::parse_a1("A1:B2").unwrap()),
        };
        xlsx_ops::apply_in_place(&mut model, &op).unwrap();
        assert_eq!(model.defined_names.len(), 1);
        authority
            .sync_model(&model, &[op], SyncOrigin::User, &HashMap::new(), true)
            .unwrap();
        let txn = authority.doc.transact();
        assert!(txn.root_refs().all(|(key, _)| key != DEFINED_NAMES));
        let sheet = txn
            .get_map(SHEETS)
            .unwrap()
            .get(&txn, "sheet:0")
            .unwrap()
            .cast::<MapRef>()
            .unwrap();
        assert!(sheet.get(&txn, AUTO_FILTER).is_some());
    }

    #[test]
    fn a_shared_frame_write_leaves_the_filter_to_its_own_write() {
        let mut model = WorkbookModel::default();
        let mut data = Sheet::new("Data");
        data.auto_filter = Some(bare("A1:B2"));
        model.sheets.push(data);
        let authority = WorkbookAuthority::from_model_with_client_id(&model, 107).unwrap();
        let mut txn = authority.doc.transact_mut();
        let sheet = txn
            .get_map(SHEETS)
            .unwrap()
            .get(&txn, "sheet:0")
            .unwrap()
            .cast::<MapRef>()
            .unwrap();
        sheet.remove(&mut txn, AUTO_FILTER);
        let data = &model.sheets[0];
        sync_sheet(&sheet, &mut txn, data, "Data", &model.styles, true).unwrap();
        assert!(sheet.get(&txn, AUTO_FILTER).is_none());
        sync_sheet(&sheet, &mut txn, data, "Data", &model.styles, false).unwrap();
        assert!(sheet.get(&txn, AUTO_FILTER).is_some());
    }

    #[test]
    fn the_source_value_is_exempt_whatever_it_holds() {
        let odd = AutoFilter {
            source_count: 2,
            ..sourced("A1:C9")
        };
        assert!(validate_stored_auto_filter(Some(&odd), Some(&odd)).is_ok());
    }

    #[test]
    fn a_value_keeps_to_the_source_or_to_a_bare_range() {
        let source = sourced("A1:C9");
        let wider = AutoFilter {
            range: CellRange::parse_a1("A1:D9").unwrap(),
            ..source.clone()
        };
        assert!(validate_stored_auto_filter(Some(&wider), Some(&source)).is_ok());
        assert!(validate_stored_auto_filter(Some(&bare("D1:F9")), Some(&source)).is_ok());
        assert!(validate_stored_auto_filter(None, Some(&source)).is_ok());
        assert!(validate_stored_auto_filter(Some(&bare("D1:F9")), None).is_ok());

        let foreign = AutoFilter {
            raw_xml: Some(r#"<autoFilter ref="A1:C9"><x/></autoFilter>"#.to_string()),
            ..source.clone()
        };
        let counted = AutoFilter {
            source_count: 2,
            ..source.clone()
        };
        let columns = AutoFilter {
            filter_columns: vec![0],
            ..source.clone()
        };
        let sorted = AutoFilter {
            has_sort_state: true,
            ..source.clone()
        };
        let moved = AutoFilter {
            range: CellRange::parse_a1("B1:D9").unwrap(),
            ..source.clone()
        };
        let narrow = AutoFilter {
            range: CellRange::parse_a1("A1:A9").unwrap(),
            ..source.clone()
        };
        let reversed = AutoFilter {
            range: CellRange {
                start: CellRef::new(9, 0),
                end: CellRef::new(0, 2),
            },
            ..bare("A1:C9")
        };
        let anchored = bare("$A$1:C9");
        for filter in [
            &foreign, &counted, &columns, &sorted, &moved, &narrow, &reversed, &anchored,
        ] {
            assert!(
                validate_stored_auto_filter(Some(filter), Some(&source)).is_err(),
                "{filter:?}"
            );
        }
        assert!(validate_stored_auto_filter(Some(&source), None).is_err());
    }

    #[test]
    fn a_repeated_source_filter_stays_as_it_is() {
        let source = AutoFilter {
            source_count: 2,
            ..bare("A1:C9")
        };
        assert!(validate_stored_auto_filter(None, Some(&source)).is_err());
        assert!(validate_stored_auto_filter(Some(&bare("A1:D9")), Some(&source)).is_err());
    }
}
