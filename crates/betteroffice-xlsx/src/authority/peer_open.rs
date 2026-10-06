use super::*;
use ooxml_opc::WorkBudget;

pub(super) async fn json_list<T: Serialize>(
    values: &[T],
    work: &WorkBudget,
) -> Result<Vec<u8>, String> {
    let mut bytes = vec![b'['];
    for (index, value) in values.iter().enumerate() {
        work.step().await;
        if index != 0 {
            bytes.push(b',');
        }
        let encoded = serde_json::to_vec(value).map_err(|error| error.to_string())?;
        work.append_bytes(&mut bytes, &encoded).await;
    }
    bytes.push(b']');
    Ok(bytes)
}

pub(super) async fn fingerprint_styles_json(
    styles: FingerprintStyles<'_>,
    work: &WorkBudget,
) -> Result<Vec<u8>, String> {
    let mut bytes = b"{\"fonts\":".to_vec();
    work.append_bytes(&mut bytes, &json_list(&styles.styles.fonts, work).await?)
        .await;
    bytes.extend_from_slice(b",\"fills\":");
    work.append_bytes(&mut bytes, &json_list(&styles.styles.fills, work).await?)
        .await;
    bytes.extend_from_slice(b",\"borders\":");
    work.append_bytes(&mut bytes, &json_list(&styles.styles.borders, work).await?)
        .await;
    bytes.extend_from_slice(b",\"cell_xfs\":");
    work.append_bytes(&mut bytes, &json_list(&styles.styles.cell_xfs, work).await?)
        .await;
    bytes.extend_from_slice(b",\"num_fmts\":");
    work.append_bytes(&mut bytes, &json_list(&styles.styles.num_fmts, work).await?)
        .await;
    bytes.extend_from_slice(b",\"theme\":");
    bytes.extend(serde_json::to_vec(&styles.styles.theme).map_err(|error| error.to_string())?);
    if !styles.indexed_colors.is_empty() {
        bytes.extend_from_slice(b",\"indexed_colors\":");
        work.append_bytes(&mut bytes, &json_list(styles.indexed_colors, work).await?)
            .await;
    }
    bytes.push(b'}');
    Ok(bytes)
}

pub(super) async fn hash_bytes_sliced(hasher: &mut Sha256, bytes: &[u8], work: &WorkBudget) {
    hash_u64(hasher, bytes.len() as u64);
    hash_payload_sliced(hasher, bytes, work).await;
}

pub(super) async fn hash_payload_sliced(hasher: &mut Sha256, bytes: &[u8], work: &WorkBudget) {
    let mut offset = 0;
    while offset < bytes.len() {
        let count = (work.take(256).await * 64).min(bytes.len() - offset);
        hasher.update(&bytes[offset..offset + count]);
        offset += count;
    }
}

pub(super) async fn hash_cell_value_sliced(
    hasher: &mut Sha256,
    value: &CellValue,
    work: &WorkBudget,
) {
    match value {
        CellValue::Text { value } => {
            hasher.update([2]);
            hash_bytes_sliced(hasher, value.as_bytes(), work).await;
        }
        _ => hash_cell_value(hasher, value),
    }
}

fn replicate(bootstrap: &Doc, local: &Doc, state: &mut StateVector) -> Result<(), String> {
    let txn = bootstrap.transact();
    let update = txn.encode_state_as_update_v1(state);
    *state = txn.state_vector();
    drop(txn);
    hydrate_local_doc(local, &update)
}

async fn seed_values(
    bootstrap: &Doc,
    local: &Doc,
    state: &mut StateVector,
    map: &MapRef,
    values: impl IntoIterator<Item = (String, Any)>,
    work: &WorkBudget,
) -> Result<(), String> {
    let mut values = values.into_iter().peekable();
    while values.peek().is_some() {
        let count = work.take(256).await;
        let mut txn = bootstrap.transact_mut_with(BOOTSTRAP_ORIGIN);
        for (key, value) in values.by_ref().take(count) {
            map.try_update(&mut txn, key, value);
        }
        drop(txn);
        replicate(bootstrap, local, state)?;
    }
    Ok(())
}

pub(super) async fn seed_sliced(
    bootstrap: &Doc,
    local: &Doc,
    base: &WorkbookBase,
    model: &WorkbookModel,
    keys: &[String],
    work: &WorkBudget,
) -> Result<(), String> {
    let mut state = StateVector::default();
    work.step().await;
    let cell_formats = bootstrap
        .transact_mut_with(BOOTSTRAP_ORIGIN)
        .get_or_insert_map(CELL_FORMATS);
    replicate(bootstrap, local, &mut state)?;
    let default = cell_format_entry(&CellFormat::default())?;
    seed_values(
        bootstrap,
        local,
        &mut state,
        &cell_formats,
        [(default.0, Any::from(default.1))],
        work,
    )
    .await?;
    for index in 0..model.styles.cell_xfs.len() {
        work.step().await;
        let index =
            u32::try_from(index).map_err(|_| "cell format table is too large".to_owned())?;
        let (key, payload) = cell_format_entry(&model.styles.cell_format(Some(index)))?;
        seed_values(
            bootstrap,
            local,
            &mut state,
            &cell_formats,
            [(key, Any::from(payload))],
            work,
        )
        .await?;
    }
    work.step().await;
    {
        let mut txn = bootstrap.transact_mut_with(BOOTSTRAP_ORIGIN);
        let meta = txn.get_or_insert_map(META);
        meta.insert(&mut txn, BASE_FINGERPRINT, base.fingerprint.as_str());
        meta.insert(&mut txn, "schemaVersion", SCHEMA_VERSION);
        meta.insert(&mut txn, STRUCTURE_GENERATION, 0_i64);
        txn.get_or_insert_array(SHEET_ORDER);
    }
    replicate(bootstrap, local, &mut state)?;
    for (index, key) in keys.iter().enumerate() {
        work.step().await;
        let mut txn = bootstrap.transact_mut_with(BOOTSTRAP_ORIGIN);
        txn.get_or_insert_array(SHEET_ORDER)
            .insert(&mut txn, index as u32, key.clone());
        drop(txn);
        replicate(bootstrap, local, &mut state)?;
    }
    work.step().await;
    let sheets = bootstrap
        .transact_mut_with(BOOTSTRAP_ORIGIN)
        .get_or_insert_map(SHEETS);
    replicate(bootstrap, local, &mut state)?;
    for (key, sheet) in keys.iter().zip(&model.sheets) {
        let hyperlinks = String::from_utf8(json_list(&sheet.hyperlinks, work).await?)
            .map_err(|error| error.to_string())?;
        let charts = String::from_utf8(json_list(&sheet.charts, work).await?)
            .map_err(|error| error.to_string())?;
        work.step().await;
        let (col_widths, contents, row_heights, styles) = {
            let mut txn = bootstrap.transact_mut_with(BOOTSTRAP_ORIGIN);
            let map = sheets.insert(&mut txn, key.as_str(), MapPrelim::default());
            let col_widths: MapRef = map.get_or_init(&mut txn, COL_WIDTHS);
            let contents: MapRef = map.get_or_init(&mut txn, CONTENTS);
            map.try_update(&mut txn, FREEZE_PANE, freeze_pane_to_any(sheet.freeze_pane));
            map.try_update(&mut txn, HYPERLINKS, hyperlinks);
            map.try_update(&mut txn, CHARTS, charts);
            map.try_update(&mut txn, MERGES, merges_to_any(&sheet.merges));
            map.try_update(&mut txn, NAME, sheet.name.as_str());
            let row_heights: MapRef = map.get_or_init(&mut txn, ROW_HEIGHTS);
            let styles: MapRef = map.get_or_init(&mut txn, STYLES);
            (col_widths, contents, row_heights, styles)
        };
        replicate(bootstrap, local, &mut state)?;
        seed_values(
            bootstrap,
            local,
            &mut state,
            &col_widths,
            sheet
                .col_widths
                .iter()
                .map(|(&at, &value)| (at.to_string(), Any::from(value))),
            work,
        )
        .await?;
        let mut values = BTreeMap::new();
        let mut style_values = BTreeMap::new();
        for (at, cell) in sheet.iter_cells() {
            work.step().await;
            if let Some(value) = content_to_any(cell) {
                values.insert(cell_key(at), value);
            }
            if let Some(style) = cell.style {
                style_values.insert(cell_key(at), Any::from(style_key(&model.styles, style)?));
            }
        }
        seed_values(bootstrap, local, &mut state, &contents, values, work).await?;
        seed_values(
            bootstrap,
            local,
            &mut state,
            &row_heights,
            sheet
                .row_heights
                .iter()
                .map(|(&at, &value)| (at.to_string(), Any::from(value))),
            work,
        )
        .await?;
        seed_values(bootstrap, local, &mut state, &styles, style_values, work).await?;
    }
    Ok(())
}
