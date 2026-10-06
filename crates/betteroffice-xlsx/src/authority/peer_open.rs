use super::*;
use ooxml_opc::WorkBudget;
use yrs::block::HAS_PARENT_SUB;

const SEED_ENTRIES_PER_UNIT: usize = 4;
const SEED_CHUNK_ENTRIES: usize = 256 * SEED_ENTRIES_PER_UNIT;

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
        let remaining = bytes.len() - offset;
        let count = work.take(remaining.div_ceil(64).min(256)).await * 64;
        let count = count.min(remaining);
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

#[derive(Clone, Copy)]
enum SeedParent {
    Root(&'static str),
    Item(ID),
}

enum SeedContent {
    Value(Any),
    Map,
    Order(Vec<Any>),
}

impl SeedContent {
    fn len(&self) -> usize {
        match self {
            Self::Order(values) => values.len(),
            _ => 1,
        }
    }
}

struct SeedEntry {
    parent: SeedParent,
    key: Option<String>,
    content: SeedContent,
}

struct SeedEncoder<'a> {
    local: &'a Doc,
    client: ClientID,
    clock: u32,
    imported_clock: u32,
    entries: Vec<SeedEntry>,
}

impl<'a> SeedEncoder<'a> {
    fn new(local: &'a Doc, client: u64) -> Self {
        Self {
            local,
            client: ClientID::new(client),
            clock: 0,
            imported_clock: 0,
            entries: Vec::with_capacity(SEED_CHUNK_ENTRIES),
        }
    }

    async fn push(
        &mut self,
        parent: SeedParent,
        key: Option<String>,
        content: SeedContent,
        work: &WorkBudget,
    ) -> Result<ID, String> {
        let id = ID::new(self.client, self.clock);
        let len = u32::try_from(content.len())
            .map_err(|_| "bootstrap content is too large".to_owned())?;
        self.clock = self
            .clock
            .checked_add(len)
            .ok_or_else(|| "bootstrap clock overflow".to_owned())?;
        self.entries.push(SeedEntry {
            parent,
            key,
            content,
        });
        if self.entries.len() == SEED_CHUNK_ENTRIES {
            self.flush(work).await?;
        }
        Ok(id)
    }

    async fn value(
        &mut self,
        parent: SeedParent,
        key: impl Into<String>,
        value: impl Into<Any>,
        work: &WorkBudget,
    ) -> Result<(), String> {
        self.push(
            parent,
            Some(key.into()),
            SeedContent::Value(value.into()),
            work,
        )
        .await?;
        Ok(())
    }

    async fn map(
        &mut self,
        parent: SeedParent,
        key: &str,
        work: &WorkBudget,
    ) -> Result<SeedParent, String> {
        self.push(parent, Some(key.to_owned()), SeedContent::Map, work)
            .await
            .map(SeedParent::Item)
    }

    async fn flush(&mut self, work: &WorkBudget) -> Result<(), String> {
        let mut entries = std::mem::take(&mut self.entries).into_iter();
        while !entries.as_slice().is_empty() {
            let count = (work.take(256).await * SEED_ENTRIES_PER_UNIT).min(entries.len());
            let mut encoder = EncoderV1::new();
            encoder.write_var(1_u32);
            encoder.write_var(count);
            encoder.write_client(self.client);
            encoder.write_var(self.imported_clock);
            for entry in entries.by_ref().take(count) {
                let content_ref = match &entry.content {
                    SeedContent::Map => BLOCK_ITEM_TYPE_REF_NUMBER,
                    _ => BLOCK_ITEM_ANY_REF_NUMBER,
                };
                let info = content_ref
                    | if entry.key.is_some() {
                        HAS_PARENT_SUB
                    } else {
                        0
                    };
                encoder.write_info(info);
                match entry.parent {
                    SeedParent::Root(name) => {
                        encoder.write_parent_info(true);
                        encoder.write_string(name);
                    }
                    SeedParent::Item(id) => {
                        encoder.write_parent_info(false);
                        encoder.write_left_id(&id);
                    }
                }
                if let Some(key) = entry.key {
                    encoder.write_string(&key);
                }
                self.imported_clock += entry.content.len() as u32;
                match entry.content {
                    SeedContent::Value(value) => {
                        encoder.write_len(1);
                        encoder.write_any(&value);
                    }
                    SeedContent::Map => encoder.write_type_ref(TYPE_REFS_MAP),
                    SeedContent::Order(values) => {
                        encoder.write_len(values.len() as u32);
                        for value in values {
                            encoder.write_any(&value);
                        }
                    }
                }
            }
            encoder.write_var(0_u32);
            hydrate_local_doc(self.local, &encoder.to_vec())?;
        }
        Ok(())
    }
}

pub(super) async fn seed_sliced(
    local: &Doc,
    base: &WorkbookBase,
    model: &WorkbookModel,
    keys: &[String],
    work: &WorkBudget,
) -> Result<(), String> {
    let mut encoder = SeedEncoder::new(local, base.bootstrap_client_id);
    let cell_formats = SeedParent::Root(CELL_FORMATS);
    let mut formats = HashMap::new();
    work.step().await;
    let (key, payload) = cell_format_entry(&CellFormat::default())?;
    encoder
        .value(cell_formats, key.as_str(), payload.as_str(), work)
        .await?;
    formats.insert(key, payload);
    for index in 0..model.styles.cell_xfs.len() {
        work.step().await;
        let index =
            u32::try_from(index).map_err(|_| "cell format table is too large".to_owned())?;
        let (key, payload) = cell_format_entry(&model.styles.cell_format(Some(index)))?;
        if let Some(existing) = formats.get(&key) {
            if existing != &payload {
                return Err(format!("conflicting cell format payload for {key}"));
            }
        } else {
            encoder
                .value(cell_formats, key.as_str(), payload.as_str(), work)
                .await?;
            formats.insert(key, payload);
        }
    }
    work.step().await;
    let meta = SeedParent::Root(META);
    encoder
        .value(meta, BASE_FINGERPRINT, base.fingerprint.as_str(), work)
        .await?;
    encoder
        .value(meta, "schemaVersion", SCHEMA_VERSION, work)
        .await?;
    encoder
        .value(meta, STRUCTURE_GENERATION, 0_i64, work)
        .await?;
    if !keys.is_empty() {
        let mut order = Vec::with_capacity(keys.len());
        for key in keys {
            work.step().await;
            order.push(Any::from(key.clone()));
        }
        encoder
            .push(
                SeedParent::Root(SHEET_ORDER),
                None,
                SeedContent::Order(order),
                work,
            )
            .await?;
    }
    let sheets = SeedParent::Root(SHEETS);
    for (key, sheet) in keys.iter().zip(&model.sheets) {
        let hyperlinks = String::from_utf8(json_list(&sheet.hyperlinks, work).await?)
            .map_err(|error| error.to_string())?;
        let charts = String::from_utf8(json_list(&sheet.charts, work).await?)
            .map_err(|error| error.to_string())?;
        work.step().await;
        let map = encoder.map(sheets, key, work).await?;
        let col_widths = encoder.map(map, COL_WIDTHS, work).await?;
        let contents = encoder.map(map, CONTENTS, work).await?;
        encoder
            .value(
                map,
                FREEZE_PANE,
                freeze_pane_to_any(sheet.freeze_pane),
                work,
            )
            .await?;
        encoder.value(map, HYPERLINKS, hyperlinks, work).await?;
        encoder.value(map, CHARTS, charts, work).await?;
        encoder
            .value(map, MERGES, merges_to_any(&sheet.merges), work)
            .await?;
        encoder.value(map, NAME, sheet.name.as_str(), work).await?;
        let row_heights = encoder.map(map, ROW_HEIGHTS, work).await?;
        let styles = encoder.map(map, STYLES, work).await?;
        for (&at, &value) in &sheet.col_widths {
            encoder
                .value(col_widths, at.to_string(), value, work)
                .await?;
        }
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
        for (key, value) in values {
            encoder.value(contents, key, value, work).await?;
        }
        for (&at, &value) in &sheet.row_heights {
            encoder
                .value(row_heights, at.to_string(), value, work)
                .await?;
        }
        for (key, value) in style_values {
            encoder.value(styles, key, value, work).await?;
        }
    }
    encoder.flush(work).await
}
