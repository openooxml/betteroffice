use std::ops::{Deref, DerefMut};
use std::sync::OnceLock;

use super::*;

type Accepted = BTreeMap<i64, Vec<String>>;

#[derive(Clone)]
enum Source {
    Ready(Arc<Accepted>),
    Planned(Arc<Plan>),
}

#[derive(Clone, Default)]
pub(super) struct Fingerprints {
    values: OnceLock<Accepted>,
    sources: Vec<Source>,
}

impl Fingerprints {
    pub(super) fn accepts(&self, version: i64, fingerprint: &str) -> bool {
        if let Some(values) = self.values.get() {
            return values
                .get(&version)
                .is_some_and(|values| values.iter().any(|value| value == fingerprint));
        }
        if self.sources.iter().any(|source| match source {
            Source::Ready(values) => values
                .get(&version)
                .is_some_and(|values| values.iter().any(|value| value == fingerprint)),
            Source::Planned(plan) => plan.accepted.get(&version).is_some_and(|keys| {
                keys.iter().any(|&index| {
                    plan.variants[index]
                        .hash
                        .get()
                        .is_some_and(|hash| hash.0 == fingerprint)
                })
            }),
        }) {
            return true;
        }
        self.sources.iter().any(|source| match source {
            Source::Ready(_) => false,
            Source::Planned(plan) => plan.accepted.get(&version).is_some_and(|keys| {
                keys.iter().any(|&index| {
                    plan.variants[index].key.6.tag() == fingerprint_reading(fingerprint).tag()
                        && plan.hash(index).0 == fingerprint
                })
            }),
        })
    }

    pub(super) fn contains_client(&self, client: u64) -> bool {
        if let Some(values) = self.values.get() {
            return values
                .values()
                .flatten()
                .any(|value| fingerprint_client_id(value) == Some(client));
        }
        self.sources.iter().any(|source| match source {
            Source::Ready(values) => values
                .values()
                .flatten()
                .any(|value| fingerprint_client_id(value) == Some(client)),
            Source::Planned(plan) => plan
                .accepted
                .values()
                .flatten()
                .any(|&index| plan.hash(index).1 == client),
        })
    }

    pub(super) fn merge(&mut self, mut other: Self) {
        if let Some(values) = self.values.take() {
            self.sources = vec![Source::Ready(Arc::new(values))];
        }
        if let Some(values) = other.values.take() {
            self.sources.push(Source::Ready(Arc::new(values)));
        } else {
            self.sources.extend(other.sources);
        }
    }
}

impl From<Accepted> for Fingerprints {
    fn from(values: Accepted) -> Self {
        Self {
            values: OnceLock::from(values),
            sources: Vec::new(),
        }
    }
}

impl Deref for Fingerprints {
    type Target = Accepted;

    fn deref(&self) -> &Accepted {
        self.values.get_or_init(|| {
            let mut values = Accepted::new();
            for (index, source) in self.sources.iter().enumerate() {
                match source {
                    Source::Ready(accepted) if index == 0 => values = accepted.as_ref().clone(),
                    Source::Ready(accepted) => {
                        for (&version, accepted) in accepted.iter() {
                            for fingerprint in accepted {
                                append_fingerprint(&mut values, version, fingerprint.clone());
                            }
                        }
                    }
                    Source::Planned(plan) => {
                        for (&version, keys) in &plan.accepted {
                            for &index in keys {
                                append_fingerprint(
                                    &mut values,
                                    version,
                                    plan.hash(index).0.clone(),
                                );
                            }
                        }
                    }
                }
            }
            values
        })
    }
}

impl DerefMut for Fingerprints {
    fn deref_mut(&mut self) -> &mut Accepted {
        let _ = Self::deref(self);
        self.values.get_mut().unwrap()
    }
}

impl IntoIterator for Fingerprints {
    type Item = (i64, Vec<String>);
    type IntoIter = std::collections::btree_map::IntoIter<i64, Vec<String>>;

    fn into_iter(mut self) -> Self::IntoIter {
        let _ = Self::deref(&self);
        self.values.take().unwrap().into_iter()
    }
}

impl std::fmt::Debug for Fingerprints {
    fn fmt(&self, f: &mut std::fmt::Formatter<'_>) -> std::fmt::Result {
        std::fmt::Debug::fmt(self.deref(), f)
    }
}

impl PartialEq for Fingerprints {
    fn eq(&self, other: &Self) -> bool {
        self.deref() == other.deref()
    }
}

impl PartialEq<Fingerprints> for Accepted {
    fn eq(&self, other: &Fingerprints) -> bool {
        self == other.deref()
    }
}

impl snapshot::Codec for Fingerprints {
    fn write(&self, writer: &mut crate::snapshot::wire::Writer) {
        snapshot::Codec::write(self.deref(), writer);
    }

    fn read(
        reader: &mut crate::snapshot::wire::Reader<'_>,
    ) -> crate::snapshot::SnapshotResult<Self> {
        Ok(<Accepted as snapshot::Codec>::read(reader)?.into())
    }
}

struct Variant {
    key: FingerprintKey,
    hash: OnceLock<(String, u64)>,
}

struct Plan {
    accepted: BTreeMap<i64, Vec<usize>>,
    variants: Vec<Variant>,
    data: Data,
}

impl Plan {
    fn hash(&self, index: usize) -> &(String, u64) {
        let variant = &self.variants[index];
        variant.hash.get_or_init(|| {
            xlsx_model::open_timing! {
                let _timing = xlsx_model::open_timing::span("fingerprint_compute");
                xlsx_model::open_timing::count("fingerprint_planned_keys", 1);
            }
            self.data.hash(variant.key)
        })
    }
}

struct EncodedSheet {
    shared: Vec<u8>,
    dimensions: [Arc<Vec<u8>>; 2],
    freeze_pane: Vec<u8>,
    hyperlinks: Vec<u8>,
    charts: Vec<u8>,
    col_styles: Vec<u8>,
    row_styles: Vec<u8>,
}

struct Data {
    prefixes: [Vec<u8>; 2],
    styles: Vec<Vec<u8>>,
    full_precision: bool,
    sheets: Vec<EncodedSheet>,
}

impl Data {
    fn new(builder: &FingerprintBuilder<'_>) -> Result<Self, String> {
        #[cfg(test)]
        FINGERPRINT_PASSES.set(FINGERPRINT_PASSES.get() + 1);
        let model = builder.model;
        let encode_error = |error| format!("cannot fingerprint workbook base: {error}");
        let mut without_names = serde_json::to_vec(&(model.date_system, &model.shared_strings))
            .map_err(encode_error)?;
        let mut with_names = serde_json::to_vec(&(
            model.date_system,
            &model.defined_names,
            &model.shared_strings,
        ))
        .map_err(encode_error)?;
        for prefix in [&mut without_names, &mut with_names] {
            prefix.pop();
            prefix.push(b',');
        }
        let has_columns = model
            .sheets
            .iter()
            .any(|sheet| !sheet.col_styles.is_empty());
        let has_rows = model
            .sheets
            .iter()
            .any(|sheet| !sheet.row_styles.is_empty());
        let mut dimensions = Vec::<Arc<Vec<u8>>>::new();
        let mut sheets = Vec::with_capacity(model.sheets.len());
        for (index, sheet) in model.sheets.iter().enumerate() {
            let mut shared = Vec::new();
            hash_bytes(&mut shared, sheet.name.as_bytes());
            hash_u64(&mut shared, sheet.iter_cells().count() as u64);
            for (at, cell) in sheet.iter_cells() {
                xlsx_model::open_timing! {
                    xlsx_model::open_timing::count("fingerprint_cells", 1);
                }
                xlsx_model::progress::tick();
                hash_u32(&mut shared, at.row);
                hash_u32(&mut shared, at.col);
                hash_cell_value(&mut shared, &cell.value);
                match &cell.formula {
                    Some(formula) => {
                        shared.push(1);
                        hash_bytes(&mut shared, formula.as_bytes());
                    }
                    None => shared.push(0),
                }
                match cell.style {
                    Some(style) => {
                        shared.push(1);
                        hash_u32(&mut shared, style);
                    }
                    None => shared.push(0),
                }
            }
            hash_u64(&mut shared, sheet.merges.len() as u64);
            for range in &sheet.merges {
                hash_cell_ref(&mut shared, range.start);
                hash_cell_ref(&mut shared, range.end);
            }
            let current = encode_dimensions(&sheet.col_widths, &sheet.row_heights, &mut dimensions);
            let legacy = match builder.legacy_dimensions.get(index) {
                Some(legacy) => {
                    encode_dimensions(&legacy.col_widths, &legacy.row_heights, &mut dimensions)
                }
                None => current.clone(),
            };
            let mut freeze_pane = Vec::new();
            match sheet.freeze_pane {
                Some(pane) => {
                    freeze_pane.push(1);
                    hash_u32(&mut freeze_pane, pane.rows);
                    hash_u32(&mut freeze_pane, pane.cols);
                    hash_cell_ref(&mut freeze_pane, pane.top_left);
                }
                None => freeze_pane.push(0),
            }
            sheets.push(EncodedSheet {
                shared,
                dimensions: [current, legacy],
                freeze_pane,
                hyperlinks: sheet_json(true, &sheet.hyperlinks, "hyperlinks")?.unwrap(),
                charts: sheet_json(true, &sheet.charts, "charts")?.unwrap(),
                col_styles: sheet_json(has_columns, &sheet.col_styles, "column styles")?
                    .unwrap_or_default(),
                row_styles: sheet_json(has_rows, &sheet.row_styles, "row styles")?
                    .unwrap_or_default(),
            });
        }
        Ok(Self {
            prefixes: [without_names, with_names],
            styles: builder.style_bytes.clone(),
            full_precision: model.full_precision,
            sheets,
        })
    }

    fn hash(&self, key: FingerprintKey) -> (String, u64) {
        let (style, version, names, columns, rows, legacy, reading) = key;
        let mut hasher = Sha256::new();
        hasher.update(fingerprint_domain(version, reading));
        if !self.full_precision {
            hasher.update(b"fullPrecision=0");
        }
        let prefix = &self.prefixes[usize::from(names)];
        let styles = &self.styles[style];
        hash_u64(&mut hasher, (prefix.len() + styles.len() + 1) as u64);
        hasher.update(prefix);
        hasher.update(styles);
        hasher.update(b"]");
        hash_u64(&mut hasher, self.sheets.len() as u64);
        for sheet in &self.sheets {
            hasher.update(&sheet.shared);
            hasher.update(sheet.dimensions[usize::from(legacy)].as_slice());
            if version >= FREEZE_PANE_SCHEMA_VERSION {
                hasher.update(&sheet.freeze_pane);
            }
            if version >= HYPERLINK_SCHEMA_VERSION {
                hash_bytes(&mut hasher, &sheet.hyperlinks);
            }
            if version >= CHARTS_SCHEMA_VERSION {
                hash_bytes(&mut hasher, &sheet.charts);
            }
            if columns && !sheet.col_styles.is_empty() {
                hash_bytes(&mut hasher, &sheet.col_styles);
            }
            if rows && !sheet.row_styles.is_empty() {
                hash_bytes(&mut hasher, &sheet.row_styles);
            }
        }
        let digest = hasher.finalize();
        let fingerprint = fingerprint_text(reading, format!("{digest:x}"));
        let mut head = [0; 8];
        head.copy_from_slice(&digest[..8]);
        (fingerprint, client_id_of_digest(u64::from_be_bytes(head)))
    }
}

fn encode_dimensions(
    columns: &BTreeMap<u32, f64>,
    rows: &BTreeMap<u32, f64>,
    known: &mut Vec<Arc<Vec<u8>>>,
) -> Arc<Vec<u8>> {
    xlsx_model::open_timing! {
        xlsx_model::open_timing::count("fingerprint_dimension_cols", columns.len() as u64);
        xlsx_model::open_timing::count("fingerprint_dimension_rows", rows.len() as u64);
    }
    let mut bytes = Vec::with_capacity(16 + 12 * (columns.len() + rows.len()));
    for values in [columns, rows] {
        hash_u64(&mut bytes, values.len() as u64);
        for (&at, &value) in values {
            hash_u32(&mut bytes, at);
            hash_u64(&mut bytes, value.to_bits());
        }
    }
    match known.iter().find(|known| known.as_slice() == bytes) {
        Some(known) => known.clone(),
        None => {
            let bytes = Arc::new(bytes);
            known.push(bytes.clone());
            bytes
        }
    }
}

impl FingerprintBuilder<'_> {
    pub(super) fn finish(self, accepted: Accepted) -> Result<(String, u64, Fingerprints), String> {
        let data = {
            xlsx_model::open_timing! {
                let _timing = xlsx_model::open_timing::span("fingerprint_compute");
            }
            Data::new(&self)?
        };
        xlsx_model::open_timing! {
            let timing_build = xlsx_model::open_timing::span("fingerprint_build");
        }
        let variants = self
            .planned
            .unwrap_or_default()
            .into_iter()
            .map(|(key, _)| Variant {
                key,
                hash: OnceLock::new(),
            })
            .collect();
        let accepted = accepted
            .into_iter()
            .map(|(version, keys)| {
                (
                    version,
                    keys.into_iter()
                        .map(|key| key.parse::<usize>().unwrap())
                        .collect(),
                )
            })
            .collect();
        let plan = Arc::new(Plan {
            accepted,
            variants,
            data,
        });
        xlsx_model::open_timing! { drop(timing_build); }
        let (fingerprint, client) = plan.hash(0).clone();
        Ok((
            fingerprint,
            client,
            Fingerprints {
                values: OnceLock::new(),
                sources: vec![Source::Planned(plan)],
            },
        ))
    }
}

pub(super) fn compute(
    builder: &FingerprintBuilder<'_>,
    planned: &[(FingerprintKey, FingerprintStyles<'_>)],
) -> Result<Vec<(String, u64)>, String> {
    let data = Data::new(builder)?;
    Ok(planned.iter().map(|(key, _)| data.hash(*key)).collect())
}
