#[cfg(test)]
use std::mem::size_of;
use std::ops::{Deref, DerefMut};
use std::sync::{Mutex, OnceLock};

use super::*;

type Accepted = BTreeMap<i64, Vec<String>>;
type Bytes = Box<[u8]>;

pub(super) const MAX_LAZY_BYTES: usize = 8 * 1024 * 1024;
const CHUNK_BYTES: usize = 64 * 1024;

#[derive(Clone)]
enum Source {
    Ready(Arc<Accepted>),
    Planned(Arc<Plan>),
}

#[derive(Default)]
pub(super) struct Fingerprints {
    values: OnceLock<Accepted>,
    sources: Mutex<Vec<Source>>,
}

impl Fingerprints {
    #[cfg(test)]
    pub(super) fn retained_bytes(&self) -> usize {
        let sources = self.sources.lock().unwrap();
        let mut bytes = self.materialized_bytes().unwrap_or(size_of::<Self>())
            + sources.capacity() * size_of::<Source>();
        let mut seen = HashSet::new();
        for source in sources.iter() {
            match source {
                Source::Ready(values) if seen.insert(Arc::as_ptr(values) as usize) => {
                    bytes +=
                        2 * size_of::<usize>() + size_of::<Accepted>() + accepted_bytes(values);
                }
                Source::Planned(plan) if seen.insert(Arc::as_ptr(plan) as usize) => {
                    bytes += 2 * size_of::<usize>()
                        + size_of::<Plan>()
                        + tree_bytes(&plan.accepted)
                        + plan
                            .accepted
                            .values()
                            .map(|keys| keys.capacity() * size_of::<usize>())
                            .sum::<usize>()
                        + plan.variants.capacity() * size_of::<Variant>()
                        + plan
                            .variants
                            .iter()
                            .filter_map(|variant| variant.hash.get())
                            .map(|hash| hash.0.capacity())
                            .sum::<usize>();
                    if let Some(data) = plan.data.lock().unwrap().as_ref() {
                        bytes += data.retained_bytes();
                    }
                }
                _ => {}
            }
        }
        bytes
    }

    #[cfg(test)]
    pub(super) fn materialized_bytes(&self) -> Option<usize> {
        self.values
            .get()
            .map(|values| size_of::<Self>() + accepted_bytes(values))
    }

    pub(super) fn accepts(&self, version: i64, fingerprint: &str) -> bool {
        let sources = self.sources.lock().unwrap();
        if let Some(values) = self.values.get() {
            return values
                .get(&version)
                .is_some_and(|values| values.iter().any(|value| value == fingerprint));
        }
        if sources.iter().any(|source| match source {
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
        sources.iter().any(|source| match source {
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
        let sources = self.sources.lock().unwrap();
        if let Some(values) = self.values.get() {
            return values
                .values()
                .flatten()
                .any(|value| fingerprint_client_id(value) == Some(client));
        }
        sources.iter().any(|source| match source {
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
        let sources = self.sources.get_mut().unwrap();
        if let Some(values) = self.values.take() {
            *sources = vec![Source::Ready(Arc::new(values))];
        }
        if let Some(values) = other.values.take() {
            sources.push(Source::Ready(Arc::new(values)));
        } else {
            sources.extend(other.sources.into_inner().unwrap());
        }
        let mut seen = HashSet::new();
        let payload = sources
            .iter()
            .filter_map(|source| match source {
                Source::Planned(plan) if seen.insert(Arc::as_ptr(plan)) => Some(
                    plan.data
                        .lock()
                        .unwrap()
                        .as_ref()
                        .map_or(0, |data| data.encoded_bytes),
                ),
                _ => None,
            })
            .sum::<usize>();
        if payload > MAX_LAZY_BYTES {
            let _ = Self::deref(self);
        }
    }
}

impl Clone for Fingerprints {
    fn clone(&self) -> Self {
        let sources = self.sources.lock().unwrap();
        match self.values.get() {
            Some(values) => values.clone().into(),
            None => Self {
                values: OnceLock::new(),
                sources: Mutex::new(sources.clone()),
            },
        }
    }
}

impl From<Accepted> for Fingerprints {
    fn from(values: Accepted) -> Self {
        Self {
            values: OnceLock::from(values),
            sources: Mutex::new(Vec::new()),
        }
    }
}

impl Deref for Fingerprints {
    type Target = Accepted;

    fn deref(&self) -> &Accepted {
        let values = self.values.get_or_init(|| {
            let sources = self.sources.lock().unwrap();
            let mut values = Accepted::new();
            for (index, source) in sources.iter().enumerate() {
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
        });
        *self.sources.lock().unwrap() = Vec::new();
        values
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
    data: Mutex<Option<Data>>,
}

impl Plan {
    fn hash(&self, index: usize) -> &(String, u64) {
        let variant = &self.variants[index];
        let mut computed = false;
        let hash = variant.hash.get_or_init(|| {
            computed = true;
            xlsx_model::open_timing! {
                let _timing = xlsx_model::open_timing::span("fingerprint_compute");
                xlsx_model::open_timing::count("fingerprint_planned_keys", 1);
            }
            self.data
                .lock()
                .unwrap()
                .as_ref()
                .unwrap()
                .hash(variant.key)
        });
        if computed
            && self
                .variants
                .iter()
                .all(|variant| variant.hash.get().is_some())
        {
            self.data.lock().unwrap().take();
        }
        hash
    }
}

struct EncodedSheet {
    shared: Bytes,
    dimensions: [Arc<Bytes>; 2],
    freeze_pane: Bytes,
    hyperlinks: Bytes,
    charts: Bytes,
    col_styles: Bytes,
    row_styles: Bytes,
}

struct Data {
    prefixes: [Bytes; 2],
    styles: Vec<Bytes>,
    full_precision: bool,
    sheets: Vec<EncodedSheet>,
    encoded_bytes: usize,
}

impl Data {
    fn new(builder: &FingerprintBuilder<'_>, limit: usize) -> Result<Self, EncodeError> {
        #[cfg(test)]
        FINGERPRINT_PASSES.set(FINGERPRINT_PASSES.get() + 1);
        let model = builder.model;
        let mut budget = Budget(limit);
        let encode_error = |error| format!("cannot fingerprint workbook base: {error}");
        let mut without_names = budget.encode(|sink| {
            serde_json::to_writer(sink, &(model.date_system, &model.shared_strings))
                .map_err(encode_error)
        })?;
        let mut with_names = budget.encode(|sink| {
            serde_json::to_writer(
                sink,
                &(
                    model.date_system,
                    &model.defined_names,
                    &model.shared_strings,
                ),
            )
            .map_err(encode_error)
        })?;
        for prefix in [&mut without_names, &mut with_names] {
            let last = prefix.len() - 1;
            prefix[last] = b',';
        }
        let styles = builder
            .style_bytes
            .iter()
            .map(|bytes| {
                budget.encode(|sink| {
                    sink.feed(bytes);
                    Ok(())
                })
            })
            .collect::<Result<_, _>>()?;
        let has_columns = model
            .sheets
            .iter()
            .any(|sheet| !sheet.col_styles.is_empty());
        let has_rows = model
            .sheets
            .iter()
            .any(|sheet| !sheet.row_styles.is_empty());
        let mut dimensions = Vec::<Arc<Bytes>>::new();
        let mut sheets = Vec::with_capacity(model.sheets.len());
        for (index, sheet) in model.sheets.iter().enumerate() {
            let shared = budget.encode(|sink| {
                encode_shared(sheet, sink);
                Ok(())
            })?;
            let current = encode_dimensions(
                &sheet.col_widths,
                &sheet.row_heights,
                &mut dimensions,
                &mut budget,
            )?;
            let legacy = match builder.legacy_dimensions.get(index) {
                Some(legacy) => encode_dimensions(
                    &legacy.col_widths,
                    &legacy.row_heights,
                    &mut dimensions,
                    &mut budget,
                )?,
                None => current.clone(),
            };
            let freeze_pane = budget.encode(|sink| {
                encode_freeze_pane(sheet.freeze_pane, sink);
                Ok(())
            })?;
            sheets.push(EncodedSheet {
                shared,
                dimensions: [current, legacy],
                freeze_pane,
                hyperlinks: budget.json(&sheet.hyperlinks, "hyperlinks")?,
                charts: budget.json(&sheet.charts, "charts")?,
                col_styles: if has_columns {
                    budget.json(&sheet.col_styles, "column styles")?
                } else {
                    Bytes::default()
                },
                row_styles: if has_rows {
                    budget.json(&sheet.row_styles, "row styles")?
                } else {
                    Bytes::default()
                },
            });
        }
        Ok(Self {
            prefixes: [without_names, with_names],
            styles,
            full_precision: model.full_precision,
            sheets,
            encoded_bytes: limit - budget.0,
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
            hasher.update(sheet.dimensions[usize::from(legacy)].as_ref().as_ref());
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
        finish_hash(hasher, reading)
    }

    #[cfg(test)]
    fn retained_bytes(&self) -> usize {
        let mut bytes = self.prefixes.iter().map(|bytes| bytes.len()).sum::<usize>()
            + self.styles.capacity() * size_of::<Bytes>()
            + self.styles.iter().map(|bytes| bytes.len()).sum::<usize>()
            + self.sheets.capacity() * size_of::<EncodedSheet>();
        let mut dimensions = HashSet::new();
        for sheet in &self.sheets {
            bytes += sheet.shared.len()
                + sheet.freeze_pane.len()
                + sheet.hyperlinks.len()
                + sheet.charts.len()
                + sheet.col_styles.len()
                + sheet.row_styles.len();
            for values in &sheet.dimensions {
                if dimensions.insert(Arc::as_ptr(values)) {
                    bytes += 2 * size_of::<usize>() + size_of::<Bytes>() + values.len();
                }
            }
        }
        bytes
    }
}

#[cfg(test)]
fn tree_bytes<K, V>(map: &BTreeMap<K, V>) -> usize {
    map.len() * (11 * size_of::<(K, V)>() + 16 * size_of::<usize>())
}

#[cfg(test)]
fn accepted_bytes(values: &Accepted) -> usize {
    tree_bytes(values)
        + values
            .values()
            .map(|accepted| {
                accepted.capacity() * size_of::<String>()
                    + accepted.iter().map(String::capacity).sum::<usize>()
            })
            .sum::<usize>()
}

fn encode_dimensions(
    columns: &BTreeMap<u32, f64>,
    rows: &BTreeMap<u32, f64>,
    known: &mut Vec<Arc<Bytes>>,
    budget: &mut Budget,
) -> Result<Arc<Bytes>, EncodeError> {
    xlsx_model::open_timing! {
        xlsx_model::open_timing::count("fingerprint_dimension_cols", columns.len() as u64);
        xlsx_model::open_timing::count("fingerprint_dimension_rows", rows.len() as u64);
    }
    let bytes = budget.encode(|sink| {
        write_dimensions(columns, rows, sink);
        Ok(())
    })?;
    Ok(
        match known
            .iter()
            .find(|known| known.as_ref().as_ref() == bytes.as_ref())
        {
            Some(known) => {
                budget.0 += bytes.len();
                known.clone()
            }
            None => {
                let bytes = Arc::new(bytes);
                known.push(bytes.clone());
                bytes
            }
        },
    )
}

impl FingerprintBuilder<'_> {
    pub(super) fn finish(self, accepted: Accepted) -> Result<(String, u64, Fingerprints), String> {
        self.finish_with_budget(accepted, MAX_LAZY_BYTES)
    }

    pub(super) fn finish_with_budget(
        self,
        accepted: Accepted,
        limit: usize,
    ) -> Result<(String, u64, Fingerprints), String> {
        let data = {
            xlsx_model::open_timing! {
                let _timing = xlsx_model::open_timing::span("fingerprint_compute");
            }
            Data::new(&self, limit)
        };
        let data = match data {
            Ok(data) => data,
            Err(EncodeError::Invalid(error)) => return Err(error),
            Err(EncodeError::Budget) => {
                let planned = self.planned.as_deref().unwrap_or_default();
                xlsx_model::open_timing! {
                    let _timing = xlsx_model::open_timing::span("fingerprint_compute");
                    xlsx_model::open_timing::count("fingerprint_payload_fallbacks", 1);
                    xlsx_model::open_timing::count("fingerprint_planned_keys", planned.len() as u64);
                }
                let hashes = compute(&self, planned)?;
                let (fingerprint, client) = hashes[0].clone();
                let mut values = Accepted::new();
                for (version, keys) in accepted {
                    for key in keys {
                        append_fingerprint(
                            &mut values,
                            version,
                            hashes[key.parse::<usize>().unwrap()].0.clone(),
                        );
                    }
                }
                return Ok((fingerprint, client, values.into()));
            }
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
            data: Mutex::new(Some(data)),
        });
        xlsx_model::open_timing! { drop(timing_build); }
        let (fingerprint, client) = plan.hash(0).clone();
        Ok((
            fingerprint,
            client,
            Fingerprints {
                values: OnceLock::new(),
                sources: Mutex::new(vec![Source::Planned(plan)]),
            },
        ))
    }
}

pub(super) fn compute(
    builder: &FingerprintBuilder<'_>,
    planned: &[(FingerprintKey, FingerprintStyles<'_>)],
) -> Result<Vec<(String, u64)>, String> {
    #[cfg(test)]
    FINGERPRINT_PASSES.set(FINGERPRINT_PASSES.get() + 1);
    let model = builder.model;
    let mut hashers = planned
        .iter()
        .map(|&(key, styles)| {
            validate_schema_version(key.1)?;
            let mut hasher = Sha256::new();
            hasher.update(fingerprint_domain(key.1, key.6));
            if !model.full_precision {
                hasher.update(b"fullPrecision=0");
            }
            let base = if key.2 {
                serde_json::to_vec(&(
                    model.date_system,
                    &model.defined_names,
                    &model.shared_strings,
                    &styles,
                ))
            } else {
                serde_json::to_vec(&(model.date_system, &model.shared_strings, &styles))
            }
            .map_err(|error| format!("cannot fingerprint workbook base: {error}"))?;
            hash_bytes(&mut hasher, &base);
            hash_u64(&mut hasher, model.sheets.len() as u64);
            Ok(hasher)
        })
        .collect::<Result<Vec<_>, String>>()?;
    let has_columns = model
        .sheets
        .iter()
        .any(|sheet| !sheet.col_styles.is_empty());
    let has_rows = model
        .sheets
        .iter()
        .any(|sheet| !sheet.row_styles.is_empty());
    for (index, sheet) in model.sheets.iter().enumerate() {
        let mut stream = Stream {
            hashers: &mut hashers,
            bytes: Vec::with_capacity(CHUNK_BYTES),
        };
        encode_shared(sheet, &mut stream);
        stream.flush();
        drop(stream);
        let hyperlinks = sheet_json(
            planned
                .iter()
                .any(|(key, _)| key.1 >= HYPERLINK_SCHEMA_VERSION),
            &sheet.hyperlinks,
            "hyperlinks",
        )?;
        let charts = sheet_json(
            planned
                .iter()
                .any(|(key, _)| key.1 >= CHARTS_SCHEMA_VERSION),
            &sheet.charts,
            "charts",
        )?;
        let col_styles = sheet_json(
            has_columns && planned.iter().any(|(key, _)| key.3),
            &sheet.col_styles,
            "column styles",
        )?;
        let row_styles = sheet_json(
            has_rows && planned.iter().any(|(key, _)| key.4),
            &sheet.row_styles,
            "row styles",
        )?;
        for (hasher, &(key, _)) in hashers.iter_mut().zip(planned) {
            let (columns, rows) = match builder.legacy_dimensions.get(index).filter(|_| key.5) {
                Some(legacy) => (&legacy.col_widths, &legacy.row_heights),
                None => (&sheet.col_widths, &sheet.row_heights),
            };
            write_dimensions(columns, rows, hasher);
            if key.1 >= FREEZE_PANE_SCHEMA_VERSION {
                encode_freeze_pane(sheet.freeze_pane, hasher);
            }
            if let Some(bytes) = hyperlinks
                .as_ref()
                .filter(|_| key.1 >= HYPERLINK_SCHEMA_VERSION)
            {
                hash_bytes(hasher, bytes);
            }
            if let Some(bytes) = charts.as_ref().filter(|_| key.1 >= CHARTS_SCHEMA_VERSION) {
                hash_bytes(hasher, bytes);
            }
            if let Some(bytes) = col_styles.as_ref().filter(|_| key.3) {
                hash_bytes(hasher, bytes);
            }
            if let Some(bytes) = row_styles.as_ref().filter(|_| key.4) {
                hash_bytes(hasher, bytes);
            }
        }
    }
    Ok(hashers
        .into_iter()
        .zip(planned)
        .map(|(hasher, (key, _))| finish_hash(hasher, key.6))
        .collect())
}

fn encode_shared(sheet: &Sheet, sink: &mut impl FingerprintSink) {
    hash_bytes(sink, sheet.name.as_bytes());
    hash_u64(sink, sheet.iter_cells().count() as u64);
    for (at, cell) in sheet.iter_cells() {
        xlsx_model::open_timing! {
            xlsx_model::open_timing::count("fingerprint_cells", 1);
        }
        xlsx_model::progress::tick();
        hash_u32(sink, at.row);
        hash_u32(sink, at.col);
        hash_cell_value(sink, &cell.value);
        match &cell.formula {
            Some(formula) => {
                sink.feed(&[1]);
                hash_bytes(sink, formula.as_bytes());
            }
            None => sink.feed(&[0]),
        }
        match cell.style {
            Some(style) => {
                sink.feed(&[1]);
                hash_u32(sink, style);
            }
            None => sink.feed(&[0]),
        }
    }
    hash_u64(sink, sheet.merges.len() as u64);
    for range in &sheet.merges {
        hash_cell_ref(sink, range.start);
        hash_cell_ref(sink, range.end);
    }
}

fn write_dimensions(
    columns: &BTreeMap<u32, f64>,
    rows: &BTreeMap<u32, f64>,
    sink: &mut impl FingerprintSink,
) {
    for values in [columns, rows] {
        hash_u64(sink, values.len() as u64);
        for (&at, &value) in values {
            hash_u32(sink, at);
            hash_u64(sink, value.to_bits());
        }
    }
}

fn encode_freeze_pane(pane: Option<FreezePane>, sink: &mut impl FingerprintSink) {
    match pane {
        Some(pane) => {
            sink.feed(&[1]);
            hash_u32(sink, pane.rows);
            hash_u32(sink, pane.cols);
            hash_cell_ref(sink, pane.top_left);
        }
        None => sink.feed(&[0]),
    }
}

fn finish_hash(hasher: Sha256, reading: FingerprintReading) -> (String, u64) {
    let digest = hasher.finalize();
    let fingerprint = fingerprint_text(reading, format!("{digest:x}"));
    let mut head = [0; 8];
    head.copy_from_slice(&digest[..8]);
    (fingerprint, client_id_of_digest(u64::from_be_bytes(head)))
}

struct Stream<'a> {
    hashers: &'a mut [Sha256],
    bytes: Vec<u8>,
}

impl Stream<'_> {
    fn flush(&mut self) {
        for hasher in self.hashers.iter_mut() {
            hasher.update(&self.bytes);
        }
        self.bytes.clear();
    }
}

impl FingerprintSink for Stream<'_> {
    fn feed(&mut self, mut bytes: &[u8]) {
        while !bytes.is_empty() {
            let count = bytes.len().min(CHUNK_BYTES - self.bytes.len());
            self.bytes.extend_from_slice(&bytes[..count]);
            bytes = &bytes[count..];
            if self.bytes.len() == CHUNK_BYTES {
                self.flush();
            }
        }
    }
}

enum EncodeError {
    Budget,
    Invalid(String),
}

struct Budget(usize);

impl Budget {
    fn encode(
        &mut self,
        encode: impl FnOnce(&mut Limited) -> Result<(), String>,
    ) -> Result<Bytes, EncodeError> {
        let mut sink = Limited {
            bytes: Vec::new(),
            limit: self.0,
            exceeded: false,
        };
        let result = encode(&mut sink);
        if sink.exceeded {
            return Err(EncodeError::Budget);
        }
        result.map_err(EncodeError::Invalid)?;
        self.0 -= sink.bytes.len();
        Ok(sink.bytes.into_boxed_slice())
    }

    fn json(&mut self, value: &impl Serialize, what: &str) -> Result<Bytes, EncodeError> {
        self.encode(|sink| {
            serde_json::to_writer(sink, value)
                .map_err(|error| format!("cannot fingerprint sheet {what}: {error}"))
        })
    }
}

struct Limited {
    bytes: Vec<u8>,
    limit: usize,
    exceeded: bool,
}

impl FingerprintSink for Limited {
    fn feed(&mut self, bytes: &[u8]) {
        if bytes.len() > self.limit - self.bytes.len() {
            self.exceeded = true;
        }
        if !self.exceeded {
            self.bytes.extend_from_slice(bytes);
        }
    }
}

impl std::io::Write for Limited {
    fn write(&mut self, bytes: &[u8]) -> std::io::Result<usize> {
        self.feed(bytes);
        if self.exceeded {
            Err(std::io::Error::other("fingerprint payload budget exceeded"))
        } else {
            Ok(bytes.len())
        }
    }

    fn flush(&mut self) -> std::io::Result<()> {
        Ok(())
    }
}
