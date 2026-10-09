//! The images of an opened package, by `media:{n}` token.
//!
//! Opening parses images as tokens naming their package part
//! ([`docx_parse::media::MediaTable`]). Seeding writes those tokens into the
//! stories only when asked to; by default it writes the parts' `data:` URLs,
//! which every replica can read, and remembers a keyed digest of each so that
//! layout, display lists and frames still carry the short token. A host
//! resolving tokens reads the bytes from the same package.

use std::collections::HashMap;
use std::hash::Hasher;
use std::sync::{Arc, Mutex};

use docx_parse::media::{MediaTable, media_token, media_token_index};
use yrs::Any;

use crate::raw::RawOp;
use crate::{SegmentContent, StorySegment};

/// Digests of the `data:` URLs seeding wrote, each to the index of the part
/// it came from.
#[derive(Clone, Debug, Default)]
pub struct MediaSources(Arc<Sources>);

#[derive(Debug, Default)]
struct Sources {
    key: (u64, u64),
    parts: HashMap<u64, u32>,
    seen: Mutex<Seen>,
}

/// Sources found to match, by address. Holding each keeps its address from
/// naming another string; an entry only the cache still holds is dropped once
/// the cache has doubled since it was last pruned.
#[derive(Debug, Default)]
struct Seen {
    sources: HashMap<usize, (Arc<str>, u32)>,
    prune_at: usize,
}

impl Seen {
    fn new(sources: HashMap<usize, (Arc<str>, u32)>) -> Self {
        let prune_at = Self::next_prune(sources.len());
        Self { sources, prune_at }
    }

    fn next_prune(len: usize) -> usize {
        (len * 2).max(64)
    }

    fn insert(&mut self, address: usize, source: Arc<str>, index: u32) {
        if self.sources.len() >= self.prune_at {
            self.sources
                .retain(|_, (source, _)| Arc::strong_count(source) > 1);
            self.prune_at = Self::next_prune(self.sources.len());
        }
        self.sources.insert(address, (source, index));
    }
}

impl PartialEq for MediaSources {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
            || (self.0.key == other.0.key && self.0.parts == other.0.parts)
    }
}

impl MediaSources {
    pub fn is_empty(&self) -> bool {
        self.0.parts.is_empty()
    }

    /// The token of the part a `data:` image source seeding wrote came from.
    pub fn token(&self, src: &Arc<str>) -> Option<String> {
        if self.is_empty() || !src.starts_with("data:") {
            return None;
        }
        let address = Arc::as_ptr(src) as *const u8 as usize;
        let mut seen = self.0.seen.lock().unwrap();
        let index = match seen.sources.get(&address) {
            Some((_, index)) => *index,
            None => {
                let index = *self.0.parts.get(&self.digest(src))?;
                seen.insert(address, Arc::clone(src), index);
                index
            }
        };
        Some(media_token(index as usize))
    }

    /// [`Self::token`] for a source read out of JSON.
    pub fn token_of(&self, src: &str) -> Option<String> {
        if self.is_empty() || !src.starts_with("data:") {
            return None;
        }
        self.0
            .parts
            .get(&self.digest(src))
            .map(|index| media_token(*index as usize))
    }

    fn digest(&self, src: &str) -> u64 {
        digest(self.0.key, src)
    }

    /// The key and digests, for a replica that did not seed.
    pub fn to_json(&self) -> String {
        let parts: Vec<(String, u32)> = self
            .0
            .parts
            .iter()
            .map(|(digest, index)| (format!("{digest:016x}"), *index))
            .collect();
        serde_json::json!({
            "key": [format!("{:016x}", self.0.key.0), format!("{:016x}", self.0.key.1)],
            "parts": parts,
        })
        .to_string()
    }

    pub fn from_json(json: &str) -> Result<Self, String> {
        #[derive(serde::Deserialize)]
        struct Wire {
            key: (String, String),
            parts: Vec<(String, u32)>,
        }
        let hex = |value: &str| u64::from_str_radix(value, 16).map_err(|error| error.to_string());
        let wire: Wire = serde_json::from_str(json).map_err(|error| error.to_string())?;
        Ok(Self(Arc::new(Sources {
            key: (hex(&wire.key.0)?, hex(&wire.key.1)?),
            parts: wire
                .parts
                .iter()
                .map(|(digest, index)| Ok((hex(digest)?, *index)))
                .collect::<Result<_, String>>()?,
            seen: Mutex::default(),
        })))
    }
}

/// SipHash-2-4 of the whole source under a per-document random key: equal
/// sources share a digest, and no peer can make another source match one.
#[allow(deprecated)]
fn digest(key: (u64, u64), src: &str) -> u64 {
    let mut hasher = std::hash::SipHasher::new_with_keys(key.0, key.1);
    hasher.write(src.as_bytes());
    hasher.finish()
}

/// Payload keys whose string is the JSON of a parsed object that can hold
/// images.
const JSON_PAYLOADS: &[&str] = &["shapeJson", "chartJson", "fieldData", "propertiesJson"];

/// Replaces the tokens of seed `ops` with the `data:` URLs of their parts,
/// each built once and shared. With `layout_tokens`, returns the sources that
/// let layout carry the tokens again. Fails when a part cannot be read, as
/// inflating the whole package would.
pub(crate) fn write_data_urls<'a>(
    ops: impl IntoIterator<Item = &'a mut RawOp>,
    table: &MediaTable,
    layout_tokens: bool,
) -> Result<MediaSources, String> {
    let mut writer = DataUrls::new(table);
    for op in ops {
        match op {
            RawOp::InsertEmbed { payload, .. } => {
                for (key, value) in payload.iter_mut() {
                    writer.any(key, value);
                }
            }
            RawOp::SetEmbedAttr { key, value, .. } => writer.any(key, value),
            RawOp::SetComment { body, .. } => writer.any("", body),
            RawOp::Insert { .. }
            | RawOp::Delete { .. }
            | RawOp::Format { .. }
            | RawOp::RemoveComment { .. } => {}
        }
    }
    if let Some(error) = writer.error {
        return Err(error);
    }
    if !layout_tokens {
        return Ok(MediaSources::default());
    }
    let [k0, k1] = crate::identity::entropy();
    let key = (k0, k1);
    let mut parts = HashMap::new();
    let mut seen = HashMap::new();
    for (index, url) in writer.urls {
        if let Some(url) = url {
            parts.insert(digest(key, &url), index as u32);
            seen.insert(Arc::as_ptr(&url) as *const u8 as usize, (url, index as u32));
        }
    }
    Ok(MediaSources(Arc::new(Sources {
        key,
        parts,
        seen: Mutex::new(Seen::new(seen)),
    })))
}

pub(crate) fn write_segment_data_urls(
    segments: &mut [StorySegment],
    table: &MediaTable,
) -> Result<(), String> {
    let mut writer = DataUrls::new(table);
    for segment in segments {
        let payload = match &mut segment.content {
            SegmentContent::Pilcrow(properties) => Some(&mut properties.values),
            SegmentContent::OtherEmbed { payload, .. } => Some(payload),
            SegmentContent::Text(_) => None,
        };
        if let Some(payload) = payload {
            for (key, value) in payload.iter_mut() {
                writer.any(key, value);
            }
        }
        for (key, value) in segment.attributes.iter_mut() {
            writer.any(key, value);
        }
    }
    match writer.error {
        Some(error) => Err(error),
        None => Ok(()),
    }
}

struct DataUrls<'a> {
    table: &'a MediaTable,
    urls: HashMap<usize, Option<Arc<str>>>,
    error: Option<String>,
}

impl<'a> DataUrls<'a> {
    fn new(table: &'a MediaTable) -> Self {
        Self {
            table,
            urls: HashMap::new(),
            error: None,
        }
    }

    fn url(&mut self, index: usize) -> Option<Arc<str>> {
        let table = self.table;
        let error = &mut self.error;
        self.urls
            .entry(index)
            .or_insert_with(|| match table.data_url(index) {
                Ok(url) => Some(Arc::from(url)),
                Err(message) => {
                    error.get_or_insert(message);
                    None
                }
            })
            .clone()
    }

    fn any(&mut self, key: &str, value: &mut Any) {
        if let Some(written) = self.written(key, value) {
            *value = written;
        }
    }

    /// `value` with its tokens written as URLs, or `None` when it has none.
    fn written(&mut self, key: &str, value: &Any) -> Option<Any> {
        match value {
            Any::String(text) if key == "src" => media_token_index(text)
                .and_then(|index| self.url(index))
                .map(Any::String),
            Any::String(text)
                if JSON_PAYLOADS.contains(&key) && text.contains("\"src\":\"media:") =>
            {
                Some(Any::String(self.json(text).into()))
            }
            Any::Array(items) => {
                let written: Vec<Option<Any>> =
                    items.iter().map(|item| self.written("", item)).collect();
                written.iter().any(Option::is_some).then(|| {
                    Any::Array(
                        written
                            .into_iter()
                            .zip(items.iter())
                            .map(|(written, item)| written.unwrap_or_else(|| item.clone()))
                            .collect::<Vec<_>>()
                            .into(),
                    )
                })
            }
            Any::Map(map) => {
                let written: Vec<(&String, Any)> = map
                    .iter()
                    .filter_map(|(key, item)| Some((key, self.written(key, item)?)))
                    .collect();
                (!written.is_empty()).then(|| {
                    let mut map = (**map).clone();
                    for (key, item) in written {
                        map.insert(key.clone(), item);
                    }
                    Any::Map(Arc::new(map))
                })
            }
            _ => None,
        }
    }

    /// `json` with each `"src":"media:{n}"` holding the part's `data:` URL.
    fn json(&mut self, json: &str) -> String {
        const KEY: &str = "\"src\":\"";
        let mut sources = Vec::new();
        let mut len = json.len();
        let mut rest = json;
        while let Some(at) = rest.find(KEY) {
            let value = &rest[at + KEY.len()..];
            let end = value.find('"').unwrap_or(value.len());
            let url = media_token_index(&value[..end]).and_then(|index| self.url(index));
            if let Some(url) = &url {
                len = len - end + url.len();
            }
            sources.push((&rest[..at + KEY.len()], &value[..end], url));
            rest = &value[end..];
        }
        let mut out = String::with_capacity(len);
        for (before, token, url) in sources {
            out.push_str(before);
            out.push_str(url.as_deref().unwrap_or(token));
        }
        out.push_str(rest);
        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn sources(entries: &[(&str, u32)]) -> MediaSources {
        let key = (7, 11);
        MediaSources(Arc::new(Sources {
            key,
            parts: entries
                .iter()
                .map(|(src, index)| (digest(key, src), *index))
                .collect(),
            seen: Mutex::default(),
        }))
    }

    #[test]
    fn a_source_maps_only_when_every_byte_matches() {
        let source = format!("data:image/bmp;base64,{}", "A".repeat(10_000));
        let mut changed = source.clone().into_bytes();
        changed[7_777] = b'B';
        let changed = String::from_utf8(changed).unwrap();
        let sources = sources(&[(&source, 3)]);
        assert_eq!(
            sources.token(&Arc::from(source.as_str())).as_deref(),
            Some("media:3")
        );
        assert_eq!(sources.token(&Arc::from(changed.as_str())), None);
        assert_eq!(sources.token_of(&source).as_deref(), Some("media:3"));
        assert_eq!(sources.token_of("media:3"), None);
    }

    #[test]
    fn matched_sources_no_longer_held_elsewhere_are_released() {
        let source = "data:image/png;base64,AAAA";
        let sources = sources(&[(source, 2)]);
        let kept: Arc<str> = Arc::from(source);
        assert_eq!(sources.token(&kept).as_deref(), Some("media:2"));
        for _ in 0..1000 {
            let copy: Arc<str> = Arc::from(source);
            assert_eq!(sources.token(&copy).as_deref(), Some("media:2"));
        }
        let seen = sources.0.seen.lock().unwrap();
        assert!(seen.sources.len() <= 64);
        assert!(
            seen.sources
                .values()
                .any(|(source, _)| Arc::ptr_eq(source, &kept))
        );
    }

    #[test]
    fn sources_round_trip_and_compare_by_content() {
        let sources = sources(&[("data:a", 3), ("data:b", 4)]);
        let loaded = MediaSources::from_json(&sources.to_json()).unwrap();
        assert_eq!(loaded, sources);
        assert_eq!(loaded.token_of("data:b").as_deref(), Some("media:4"));
        assert_eq!(MediaSources::default(), MediaSources::default());
        assert!(MediaSources::from_json(r#"{"key":["zz","0"],"parts":[]}"#).is_err());
    }

    #[test]
    fn json_payload_sources_become_data_urls() {
        let bytes = ooxml_opc::rezip_parts(&[
            ("word/media/a.png".to_owned(), vec![1, 2, 3, 4]),
            ("word/media/b.png".to_owned(), vec![5]),
        ])
        .unwrap();
        let table =
            MediaTable::new(ooxml_opc::RetainedPackage::new(Arc::from(bytes)).unwrap()).unwrap();
        let index = |path| {
            (0..table.len())
                .find(|&i| table.path(i) == Some(path))
                .unwrap()
        };
        let (a, b) = (index("word/media/a.png"), index("word/media/b.png"));
        let (url_a, url_b) = (
            "data:image/png;base64,AQIDBA==",
            "data:image/png;base64,BQ==",
        );
        assert_eq!(table.data_url(a).unwrap(), url_a);
        assert_eq!(table.data_url(b).unwrap(), url_b);

        let mut writer = DataUrls::new(&table);
        let json = format!(
            r#"{{"a":{{"src":"media:{a}"}},"b":[{{"src":"media:{b}","label":"media:{a}"}},{{"src":"https://x"}}],"c":{{"src":"media:{a}"}}}}"#
        );
        assert_eq!(
            writer.json(&json),
            format!(
                r#"{{"a":{{"src":"{url_a}"}},"b":[{{"src":"{url_b}","label":"media:{a}"}},{{"src":"https://x"}}],"c":{{"src":"{url_a}"}}}}"#
            )
        );
        assert!(writer.error.is_none());
        assert_eq!(writer.json(r#"{"src":"media:9"}"#), r#"{"src":"media:9"}"#);
        assert!(writer.error.is_some());
    }
}
