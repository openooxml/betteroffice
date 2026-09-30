//! The images of an opened package, by `media:{n}` token.
//!
//! Opening parses images as tokens naming their package part
//! ([`docx_parse::media::MediaTable`]). Seeding writes those tokens into the
//! stories only when asked to; by default it writes the parts' `data:` URLs,
//! which every replica can read, and remembers a fingerprint of each so that
//! layout, display lists and frames still carry the short token. A host
//! resolving tokens reads the bytes from the same package.

use std::borrow::Cow;
use std::collections::HashMap;
use std::sync::Arc;

use docx_parse::media::{MediaTable, media_token, media_token_index};
use yrs::Any;

use crate::raw::RawOp;

/// Fingerprints of the `data:` URLs seeding wrote, each to the index of the
/// part it came from. Equal when shared.
#[derive(Clone, Debug, Default)]
pub struct MediaSources(Arc<HashMap<u64, u32>>);

impl PartialEq for MediaSources {
    fn eq(&self, other: &Self) -> bool {
        Arc::ptr_eq(&self.0, &other.0)
    }
}

impl MediaSources {
    pub fn is_empty(&self) -> bool {
        self.0.is_empty()
    }

    /// The token layout carries for an image source: a token as it is, and a
    /// `data:` URL seeding wrote as the token of its part.
    pub fn token<'a>(&self, src: &'a str) -> Cow<'a, str> {
        if src.starts_with("data:")
            && let Some(index) = self.0.get(&fingerprint(src))
        {
            return Cow::Owned(media_token(*index as usize));
        }
        Cow::Borrowed(src)
    }

    /// `[[fingerprint hex, index], …]`, for a replica that did not seed.
    pub fn to_json(&self) -> String {
        let entries: Vec<(String, u32)> = self
            .0
            .iter()
            .map(|(fingerprint, index)| (format!("{fingerprint:016x}"), *index))
            .collect();
        serde_json::to_string(&entries).expect("fingerprints serialize")
    }

    pub fn from_json(json: &str) -> Result<Self, String> {
        let entries: Vec<(String, u32)> =
            serde_json::from_str(json).map_err(|error| error.to_string())?;
        entries
            .into_iter()
            .map(|(fingerprint, index)| {
                u64::from_str_radix(&fingerprint, 16)
                    .map(|fingerprint| (fingerprint, index))
                    .map_err(|error| error.to_string())
            })
            .collect::<Result<HashMap<_, _>, _>>()
            .map(|entries| Self(Arc::new(entries)))
    }
}

/// Bytes sampled from each end and across the middle of a source.
const ENDS: usize = 256;
const SAMPLES: usize = 64;
const SAMPLE: usize = 16;

/// A 64-bit FNV-1a digest of a source's length, both ends and evenly spaced
/// samples: equal for one `data:` URL, and for two different images only by
/// matching byte for byte at every sampled offset.
fn fingerprint(src: &str) -> u64 {
    let bytes = src.as_bytes();
    let mut hash = 0xcbf2_9ce4_8422_2325_u64;
    let mut feed = |chunk: &[u8]| {
        for byte in chunk {
            hash = (hash ^ u64::from(*byte)).wrapping_mul(0x0000_0100_0000_01b3);
        }
    };
    feed(&(bytes.len() as u64).to_le_bytes());
    feed(&bytes[..bytes.len().min(ENDS)]);
    feed(&bytes[bytes.len().saturating_sub(ENDS)..]);
    for sample in 0..SAMPLES {
        let at = bytes.len() * sample / SAMPLES;
        feed(&bytes[at..(at + SAMPLE).min(bytes.len())]);
    }
    hash
}

/// Replaces the tokens of seed `ops` with the `data:` URLs of their parts,
/// each built once and shared, and returns the fingerprints of those URLs.
/// A token inside a JSON string, such as a shape's source, is replaced in
/// place under its `"src"` key.
pub(crate) fn write_data_urls<'a>(
    ops: impl IntoIterator<Item = &'a mut RawOp>,
    table: &MediaTable,
) -> MediaSources {
    let mut writer = DataUrls {
        table,
        urls: HashMap::new(),
    };
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
    MediaSources(Arc::new(
        writer
            .urls
            .into_iter()
            .filter_map(|(index, url)| Some((fingerprint(&url?), index as u32)))
            .collect(),
    ))
}

struct DataUrls<'a> {
    table: &'a MediaTable,
    urls: HashMap<usize, Option<Arc<str>>>,
}

impl DataUrls<'_> {
    fn url(&mut self, index: usize) -> Option<Arc<str>> {
        self.urls
            .entry(index)
            .or_insert_with(|| self.table.data_url(index).ok().map(Arc::from))
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
            Any::String(text) if text.contains("\"src\":\"media:") => {
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
        let mut out = String::with_capacity(json.len());
        let mut rest = json;
        while let Some(at) = rest.find(KEY) {
            let value = &rest[at + KEY.len()..];
            let end = value.find('"').unwrap_or(value.len());
            out.push_str(&rest[..at + KEY.len()]);
            match media_token_index(&value[..end]).and_then(|index| self.url(index)) {
                Some(url) => out.push_str(&url),
                None => out.push_str(&value[..end]),
            }
            rest = &value[end..];
        }
        out.push_str(rest);
        out
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn fingerprints_tell_sources_apart_by_any_sampled_byte() {
        let source = format!("data:image/png;base64,{}", "A".repeat(10_000));
        let mut changed = source.clone().into_bytes();
        changed[5_000] = b'B';
        let changed = String::from_utf8(changed).unwrap();
        assert_eq!(fingerprint(&source), fingerprint(&source.clone()));
        assert_ne!(fingerprint(&source), fingerprint(&changed));
        assert_ne!(fingerprint(&source), fingerprint(&format!("{source}A")));
        assert_eq!(fingerprint(""), fingerprint(""));
    }

    #[test]
    fn sources_round_trip_and_map_only_what_seeding_wrote() {
        let sources = MediaSources(Arc::new(HashMap::from([(fingerprint("data:a"), 3)])));
        let loaded = MediaSources::from_json(&sources.to_json()).unwrap();
        assert_eq!(loaded.token("data:a"), "media:3");
        assert_eq!(loaded.token("data:b"), "data:b");
        assert_eq!(loaded.token("media:7"), "media:7");
        assert!(MediaSources::from_json("[[\"zz\",1]]").is_err());
    }
}
