//! Embedded media table and the incumbent image-resolution aliases.

use base64::Engine as _;
use indexmap::IndexMap;
use serde::{Deserialize, Serialize};

use crate::relationships::RelationshipMap;

pub type MediaMap = IndexMap<String, MediaFile>;

#[derive(Clone, Debug, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct MediaFile {
    pub path: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filename: Option<String>,
    pub mime_type: String,
    pub base64: String,
    pub data_url: String,
}

#[derive(Clone, Debug, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ResolvedImageData {
    #[serde(skip_serializing_if = "Option::is_none")]
    pub src: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mime_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filename: Option<String>,
}

pub fn build_media_map(parts: &[(String, Vec<u8>)]) -> MediaMap {
    let mut media = MediaMap::new();
    for (path, data) in parts {
        if !path.to_ascii_lowercase().starts_with("word/media/") {
            continue;
        }
        let filename = path.rsplit('/').next().unwrap_or(path).to_owned();
        let mime_type = media_mime_type(path).to_owned();
        let base64 = base64::engine::general_purpose::STANDARD.encode(data);
        let file = MediaFile {
            path: path.clone(),
            filename: Some(filename),
            mime_type: mime_type.clone(),
            data_url: format!("data:{mime_type};base64,{base64}"),
            base64,
        };
        media.insert(path.clone(), file.clone());
        if let Some(normalized) = path.strip_prefix("word/") {
            media.insert(normalized.to_owned(), file);
        }
    }
    media
}

pub fn resolve_image_data(
    relationship_id: &str,
    relationships: Option<&RelationshipMap>,
    media: Option<&MediaMap>,
) -> ResolvedImageData {
    if relationship_id.is_empty() {
        return ResolvedImageData::default();
    }
    let Some(relationship) = relationships.and_then(|map| map.get(relationship_id)) else {
        return ResolvedImageData::default();
    };
    if relationship.target.is_empty() {
        return ResolvedImageData::default();
    }
    let target = &relationship.target;
    let normalized = normalize_media_path(target);
    let filename = target.rsplit('/').next().map(str::to_owned);
    if let Some(media) = media {
        for candidate in [
            normalized,
            target.trim_start_matches('/').to_owned(),
            format!("word/{}", target.trim_start_matches('/')),
        ] {
            if let Some(file) = find_case_insensitive(media, &candidate) {
                return ResolvedImageData {
                    src: Some(if file.data_url.is_empty() {
                        file.base64.clone()
                    } else {
                        file.data_url.clone()
                    }),
                    mime_type: Some(file.mime_type.clone()),
                    filename,
                };
            }
        }
    }
    ResolvedImageData {
        src: None,
        mime_type: Some(media_mime_type(target).to_owned()),
        filename,
    }
}

pub fn media_mime_type(path: &str) -> &'static str {
    match path
        .rsplit('.')
        .next()
        .unwrap_or_default()
        .to_ascii_lowercase()
        .as_str()
    {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        "tif" | "tiff" => "image/tiff",
        "webp" => "image/webp",
        "svg" => "image/svg+xml",
        "emf" => "image/x-emf",
        "wmf" => "image/x-wmf",
        _ => "application/octet-stream",
    }
}

fn normalize_media_path(path: &str) -> String {
    let path = path.trim_start_matches('/');
    if path.starts_with("media/") {
        format!("word/{path}")
    } else if path.starts_with("word/") {
        path.to_owned()
    } else {
        format!("word/{path}")
    }
}

fn find_case_insensitive<'a>(media: &'a MediaMap, path: &str) -> Option<&'a MediaFile> {
    media
        .iter()
        .find(|(candidate, _)| candidate.eq_ignore_ascii_case(path))
        .map(|(_, file)| file)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::relationships::{Relationship, TargetMode};

    #[test]
    fn aliases_bytes_and_case_insensitive_resolution_match_the_incumbent() {
        let parts = vec![("word/media/IMAGE.PNG".to_owned(), vec![0, 255, 16])];
        let media = build_media_map(&parts);
        assert_eq!(
            media.keys().collect::<Vec<_>>(),
            ["word/media/IMAGE.PNG", "media/IMAGE.PNG"]
        );
        let relationships = RelationshipMap::from([(
            "rId1".into(),
            Relationship {
                id: "rId1".into(),
                relationship_type: "image".into(),
                target: "media/image.png".into(),
                target_mode: Some(TargetMode::External),
            },
        )]);
        let resolved = resolve_image_data("rId1", Some(&relationships), Some(&media));
        // TargetMode is deliberately not used as a fetch signal. Resolution
        // only consults already embedded package bytes and performs no I/O.
        assert_eq!(resolved.mime_type.as_deref(), Some("image/png"));
        assert_eq!(resolved.filename.as_deref(), Some("image.png"));
        assert_eq!(resolved.src.as_deref(), Some("data:image/png;base64,AP8Q"));
    }

    #[test]
    fn missing_media_returns_only_extension_metadata() {
        let relationships = RelationshipMap::from([(
            "rId2".into(),
            Relationship {
                id: "rId2".into(),
                relationship_type: "image".into(),
                target: "../outside.WMF".into(),
                target_mode: None,
            },
        )]);
        assert_eq!(
            resolve_image_data("rId2", Some(&relationships), None),
            ResolvedImageData {
                src: None,
                mime_type: Some("image/x-wmf".into()),
                filename: Some("outside.WMF".into()),
            }
        );
    }
}
