use std::collections::HashMap;

use anyhow::Result;
use docx_parse::{
    RelationshipTarget, is_footer_relationship, is_header_relationship, is_image_relationship,
    parse_docx_relationship_parts, resolve_relationship_target,
};
use docx_raster::{ImageMap, ImageScope, NoteKind, scoped_image_key};
use vello::peniko::{Blob, ImageAlphaType, ImageData, ImageFormat};

pub struct ImageRegistry {
    pub raw: ImageMap,
    decoded: HashMap<String, ImageData>,
}

impl ImageRegistry {
    pub fn load(docx: &[u8]) -> Result<Self> {
        let parts = ooxml_opc::unzip_parts(docx).map_err(anyhow::Error::msg)?;
        let package = parts
            .into_iter()
            .map(|(path, bytes)| (path.to_lowercase(), bytes))
            .collect::<HashMap<_, _>>();
        let relationships = parse_docx_relationship_parts(docx)?;
        let body_part = relationships.relationship_parts.iter().find(|part| {
            part.path
                .eq_ignore_ascii_case("word/_rels/document.xml.rels")
        });
        let mut header_footer_ids = HashMap::new();
        if let Some(body_part) = body_part {
            for (_, relationship) in &body_part.relationships {
                if !(is_header_relationship(relationship) || is_footer_relationship(relationship)) {
                    continue;
                }
                if let RelationshipTarget::Internal(target) =
                    resolve_relationship_target(&body_part.path, relationship)?
                {
                    header_footer_ids.insert(target.to_lowercase(), relationship.id.clone());
                }
            }
        }

        let mut raw = HashMap::new();
        let mut decoded = HashMap::new();
        for part in &relationships.relationship_parts {
            let Some(owner) = relationship_owner(&part.path) else {
                continue;
            };
            let owner_lower = owner.to_lowercase();
            for (_, relationship) in &part.relationships {
                if !is_image_relationship(relationship) {
                    continue;
                }
                let RelationshipTarget::Internal(target) =
                    resolve_relationship_target(&part.path, relationship)?
                else {
                    continue;
                };
                let Some(bytes) = package.get(&target.to_lowercase()) else {
                    continue;
                };
                let key = if owner_lower == "word/document.xml" {
                    scoped_image_key(ImageScope::Body, &relationship.id)
                } else if owner_lower == "word/footnotes.xml" {
                    scoped_image_key(ImageScope::Notes(NoteKind::Footnote), &relationship.id)
                } else if owner_lower == "word/endnotes.xml" {
                    scoped_image_key(ImageScope::Notes(NoteKind::Endnote), &relationship.id)
                } else if let Some(r_id) = header_footer_ids.get(&owner_lower) {
                    scoped_image_key(ImageScope::HeaderFooter(r_id), &relationship.id)
                } else {
                    continue;
                };
                raw.insert(key.clone(), bytes.clone());
                if let Ok(image) = image::load_from_memory(bytes) {
                    let rgba = image.into_rgba8();
                    let (width, height) = rgba.dimensions();
                    decoded.insert(
                        key,
                        ImageData {
                            data: Blob::from(rgba.into_raw()),
                            format: ImageFormat::Rgba8,
                            alpha_type: ImageAlphaType::Alpha,
                            width,
                            height,
                        },
                    );
                }
            }
        }
        Ok(Self { raw, decoded })
    }

    pub fn get(&self, scope: ImageScope<'_>, rel_id: &str) -> Option<&ImageData> {
        let key = scoped_image_key(scope, rel_id);
        self.decoded.get(&key)
    }
}

fn relationship_owner(path: &str) -> Option<String> {
    let normalized = path.replace('\\', "/");
    if let Some((directory, file)) = normalized.rsplit_once("/_rels/") {
        return Some(format!("{directory}/{}", file.strip_suffix(".rels")?));
    }
    let file = normalized.strip_prefix("_rels/")?;
    Some(file.strip_suffix(".rels")?.to_owned())
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn resolves_relationship_owner_parts() {
        assert_eq!(
            relationship_owner("word/_rels/header1.xml.rels").as_deref(),
            Some("word/header1.xml")
        );
        assert_eq!(relationship_owner("_rels/.rels").as_deref(), Some(""));
    }
}
