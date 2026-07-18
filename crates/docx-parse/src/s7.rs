//! S7 body projection with complete recursively typed table content.

use base64::Engine as _;
use serde::{Deserialize, Serialize};

use crate::canonical::{canonical_sha256, from_serializable, to_canonical_bytes};
use crate::s6::{S6Projection, parse_docx_story_projection};
use crate::xml::ParseError;

pub type S7Projection = S6Projection;

#[derive(Clone, Debug, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct S7WireEnvelope {
    pub wire_version: u8,
    pub projection: S7Projection,
    pub canonical_base64: String,
    pub canonical_sha256: String,
}

pub fn parse_docx_s7_projection(data: &[u8]) -> Result<S7Projection, ParseError> {
    parse_docx_story_projection(data)
}

pub fn s7_wire_envelope(projection: S7Projection) -> Result<S7WireEnvelope, ParseError> {
    let canonical =
        from_serializable(&projection).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let bytes =
        to_canonical_bytes(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    let sha =
        canonical_sha256(&canonical).map_err(|error| ParseError::Canonical(error.to_string()))?;
    Ok(S7WireEnvelope {
        wire_version: 1,
        projection,
        canonical_base64: base64::engine::general_purpose::STANDARD.encode(bytes),
        canonical_sha256: sha,
    })
}

pub fn parse_docx_s7_wire(data: &[u8]) -> Result<S7WireEnvelope, ParseError> {
    s7_wire_envelope(parse_docx_s7_projection(data)?)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::block::BlockContent;

    #[test]
    fn package_projection_keeps_typed_nested_tables_and_s6_stays_a_boundary() {
        let package = ooxml_opc::rezip_parts(&[(
            "word/document.xml".to_owned(),
            br#"<w:document xmlns:w="w"><w:body><w:tbl><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>{tableName}</w:t></w:r></w:p><w:tbl><w:tr><w:tc><w:p><w:r><w:t>{nestedName}</w:t></w:r></w:p></w:tc></w:tr></w:tbl></w:tc></w:tr></w:tbl></w:body></w:document>"#.to_vec(),
        )])
        .unwrap();
        let s7 = parse_docx_s7_wire(&package).unwrap();
        let BlockContent::Table(table) = &s7.projection.body.content[0] else {
            panic!("table")
        };
        assert_eq!(table.column_widths.as_deref(), Some(&[1000.0][..]));
        assert!(matches!(
            table.rows[0].cells[0].content[1],
            BlockContent::Table(_)
        ));
        assert_eq!(
            s7.projection.template_variables,
            ["tableName", "nestedName"]
        );

        let s6 = crate::s6::parse_docx_s6_projection(&package).unwrap();
        let BlockContent::Table(boundary) = &s6.body.content[0] else {
            panic!("table boundary")
        };
        assert!(boundary.rows.is_empty());
        assert!(boundary.formatting.is_none());
        assert!(boundary.column_widths.is_none());
        assert!(s6.template_variables.is_empty());
    }
}
