use thiserror::Error;

#[derive(Debug, Error, PartialEq, Eq)]
pub enum VsdxError {
    #[error("VSDX container error: {0}")]
    Container(String),
    #[error("missing required VSDX part {0}")]
    MissingPart(String),
    #[error("unsupported VSDX document kind {0:?}")]
    UnsupportedDocumentKind(ooxml_opc::DocumentKind),
    #[error("conflicting main document relationships: {0:?}")]
    ConflictingMainDocumentRelationships(Vec<String>),
    #[error("malformed XML in {part} at byte {offset}: {message}")]
    MalformedXml {
        part: String,
        offset: u64,
        message: String,
    },
    #[error("unsafe XML in {part}: {kind}")]
    UnsafeXml { part: String, kind: &'static str },
    #[error("VSDX resource limit {kind} exceeded in {part}")]
    ResourceLimit { part: String, kind: &'static str },
    #[error("invalid relationship target {target} from {source_part}")]
    InvalidRelationship { source_part: String, target: String },
    #[error("invalid lexical patch span")]
    InvalidSpan,
    #[error("attribute value contains a character forbidden by XML 1.0")]
    InvalidXmlCharacter,
    #[error("lexical patch limit {kind} exceeded")]
    PatchLimit { kind: &'static str },
    #[error("invalid cell edit in {part}: {message}")]
    InvalidCellEdit { part: String, message: String },
}
