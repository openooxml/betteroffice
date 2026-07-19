use thiserror::Error;

#[derive(Debug, Error, PartialEq, Eq)]
pub enum PptxError {
    #[error("PPTX container error: {0}")]
    Container(String),
    #[error("missing required PPTX part {0}")]
    MissingPart(String),
    #[error("malformed XML in {part} at byte {offset}: {message}")]
    MalformedXml {
        part: String,
        offset: u64,
        message: String,
    },
    #[error("unsafe XML in {part}: {kind}")]
    UnsafeXml { part: String, kind: &'static str },
    #[error("PPTX resource limit {kind} exceeded in {part}")]
    ResourceLimit { part: String, kind: &'static str },
    #[error("invalid relationship target {target} from {source_part}")]
    InvalidRelationship { source_part: String, target: String },
}
