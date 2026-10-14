use std::fmt;

#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    Parse(pptx_parse::PptxError),
    Edit(pptx_edit::EditError),
    Proposal(pptx_edit::ProposalError),
    Render(pptx_render::RenderError),
    /// A headless export read unusable bytes or refused its options.
    Export(pptx_edit::structured::ExportError),
    /// The raster backend refused a surface or could not paint a primitive.
    Raster(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Parse(error) => error.fmt(formatter),
            Self::Edit(error) => error.fmt(formatter),
            Self::Proposal(error) => error.fmt(formatter),
            Self::Render(error) => error.fmt(formatter),
            Self::Export(error) => error.fmt(formatter),
            Self::Raster(message) => formatter.write_str(message),
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Parse(error) => Some(error),
            Self::Edit(error) => Some(error),
            Self::Proposal(error) => Some(error),
            Self::Render(error) => Some(error),
            Self::Export(error) => Some(error),
            Self::Raster(_) => None,
        }
    }
}

impl From<pptx_parse::PptxError> for Error {
    fn from(error: pptx_parse::PptxError) -> Self {
        Self::Parse(error)
    }
}

impl From<pptx_edit::EditError> for Error {
    fn from(error: pptx_edit::EditError) -> Self {
        Self::Edit(error)
    }
}

impl From<pptx_edit::ProposalError> for Error {
    fn from(error: pptx_edit::ProposalError) -> Self {
        Self::Proposal(error)
    }
}

impl From<pptx_render::RenderError> for Error {
    fn from(error: pptx_render::RenderError) -> Self {
        Self::Render(error)
    }
}

impl From<pptx_edit::structured::ExportError> for Error {
    fn from(error: pptx_edit::structured::ExportError) -> Self {
        Self::Export(error)
    }
}

impl From<pptx_edit::structured::ExportFailure> for Error {
    fn from(failure: pptx_edit::structured::ExportFailure) -> Self {
        Self::Export(pptx_edit::structured::ExportError::Refused(failure))
    }
}
