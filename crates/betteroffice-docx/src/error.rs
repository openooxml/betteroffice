use std::fmt;

#[derive(Debug)]
#[non_exhaustive]
pub enum Error {
    Parse(docx_parse::ParseError),
    Edit(docx_edit::EditError),
    Operation(docx_edit::OpError),
    Layout(docx_layout::LayoutError),
    DisplayList(String),
    ParagraphNotFound(String),
    UnsupportedParagraphEdit(String),
}

impl fmt::Display for Error {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        match self {
            Self::Parse(error) => error.fmt(formatter),
            Self::Edit(error) => error.fmt(formatter),
            Self::Operation(error) => error.fmt(formatter),
            Self::Layout(error) => error.fmt(formatter),
            Self::DisplayList(error) => formatter.write_str(error),
            Self::ParagraphNotFound(id) => write!(formatter, "paragraph {id:?} was not found"),
            Self::UnsupportedParagraphEdit(id) => {
                write!(
                    formatter,
                    "paragraph {id:?} is not a plain single-run paragraph"
                )
            }
        }
    }
}

impl std::error::Error for Error {
    fn source(&self) -> Option<&(dyn std::error::Error + 'static)> {
        match self {
            Self::Parse(error) => Some(error),
            Self::Edit(error) => Some(error),
            Self::Operation(error) => Some(error),
            Self::Layout(error) => Some(error),
            Self::DisplayList(_)
            | Self::ParagraphNotFound(_)
            | Self::UnsupportedParagraphEdit(_) => None,
        }
    }
}

impl From<docx_parse::ParseError> for Error {
    fn from(error: docx_parse::ParseError) -> Self {
        Self::Parse(error)
    }
}

impl From<docx_edit::EditError> for Error {
    fn from(error: docx_edit::EditError) -> Self {
        Self::Edit(error)
    }
}

impl From<docx_edit::OpError> for Error {
    fn from(error: docx_edit::OpError) -> Self {
        Self::Operation(error)
    }
}

impl From<docx_layout::LayoutError> for Error {
    fn from(error: docx_layout::LayoutError) -> Self {
        Self::Layout(error)
    }
}
