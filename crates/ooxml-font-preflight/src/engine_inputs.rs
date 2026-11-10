//! Every part and relationship type the engines' loaders name, and how
//! preflight answers it. `scripts/font-preflight-inputs.test.ts` scans the
//! engine crates and fails when a loader names one this table does not list.

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Answer {
    /// Preflight reads it the way the engine does.
    Modelled,
    /// Preflight answers `complete: false` when the document uses it.
    Incomplete,
    /// Preflight refuses the document with an error instead of answering.
    Refused,
    /// The engine reads it but draws no text with document faces on the opening view.
    NotPainted,
    /// The engine reads it only to drop requested faces, so preflight lists a superset.
    Superset,
}

use Answer::*;

pub(crate) const PPTX: &[(&str, Answer)] = &[
    ("officeDocument", Refused),
    ("ppt/presentation.xml", Modelled),
    ("slide", Modelled),
    ("slideLayout", Modelled),
    ("slideMaster", Modelled),
    ("theme", Modelled),
    ("tableStyles", Modelled),
    ("ppt/tableStyles.xml", Modelled),
    ("chart", Incomplete),
    ("diagramData", Incomplete),
    ("diagramDrawing", Incomplete),
    ("diagramLayout", Incomplete),
    ("image", NotPainted),
    ("vmlDrawing", NotPainted),
    ("notesMaster", NotPainted),
    ("notesSlide", NotPainted),
    ("comment", NotPainted),
    ("comments", NotPainted),
    ("commentAuthors", NotPainted),
    ("authors", NotPainted),
    ("ppt/authors.xml", NotPainted),
    ("ppt/commentAuthors.xml", NotPainted),
];

pub(crate) const XLSX: &[(&str, Answer)] = &[
    ("officeDocument", Refused),
    ("xl/workbook.xml", Modelled),
    ("worksheet", Modelled),
    ("styles", Modelled),
    ("xl/styles.xml", Modelled),
    ("sharedStrings", Modelled),
    ("xl/sharedStrings.xml", Modelled),
    ("theme", Incomplete),
    ("drawing", Incomplete),
    ("chart", Incomplete),
    ("hyperlink", Incomplete),
    ("pivotTable", NotPainted),
    ("pivotCacheRecords", NotPainted),
    ("xl/calcChain.xml", NotPainted),
];

pub(crate) const DOCX: &[(&str, Answer)] = &[
    ("officeDocument", Refused),
    ("word/document.xml", Modelled),
    ("styles", Modelled),
    ("word/styles.xml", Modelled),
    ("numbering", Modelled),
    ("word/numbering.xml", Modelled),
    ("theme", Modelled),
    ("settings", Modelled),
    ("word/settings.xml", Modelled),
    ("header", Modelled),
    ("footer", Modelled),
    ("footnotes", Modelled),
    ("word/footnotes.xml", Modelled),
    ("endnotes", Modelled),
    ("word/endnotes.xml", Modelled),
    ("fontTable", Superset),
    ("word/fontTable.xml", Superset),
    ("chart", Incomplete),
    ("diagramData", Incomplete),
    ("oleObject", Incomplete),
    ("wordprocessingCanvas", Incomplete),
    ("wordprocessingGroup", Incomplete),
    ("image", NotPainted),
    ("hyperlink", NotPainted),
    ("customXml", NotPainted),
    ("webSettings", NotPainted),
    ("comments", NotPainted),
    ("word/comments.xml", NotPainted),
    ("commentsExtended", NotPainted),
    ("word/commentsExtended.xml", NotPainted),
    ("commentsExtensible", NotPainted),
    ("word/commentsExtensible.xml", NotPainted),
    ("commentsIds", NotPainted),
    ("word/commentsIds.xml", NotPainted),
];

#[test]
fn tables_name_each_input_once() {
    for table in [PPTX, XLSX, DOCX] {
        let mut names: Vec<_> = table.iter().map(|(name, _)| *name).collect();
        names.sort_unstable();
        let count = names.len();
        names.dedup();
        assert_eq!(names.len(), count);
    }
}
