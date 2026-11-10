//! Every part and relationship type the engines' loaders name, how preflight
//! answers it, and the tests that show that answer.
//! `scripts/font-preflight-inputs.test.ts` fails when a loader names an input
//! this table does not list, and when a Modelled, Incomplete or Refused row
//! names no existing test. The scan is partial: it sees only string literals
//! of the forms `relationships/<kind>`, `"/<kind>"` and `"ppt|xl|word/<path>.xml"`
//! without digits, in the listed crates. Rows below marked "by hand" are
//! inputs it cannot see. `NotPainted` is an engine-side claim checked by
//! review, not by a test.

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
    /// No test shows preflight's answer for it yet.
    Unverified,
}

use Answer::*;

pub(crate) const PPTX: &[(&str, Answer, &[&str])] = &[
    (
        "officeDocument",
        Modelled,
        &["main_part_named_by_the_package_relationship_is_read_and_incomplete"],
    ),
    // By hand: read only for the content types of media parts.
    ("[Content_Types].xml", NotPainted, &[]),
    // By hand: embedded font parts, registered as available faces.
    ("font", Unverified, &[]),
    (
        "ppt/presentation.xml",
        Modelled,
        &["presentation_default_text_style_names_a_face"],
    ),
    (
        "slide",
        Modelled,
        &["active_slide_excludes_unused_slides_and_theme_lists"],
    ),
    (
        "slideLayout",
        Modelled,
        &["inherited_shapes_hidden_groups_and_show_master_shapes"],
    ),
    (
        "slideMaster",
        Modelled,
        &["inherited_shapes_hidden_groups_and_show_master_shapes"],
    ),
    (
        "theme",
        Modelled,
        &["active_slide_excludes_unused_slides_and_theme_lists"],
    ),
    (
        "tableStyles",
        Modelled,
        &["table_styles_apply_only_to_painted_cells"],
    ),
    (
        "ppt/tableStyles.xml",
        Modelled,
        &["table_styles_load_the_conventional_part_without_a_relationship"],
    ),
    ("chart", Incomplete, &["chart_relationship_is_incomplete"]),
    (
        "diagramData",
        Incomplete,
        &["diagram_relationships_are_incomplete"],
    ),
    (
        "diagramDrawing",
        Incomplete,
        &["diagram_relationships_are_incomplete"],
    ),
    (
        "diagramLayout",
        Incomplete,
        &["diagram_relationships_are_incomplete"],
    ),
    ("image", NotPainted, &[]),
    ("vmlDrawing", NotPainted, &[]),
    ("notesMaster", NotPainted, &[]),
    ("notesSlide", NotPainted, &[]),
    ("comment", NotPainted, &[]),
    ("comments", NotPainted, &[]),
    ("commentAuthors", NotPainted, &[]),
    ("authors", NotPainted, &[]),
    ("ppt/authors.xml", NotPainted, &[]),
    ("ppt/commentAuthors.xml", NotPainted, &[]),
    // Built by format! in pptx-parse comments.rs; not seen by the scan.
    ("ppt/comments/comment{n}.xml", NotPainted, &[]),
    ("ppt/comments/modernComment{n}.xml", NotPainted, &[]),
];

pub(crate) const XLSX: &[(&str, Answer, &[&str])] = &[
    // The engine reads xl/workbook.xml by path, as preflight does.
    (
        "officeDocument",
        Modelled,
        &["package_main_relationship_is_ignored"],
    ),
    // By hand: the theme fallback path. The engine reads only its colours.
    (
        "xl/theme/theme1.xml",
        Incomplete,
        &["theme_font_without_a_cached_name_is_incomplete"],
    ),
    (
        "chartsheet",
        Incomplete,
        &["active_chartsheet_never_reads_a_worksheet_by_position"],
    ),
    (
        "dialogsheet",
        Incomplete,
        &["active_chartsheet_never_reads_a_worksheet_by_position"],
    ),
    ("table", NotPainted, &[]),
    ("comments", NotPainted, &[]),
    ("threadedComment", NotPainted, &[]),
    (
        "xl/workbook.xml",
        Modelled,
        &["active_visible_sheet_values_and_cell_styles_only"],
    ),
    (
        "worksheet",
        Modelled,
        &["active_visible_sheet_values_and_cell_styles_only"],
    ),
    // The positional fallback, built by format! in xlsx-parse package.rs and
    // read.rs; not seen by the scan. Preflight falls back the same way.
    (
        "xl/worksheets/sheet{n}.xml",
        Modelled,
        &["styles_fall_back_to_the_canonical_part_without_a_relationship"],
    ),
    (
        "styles",
        Modelled,
        &["active_visible_sheet_values_and_cell_styles_only"],
    ),
    (
        "xl/styles.xml",
        Modelled,
        &["styles_fall_back_to_the_canonical_part_without_a_relationship"],
    ),
    (
        "sharedStrings",
        Modelled,
        &["shared_and_inline_strings_attach_only_used_scripts"],
    ),
    (
        "xl/sharedStrings.xml",
        Modelled,
        &["shared_strings_fall_back_to_the_canonical_part_without_a_relationship"],
    ),
    (
        "theme",
        Incomplete,
        &["theme_font_without_a_cached_name_is_incomplete"],
    ),
    ("drawing", Incomplete, &["sheet_drawings_are_incomplete"]),
    (
        "chart",
        Incomplete,
        &["chart_relationship_in_a_sheet_drawing_is_incomplete"],
    ),
    (
        "hyperlink",
        Incomplete,
        &["hyperlink_relationship_is_incomplete"],
    ),
    ("pivotTable", NotPainted, &[]),
    ("pivotCacheRecords", NotPainted, &[]),
    ("xl/calcChain.xml", NotPainted, &[]),
];

pub(crate) const DOCX: &[(&str, Answer, &[&str])] = &[
    (
        "officeDocument",
        Modelled,
        &["main_part_named_by_the_package_relationship"],
    ),
    (
        "word/document.xml",
        Modelled,
        &["doc_defaults_stored_and_deflated"],
    ),
    (
        "styles",
        Modelled,
        &["duplicate_style_id_takes_the_default_flag_of_the_last_definition"],
    ),
    (
        "word/styles.xml",
        Modelled,
        &["doc_defaults_stored_and_deflated"],
    ),
    (
        "numbering",
        Modelled,
        &["duplicate_numbering_ids_take_the_last_definition"],
    ),
    (
        "word/numbering.xml",
        Modelled,
        &["numbering_reads_only_used_levels_and_instances"],
    ),
    ("theme", Modelled, &["canonical_theme_is_complete"]),
    (
        "settings",
        Modelled,
        &["duplicate_theme_script_fonts_take_the_last_entry"],
    ),
    (
        "word/settings.xml",
        Modelled,
        &["theme_language_fills_empty_slots"],
    ),
    (
        "header",
        Modelled,
        &["header_relationship_and_unused_header"],
    ),
    ("footer", Modelled, &["footer_relationship_names_a_face"]),
    (
        "footnotes",
        Modelled,
        &["normal_notes_and_separators_match_the_editor_request"],
    ),
    (
        "word/footnotes.xml",
        Modelled,
        &["notes_at_the_conventional_path_without_a_relationship"],
    ),
    (
        "endnotes",
        Modelled,
        &["endnotes_relationship_names_a_face"],
    ),
    (
        "word/endnotes.xml",
        Modelled,
        &["notes_at_the_conventional_path_without_a_relationship"],
    ),
    // Also the owner of embedded font parts, registered as available faces.
    ("fontTable", Unverified, &[]),
    ("word/fontTable.xml", Unverified, &[]),
    ("font", Unverified, &[]),
    // By hand: the engine reads the theme at this fixed path.
    (
        "word/theme/theme1.xml",
        Modelled,
        &["canonical_theme_is_complete"],
    ),
    (
        "diagramLayout",
        Incomplete,
        &["unsupported_drawing_relationships_are_incomplete"],
    ),
    (
        "diagramDrawing",
        Incomplete,
        &["unsupported_drawing_relationships_are_incomplete"],
    ),
    // By hand: built by format! in docx-parse smart_art.rs; reached only
    // through w:drawing, which preflight always answers incomplete.
    (
        "word/diagrams/drawing{n}.xml",
        Incomplete,
        &["unsupported_drawing_relationships_are_incomplete"],
    ),
    (
        "chart",
        Incomplete,
        &["unsupported_drawing_relationships_are_incomplete"],
    ),
    (
        "diagramData",
        Incomplete,
        &["unsupported_drawing_relationships_are_incomplete"],
    ),
    (
        "oleObject",
        Incomplete,
        &["unsupported_drawing_relationships_are_incomplete"],
    ),
    (
        "wordprocessingCanvas",
        Incomplete,
        &["unsupported_drawing_namespaces_are_incomplete"],
    ),
    (
        "wordprocessingGroup",
        Incomplete,
        &["unsupported_drawing_namespaces_are_incomplete"],
    ),
    ("image", NotPainted, &[]),
    ("hyperlink", NotPainted, &[]),
    ("customXml", NotPainted, &[]),
    ("webSettings", NotPainted, &[]),
    ("comments", NotPainted, &[]),
    ("word/comments.xml", NotPainted, &[]),
    ("commentsExtended", NotPainted, &[]),
    ("word/commentsExtended.xml", NotPainted, &[]),
    ("commentsExtensible", NotPainted, &[]),
    ("word/commentsExtensible.xml", NotPainted, &[]),
    ("commentsIds", NotPainted, &[]),
    ("word/commentsIds.xml", NotPainted, &[]),
];

#[test]
fn tables_name_each_input_once() {
    for table in [PPTX, XLSX, DOCX] {
        let mut names: Vec<_> = table.iter().map(|(name, _, _)| *name).collect();
        names.sort_unstable();
        let count = names.len();
        names.dedup();
        assert_eq!(names.len(), count);
        for (_, answer, tests) in table {
            if [Modelled, Incomplete, Refused].contains(answer) {
                assert!(!tests.is_empty());
            } else if *answer == Unverified {
                assert!(tests.is_empty());
            }
        }
    }
}
