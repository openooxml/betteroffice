//! Sheet renames and deletions in the formulas a worksheet keeps as the file
//! wrote them: data validations, which the model does not hold, and the
//! extension list, whose `xm:f` formulas back x14 data validations, x14
//! conditional formats and sparklines. Each source sheet's name maps to the
//! current name of that same sheet, so the patch follows sheet identity, not
//! rename history; a reference to a deleted source sheet becomes `#REF!`, or
//! a 3-D span narrows, as Excel writes it. Only formula text that names an
//! edited sheet changes.

use quick_xml::Reader;
use quick_xml::events::Event;
use xlsx_model::sheet_qualifier::{
    remove_sheet_references_noting_spans, rename_sheets, sheet_names_equal,
};

use crate::package::XmlTemplate;
use crate::xml::{formula_text, resolve_entity};

/// What a preserved formula belongs to, for a refusal or a save warning.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FragmentKind {
    DataValidation,
    X14DataValidation,
    X14ConditionalFormat,
    Sparkline,
    Extension,
}

impl FragmentKind {
    pub fn element(self) -> &'static str {
        match self {
            Self::DataValidation => "data-validation",
            Self::X14DataValidation => "x14-data-validation",
            Self::X14ConditionalFormat => "x14-conditional-format",
            Self::Sparkline => "sparkline",
            Self::Extension => "extension",
        }
    }
}

/// Why a fragment cannot follow the rename.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum FragmentFailureReason {
    /// The new name holds a character xml cannot carry in formula text.
    UnwritableName,
    /// The fragment does not read as formulas this patch can rewrite.
    Unreadable,
    /// A deleted sheet ends a 3-D span in the fragment, and sheets moved or
    /// were added since the file was read, so the order the span narrows
    /// against at deletion is not the file's.
    OrderChanged,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FragmentFailure {
    pub kind: FragmentKind,
    pub reason: FragmentFailureReason,
}

/// The sheet edits since the package was read that preserved formulas follow:
/// the source sheets in their order in the file, those deleted since, and each
/// kept source sheet's name in the file with its current name.
#[derive(Clone, Debug, Default)]
pub(crate) struct SheetEdits {
    pub(crate) source_order: Vec<String>,
    pub(crate) removed: Vec<String>,
    pub(crate) renamed: Vec<(String, String)>,
    /// The sheet order is no longer the file's: a kept source sheet moved, or
    /// a sheet was added, either of which a 3-D span may now cover.
    pub(crate) moved: bool,
}

impl SheetEdits {
    pub(crate) fn is_empty(&self) -> bool {
        self.removed.is_empty() && self.renamed.is_empty()
    }

    /// `text` with the deletions applied first, each against the sheets the
    /// ones before it left, then the renames; plus which renames it wrote.
    /// `None` when a deletion would narrow a 3-D span after sheets moved,
    /// since the order that span narrows against is no longer known.
    fn apply(&self, text: &str) -> Option<(String, Vec<usize>)> {
        let mut order = self.source_order.clone();
        let mut current = text.to_owned();
        for removed in &self.removed {
            if let Some((next, narrowed)) =
                remove_sheet_references_noting_spans(&current, removed, &order)
            {
                if narrowed && self.moved {
                    return None;
                }
                current = next;
            }
            order.retain(|sheet| !sheet_names_equal(sheet, removed));
        }
        Some(rename_sheets(&current, &self.renamed))
    }
}

/// `template` with `edits` applied to its preserved formulas, or `None` when
/// no formula names an edited sheet.
pub(crate) fn rename_preserved_fragments(
    template: &XmlTemplate,
    edits: &SheetEdits,
) -> Result<Option<XmlTemplate>, FragmentFailure> {
    if edits.is_empty() {
        return Ok(None);
    }
    let mut patched = None;
    for (index, child) in template.children.iter().enumerate() {
        let in_extension = match child.local_name.as_str() {
            "dataValidations" => false,
            "extLst" => true,
            _ => continue,
        };
        if !template.is_root_child_named(child, &child.local_name) {
            continue;
        }
        if let Some(bytes) = rename_fragment(&child.bytes, in_extension, edits)? {
            patched.get_or_insert_with(|| template.clone()).children[index].bytes = bytes;
        }
    }
    Ok(patched)
}

fn kind_of(stack: &[Vec<u8>], in_extension: bool) -> FragmentKind {
    if !in_extension {
        return FragmentKind::DataValidation;
    }
    let has = |name: &[u8]| stack.iter().any(|entry| entry.as_slice() == name);
    if has(b"dataValidation") {
        FragmentKind::X14DataValidation
    } else if has(b"cfRule") || has(b"conditionalFormatting") {
        FragmentKind::X14ConditionalFormat
    } else if has(b"sparklineGroup") {
        FragmentKind::Sparkline
    } else {
        FragmentKind::Extension
    }
}

fn is_formula(name: &[u8], stack: &[Vec<u8>], in_extension: bool) -> bool {
    if in_extension {
        name == b"f"
    } else {
        matches!(name, b"formula1" | b"formula2")
            && stack
                .last()
                .is_some_and(|parent| parent.as_slice() == b"dataValidation")
    }
}

/// `bytes`, one preserved child, with each formula naming a renamed sheet
/// rewritten in place; `None` when nothing changes.
fn rename_fragment(
    bytes: &[u8],
    in_extension: bool,
    sheet_edits: &SheetEdits,
) -> Result<Option<Vec<u8>>, FragmentFailure> {
    let unreadable = |kind| FragmentFailure {
        kind,
        reason: FragmentFailureReason::Unreadable,
    };
    let fallback = if in_extension {
        FragmentKind::Extension
    } else {
        FragmentKind::DataValidation
    };
    let Ok(source) = std::str::from_utf8(bytes) else {
        return Err(unreadable(fallback));
    };
    let mut reader = Reader::from_str(source);
    let mut stack: Vec<Vec<u8>> = Vec::new();
    let mut formula: Option<(usize, String, FragmentKind)> = None;
    let mut edits = Vec::new();
    loop {
        let start = reader.buffer_position() as usize;
        let event = match reader.read_event() {
            Ok(event) => event,
            Err(_) => return Err(unreadable(kind_of(&stack, in_extension))),
        };
        match event {
            Event::Start(element) => {
                let name = element.local_name().as_ref().to_vec();
                if let Some((_, _, kind)) = formula {
                    return Err(unreadable(kind));
                }
                if is_formula(&name, &stack, in_extension) {
                    formula = Some((
                        reader.buffer_position() as usize,
                        String::new(),
                        kind_of(&stack, in_extension),
                    ));
                }
                stack.push(name);
            }
            Event::Text(text) => {
                if let Some((_, value, kind)) = formula.as_mut() {
                    let decoded = text.xml10_content().map_err(|_| unreadable(*kind))?;
                    value.push_str(&decoded);
                }
            }
            Event::CData(text) => {
                if let Some((_, value, kind)) = formula.as_mut() {
                    let raw = text.into_inner();
                    value.push_str(std::str::from_utf8(&raw).map_err(|_| unreadable(*kind))?);
                }
            }
            Event::GeneralRef(reference) => {
                if let Some((_, value, kind)) = formula.as_mut() {
                    let name = reference.decode().map_err(|_| unreadable(*kind))?;
                    value.push_str(&resolve_entity(&name).map_err(|_| unreadable(*kind))?);
                }
            }
            Event::End(element) => {
                let name = element.local_name().as_ref().to_vec();
                if stack.pop().as_deref() != Some(name.as_slice()) {
                    return Err(unreadable(kind_of(&stack, in_extension)));
                }
                if let Some((content, value, kind)) = formula.take() {
                    let Some((rewritten, used)) = sheet_edits.apply(&value) else {
                        return Err(FragmentFailure {
                            kind,
                            reason: FragmentFailureReason::OrderChanged,
                        });
                    };
                    if rewritten != value {
                        if used
                            .iter()
                            .any(|&at| !xml_writable(&sheet_edits.renamed[at].1))
                        {
                            return Err(FragmentFailure {
                                kind,
                                reason: FragmentFailureReason::UnwritableName,
                            });
                        }
                        edits.push((content..start, formula_text(&rewritten)));
                    }
                }
            }
            Event::Empty(_) => {
                if let Some((_, _, kind)) = formula {
                    return Err(unreadable(kind));
                }
            }
            Event::Eof => break,
            _ => {}
        }
    }
    if formula.is_some() || !stack.is_empty() {
        return Err(unreadable(fallback));
    }
    if edits.is_empty() {
        return Ok(None);
    }
    let mut out = source.to_owned();
    for (span, replacement) in edits.into_iter().rev() {
        out.replace_range(span, &replacement);
    }
    Ok(Some(out.into_bytes()))
}

fn xml_writable(name: &str) -> bool {
    name.chars().all(|ch| {
        matches!(ch, '\t' | '\n' | '\r') || (ch >= ' ' && ch != '\u{fffe}' && ch != '\u{ffff}')
    })
}

#[cfg(test)]
mod tests {
    use super::{FragmentFailureReason, FragmentKind, SheetEdits, rename_fragment};

    fn renamed(pairs: &[(&str, &str)]) -> SheetEdits {
        SheetEdits {
            renamed: pairs
                .iter()
                .map(|(old, new)| ((*old).to_owned(), (*new).to_owned()))
                .collect(),
            ..SheetEdits::default()
        }
    }

    fn patched(xml: &str, extension: bool, pairs: &[(&str, &str)]) -> Option<String> {
        rename_fragment(xml.as_bytes(), extension, &renamed(pairs))
            .unwrap()
            .map(|bytes| String::from_utf8(bytes).unwrap())
    }

    #[test]
    fn a_validation_formula_follows_the_rename_and_the_rest_keeps_its_bytes() {
        let xml = r#"<dataValidations count="1"><dataValidation type="list" sqref="B1" prompt="Data!A1"><formula1>Data!$A$1:$A$5</formula1></dataValidation></dataValidations>"#;
        assert_eq!(
            patched(xml, false, &[("Data", "New Data")]).as_deref(),
            Some(
                xml.replace(">Data!$A$1:$A$5<", ">&apos;New Data&apos;!$A$1:$A$5<")
                    .as_str()
            )
        );
        assert_eq!(patched(xml, false, &[("Other", "New")]), None);
    }

    #[test]
    fn every_extension_formula_follows_the_rename() {
        let xml = r#"<extLst><ext uri="u" xmlns:x14="x" xmlns:xm="m"><x14:sparklineGroups><x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Data!A1:C1</xm:f><xm:sqref>D1</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup></x14:sparklineGroups></ext></extLst>"#;
        assert_eq!(
            patched(xml, true, &[("Data", "B")]).as_deref(),
            Some(xml.replace("<xm:f>Data!", "<xm:f>B!").as_str())
        );
    }

    #[test]
    fn a_name_xml_cannot_carry_and_a_fragment_that_does_not_read_fail() {
        let xml = r#"<dataValidations><dataValidation><formula1>Data!A1</formula1></dataValidation></dataValidations>"#;
        let error =
            rename_fragment(xml.as_bytes(), false, &renamed(&[("Data", "D\u{1}")])).unwrap_err();
        assert_eq!(error.reason, FragmentFailureReason::UnwritableName);
        assert_eq!(error.kind, FragmentKind::DataValidation);
        let broken = r#"<dataValidations><dataValidation><formula1>Data!A1<x/></formula1></dataValidation></dataValidations>"#;
        let error =
            rename_fragment(broken.as_bytes(), false, &renamed(&[("Data", "New")])).unwrap_err();
        assert_eq!(error.reason, FragmentFailureReason::Unreadable);
    }

    #[test]
    fn a_deleted_sheet_becomes_ref_and_a_span_narrows_before_renames_apply() {
        let edits = SheetEdits {
            source_order: ["Jan", "Feb", "Mar", "Data"].map(str::to_owned).to_vec(),
            removed: vec!["Jan".to_owned(), "Data".to_owned()],
            renamed: vec![("Feb".to_owned(), "February".to_owned())],
            moved: false,
        };
        let xml = r#"<dataValidations><dataValidation><formula1>SUM(Jan:Mar!A1)+Data!$A$1:$A$5</formula1></dataValidation></dataValidations>"#;
        let patched = String::from_utf8(
            rename_fragment(xml.as_bytes(), false, &edits)
                .unwrap()
                .unwrap(),
        )
        .unwrap();
        assert_eq!(
            patched,
            xml.replace(
                "SUM(Jan:Mar!A1)+Data!$A$1:$A$5",
                "SUM(February:Mar!A1)+#REF!"
            )
        );
    }

    #[test]
    fn a_span_a_deletion_would_narrow_after_sheets_moved_is_refused() {
        let edits = SheetEdits {
            source_order: ["Jan", "Feb", "Mar"].map(str::to_owned).to_vec(),
            removed: vec!["Jan".to_owned()],
            renamed: Vec::new(),
            moved: true,
        };
        let span = r#"<dataValidations><dataValidation><formula1>SUM(Jan:Mar!A1)</formula1></dataValidation></dataValidations>"#;
        let error = rename_fragment(span.as_bytes(), false, &edits).unwrap_err();
        assert_eq!(error.reason, FragmentFailureReason::OrderChanged);
        let single = r#"<dataValidations><dataValidation><formula1>Jan!A1</formula1></dataValidation></dataValidations>"#;
        assert!(
            rename_fragment(single.as_bytes(), false, &edits)
                .unwrap()
                .is_some()
        );
    }
}
