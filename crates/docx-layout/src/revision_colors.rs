//! Word's tracked-change colors.

/// Word's 20 by-author colors, taken in the order authors first appear.
pub const AUTHOR_PALETTE: [&str; 20] = [
    "#D13438", "#0078D4", "#5C2E91", "#498205", "#CC3595", "#7160E8", "#038387", "#6D5700",
    "#CF0F1F", "#4E6AED", "#B146C2", "#394146", "#0B6A0B", "#CA5010", "#750B1C", "#5D5A58",
    "#881798", "#69797E", "#005B70", "#8E562E",
];

/// Moves ignore the author and paint in this color.
pub const MOVE_COLOR: &str = "#2C6234";

/// Change bars use the automatic (black) color.
pub const CHANGE_BAR_COLOR: &str = "#000000";

/// Word draws change bars this far left of the left page margin: 36pt.
pub const CHANGE_BAR_MARGIN_OFFSET_PX: f64 = 48.0;

/// 0.72pt.
pub const CHANGE_BAR_WIDTH_PX: f64 = 0.96;

/// The color of `author` given the document's authors in first-appearance order. An author
/// the list lacks takes a stable slot after the listed ones.
pub fn author_color(authors: &[String], author: Option<&str>) -> &'static str {
    let author = author.filter(|name| !name.is_empty()).unwrap_or("Unknown");
    let index = authors
        .iter()
        .position(|listed| listed == author)
        .unwrap_or_else(|| authors.len() + fnv1a(author) as usize % AUTHOR_PALETTE.len());
    AUTHOR_PALETTE[index % AUTHOR_PALETTE.len()]
}

/// Word attributes list renumbering to "Unknown", registered after the document's authors.
pub fn renumbering_color(authors: &[String]) -> &'static str {
    if authors.iter().any(|author| author == "Unknown") {
        author_color(authors, Some("Unknown"))
    } else {
        AUTHOR_PALETTE[authors.len() % AUTHOR_PALETTE.len()]
    }
}

fn fnv1a(text: &str) -> u32 {
    text.bytes().fold(0x811c_9dc5, |hash, byte| {
        (hash ^ u32::from(byte)).wrapping_mul(0x0100_0193)
    })
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn authors_take_palette_slots_in_order_and_cycle() {
        let authors: Vec<String> = (0..21).map(|index| format!("A{index}")).collect();
        assert_eq!(author_color(&authors, Some("A0")), "#D13438");
        assert_eq!(author_color(&authors, Some("A1")), "#0078D4");
        assert_eq!(author_color(&authors, Some("A19")), "#8E562E");
        assert_eq!(author_color(&authors, Some("A20")), "#D13438");
    }

    #[test]
    fn missing_and_unlisted_authors_resolve_stably() {
        let authors = vec!["Unknown".to_owned(), "B".to_owned()];
        assert_eq!(author_color(&authors, None), "#D13438");
        assert_eq!(author_color(&authors, Some("")), "#D13438");
        assert_eq!(
            author_color(&authors, Some("Zed")),
            author_color(&authors, Some("Zed"))
        );
    }

    #[test]
    fn renumbering_takes_the_slot_after_the_document_authors() {
        assert_eq!(renumbering_color(&["A".to_owned()]), "#0078D4");
        assert_eq!(
            renumbering_color(&["Unknown".to_owned(), "A".to_owned()]),
            "#D13438"
        );
    }
}
