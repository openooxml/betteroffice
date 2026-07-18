//! Pure OOXML image wrap taxonomy shared by drawing/story consumers.

pub fn is_floating_wrap_type(value: Option<&str>) -> bool {
    matches!(
        value,
        Some("square" | "tight" | "through" | "behind" | "inFront")
    )
}

pub fn is_wrap_none(value: Option<&str>) -> bool {
    matches!(value, Some("behind" | "inFront"))
}

pub fn wraps_around_text(value: Option<&str>) -> bool {
    matches!(value, Some("square" | "tight" | "through"))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn matches_the_typescript_taxonomy() {
        for value in ["square", "tight", "through"] {
            assert!(is_floating_wrap_type(Some(value)));
            assert!(wraps_around_text(Some(value)));
            assert!(!is_wrap_none(Some(value)));
        }
        assert!(!is_floating_wrap_type(Some("topAndBottom")));
        assert!(!is_floating_wrap_type(Some("inline")));
        assert!(is_wrap_none(Some("behind")));
        assert!(is_wrap_none(Some("inFront")));
        assert!(!wraps_around_text(Some("garbage")));
    }
}
