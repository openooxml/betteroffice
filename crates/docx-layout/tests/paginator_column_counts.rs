use docx_layout::page_flow::Paginator;
use docx_layout::section_breaks::resolve_page_margins;
use docx_layout::types::{ColumnLayout, Size};

fn columns(count: f64) -> ColumnLayout {
    ColumnLayout {
        count,
        gap: 6.0,
        equal_width: None,
        separator: Some(true),
        columns: None,
    }
}

fn paginator(count: f64) -> Paginator {
    Paginator::new(
        Size {
            w: 816.0,
            h: 1056.0,
        },
        resolve_page_margins(None),
        columns(count),
        None,
    )
    .unwrap()
}

fn width(count: f64) -> f64 {
    paginator(count).get_column_width(0)
}

#[test]
fn paginator_clamps_column_counts() {
    assert_eq!(width(1e9), width(45.0));
    assert_eq!(width(f64::INFINITY), width(45.0));
    assert_eq!(width(f64::NAN), width(1.0));
    assert_eq!(width(0.0), width(1.0));
    assert_eq!(width(-3.0), width(1.0));
    assert_ne!(width(2.0), width(3.0));
}

#[test]
fn paginator_clamps_updated_column_counts() {
    let mut paginator = paginator(2.0);
    paginator.update_columns(columns(1e9));
    assert_eq!(paginator.get_column_width(0), width(45.0));
}
