use betteroffice_pptx::{
    CommandReadResponse, DeckSnapshot, EditCtx, EditError, Error, GraphicFrameData, HistoryReceipt,
    PictureDraft, Presentation, ReadRequest, ShapeRect, TableColumnWidth, TableEdit,
    TableInsertPosition, TableReceipt, TextParagraphDraft, TextRunDraft, TextStyle,
};
use pptx_edit::DeckSession;

const FIXTURE: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const CLIENT_ID: u64 = 73;

struct Pair {
    facade: Presentation,
    session: DeckSession,
    context: EditCtx,
}

impl Pair {
    fn open(bytes: &[u8]) -> Self {
        let pair = Self {
            facade: Presentation::open_collaborative(bytes, CLIENT_ID).unwrap(),
            session: DeckSession::open(bytes, CLIENT_ID).unwrap(),
            context: EditCtx::local("edit-parity"),
        };
        pair.assert_snapshot();
        pair
    }

    fn demo() -> Self {
        Self::open(FIXTURE)
    }

    fn table() -> Self {
        let mut parts = ooxml_opc::unzip_parts(FIXTURE).unwrap();
        let (_, bytes) = parts
            .iter_mut()
            .find(|(name, _)| name == "ppt/slides/slide1.xml")
            .unwrap();
        let mut xml = String::from_utf8(bytes.clone()).unwrap();
        let mut table = String::from(
            r#"<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="10000" name="Parity table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="800000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr/><a:tblGrid><a:gridCol w="1000000"/><a:gridCol w="1000000"/></a:tblGrid>"#,
        );
        for _ in 0..2 {
            table.push_str(r#"<a:tr h="400000">"#);
            for _ in 0..2 {
                table.push_str(
                    "<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Cell</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc>",
                );
            }
            table.push_str("</a:tr>");
        }
        table.push_str("</a:tbl></a:graphicData></a:graphic></p:graphicFrame>");
        let index = xml.find("</p:spTree>").unwrap();
        xml.insert_str(index, &table);
        *bytes = xml.into_bytes();
        Self::open(&ooxml_opc::rezip_parts(&parts).unwrap())
    }

    fn snapshot(&self) -> DeckSnapshot {
        self.facade.snapshot().unwrap()
    }

    fn assert_snapshot(&self) {
        assert_eq!(self.snapshot(), self.session.snapshot().unwrap());
    }

    fn table_target(&self) -> (String, String) {
        let snapshot = self.snapshot();
        let slide = &snapshot.slides[0];
        let shape = slide
            .shapes
            .iter()
            .find(|shape| matches!(&shape.graphic, Some(GraphicFrameData::Table(_))))
            .unwrap();
        (slide.id.clone(), shape.id.clone())
    }

    fn assert_refusal<T, U>(
        &self,
        facade: betteroffice_pptx::Result<T>,
        session: Result<U, EditError>,
    ) {
        let Err(Error::Edit(actual)) = facade else {
            panic!("expected a facade edit error");
        };
        let Err(expected) = session else {
            panic!("expected a session edit error");
        };
        assert_eq!(
            std::mem::discriminant(&actual),
            std::mem::discriminant(&expected)
        );
        assert_eq!(actual.to_string(), expected.to_string());
        self.assert_snapshot();
    }
}

fn assert_table_receipt(actual: TableReceipt, expected: TableReceipt) {
    assert_eq!(actual.slide_id, expected.slide_id);
    assert_eq!(actual.shape_id, expected.shape_id);
    assert_eq!(actual.row_count, expected.row_count);
    assert_eq!(actual.column_count, expected.column_count);
    assert_eq!(actual.new_cell_refs, expected.new_cell_refs);
}

fn assert_history_receipt(actual: &HistoryReceipt, expected: &HistoryReceipt) {
    assert_eq!(actual.applied, expected.applied);
    assert_eq!(actual.changed_slide_ids, expected.changed_slide_ids);
    assert_eq!(actual.deck_size_changed, expected.deck_size_changed);
    assert_eq!(actual.slide_order_changed, expected.slide_order_changed);
}

fn assert_command_read(actual: &CommandReadResponse, expected: &CommandReadResponse) {
    assert_eq!(actual.read.slides, expected.read.slides);
    assert_eq!(actual.read.stories, expected.read.stories);
    assert_eq!(actual.sources.len(), expected.sources.len());
    for (actual, expected) in actual.sources.iter().zip(&expected.sources) {
        assert_eq!(actual.story_id, expected.story_id);
        assert_eq!(actual.paragraphs.len(), expected.paragraphs.len());
        for (actual, expected) in actual.paragraphs.iter().zip(&expected.paragraphs) {
            assert_eq!(actual.paragraph_id, expected.paragraph_id);
            assert_eq!(actual.text, expected.text);
            assert_eq!(actual.fields, expected.fields);
            assert_eq!(actual.editable, expected.editable);
        }
    }
}

#[test]
fn duplicate_slide_matches_session() {
    let pair = Pair::demo();
    let before = pair.snapshot();
    let slide_id = &before.slides[0].id;
    assert_eq!(
        pair.facade
            .duplicate_slide(&pair.context, slide_id, 1)
            .unwrap(),
        pair.session
            .duplicate_slide(&pair.context, slide_id, 1)
            .unwrap()
    );
    pair.assert_snapshot();
    assert_eq!(pair.snapshot().slides.len(), before.slides.len() + 1);
    let before = pair.snapshot();
    pair.assert_refusal(
        pair.facade
            .duplicate_slide(&pair.context, "missing-slide", 0),
        pair.session
            .duplicate_slide(&pair.context, "missing-slide", 0),
    );
    assert_eq!(pair.snapshot(), before);
}

#[test]
fn add_picture_matches_session() {
    let pair = Pair::demo();
    let slide_id = pair.snapshot().slides[0].id.clone();
    let mut media_bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut media_bytes, 1, 1);
        encoder.set_color(png::ColorType::Rgba);
        encoder.set_depth(png::BitDepth::Eight);
        encoder
            .write_header()
            .unwrap()
            .write_image_data(&[255, 0, 0, 255])
            .unwrap();
    }
    let draft = PictureDraft {
        name: "Parity picture".into(),
        rect: ShapeRect {
            x: 0,
            y: 0,
            width: 100_000,
            height: 100_000,
        },
        content_type: "image/png".into(),
        media_bytes,
    };
    assert_eq!(
        pair.facade
            .add_picture(&pair.context, &slide_id, &draft)
            .unwrap(),
        pair.session
            .add_picture(&pair.context, &slide_id, &draft)
            .unwrap()
    );
    pair.assert_snapshot();
    assert_eq!(pair.facade.save().unwrap(), pair.session.save().unwrap());
    let before = pair.snapshot();
    let invalid = PictureDraft {
        media_bytes: Vec::new(),
        ..draft
    };
    pair.assert_refusal(
        pair.facade.add_picture(&pair.context, &slide_id, &invalid),
        pair.session.add_picture(&pair.context, &slide_id, &invalid),
    );
    assert_eq!(pair.snapshot(), before);
}

macro_rules! z_order_test {
    ($name:ident, $facade:ident, $session:ident, $last:expr) => {
        #[test]
        fn $name() {
            let pair = Pair::demo();
            let before = pair.snapshot();
            let slide = &before.slides[0];
            let shape = if $last {
                slide.shapes.last().unwrap()
            } else {
                &slide.shapes[0]
            };
            let actual = pair
                .facade
                .$facade(&pair.context, &slide.id, &shape.id)
                .unwrap();
            let expected = pair
                .session
                .$session(&pair.context, &slide.id, &shape.id)
                .unwrap();
            assert_eq!(actual, expected);
            assert_ne!(actual.from_index, actual.to_index);
            pair.assert_snapshot();
            let before = pair.snapshot();
            pair.assert_refusal(
                pair.facade
                    .$facade(&pair.context, &slide.id, "missing-shape"),
                pair.session
                    .$session(&pair.context, &slide.id, "missing-shape"),
            );
            assert_eq!(pair.snapshot(), before);
        }
    };
}

z_order_test!(
    bring_shape_to_front_matches_session,
    bring_shape_to_front,
    bring_to_front,
    false
);
z_order_test!(
    send_shape_to_back_matches_session,
    send_shape_to_back,
    send_to_back,
    true
);
z_order_test!(
    bring_shape_forward_matches_session,
    bring_shape_forward,
    bring_forward,
    false
);
z_order_test!(
    send_shape_backward_matches_session,
    send_shape_backward,
    send_backward,
    true
);

#[test]
fn set_story_paragraphs_matches_session() {
    let pair = Pair::demo();
    let snapshot = pair.snapshot();
    let story_id = &snapshot.slides[0]
        .shapes
        .iter()
        .flat_map(|shape| &shape.text_stories)
        .next()
        .unwrap()
        .id;
    let paragraphs = [
        TextParagraphDraft {
            alignment: Some("ctr".into()),
            bullet: Some(true),
            runs: vec![TextRunDraft {
                text: "Parity 😀".into(),
                style: TextStyle {
                    bold: Some(true),
                    ..TextStyle::default()
                },
            }],
        },
        TextParagraphDraft {
            alignment: None,
            bullet: Some(false),
            runs: vec![TextRunDraft {
                text: "Second paragraph".into(),
                style: TextStyle::default(),
            }],
        },
    ];
    assert_eq!(
        pair.facade
            .set_story_paragraphs(&pair.context, story_id, &paragraphs)
            .unwrap(),
        pair.session
            .set_story_paragraphs(&pair.context, story_id, &paragraphs)
            .unwrap()
    );
    pair.assert_snapshot();
    let before = pair.snapshot();
    pair.assert_refusal(
        pair.facade
            .set_story_paragraphs(&pair.context, story_id, &[]),
        pair.session
            .set_story_paragraphs(&pair.context, story_id, &[]),
    );
    assert_eq!(pair.snapshot(), before);
}

#[test]
fn set_comment_position_matches_session() {
    let pair = Pair::demo();
    let slide_id = pair.snapshot().slides[0].id.clone();
    let comment = pair
        .facade
        .add_comment(
            &pair.context,
            &slide_id,
            "Parity",
            "P",
            "Position",
            "2026-10-29T00:00:00Z",
            0,
            0,
        )
        .unwrap();
    assert_eq!(
        comment,
        pair.session
            .add_comment(
                &pair.context,
                &slide_id,
                "Parity",
                "P",
                "Position",
                "2026-10-29T00:00:00Z",
                0,
                0,
            )
            .unwrap()
    );
    assert_eq!(
        pair.facade
            .set_comment_position(&pair.context, &comment.comment_id, 123_456, -234_567)
            .unwrap(),
        pair.session
            .set_comment_position(&pair.context, &comment.comment_id, 123_456, -234_567)
            .unwrap()
    );
    pair.assert_snapshot();
    let before = pair.snapshot();
    pair.assert_refusal(
        pair.facade
            .set_comment_position(&pair.context, &comment.comment_id, i64::MAX, 0),
        pair.session
            .set_comment_position(&pair.context, &comment.comment_id, i64::MAX, 0),
    );
    assert_eq!(pair.snapshot(), before);
}

macro_rules! table_edit_test {
    ($name:ident, $method:ident $(, $arg:expr)*) => {
        #[test]
        fn $name() {
            let pair = Pair::table();
            let (slide, shape) = pair.table_target();
            assert_table_receipt(
                pair.facade.$method(&pair.context, &slide, &shape, 0 $(, $arg)*).unwrap(),
                pair.session.$method(&pair.context, &slide, &shape, 0 $(, $arg)*).unwrap(),
            );
            pair.assert_snapshot();
            assert_eq!(pair.facade.save().unwrap(), pair.session.save().unwrap());
            let pair = Pair::demo();
            let before = pair.snapshot();
            let slide = &before.slides[0];
            let shape = &slide.shapes[0];
            pair.assert_refusal(
                pair.facade.$method(&pair.context, &slide.id, &shape.id, 0 $(, $arg)*),
                pair.session.$method(&pair.context, &slide.id, &shape.id, 0 $(, $arg)*),
            );
            assert_eq!(pair.snapshot(), before);
        }
    };
}

table_edit_test!(
    insert_table_row_matches_session,
    insert_table_row,
    TableInsertPosition::After,
    &["A".into(), "B".into()]
);
table_edit_test!(delete_table_row_matches_session, delete_table_row);
table_edit_test!(
    insert_table_column_matches_session,
    insert_table_column,
    TableInsertPosition::Before,
    &["A".into(), "B".into()]
);
table_edit_test!(
    insert_table_column_with_width_matches_session,
    insert_table_column_with_width,
    TableInsertPosition::After,
    &["A".into(), "B".into()],
    TableColumnWidth::Copy
);
table_edit_test!(delete_table_column_matches_session, delete_table_column);

#[test]
fn table_prechecks_match_session_without_edits() {
    for edit in [
        TableEdit::InsertRow,
        TableEdit::DeleteRow,
        TableEdit::InsertColumn,
        TableEdit::DeleteColumn,
    ] {
        for position in [TableInsertPosition::Before, TableInsertPosition::After] {
            let pair = Pair::table();
            let before = pair.snapshot();
            let (slide, shape) = pair.table_target();
            let version = pair.facade.version();
            let session_version = pair.session.version();
            pair.facade
                .check_table_edit(edit, &slide, &shape, 0, position, &[])
                .unwrap();
            pair.session
                .check_table_edit(edit, &slide, &shape, 0, position, &[])
                .unwrap();
            for width in [TableColumnWidth::Split, TableColumnWidth::Copy] {
                pair.facade
                    .check_table_edit_with_width(edit, &slide, &shape, 0, position, &[], width)
                    .unwrap();
                pair.session
                    .check_table_edit_with_width(edit, &slide, &shape, 0, position, &[], width)
                    .unwrap();
            }
            pair.assert_snapshot();
            assert_eq!(pair.snapshot(), before);
            assert_eq!(pair.facade.version(), version);
            assert_eq!(pair.session.version(), session_version);

            let pair = Pair::demo();
            let before = pair.snapshot();
            let slide = &before.slides[0];
            let shape = &slide.shapes[0];
            pair.assert_refusal(
                pair.facade
                    .check_table_edit(edit, &slide.id, &shape.id, 0, position, &[]),
                pair.session
                    .check_table_edit(edit, &slide.id, &shape.id, 0, position, &[]),
            );
            pair.assert_refusal(
                pair.facade.check_table_edit_with_width(
                    edit,
                    &slide.id,
                    &shape.id,
                    0,
                    position,
                    &[],
                    TableColumnWidth::Copy,
                ),
                pair.session.check_table_edit_with_width(
                    edit,
                    &slide.id,
                    &shape.id,
                    0,
                    position,
                    &[],
                    TableColumnWidth::Copy,
                ),
            );
            assert_eq!(pair.snapshot(), before);
        }
    }
}

#[test]
fn history_step_matches_session() {
    let pair = Pair::demo();
    for redo in [false, true] {
        let actual = pair.facade.history_step(redo).unwrap();
        let expected = pair.session.history_step(redo).unwrap();
        assert_history_receipt(&actual, &expected);
        assert!(!actual.applied);
    }
    let before = pair.snapshot();
    let slide_id = &before.slides[0].id;
    pair.facade
        .duplicate_slide(&pair.context, slide_id, 1)
        .unwrap();
    pair.session
        .duplicate_slide(&pair.context, slide_id, 1)
        .unwrap();
    pair.facade.add_undo_barrier();
    pair.session.add_undo_barrier();
    let edited = pair.snapshot();
    for redo in [false, true] {
        let actual = pair.facade.history_step(redo).unwrap();
        let expected = pair.session.history_step(redo).unwrap();
        assert_history_receipt(&actual, &expected);
        assert!(actual.applied);
        assert!(actual.slide_order_changed);
        pair.assert_snapshot();
        assert_eq!(
            pair.snapshot(),
            if redo { edited.clone() } else { before.clone() }
        );
    }
}

#[test]
fn read_command_content_matches_session() {
    let pair = Pair::demo();
    let before = pair.snapshot();
    for slide_ids in [None, Some(vec![before.slides[1].id.clone()])] {
        let request = ReadRequest { slide_ids };
        let actual = pair.facade.read_command_content(&request).unwrap().unwrap();
        let expected = pair
            .session
            .read_command_content(&request)
            .unwrap()
            .unwrap();
        assert_eq!(actual.read.version, pair.facade.version());
        assert_eq!(expected.read.version, pair.session.version());
        assert_command_read(&actual, &expected);
    }
    let request = ReadRequest {
        slide_ids: Some(vec!["x".repeat(betteroffice_pptx::MAX_REQUEST_BYTES + 1)]),
    };
    let Err(actual) = pair.facade.read_command_content(&request).unwrap() else {
        panic!("expected a command read refusal");
    };
    let Err(expected) = pair.session.read_command_content(&request).unwrap() else {
        panic!("expected a session command read refusal");
    };
    assert_eq!(actual.failure, expected.failure);
    assert_eq!(actual.version, pair.facade.version());
    assert_eq!(expected.version, pair.session.version());
    pair.assert_snapshot();
    assert_eq!(pair.snapshot(), before);
}
