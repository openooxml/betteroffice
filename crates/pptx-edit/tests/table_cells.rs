use pptx_edit::{DeckSession, EditCtx, TextStyle};

const DECK: &[u8] = include_bytes!("../../pptx-render/tests/fixtures/table-basic.pptx");
const DEMO: &[u8] = include_bytes!("../../../apps/demo/public/betteroffice-demo.pptx");
const V1_UPDATE: &[u8] = include_bytes!("fixtures/deck-schema-v1.update.bin");
const STORY: &str = "story:slide:0:256:shape:0:table:0:0";

#[test]
fn editing_a_cell_leaves_its_cell_properties_in_place() {
    let session = DeckSession::open(DECK, 34101).unwrap();
    session
        .insert_text(
            &EditCtx::local("test"),
            STORY,
            0,
            "New ",
            &TextStyle::default(),
        )
        .unwrap();
    let saved = session.save().unwrap();
    let parts = ooxml_opc::unzip_parts(&saved).unwrap();
    let slide = parts
        .iter()
        .find(|(path, _)| path == "ppt/slides/slide1.xml")
        .map(|(_, bytes)| String::from_utf8(bytes.clone()).unwrap())
        .unwrap();

    assert!(slide.contains("New Header spans two"));
    assert_eq!(
        slide
            .matches(r#"<a:tcPr anchor="ctr" marL="91440" marR="91440"/>"#)
            .count(),
        7
    );
    assert!(slide.contains(r#"<a:solidFill><a:srgbClr val="DDEBF7"/></a:solidFill>"#));
    assert!(slide.contains(r#"<a:lnB w="19050">"#));
    assert!(slide.contains(r#"<a:gridCol w="1828800"/>"#));
    assert!(slide.contains(r#"<a:tc rowSpan="2">"#));
}

#[test]
fn every_cell_is_addressable_in_row_major_order() {
    let session = DeckSession::open(DECK, 34102).unwrap();
    let snapshot = session.snapshot().unwrap();
    let stories: Vec<_> = snapshot.slides[0].shapes[0]
        .text_stories
        .iter()
        .map(|story| story.id.clone())
        .collect();
    assert_eq!(stories.len(), 9);
    assert_eq!(stories[1], "story:slide:0:256:shape:0:table:0:1");
    assert_eq!(stories[8], "story:slide:0:256:shape:0:table:2:2");
    assert_eq!(session.story(&stories[3]).unwrap().plain_text(), "Tall");
}

#[test]
fn a_reattached_source_restores_the_geometry_a_released_snapshot_never_stored() {
    let detached = DeckSession::open_from_update(V1_UPDATE, 34103).unwrap();
    assert!(stored_table(&detached).grid.is_empty());
    assert!(seeded_table(&detached).grid.is_empty());

    let attached = DeckSession::open_from_update_with_source(V1_UPDATE, DEMO, 34104).unwrap();
    let table = stored_table(&attached);
    assert_eq!(table.grid, [1_981_200, 2_438_400, 2_895_600, 3_352_800]);
    assert!(table.properties.first_row && table.properties.band_row);
    assert_eq!(
        table.rows.iter().map(|row| row.height).collect::<Vec<_>>(),
        [609_600, 762_000, 762_000, 762_000, 762_000]
    );
    assert_eq!(
        table
            .rows
            .iter()
            .flat_map(|row| row.cells.iter())
            .filter(|cell| cell.fill.is_some())
            .count(),
        20
    );
    assert_eq!(seeded_table(&attached).grid, table.grid);
    assert_eq!(
        table,
        stored_table(&DeckSession::open(DEMO, 34105).unwrap())
    );

    let update = attached.encode_state_as_update_v1();
    let reattached = DeckSession::open_from_update_with_source(&update, DEMO, 34106).unwrap();
    assert_eq!(reattached.encode_state_as_update_v1(), update);
}

fn stored_table(session: &DeckSession) -> pptx_parse::Table {
    session
        .package()
        .slides
        .iter()
        .flat_map(|slide| slide.shapes.iter())
        .find_map(|node| match node {
            pptx_parse::ShapeNode::GraphicFrame(frame) => match &frame.data {
                pptx_parse::GraphicFrameData::Table(table) => Some(table.clone()),
                _ => None,
            },
            _ => None,
        })
        .expect("the demo deck has a table")
}

fn seeded_table(session: &DeckSession) -> pptx_parse::Table {
    session
        .snapshot()
        .unwrap()
        .slides
        .iter()
        .flat_map(|slide| slide.shapes.clone())
        .find_map(|shape| match shape.graphic {
            Some(pptx_parse::GraphicFrameData::Table(table)) => Some(table),
            _ => None,
        })
        .expect("the demo deck seeds a table")
}
