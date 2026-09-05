use std::{env, fs, path::Path};

use pptx_edit::{DeckSession, EditCtx, ShapeDraft, ShapeRect, TextStyle};
use pptx_parse::{PptxPackage, ShapeNode};
use yrs::{Any, Map, Out, ReadTxn, Transact};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args: Vec<String> = env::args().collect();
    let branch = Path::new(&args[1]);
    let fixtures = branch.join("crates/pptx-edit/tests/fixtures");
    let source = fs::read(fixtures.join("hidden-shapes.pptx"))?;
    for (version, package) in [
        (2, legacy_package(&source)),
        (5, pptx_parse::parse_pptx(&source)?),
    ] {
        let session = DeckSession::from_package_with_source(package, &source, 4343)?;
        assert_main(&session);
        let ctx = EditCtx::local("fixture");
        session.add_text_box(
            &ctx,
            "slide:1:257",
            &ShapeDraft {
                name: "Persisted v2 textbox".to_owned(),
                rect: ShapeRect {
                    x: 100_000,
                    y: 100_000,
                    width: 2_000_000,
                    height: 600_000,
                },
                text: "persisted on v2".to_owned(),
                style: TextStyle::default(),
            },
        )?;
        session.insert_text(
            &ctx,
            "story:shape:4343:0:0",
            0,
            "edited ",
            &TextStyle::default(),
        )?;
        session.remove_shape(&ctx, "slide:1:257", "slide:1:257:shape:4")?;
        session.move_slide(&ctx, "slide:2:258", 0)?;
        if version == 2 {
            restamp_v2(&session);
        }
        fs::write(
            fixtures.join(format!("deck-schema-v{version}-hidden.update.bin")),
            session.encode_state_as_update_v1(),
        )?;
    }
    let demo = fs::read(branch.join("apps/demo/public/betteroffice-demo.pptx"))?;
    let session = DeckSession::open(&demo, 4343)?;
    assert_main(&session);
    fs::write(
        fixtures.join("deck-schema-v5.snapshot.json"),
        serde_json::to_string(&session.snapshot()?)?,
    )?;
    let source = fs::read(branch.join("crates/pptx-parse/tests/fixtures/style-matrix-deck.pptx"))?;
    let mut package = pptx_parse::parse_pptx(&source)?;
    package.presentation.first_slide_num = 10;
    let ShapeNode::Shape(shape) = &mut package.slides[0].shapes[0] else {
        panic!("expected a shape");
    };
    shape.base.hidden = true;
    let session = DeckSession::from_package(package, 4343)?;
    assert_main(&session);
    assert!(!session.package().themes[0].format_scheme.is_empty());
    fs::write(
        fixtures.join("deck-schema-v5-theme-hidden.update.bin"),
        session.encode_state_as_update_v1(),
    )?;
    let session = DeckSession::from_package_with_source(legacy_package(&source), &source, 9300)?;
    assert_main(&session);
    let story = session.snapshot()?.slides[0]
        .shapes
        .iter()
        .find_map(|shape| shape.text_stories.first())
        .unwrap()
        .id
        .clone();
    session.insert_text(
        &EditCtx::local("fixture"),
        &story,
        0,
        "persisted-v2 ",
        &TextStyle::default(),
    )?;
    restamp_v2(&session);
    fs::write(
        fixtures.join("deck-schema-v2.update.bin"),
        session.encode_state_as_update_v1(),
    )?;
    for name in ["connectors", "nested-connectors"] {
        let source = fs::read(fixtures.join(format!("deck-schema-v2-{name}.pptx")))?;
        let session = DeckSession::from_package_with_source(legacy_package(&source), &source, 273)?;
        assert_main(&session);
        restamp_v2(&session);
        fs::write(
            fixtures.join(format!("deck-schema-v2-{name}.update.bin")),
            session.encode_state_as_update_v1(),
        )?;
        if name == "connectors" {
            session.move_shape(
                &EditCtx::local("fixture"),
                "slide:0:256",
                "slide:0:256:shape:1",
                952_500,
                1_047_750,
            )?;
            fs::write(
                fixtures.join("deck-schema-v2-connectors-moved.update.bin"),
                session.encode_state_as_update_v1(),
            )?;
        }
    }
    Ok(())
}

fn assert_main(session: &DeckSession) {
    let txn = session.yrs_doc().transact();
    let meta = txn.get_map("pptx:meta").unwrap();
    assert_eq!(
        meta.get(&txn, "schemaVersion"),
        Some(Out::Any(Any::Number(5.0)))
    );
}

fn restamp_v2(session: &DeckSession) {
    let mut txn = session.yrs_doc().transact_mut();
    let meta = txn.get_map("pptx:meta").unwrap();
    meta.insert(&mut txn, "schemaVersion", 2.0);
}

fn legacy_package(source: &[u8]) -> PptxPackage {
    let mut package = pptx_parse::parse_pptx_without_connectors(source).unwrap();
    package.presentation.first_slide_num = 1;
    for theme in &mut package.themes {
        theme.format_scheme = Default::default();
    }
    for slide in &mut package.slides {
        clear_styles(&mut slide.shapes);
    }
    for layout in &mut package.layouts {
        clear_styles(&mut layout.shapes);
    }
    for master in &mut package.masters {
        clear_styles(&mut master.shapes);
    }
    package
}

fn clear_styles(shapes: &mut [ShapeNode]) {
    for shape in shapes {
        match shape {
            ShapeNode::Shape(shape) => shape.style = None,
            ShapeNode::Picture(picture) => picture.style = None,
            ShapeNode::Group(group) => clear_styles(&mut group.children),
            _ => {}
        }
    }
}
