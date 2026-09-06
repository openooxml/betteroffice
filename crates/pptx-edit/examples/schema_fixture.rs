use std::{env, fs};

use pptx_edit::{DeckSession, EditCtx, EditError, TextStyle};
use pptx_parse::ShapeNode;
use yrs::updates::decoder::Decode;
use yrs::{Any, Doc, Map, Out, ReadTxn, StateVector, Transact, Update};

fn main() {
    let args = env::args().collect::<Vec<_>>();
    let source = fs::read(&args[2]).unwrap();
    if args[1] == "reject" {
        assert!(matches!(
            DeckSession::open_from_update(&source, 9300),
            Err(EditError::InvalidState(message)) if message == "unsupported deck schema version"
        ));
        return;
    }
    let session = match args[1].as_str() {
        "migrate" => DeckSession::open_from_update(&source, 9300).unwrap(),
        "legacy-v2" => {
            let mut package = pptx_parse::parse_pptx_without_connectors(&source).unwrap();
            package.presentation.first_slide_num = 1;
            for slide in &mut package.slides {
                clear_text_styles(&mut slide.shapes);
            }
            for layout in &mut package.layouts {
                clear_text_styles(&mut layout.shapes);
            }
            for master in &mut package.masters {
                clear_text_styles(&mut master.shapes);
            }
            DeckSession::from_package_with_source(package, &source, 9300).unwrap()
        }
        "seed" => DeckSession::open(&source, 9300).unwrap(),
        _ => panic!("expected seed, legacy-v2, or migrate"),
    };
    if let Some(text) = args.get(4) {
        let snapshot = session.snapshot().unwrap();
        let story = snapshot.slides[0]
            .shapes
            .iter()
            .find_map(|shape| shape.text_stories.first())
            .unwrap();
        session
            .insert_text(
                &EditCtx::local("fixture"),
                &story.id,
                0,
                text,
                &TextStyle::default(),
            )
            .unwrap();
    }
    let mut update = session.encode_state_as_update_v1();
    if args[1] == "legacy-v2" {
        let doc = Doc::with_client_id(9301);
        let mut txn = doc.transact_mut();
        txn.apply_update(Update::decode_v1(&update).unwrap())
            .unwrap();
        let meta = txn.get_map("pptx:meta").unwrap();
        assert_eq!(
            meta.get(&txn, "schemaVersion"),
            Some(Out::Any(Any::Number(4.0)))
        );
        meta.insert(&mut txn, "schemaVersion", 2.0);
        update = txn.encode_state_as_update_v1(&StateVector::default());
    }
    fs::write(&args[3], update).unwrap();
}

fn clear_text_styles(shapes: &mut [ShapeNode]) {
    for shape in shapes {
        match shape {
            ShapeNode::Shape(shape) => shape.style = None,
            ShapeNode::Group(group) => clear_text_styles(&mut group.children),
            _ => {}
        }
    }
}
