use std::alloc::{GlobalAlloc, Layout, System};
use std::sync::atomic::{AtomicUsize, Ordering};

use pptx_edit::DeckSession;

static CURRENT: AtomicUsize = AtomicUsize::new(0);
static PEAK: AtomicUsize = AtomicUsize::new(0);

struct Counting;

impl Counting {
    fn track_alloc(size: usize) {
        let current = CURRENT.fetch_add(size, Ordering::Relaxed) + size;
        let mut peak = PEAK.load(Ordering::Relaxed);
        while current > peak {
            match PEAK.compare_exchange_weak(peak, current, Ordering::Relaxed, Ordering::Relaxed) {
                Ok(_) => break,
                Err(actual) => peak = actual,
            }
        }
    }
}

unsafe impl GlobalAlloc for Counting {
    unsafe fn alloc(&self, layout: Layout) -> *mut u8 {
        let ptr = unsafe { System.alloc(layout) };
        if !ptr.is_null() {
            Self::track_alloc(layout.size());
        }
        ptr
    }

    unsafe fn dealloc(&self, ptr: *mut u8, layout: Layout) {
        CURRENT.fetch_sub(layout.size(), Ordering::Relaxed);
        unsafe { System.dealloc(ptr, layout) }
    }
}

#[global_allocator]
static ALLOCATOR: Counting = Counting;

fn current() -> usize {
    CURRENT.load(Ordering::Relaxed)
}

fn synthetic_deck() -> Vec<u8> {
    const SLIDES: usize = 40;
    const SHAPES_PER_SLIDE: usize = 120;
    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    let mut types = String::from(
        r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>"#,
    );
    for slide in 0..SLIDES {
        types.push_str(&format!(
            r#"<Override PartName="/ppt/slides/slide{}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>"#,
            slide + 1
        ));
    }
    types.push_str("</Types>");
    parts.push(("[Content_Types].xml".to_owned(), types.into_bytes()));
    parts.push((
        "_rels/.rels".to_owned(),
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/></Relationships>"#
            .as_bytes()
            .to_vec(),
    ));
    let mut presentation = String::from(
        r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst>"#,
    );
    let mut rels = String::from(
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
    );
    for slide in 0..SLIDES {
        presentation.push_str(&format!(
            r#"<p:sldId id="{}" r:id="s{}"/>"#,
            256 + slide,
            slide
        ));
        rels.push_str(&format!(
            r#"<Relationship Id="s{}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{}.xml"/>"#,
            slide,
            slide + 1
        ));
    }
    presentation.push_str(
        r#"</p:sldIdLst><p:sldSz cx="9144000" cy="6858000"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>"#,
    );
    rels.push_str("</Relationships>");
    parts.push(("ppt/presentation.xml".to_owned(), presentation.into_bytes()));
    parts.push((
        "ppt/_rels/presentation.xml.rels".to_owned(),
        rels.into_bytes(),
    ));
    for slide in 0..SLIDES {
        let mut xml = String::from(
            r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>"#,
        );
        for shape in 0..SHAPES_PER_SLIDE {
            xml.push_str(&format!(
                r#"<p:sp><p:nvSpPr><p:cNvPr id="{}" name="Text {shape}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="{}" y="{}"/><a:ext cx="1000000" cy="1000000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Slide {slide} shape {shape} body text</a:t></a:r><a:endParaRPr/></a:p></p:txBody></p:sp>"#,
                shape + 2,
                shape * 1000,
                slide * 1000,
            ));
        }
        xml.push_str("</p:spTree></p:cSld></p:sld>");
        parts.push((
            format!("ppt/slides/slide{}.xml", slide + 1),
            xml.into_bytes(),
        ));
    }
    ooxml_opc::rezip_parts(&parts).unwrap()
}

#[test]
fn open_peak_stays_below_two_coexisting_docs() {
    let bytes = synthetic_deck();
    let reference = DeckSession::open(&bytes, 7).unwrap();
    let baseline = reference.encode_state_as_update_v1();
    let baseline_len = baseline.len();
    assert!(baseline_len > 0);
    drop(reference);

    let retained_start = current();
    let alone = DeckSession::open_from_update(&baseline, 42).unwrap();
    let one_doc_bytes = current().saturating_sub(retained_start);
    assert!(one_doc_bytes > 0);
    assert_eq!(alone.encode_state_as_update_v1(), baseline);

    let open_start = current();
    PEAK.store(open_start, Ordering::Relaxed);
    let session = DeckSession::open(&bytes, 7).unwrap();
    let open_peak = PEAK.load(Ordering::Relaxed).saturating_sub(open_start);
    let bound = baseline_len * 3 + one_doc_bytes + one_doc_bytes / 2;
    eprintln!(
        "open_peak: baseline_len={baseline_len} one_doc_bytes={one_doc_bytes} open_peak={open_peak} bound={bound}"
    );
    assert_eq!(session.client_id(), 7);
    assert_eq!(session.encode_state_as_update_v1(), baseline);
    assert_eq!(session.snapshot().unwrap(), alone.snapshot().unwrap());
    assert!(
        open_peak < bound,
        "open peak {open_peak} must stay below {bound} (baseline_len={baseline_len}, one_doc_bytes={one_doc_bytes})"
    );
}
