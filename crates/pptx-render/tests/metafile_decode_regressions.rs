use pptx_render::fuzzing::decode;

/// POLYGON16 declaring eight points in a 40-byte record that holds three (#318).
const EMF_POLYGON_READS_PAST_ITS_RECORD: &[u8] =
    include_bytes!("fixtures/metafile-decode/emf-polygon-reads-past-its-record.emf");
const EMF_POLYGON_COUNT_AT: usize = 0x70;

/// META_POLYGON declaring four points in a 10-word record that holds three (#318).
const WMF_POLYGON_READS_PAST_ITS_RECORD: &[u8] =
    include_bytes!("fixtures/metafile-decode/wmf-polygon-reads-past-its-record.wmf");
const WMF_POLYGON_COUNT_AT: usize = 0x18;

fn with_count(bytes: &[u8], at: usize, count: u8) -> Vec<u8> {
    let mut bytes = bytes.to_vec();
    bytes[at] = count;
    bytes
}

#[test]
fn an_emf_polygon_counting_past_its_record_yields_no_drawing() {
    assert!(decode(EMF_POLYGON_READS_PAST_ITS_RECORD).is_none());
    assert!(
        decode(&with_count(
            EMF_POLYGON_READS_PAST_ITS_RECORD,
            EMF_POLYGON_COUNT_AT,
            3
        ))
        .is_some()
    );
}

#[test]
fn a_wmf_polygon_counting_past_its_record_yields_no_drawing() {
    assert!(decode(WMF_POLYGON_READS_PAST_ITS_RECORD).is_none());
    assert!(
        decode(&with_count(
            WMF_POLYGON_READS_PAST_ITS_RECORD,
            WMF_POLYGON_COUNT_AT,
            3
        ))
        .is_some()
    );
}
