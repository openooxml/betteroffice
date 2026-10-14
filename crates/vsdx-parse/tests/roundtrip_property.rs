use std::fmt::Write;
use std::panic::{AssertUnwindSafe, catch_unwind};

use vsdx_parse::{parse_vsdx, write_vsdx};

const CASES: usize = 256;
const SEED: u64 = 0x5EED_C0DE_D15E_A5E5;

struct Rng(u64);

impl Rng {
    fn new(seed: u64) -> Self {
        Self(seed)
    }

    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn choose(&mut self, upper: usize) -> usize {
        (self.next() % upper as u64) as usize
    }

    fn chance(&mut self, numerator: usize, denominator: usize) -> bool {
        self.choose(denominator) < numerator
    }
}

#[test]
fn lossless_round_trip_property() {
    for case in 0..CASES {
        let source = package_for_case(case);
        let (first, written, second) = run_case(&source, case);

        for (path, expected) in &source.parts {
            assert_eq!(
                second.part_bytes(path),
                Some(expected.as_slice()),
                "round-trip diverged in {path}; seed={SEED:#018X}; case={case}"
            );
        }
        assert_eq!(
            source.parts.len(),
            second_part_count(&second, &source.parts),
            "round-trip omitted or added a part; seed={SEED:#018X}; case={case}"
        );
        assert!(
            !written.is_empty(),
            "serializer emitted an empty package; seed={SEED:#018X}; case={case}"
        );
        assert_eq!(
            first, second,
            "model changed after round-trip; seed={SEED:#018X}; case={case}"
        );
    }
}

#[test]
fn parse_idempotence_property() {
    for case in 0..CASES {
        let source = package_for_case(case);
        let (first, _, second) = run_case(&source, case);
        assert_eq!(
            first, second,
            "model changed after reparsing; seed={SEED:#018X}; case={case}"
        );
    }
}

#[test]
fn parse_and_serialize_do_not_panic_property() {
    for case in 0..CASES {
        let source = package_for_case(case);
        let result = catch_unwind(AssertUnwindSafe(|| {
            let parsed = parse_vsdx(&source.bytes).expect("generated package must parse");
            write_vsdx(&parsed).expect("generated package must serialize")
        }));
        assert!(
            result.is_ok(),
            "parse or serialize panicked; seed={SEED:#018X}; case={case}"
        );
    }
}

fn run_case(
    source: &GeneratedPackage,
    case: usize,
) -> (vsdx_parse::VsdxPackage, Vec<u8>, vsdx_parse::VsdxPackage) {
    let result = catch_unwind(AssertUnwindSafe(|| {
        let first = parse_vsdx(&source.bytes).unwrap_or_else(|error| {
            panic!("generated package did not parse; seed={SEED:#018X}; case={case}; {error}")
        });
        let written = write_vsdx(&first).unwrap_or_else(|error| {
            panic!("generated package did not serialize; seed={SEED:#018X}; case={case}; {error}")
        });
        let second = parse_vsdx(&written).unwrap_or_else(|error| {
            panic!("serialized package did not parse; seed={SEED:#018X}; case={case}; {error}")
        });
        (first, written, second)
    }));
    result.unwrap_or_else(|_| panic!("parse or serialize panicked; seed={SEED:#018X}; case={case}"))
}

fn second_part_count(package: &vsdx_parse::VsdxPackage, expected: &[(String, Vec<u8>)]) -> usize {
    expected
        .iter()
        .filter(|(path, _)| package.part_bytes(path).is_some())
        .count()
}

struct GeneratedPackage {
    bytes: Vec<u8>,
    parts: Vec<(String, Vec<u8>)>,
}

fn package_for_case(case: usize) -> GeneratedPackage {
    let mut rng = Rng::new(SEED ^ (case as u64).wrapping_mul(0x9E37_79B9_7F4A_7C15));
    let page_xml = page_contents(&mut rng);
    let q = quote(&mut rng);
    let namespace = "http://schemas.microsoft.com/office/visio/2012/main";
    let parts = vec![
        xml_part(
            "[Content_Types].xml",
            "<Types xmlns='http://schemas.openxmlformats.org/package/2006/content-types'><Default ContentType='application/xml' Extension='xml'/><Default Extension='rels' ContentType='application/vnd.openxmlformats-package.relationships+xml'/><Override ContentType='application/vnd.ms-visio.drawing.main+xml' PartName='/visio/document.xml'/></Types>",
        ),
        xml_part(
            "_rels/.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Target=\"visio/document.xml\" Type=\"http://schemas.microsoft.com/visio/2010/relationships/document\" Id=\"rId1\"/></Relationships>",
        ),
        xml_part(
            "visio/document.xml",
            &format!(
                "<VisioDocument Data='doc &amp; data' xmlns={q}{namespace}{q} Odd='before'><UnknownDocument Strange={q}α &amp; β {q}>  document &#x1F642; text  </UnknownDocument><StyleSheets><StyleSheet NameU='Random' ID='0' Other='style'><Cell V='0' N='LineColor'/><Section N='User' Extra='section'><Row IX='0' T='Value' Weird='row'><Cell N='Value' V='42'/></Row></Section></StyleSheet></StyleSheets><DocumentSheet Mystery='document-sheet'><Cell V='8.5' N='PageWidth'/><Cell N='PageHeight' V='11'/></DocumentSheet></VisioDocument>"
            ),
        ),
        xml_part(
            "visio/_rels/document.xml.rels",
            "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/pages' Target='pages/pages.xml'/></Relationships>",
        ),
        xml_part(
            "visio/pages/pages.xml",
            &format!(
                "<Pages xmlns={q}{namespace}{q} Catalog='yes'><UnknownPageLevel Flag='kept'> page catalog &#x1F642; </UnknownPageLevel><Page Name='P &amp; Q' ID='1' NameU='Page-1' r:id='rId1' xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' PageAttr='unusual'><UnknownPage Flag='opaque'>  α&#x1F642;β  </UnknownPage><PageSheet PageSheetAttr='yes'><Cell V='8.5' N='PageWidth'/><Cell N='PageHeight' V='11'/></PageSheet></Page></Pages>"
            ),
        ),
        xml_part(
            "visio/pages/_rels/pages.xml.rels",
            "<Relationships xmlns='http://schemas.openxmlformats.org/package/2006/relationships'><Relationship Target='page1.xml' Id='rId1' Type='http://schemas.microsoft.com/visio/2010/relationships/page'/></Relationships>",
        ),
        xml_part("visio/pages/page1.xml", &page_xml),
    ];
    let bytes = stored_zip(&parts);
    GeneratedPackage { bytes, parts }
}

fn xml_part(path: &str, xml: &str) -> (String, Vec<u8>) {
    (path.to_owned(), xml.as_bytes().to_vec())
}

fn page_contents(rng: &mut Rng) -> String {
    let mut output = String::new();
    let q = quote(rng);
    let root_attributes = if rng.chance(1, 2) {
        format!(
            "PageAttr={q}page &amp; value{q} xmlns={q}http://schemas.microsoft.com/office/visio/2012/main{q} RootUnknown='yes'"
        )
    } else {
        format!(
            "xmlns={q}http://schemas.microsoft.com/office/visio/2012/main{q} RootUnknown='yes' PageAttr={q}page &amp; value{q}"
        )
    };
    let _ = write!(
        output,
        "<PageContents {root_attributes}>\n  <UnknownPageContents Value='opaque'>  α&#x1F642;β &amp; γ  </UnknownPageContents>\n  <Shapes>"
    );
    let mut next_id = 1;
    for _ in 0..(1 + rng.choose(6)) {
        let depth = rng.choose(4);
        shape(&mut output, rng, &mut next_id, depth);
    }
    output.push_str("</Shapes><UnknownPageTail Flag=\"kept\"> tail &amp; &#x1F642; </UnknownPageTail></PageContents>");
    output
}

fn shape(output: &mut String, rng: &mut Rng, next_id: &mut u32, depth: usize) {
    let id = *next_id;
    *next_id += 1;
    let q = quote(rng);
    let group = depth > 0 && rng.chance(3, 4);
    let shape_type = if group { "Group" } else { "Shape" };
    let attributes = if rng.chance(1, 2) {
        format!(
            "NameU={q}Shape-{id}{q} Mystery='shape-{id}' Type={q}{shape_type}{q} ID='{id}' Name='S &amp; {id}'"
        )
    } else {
        format!(
            "ID='{id}' Name='S &amp; {id}' Type={q}{shape_type}{q} Mystery='shape-{id}' NameU={q}Shape-{id}{q}"
        )
    };
    let _ = write!(output, "<Shape {attributes}>");
    output.push_str("\n  <UnknownShape Attr='kept'> shape α&#x1F642;β &amp; γ <Inner Flag=\"yes\"/> </UnknownShape>\n");
    cells(output, rng);
    section(output, rng);
    let _ = write!(
        output,
        "<Text>  α &#x1F642; &amp; β  <cp IX='0'/>  γ  </Text>"
    );
    if group {
        output.push_str("<Shapes>");
        for _ in 0..(1 + rng.choose(3)) {
            shape(output, rng, next_id, depth - 1);
        }
        output.push_str("</Shapes>");
    }
    output.push_str("</Shape>");
}

fn cells(output: &mut String, rng: &mut Rng) {
    for duplicate in 0..2 {
        let q = quote(rng);
        if rng.chance(1, 2) {
            let _ = write!(
                output,
                "<Cell V={q}{duplicate}{q} N='Dup' UnknownCell='kept'/><Cell N={q}FOnly{q} F='Width*2' Extra='formula'/><Cell V='5' N='VOnly'/><Cell F='Height*2' N='Both' V='2' Del='1'/>"
            );
        } else {
            let _ = write!(
                output,
                "<Cell UnknownCell='kept' N='Dup' V={q}{duplicate}{q}/><Cell Extra='formula' F='Width*2' N={q}FOnly{q}/><Cell N='VOnly' V='5'/><Cell Del='1' V='2' N='Both' F='Height*2'/>"
            );
        }
    }
}

fn section(output: &mut String, rng: &mut Rng) {
    let q = quote(rng);
    if rng.chance(1, 2) {
        let _ = write!(output, "<Section Extra={q}section-{q} N='User' IX='0'>");
    } else {
        let _ = write!(output, "<Section IX='0' N='User' Extra={q}section-{q}>");
    }
    output.push_str("<UnknownSection Flag='opaque'> section α&#x1F642;β &amp; γ </UnknownSection>");
    for row in 0..(1 + rng.choose(4)) {
        let q = quote(rng);
        let deleted = if rng.chance(1, 3) { " Del='1'" } else { "" };
        if rng.chance(1, 2) {
            let _ = write!(
                output,
                "<Row LocalName={q}Row-{row}{q} T='Value' IX='{row}' UnknownRow='kept'{deleted}><UnknownRowChild Flag='yes'> row α&#x1F642;β &amp; γ </UnknownRowChild><Cell V='1' N='Value'/><Cell N='Value' F='GUARD(1)'/><Cell N='OnlyValue' V='&#x1F642;'{deleted}/></Row>"
            );
        } else {
            let _ = write!(
                output,
                "<Row UnknownRow='kept' IX='{row}' T='Value' LocalName={q}Row-{row}{q}{deleted}><UnknownRowChild Flag='yes'> row α&#x1F642;β &amp; γ </UnknownRowChild><Cell N='Value' V='1'/><Cell F='GUARD(1)' N='Value'/><Cell{deleted} V='&#x1F642;' N='OnlyValue'/></Row>"
            );
        }
    }
    output.push_str("</Section>");
}

fn quote(rng: &mut Rng) -> char {
    if rng.chance(1, 2) { '\'' } else { '"' }
}

fn stored_zip(parts: &[(String, Vec<u8>)]) -> Vec<u8> {
    let mut output = Vec::new();
    let mut central = Vec::new();
    for (path, bytes) in parts {
        let offset = output.len() as u32;
        let name = path.as_bytes();
        let checksum = crc32(bytes);
        le_u32(&mut output, 0x0403_4B50);
        le_u16(&mut output, 20);
        le_u16(&mut output, 0x0800);
        le_u16(&mut output, 0);
        le_u16(&mut output, 0);
        le_u16(&mut output, 0);
        le_u32(&mut output, checksum);
        le_u32(&mut output, bytes.len() as u32);
        le_u32(&mut output, bytes.len() as u32);
        le_u16(&mut output, name.len() as u16);
        le_u16(&mut output, 0);
        output.extend_from_slice(name);
        output.extend_from_slice(bytes);

        le_u32(&mut central, 0x0201_4B50);
        le_u16(&mut central, 20);
        le_u16(&mut central, 20);
        le_u16(&mut central, 0x0800);
        le_u16(&mut central, 0);
        le_u16(&mut central, 0);
        le_u16(&mut central, 0);
        le_u32(&mut central, checksum);
        le_u32(&mut central, bytes.len() as u32);
        le_u32(&mut central, bytes.len() as u32);
        le_u16(&mut central, name.len() as u16);
        le_u16(&mut central, 0);
        le_u16(&mut central, 0);
        le_u16(&mut central, 0);
        le_u16(&mut central, 0);
        le_u32(&mut central, 0);
        le_u32(&mut central, offset);
        central.extend_from_slice(name);
    }
    let central_offset = output.len() as u32;
    output.extend_from_slice(&central);
    le_u32(&mut output, 0x0605_4B50);
    le_u16(&mut output, 0);
    le_u16(&mut output, 0);
    le_u16(&mut output, parts.len() as u16);
    le_u16(&mut output, parts.len() as u16);
    le_u32(&mut output, central.len() as u32);
    le_u32(&mut output, central_offset);
    le_u16(&mut output, 0);
    output
}

fn le_u16(output: &mut Vec<u8>, value: u16) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn le_u32(output: &mut Vec<u8>, value: u32) {
    output.extend_from_slice(&value.to_le_bytes());
}

fn crc32(bytes: &[u8]) -> u32 {
    let mut value = !0_u32;
    for byte in bytes {
        value ^= u32::from(*byte);
        for _ in 0..8 {
            value = (value >> 1) ^ (0xEDB8_8320 & 0_u32.wrapping_sub(value & 1));
        }
    }
    !value
}
