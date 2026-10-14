pub use ooxml_redact::RedactionOptions;
use ooxml_redact::{Format, RedactionReport};
use serde::Deserialize;

pub const MAX_FILE_BYTES: usize = 64 * 1024 * 1024;
pub const DEFAULT_UPLOAD_URL: &str = "https://redact.betteroffice.dev/upload";

pub struct RedactedFile {
    bytes: Vec<u8>,
    report: RedactionReport,
}

impl RedactedFile {
    pub fn bytes(&self) -> &[u8] {
        &self.bytes
    }

    pub fn report(&self) -> &RedactionReport {
        &self.report
    }

    pub fn format(&self) -> Format {
        self.report.format
    }
}

#[derive(Debug, Deserialize, Eq, PartialEq)]
pub struct ShareResponse {
    pub id: String,
    pub url: String,
}

pub fn redact_local(input: &[u8]) -> Result<RedactedFile, String> {
    redact_local_with_options(input, &RedactionOptions::default())
}

pub fn redact_local_with_options(
    input: &[u8],
    options: &RedactionOptions,
) -> Result<RedactedFile, String> {
    enforce_size(input.len(), "input")?;
    let (bytes, report) =
        ooxml_redact::redact_with_report_and_options(input, Format::Auto, options)
            .map_err(|error| error.to_string())?;
    enforce_size(bytes.len(), "redacted output")?;
    Ok(RedactedFile { bytes, report })
}

pub fn upload_redacted(redacted: &RedactedFile, endpoint: &str) -> Result<ShareResponse, String> {
    enforce_size(redacted.bytes.len(), "redacted upload")?;
    let extension = redacted
        .format()
        .extension()
        .ok_or("redacted upload has no detected format")?;
    let response = ureq::post(endpoint)
        .header("Content-Type", content_type(redacted.format()))
        .header("X-BetterOffice-Format", extension)
        .send(&redacted.bytes)
        .map_err(|error| format!("upload failed: {error}"))?;
    let mut body = response.into_body();
    let json = body
        .read_to_string()
        .map_err(|error| format!("reading upload response: {error}"))?;
    serde_json::from_str(&json).map_err(|error| format!("invalid upload response: {error}"))
}

pub fn report_line(report: &RedactionReport) -> String {
    format!(
        "redacted {}: {} text nodes ({} characters), {} attributes, {} media parts, {} binary parts, {} XML comments",
        report.format,
        report.text_nodes,
        report.characters,
        report.attributes,
        report.media_parts,
        report.binary_parts,
        report.xml_comments
    )
}

fn enforce_size(size: usize, label: &str) -> Result<(), String> {
    if size > MAX_FILE_BYTES {
        Err(format!(
            "{label} is {size} bytes; maximum is {MAX_FILE_BYTES} bytes"
        ))
    } else {
        Ok(())
    }
}

fn content_type(format: Format) -> &'static str {
    match format {
        Format::Docx => "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        Format::Xlsx => "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        Format::Pptx => "application/vnd.openxmlformats-officedocument.presentationml.presentation",
        Format::Vsdx => "application/vnd.ms-visio.drawing",
        Format::Vstx => "application/vnd.ms-visio.template",
        Format::Auto => "application/octet-stream",
    }
}

#[cfg(test)]
mod tests {
    use std::io::{Read, Write};
    use std::net::TcpListener;
    use std::sync::mpsc;
    use std::thread;

    use super::*;

    const SECRET: &str = "CLI_SECRET_CONTENT";

    #[test]
    fn local_random_mode_uses_the_options_without_changing_the_default() {
        let input = fixture();
        let default = redact_local(&input).unwrap();
        let explicit_default =
            redact_local_with_options(&input, &RedactionOptions::default()).unwrap();
        assert_eq!(
            ooxml_opc::unzip_parts(default.bytes()).unwrap(),
            ooxml_opc::unzip_parts(explicit_default.bytes()).unwrap()
        );
        let random = redact_local_with_options(
            &input,
            &RedactionOptions {
                random_characters: true,
            },
        )
        .unwrap();
        let parts = ooxml_opc::unzip_parts(random.bytes()).unwrap();
        let document = parts
            .iter()
            .find(|(path, _)| path == "word/document.xml")
            .unwrap();
        let text = String::from_utf8_lossy(&document.1);
        assert!(!text.contains(SECRET));
        assert!(!text.contains(&"x".repeat(SECRET.len())));
        assert_eq!(random.report().format, Format::Docx);
    }

    #[test]
    fn local_redaction_removes_secret_before_upload() {
        let redacted = redact_local(&fixture()).unwrap();
        let parts = ooxml_opc::unzip_parts(redacted.bytes()).unwrap();
        assert!(
            parts
                .iter()
                .all(|(_, bytes)| { !String::from_utf8_lossy(bytes).contains(SECRET) })
        );
    }

    #[test]
    fn uploader_sends_only_redacted_bytes_and_no_filename() {
        let redacted = redact_local(&fixture()).unwrap();
        let expected = redacted.bytes().to_vec();
        let listener = TcpListener::bind("127.0.0.1:0").unwrap();
        let endpoint = format!("http://{}/upload", listener.local_addr().unwrap());
        let (sender, receiver) = mpsc::channel();
        let server = thread::spawn(move || {
            let (mut stream, _) = listener.accept().unwrap();
            let mut received = Vec::new();
            let mut buffer = [0_u8; 4096];
            let header_end = loop {
                let count = stream.read(&mut buffer).unwrap();
                assert_ne!(count, 0);
                received.extend_from_slice(&buffer[..count]);
                if let Some(index) = received.windows(4).position(|window| window == b"\r\n\r\n") {
                    break index + 4;
                }
            };
            let headers = String::from_utf8_lossy(&received[..header_end]).into_owned();
            let content_length = headers
                .lines()
                .find_map(|line| {
                    line.split_once(':').and_then(|(name, value)| {
                        name.eq_ignore_ascii_case("content-length")
                            .then(|| value.trim().parse::<usize>().unwrap())
                    })
                })
                .unwrap();
            while received.len() - header_end < content_length {
                let count = stream.read(&mut buffer).unwrap();
                received.extend_from_slice(&buffer[..count]);
            }
            sender
                .send((
                    headers,
                    received[header_end..header_end + content_length].to_vec(),
                ))
                .unwrap();
            stream
                .write_all(
                    b"HTTP/1.1 200 OK\r\nContent-Type: application/json\r\nContent-Length: 58\r\nConnection: close\r\n\r\n{\"id\":\"opaque-id\",\"url\":\"https://example.com/f/opaque-id\"}",
                )
                .unwrap();
        });

        let response = upload_redacted(&redacted, &endpoint).unwrap();
        let (headers, body) = receiver.recv().unwrap();
        server.join().unwrap();
        assert_eq!(response.id, "opaque-id");
        assert_eq!(body, expected);
        assert!(!String::from_utf8_lossy(&body).contains(SECRET));
        assert!(!headers.to_ascii_lowercase().contains("filename"));
        assert!(
            headers
                .to_ascii_lowercase()
                .contains("x-betteroffice-format: docx")
        );
    }

    #[test]
    fn local_visio_redaction_uses_the_detected_extension_and_content_type() {
        for (source, format, mime) in [
            (
                include_bytes!("../../../apps/demo/public/betteroffice-demo.vsdx").as_slice(),
                Format::Vsdx,
                "application/vnd.ms-visio.drawing",
            ),
            (
                include_bytes!("../../vsdx-parse/tests/fixtures/template.vstx").as_slice(),
                Format::Vstx,
                "application/vnd.ms-visio.template",
            ),
        ] {
            let redacted = redact_local(source).unwrap();
            assert_eq!(redacted.format(), format);
            assert_eq!(content_type(redacted.format()), mime);
            assert_ne!(redacted.bytes(), source);
            assert!(
                ooxml_opc::sanitize_package_for_format(
                    redacted.bytes(),
                    format.extension().unwrap()
                )
                .is_ok()
            );
        }
    }

    #[test]
    fn refuses_visio_without_package_metadata() {
        let source = ooxml_opc::rezip_parts(&[(
            "visio/document.xml".to_owned(),
            br#"<VisioDocument><CommentList><CommentEntry Author="PRIVATE_AUTHOR">PRIVATE_COMMENT</CommentEntry></CommentList></VisioDocument>"#.to_vec(),
        )]).unwrap();
        let error = redact_local(&source)
            .err()
            .expect("missing package metadata must be rejected");
        assert!(error.contains("unsupported or ambiguous Visio format"));
    }

    fn fixture() -> Vec<u8> {
        ooxml_opc::rezip_parts(&[
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#.to_vec(),
            ),
            (
                "word/document.xml".to_owned(),
                format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{SECRET}</w:t></w:r></w:p></w:body></w:document>"#).into_bytes(),
            ),
        ])
        .unwrap()
    }

    #[test]
    fn local_redaction_removes_xlsx_person_and_pivot_secrets() {
        let input = ooxml_opc::rezip_parts(&[
            (
                "[Content_Types].xml".to_owned(),
                br#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/></Types>"#.to_vec(),
            ),
            (
                "xl/workbook.xml".to_owned(),
                br#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>"#.to_vec(),
            ),
            (
                "xl/persons/person.xml".to_owned(),
                br#"<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"><person displayName="CLI_SECRET_PERSON" id="{11111111-1111-1111-1111-111111111111}" userId="cli.secret@example.com" providerId="AD"/></personList>"#.to_vec(),
            ),
            (
                "xl/pivotCache/pivotCacheDefinition1.xml".to_owned(),
                br#"<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cacheFields count="1"><cacheField name="CLI_SECRET_FIELD" numFmtId="0"><sharedItems count="1"><s v="CLI_SECRET_SHARED"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>"#.to_vec(),
            ),
        ])
        .unwrap();
        let redacted = redact_local(&input).unwrap();
        assert_eq!(redacted.format(), Format::Xlsx);
        let parts = ooxml_opc::unzip_parts(redacted.bytes()).unwrap();
        for secret in [
            "CLI_SECRET_PERSON",
            "cli.secret@example.com",
            "CLI_SECRET_FIELD",
            "CLI_SECRET_SHARED",
        ] {
            assert!(
                parts
                    .iter()
                    .all(|(_, bytes)| !String::from_utf8_lossy(bytes).contains(secret)),
                "secret survived: {secret}"
            );
        }
        let persons = parts
            .iter()
            .find(|(path, _)| path == "xl/persons/person.xml")
            .unwrap();
        assert!(
            String::from_utf8_lossy(&persons.1).contains("{11111111-1111-1111-1111-111111111111}")
        );
    }
}
