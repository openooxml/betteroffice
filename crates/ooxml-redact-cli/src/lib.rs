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
    enforce_size(input.len(), "input")?;
    let (bytes, report) =
        ooxml_redact::redact_with_report(input, Format::Auto).map_err(|error| error.to_string())?;
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
}
