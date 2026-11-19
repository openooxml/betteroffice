//! Revision authors in the order Word first meets them, which picks each one's color.

use quick_xml::{Reader, events::Event};

const REVISIONS: [&[u8]; 16] = [
    b"ins",
    b"del",
    b"moveFrom",
    b"moveTo",
    b"rPrChange",
    b"pPrChange",
    b"sectPrChange",
    b"tblPrChange",
    b"tblPrExChange",
    b"tblGridChange",
    b"trPrChange",
    b"tcPrChange",
    b"cellIns",
    b"cellDel",
    b"cellMerge",
    b"numberingChange",
];

/// The `w:author` of every revision in `parts`, in XML order, each once. A missing or empty
/// author is Word's "Unknown".
pub fn revision_authors<'a>(parts: impl IntoIterator<Item = &'a [u8]>) -> Vec<String> {
    let mut authors: Vec<String> = Vec::new();
    for xml in parts {
        let mut reader = Reader::from_reader(xml);
        loop {
            match reader.read_event() {
                Ok(Event::Start(element) | Event::Empty(element))
                    if REVISIONS.contains(&element.local_name().as_ref()) =>
                {
                    #[allow(deprecated)]
                    let author = element
                        .attributes()
                        .flatten()
                        .find(|attribute| attribute.key.local_name().as_ref() == b"author")
                        .and_then(|attribute| {
                            attribute.decode_and_unescape_value(reader.decoder()).ok()
                        })
                        .filter(|author| !author.trim().is_empty())
                        .map_or_else(|| "Unknown".to_owned(), |author| author.into_owned());
                    if !authors.contains(&author) {
                        authors.push(author);
                    }
                }
                Ok(Event::Eof) | Err(_) => break,
                Ok(_) => {}
            }
        }
    }
    authors
}

/// The revision authors of a package's stories: the main document first, then the header,
/// footer and note parts in package order.
pub fn package_revision_authors(parts: &[(String, Vec<u8>)], document_path: &str) -> Vec<String> {
    let story = |path: &str| {
        let name = path.rsplit('/').next().unwrap_or(path).to_ascii_lowercase();
        name.starts_with("header")
            || name.starts_with("footer")
            || name == "footnotes.xml"
            || name == "endnotes.xml"
    };
    let main = parts.iter().filter(|(path, _)| path == document_path);
    let others = parts
        .iter()
        .filter(|(path, _)| path != document_path && story(path));
    revision_authors(main.chain(others).map(|(_, xml)| xml.as_slice()))
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn authors_follow_xml_order_once_each() {
        let xml = br#"<w:body xmlns:w="w"><w:p><w:pPr><w:rPr><w:ins w:author="Mark"/></w:rPr></w:pPr><w:ins w:author="Ann"/><w:del w:author="O&apos;Neil"/><w:ins w:author="Ann"/><w:ins/><w:ins w:author=" "/></w:p></w:body>"#;
        assert_eq!(
            revision_authors([xml.as_slice()]),
            ["Mark", "Ann", "O'Neil", "Unknown"]
        );
    }
}
