# Save fidelity oracles

`roundtrip_findings` compares original and saved package parts using an XML
reader independent of the serializer. The structural fingerprint compares
resolved element names, attributes, text, namespace declarations and child
order. The WML digest compares paragraph text, containment, properties,
identities, root attributes and generic subtrees. The element census reports decreasing counts
by expanded name. Findings identify the affected part or digest field.

Non-XML parts and XML parts outside `MODELLED_XML_PARTS` must retain identical
bytes. Byte-identical parts bypass XML parsing; only changed XML parts contribute
to report census counts and digests. If either side of a changed part cannot be
parsed, the report includes `unparseable part: <name>` and still enforces its
byte rule. The reader explicitly enables quick-xml encoding support and rejects
non-ASCII-compatible encodings, DTDs and undefined entities. Text outside the
root is rejected. Missing parts and unexpected additions are findings. Serializer
exceptions are listed in `DECLARED_NORMALIZATIONS` in `src/registry.rs`;
relationship and content-type entries compare as sets, and paragraph identities
may be added but never changed or removed. Root `mc:Ignorable` values compare as
resolved URI sets: additions from the registered standard WML namespaces are
allowed; removals and custom additions are findings. `mc:ProcessContent`,
`mc:Choice/@Requires` and `xsi:type` also resolve values through in-scope
namespace bindings. Unbound prefixes are rejected.

The companion exception permits only newly added `word/commentsIds.xml`,
`word/commentsExtended.xml` and `word/commentsExtensible.xml` parts, plus
new content-type overrides and internal relationship entries resolving to those
exact part names. Existing entries and unrelated names remain significant.

Literal CRLF and CR normalize to LF; character references retain their value.
Whitespace in empty WML property containers is insignificant unless
`xml:space="preserve"` applies. Nonbreaking spaces remain content. Property child
order, including `pPr` and `rPr`, remains significant. The demo builder's
spacing-before-indentation change corrected the fixture to Word schema order.

Synthetic pairs exercise differences that a census alone cannot detect:
attribute values, text, field instructions, properties, containment, markers,
relationships and unknown subtrees. Run `cargo test -p betteroffice-ooxml-fidelity`
and the [DOCX corpus tests](../betteroffice-docx/tests/corpus/README.md).
