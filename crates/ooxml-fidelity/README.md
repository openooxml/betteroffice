# Save fidelity oracles

`roundtrip_findings` compares original and saved package parts using an XML
reader independent of the serializer. The structural fingerprint compares
resolved element names, attributes, text, namespace declarations and child
order. The WML digest compares paragraph text, containment, properties,
identities and generic subtrees. The element census reports decreasing counts
by expanded name. Findings identify the affected part or digest field.

Non-XML parts and XML parts outside `MODELLED_XML_PARTS` must retain identical
bytes. Missing parts and unexpected additions are findings. Serializer
exceptions are listed in `DECLARED_NORMALIZATIONS` in `src/registry.rs`;
relationship and content-type entries compare as sets, and paragraph identities
may be added but never changed or removed.

Synthetic pairs exercise differences that a census alone cannot detect:
attribute values, text, field instructions, properties, containment, markers,
relationships and unknown subtrees. Run `cargo test -p betteroffice-ooxml-fidelity`
and the [DOCX corpus tests](../betteroffice-docx/tests/corpus/README.md).
