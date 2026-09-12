# Redaction integrity fixture

Run `python3 generate_redaction_integrity.py` to regenerate `redaction-integrity.docx` deterministically.

The fixture uses only synthetic `SECRET_*` sentinels. It exercises same-length custom style names and references, built-in styles, custom XML item/property relationships, sensitive custom XML attributes, ordinary custom XML text, a VML `o:gfxdata` nested ZIP containing synthetic text and a tiny PNG, outer media, geometry attributes, and visible document text.

Custom XML includes `xsi:nil`, `xsi:type`, and a private schema location to verify that redaction preserves valid nil booleans and removes optional type overrides and external schema hints.
