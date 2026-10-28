# Font pack tooling

Install the release tooling in a virtual environment:

```sh
python3 -m venv /tmp/fontpack-venv
/tmp/fontpack-venv/bin/pip install fonttools==4.60.1 brotli==1.1.0
/tmp/fontpack-venv/bin/python scripts/fonts/woff2-transform.py <ttf-directory>
```

The transform writes `.uw2` (identity table stream) and standard `.woff2`
companions beside each `.ttf`, with fixed head timestamps. Both transform only
`glyf` and `loca`. It also writes `<face>.source.sha256`, containing the lowercase
hexadecimal SHA-256 of the exact TTF bytes encoded followed by a newline.
Use `fontpack build --codec woff2t` to compress and verify them; the builder
requires a matching source hash for every face. Regenerate companions with
`scripts/fonts/woff2-transform.py` whenever a TTF changes.
The default builder codec remains `zstd`; CFF/OTF faces use that codec.
This tooling runs only during development or release preparation. It never
fetches dependencies or runs at package installation.
