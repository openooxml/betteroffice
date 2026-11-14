# Third-party asset attribution

## `sRGB-v2-magic.icc`

- **File:** `assets/sRGB-v2-magic.icc`
- **SHA-256:** `af4efe28f6d311799865f325ba39184a2e978b113ca74124d60af7de22b105f4`
- **Size:** 736 bytes
- **Description:** sRGB IEC61966-2.1 ICC v2 (2.1.0) display profile, RGB data
  color space with XYZ profile connection space. This is the "magic" variant of
  the minimal sRGB profile, using a 182-point tone reproduction curve.
- **License:** [CC0 1.0 Universal](https://creativecommons.org/publicdomain/zero/1.0/)
  (public domain dedication). The profile's embedded `cprt` tag reads `CC0`.
- **Author:** Clinton Ingram ([saucecontrol](https://github.com/saucecontrol)),
  author of *Compact ICC Profiles*.
- **Upstream source:**
  <https://github.com/saucecontrol/Compact-ICC-Profiles/blob/master/profiles/sRGB-v2-magic.icc>
- **Upstream license statement:**
  <https://github.com/saucecontrol/Compact-ICC-Profiles/blob/master/readme.md>
- **Obtained from:** the [krilla](https://github.com/LaurenzV/krilla) PDF
  library, path
  [`crates/krilla/icc/sRGB-v2-magic.icc`](https://github.com/LaurenzV/krilla/blob/main/crates/krilla/icc/sRGB-v2-magic.icc).
  The copy in this repository is byte-for-byte identical to the saucecontrol
  upstream file (same SHA-256).

The profile is bundled solely as the PDF/A output intent destination profile
for sRGB job color. It is not modified.
