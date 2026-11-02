from pathlib import Path

import pytest

from betteroffice_docx import decode_font_face

FIXTURES = Path(__file__).resolve().parents[3] / "crates" / "betteroffice-fonts" / "tests" / "fixtures"


def fixture(name):
    return (FIXTURES / name).read_bytes()


def test_zstd_preserves_source_fonts_with_and_without_dictionary():
    regular = fixture("Synthetic-Regular.ttf")
    assert decode_font_face(fixture("Synthetic-Regular.sfnt.zst"), "zstd") == (regular, None)
    assert decode_font_face(
        fixture("Synthetic-Bold.sfnt.zst"), "zstd", dict=regular
    ) == (fixture("Synthetic-Bold.ttf"), None)


def test_woff2t_preserves_payloads_with_and_without_dictionary():
    font, payload = decode_font_face(fixture("Synthetic-Regular.ttf.zst"), "woff2t+zstd")
    assert isinstance(font, bytes)
    assert font.startswith(b"\x00\x01\x00\x00")
    assert payload == fixture("Synthetic-Regular.uw2")
    bold, bold_payload = decode_font_face(
        fixture("Synthetic-Bold.ttf.zst"), "woff2t+zstd", dict=payload
    )
    assert bold.startswith(b"\x00\x01\x00\x00")
    assert bold_payload == fixture("Synthetic-Bold.uw2")
    assert decode_font_face(
        fixture("Synthetic-Bold.nodict.zst"), "woff2t+zstd"
    ) == (bold, bold_payload)


@pytest.mark.parametrize("limit", [0x80000000, -1, 1 << 100, 1.5])
@pytest.mark.parametrize("codec", ["zstd", "woff2t+zstd"])
@pytest.mark.parametrize("name", ["Synthetic-Regular.ttf", "Synthetic-Regular.woff2", None])
def test_passthrough_ignores_output_limit(codec, name, limit):
    data = fixture(name) if name else b"\x28\xb5"
    assert decode_font_face(data, codec, max_output=limit) == (data, None)


@pytest.mark.parametrize("data", [b"\x00\x01", b"\x28\xb5\x2f\xfd"])
def test_unknown_codec_precedes_passthrough_and_output_limit(data):
    with pytest.raises(ValueError) as error:
        decode_font_face(data, "brotli", max_output=0x80000000)
    assert str(error.value) == "unsupported font face codec: brotli"


@pytest.mark.parametrize("limit", [0x80000000, -1, 1 << 100, 1.5])
@pytest.mark.parametrize("codec", ["zstd", "woff2t+zstd"])
def test_invalid_output_limit_uses_shared_message(codec, limit):
    with pytest.raises(ValueError) as error:
        decode_font_face(fixture("Synthetic-Regular.sfnt.zst"), codec, max_output=limit)
    assert str(error.value) == "font face maxOutput must be an integer from 0 to 2147483647"


def test_corruption_uses_shared_message():
    corrupt = bytearray(fixture("Synthetic-Regular.sfnt.zst"))
    corrupt[-1] ^= 1
    with pytest.raises(ValueError) as error:
        decode_font_face(bytes(corrupt), "zstd")
    assert str(error.value) == "font face checksum failed: corrupt face or wrong dictionary"


def test_truncation_uses_shared_message():
    with pytest.raises(ValueError) as error:
        decode_font_face(b"\x28\xb5\x2f\xfd", "zstd")
    assert str(error.value) == "font face is truncated"
