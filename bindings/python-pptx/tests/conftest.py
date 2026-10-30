from pathlib import Path

import pytest
from betteroffice_pptx import decode_font_face

ROOT = Path(__file__).resolve().parents[3]
FIXTURES = ROOT / "apps" / "demo" / "public"
FONTS = ROOT / "crates" / "ooxml-text" / "tests" / "fonts"
HEAVY_FONT = ROOT / "packages" / "fonts-cjk" / "assets" / "NotoSerifSC-Regular.otf.zst"


def _read(path: Path) -> bytes:
    if not path.exists():
        pytest.skip(f"fixture missing: {path}")
    return path.read_bytes()


@pytest.fixture(scope="session")
def sample_path() -> Path:
    path = FIXTURES / "betteroffice-demo.pptx"
    if not path.exists():
        pytest.skip(f"fixture missing: {path}")
    return path


@pytest.fixture(scope="session")
def sample_bytes(sample_path: Path) -> bytes:
    return sample_path.read_bytes()


@pytest.fixture(scope="session")
def font_bytes() -> bytes:
    return _read(FONTS / "LiberationSans-Regular.ttf")


@pytest.fixture(scope="session")
def heavy_font_path(tmp_path_factory) -> Path:
    font, _ = decode_font_face(HEAVY_FONT.read_bytes(), "zstd")
    path = tmp_path_factory.mktemp("fonts") / HEAVY_FONT.stem
    path.write_bytes(font)
    return path
