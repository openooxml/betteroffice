from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[3]
FIXTURES = ROOT / "apps" / "demo" / "public"
FONTS = ROOT / "crates" / "ooxml-text" / "tests" / "fonts"


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
