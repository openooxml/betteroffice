from pathlib import Path

import pytest

FIXTURES = Path(__file__).resolve().parents[3] / "apps" / "demo" / "public"


@pytest.fixture(scope="session")
def sample_path() -> Path:
    path = FIXTURES / "sample.xlsx"
    if not path.exists():
        pytest.skip(f"fixture missing: {path}")
    return path


@pytest.fixture(scope="session")
def sample_bytes(sample_path: Path) -> bytes:
    return sample_path.read_bytes()
