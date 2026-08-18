from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from testdata.generate_testdata import generate_all


@pytest.fixture(scope="session")
def testdata() -> dict[str, Path]:
    return generate_all()
