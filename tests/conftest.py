import importlib.util
import json
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture(scope="session")
def library(tmp_path_factory):
    spec = importlib.util.spec_from_file_location("make_library", ROOT / "examples" / "make_library.py")
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    p = tmp_path_factory.mktemp("lib") / "library.pptx"
    mod.main(str(p))
    return p


@pytest.fixture(scope="session")
def demo_plan():
    return json.loads((ROOT / "examples" / "plan_demo.json").read_text(encoding="utf-8"))


@pytest.fixture(scope="session")
def demo_out(library, demo_plan, tmp_path_factory):
    from ade.build import build
    out = tmp_path_factory.mktemp("out") / "out.pptx"
    build(library, demo_plan, out)
    return out
