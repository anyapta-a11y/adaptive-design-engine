import argparse
import json
import sys

from .build import build
from .catalog import build_catalog, dumps
from .errors import AdeError
from .qa import run_qa


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="ade")
    sub = ap.add_subparsers(dest="cmd", required=True)
    c = sub.add_parser("catalog", help="library.pptx -> catalog.json (stdout)")
    c.add_argument("library")
    c.add_argument("--out")
    b = sub.add_parser("build", help="plan + library -> pptx")
    b.add_argument("--library", required=True)
    b.add_argument("--plan", required=True)
    b.add_argument("--out", required=True)
    q = sub.add_parser("qa", help="проверка pptx -> qa.json")
    q.add_argument("pptx")
    q.add_argument("--plan")
    q.add_argument("--pdf", action="store_true")
    q.add_argument("--json", help="путь qa.json (по умолчанию рядом с pptx)")
    a = ap.parse_args(argv)
    try:
        if a.cmd == "catalog":
            text = dumps(build_catalog(a.library))
            if a.out:
                open(a.out, "w", encoding="utf-8").write(text)
            else:
                print(text)
        elif a.cmd == "build":
            rep = build(a.library, a.plan, a.out)
            print(f"OK: {len(rep['slides'])} слайдов -> {a.out}")
            for i, s in enumerate(rep["slides"], 1):
                print(f"  {i}: {s['variant']} ({s['action']})")
        else:
            r = run_qa(a.pptx, a.plan, a.pdf, a.json)
            print(f"QA {r['status']}: слайдов {r['slides']}, fails {len(r['fails'])}, warns {len(r['warns'])}")
            for f in r["fails"]:
                print("  FAIL", f)
            for w in r["warns"]:
                print("  WARN", w)
            if r["pdf"]:
                print("  PDF", r["pdf"]["path"], "страниц:", r["pdf"]["pages"])
            return 0 if r["status"] == "PASS" else 1
    except AdeError as e:
        print(f"Ошибка: {e}", file=sys.stderr)
        return 2
    return 0
