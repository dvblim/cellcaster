#!/usr/bin/env python3
"""Runs the engine's VBA modules for real and checks every cell they produce.

Excel is not needed: LibreOffice Basic has a VBA compatibility mode that runs
the modules that are pure maths (Config, World, Textures, Raycaster). The
frames they render are compared, cell by cell, with the Python reference port.
Platform, Renderer and Game depend on Excel and Windows and are not covered.

    python tests/run_vba_tests.py        (needs `soffice` on the PATH)
"""
import html
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT / "tools"))
import reference_renderer as R  # noqa: E402

MODULES = ["Config", "World", "Textures", "Raycaster"]
XBA = ('<?xml version="1.0" encoding="UTF-8"?>\n'
       '<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">\n'
       '<script:module xmlns:script="http://openoffice.org/2000/script" script:name="{name}" '
       'script:language="StarBasic">{code}</script:module>\n')
XLB = ('<?xml version="1.0" encoding="UTF-8"?>\n'
       '<!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">\n'
       '<library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" '
       'library:readonly="false" library:passwordprotected="false">\n{elements}</library:library>\n')


def run_vba(work):
    """Installs the modules into a throwaway LibreOffice profile and runs them."""
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        sys.exit("soffice not found: install LibreOffice to run the VBA tests")
    profile = "-env:UserInstallation=" + (work / "profile").as_uri()
    subprocess.run([soffice, "--headless", profile, "--terminate_after_init"],
                   timeout=180, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)

    lib = work / "profile" / "user" / "basic" / "Standard"
    lib.mkdir(parents=True, exist_ok=True)
    sources = {}
    for name in MODULES:
        lines = (ROOT / "src" / f"{name}.bas").read_text(encoding="ascii").splitlines()
        sources[name] = "Option VBASupport 1\n" + "\n".join(l for l in lines if not l.startswith("Attribute "))
    sources["Harness"] = "\n".join((ROOT / "tests" / "harness.bas").read_text(encoding="ascii").splitlines())
    for name, code in sources.items():
        (lib / f"{name}.xba").write_text(XBA.format(name=name, code=html.escape(code)), encoding="utf-8")
    elements = "".join(f' <library:element library:name="{n}"/>\n' for n in sources)
    (lib / "script.xlb").write_text(XLB.format(elements=elements), encoding="utf-8")

    out = work / "frames.txt"
    subprocess.run([soffice, "--headless", profile, "macro:///Standard.Harness.Main"], timeout=300,
                   env={**os.environ, "CELLCASTER_OUT": str(out)},
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    if not out.exists():
        sys.exit("the VBA harness produced no output (did it fail to compile?)")
    return out.read_text().splitlines()


def main():
    problems = R.check_level()
    if problems:
        sys.exit("level: " + "; ".join(problems))

    with tempfile.TemporaryDirectory() as tmp:
        lines = run_vba(Path(tmp))
    if not lines or lines[-1] != "DONE":
        sys.exit("VBA run failed: " + (lines[-1] if lines else "empty output"))

    cursor = 0
    failures = 0

    def take(n):
        nonlocal cursor
        chunk = lines[cursor:cursor + n]
        cursor += n
        return chunk

    def check(name, ok, detail=""):
        nonlocal failures
        failures += not ok
        print(f"{'ok  ' if ok else 'FAIL'} {name} {detail}")

    def check_frame(name, expected):
        got = [[int(v) for v in row.split()] for row in take(R.SCREEN_H)]
        bad = sum(a != b for ra, rb in zip(expected, got) for a, b in zip(ra, rb))
        check(name, bad == 0 and len(got) == R.SCREEN_H, f"({bad} of {R.SCREEN_W * R.SCREEN_H} cells differ)")

    palette = [r | g << 8 | b << 16 for r, g, b in map(R.palette_rgb, range(1, R.CFG["PALETTE_SIZE"] + 1))]
    check("palette", [int(v) for v in take(1)[0].split()] == palette)

    w = R.World(); w.px, w.py = 3.5, 4.5; w.rotate(0.35)
    check_frame("frame: stone hall + minimap", R.render(w, 0.4, True))
    w = R.World(); w.px, w.py = 16.6, 17.5; w.rotate(-0.25)
    check_frame("frame: gem billboard", R.render(w, 1.1))
    w = R.World(); w.px, w.py = 19.5, 14.5; w.rotate(-1.5707963)
    check_frame("frame: facing north", R.render(w, 2.3))
    w = R.World(); w.rotate(2.9)
    check_frame("frame: nose against a corner", R.render(w, 0.0))

    w = R.World()
    for _ in range(400):
        w.rotate(0.011)
        w.try_move(w.dx * 0.09 + 0.01, w.dy * 0.09 + 0.03)
    got = [int(v) for v in take(1)[0].split()]
    check("collision walk", abs(got[0] - w.px * 1e6) <= 1 and abs(got[1] - w.py * 1e6) <= 1,
          f"(ended at {w.px:.3f}, {w.py:.3f})")
    check_frame("frame: after the walk", R.render(w, 5.0, True))

    w = R.World()
    picked = 0
    for g in list(w.gems):
        w.px, w.py = g[0] + 0.1, g[1] - 0.1
        picked += w.collect()
    check("gem pickups", take(1)[0] == f"{picked} 0" and picked == len(w.gems), f"({picked} gems)")
    check_frame("frame: all gems collected", R.render(w, 0.7, True))

    print("all checks passed" if failures == 0 else f"{failures} check(s) failed")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
