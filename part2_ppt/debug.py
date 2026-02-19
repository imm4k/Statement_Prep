from __future__ import annotations

import cProfile
import io
import os
import platform
import runpy
import sys
import time
from pathlib import Path
from typing import Iterable, Optional, Tuple


def _hr() -> str:
    return "=" * 98


def _now_ms() -> float:
    return time.perf_counter() * 1000.0


def _fmt_ms(ms: float) -> str:
    if ms < 1000:
        return f"{ms:,.0f} ms"
    return f"{ms/1000.0:,.2f} s"


def _safe_stat(path: Path) -> Tuple[bool, Optional[int]]:
    try:
        st = path.stat()
        return True, int(st.st_size)
    except Exception:
        return False, None


def _time_read_head(path: Path, nbytes: int = 1024 * 1024) -> Tuple[bool, float]:
    t0 = _now_ms()
    try:
        with path.open("rb") as f:
            _ = f.read(nbytes)
        return True, _now_ms() - t0
    except Exception:
        return False, _now_ms() - t0


def _time_listdir(path: Path) -> Tuple[bool, float, int]:
    t0 = _now_ms()
    try:
        items = list(path.iterdir())
        return True, _now_ms() - t0, len(items)
    except Exception:
        return False, _now_ms() - t0, 0


def _find_candidate_pptx(output_dir: Path) -> Optional[Path]:
    if not output_dir.exists():
        return None
    try:
        # Prefer non temp decks
        pptx = sorted(
            [p for p in output_dir.rglob("*.pptx") if "__tmp_updated_" not in p.name],
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        return pptx[0] if pptx else None
    except Exception:
        return None


def _print_latency_checks(paths: Iterable[Tuple[str, Path]]) -> None:
    print(_hr())
    print("Path latency checks")
    print(_hr())

    for label, p in paths:
        exists = p.exists()
        ok_stat, size = _safe_stat(p)

        print(f"{label}: {p}")
        print(f"  exists: {exists}")

        if exists and ok_stat and size is not None and p.is_file():
            ok_read, ms = _time_read_head(p, 1024 * 1024)
            print(f"  size: {size:,} bytes")
            print(f"  read 1 MB: {('ok' if ok_read else 'failed')} in {_fmt_ms(ms)}")

        if exists and p.is_dir():
            ok_ls, ms, n = _time_listdir(p)
            print(f"  listdir: {('ok' if ok_ls else 'failed')} in {_fmt_ms(ms)}, items: {n:,}")

        print("")


def _run_main_under_profile(main_py: Path) -> str:
    pr = cProfile.Profile()

    t0 = _now_ms()
    pr.enable()
    try:
        # Run main.py exactly as if you executed it directly
        runpy.run_path(str(main_py), run_name="__main__")
    finally:
        pr.disable()
    total_ms = _now_ms() - t0

    s = io.StringIO()
    import pstats  # local import to keep top clean

    ps = pstats.Stats(pr, stream=s).strip_dirs().sort_stats("cumulative")

    s.write(_hr() + "\n")
    s.write("Profile summary\n")
    s.write(_hr() + "\n")
    s.write(f"Total runtime: {_fmt_ms(total_ms)}\n\n")

    # Overall top cumulative
    s.write("Top cumulative time, all functions\n")
    ps.print_stats(50)

    # Focused filters for the slow step you care about
    s.write("\n")
    s.write(_hr() + "\n")
    s.write("Focused filters, insertion, slides, shapes, images, save, zip, pptx\n")
    s.write(_hr() + "\n")

    for pattern in [
        "standard",
        "insert",
        "slide",
        "shapes",
        "shape",
        "image",
        "media",
        "relationship",
        "rel",
        "save",
        "zipfile",
        "pptx",
        "Presentation",
    ]:
        s.write(f"\nFilter contains: {pattern}\n")
        ps.print_stats(pattern)

    return s.getvalue()


def main() -> None:
    print("")
    print(_hr())
    print("debug.py: standard slide insertion performance diagnostics")
    print(_hr())
    print(f"Python: {platform.python_version()}")
    print(f"Executable: {sys.executable}")
    print(f"Platform: {platform.platform()}")
    print(f"CWD: {Path.cwd()}")
    script_dir = Path(__file__).resolve().parent
    print(f"Script dir: {script_dir}")
    print("")

    # Import config if present, to locate the most relevant paths for latency checks
    template_dir = None
    output_dir = None
    sqlite_path = None

    try:
        import config  # type: ignore

        template_dir = Path(str(getattr(config, "TEMPLATE_DIR", ""))).expanduser()
        output_dir = Path(str(getattr(config, "OUTPUT_LOCATION", ""))).expanduser()
        sqlite_path = Path(str(getattr(config, "SQLITE_PATH", ""))).expanduser()
    except Exception:
        pass

    # Fallbacks, still give useful checks even if config import fails
    candidates = []
    if template_dir:
        candidates.append(("Template Dir", template_dir))
    if output_dir:
        candidates.append(("Output Location", output_dir))
    if sqlite_path:
        candidates.append(("SQLite", sqlite_path))

    # If output dir exists, try to find a recent pptx to sanity check file read latency
    if output_dir and output_dir.exists():
        pptx = _find_candidate_pptx(output_dir)
        if pptx:
            candidates.append(("Recent PPTX in output", pptx))

    if candidates:
        _print_latency_checks(candidates)
    else:
        print(_hr())
        print("Config import did not yield paths to test.")
        print("This script will still run main.py under profiler.")
        print(_hr())
        print("")

    main_py = script_dir / "main.py"
    if not main_py.exists():
        print("ERROR: main.py not found in the same folder as debug.py")
        print(f"Expected: {main_py}")
        return

    # Optional, reduce noise in output
    os.environ.setdefault("PYTHONWARNINGS", "ignore")

    print(_hr())
    print("Running main.py under cProfile")
    print("This will produce a profile summary at the end.")
    print(_hr())
    print("")

    prof_text = _run_main_under_profile(main_py)

    print("")
    print(prof_text)

    print(_hr())
    print("Next step")
    print(_hr())
    print(
        "Run this twice, once when it is fast and once when it is slow, then paste both profile summaries.\n"
        "The delta will usually point to either zipfile save time, file reads from the template dir, or slide copy."
    )


if __name__ == "__main__":
    main()
