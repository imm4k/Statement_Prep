from __future__ import annotations

import subprocess
import time
from pathlib import Path

import win32com.client

import config


INVESTOR = "Nicholas Granato"


def _tasklist_powerpnt() -> str:
    p = subprocess.run(
        ["cmd", "/c", "tasklist", "/fi", "imagename eq POWERPNT.EXE"],
        capture_output=True,
        text=True,
    )
    return (p.stdout or "").strip()


def _count_powerpnt_processes() -> int:
    out = _tasklist_powerpnt()
    if "No tasks are running" in out:
        return 0
    lines = [ln for ln in out.splitlines() if ln.strip().startswith("POWERPNT.EXE")]
    return len(lines)


def main() -> None:
    template_dir = Path(config.TEMPLATE_DIR)
    standard_path = template_dir / config.STANDARD_SLIDES_FILENAME
    base_path = template_dir / config.INVESTOR_TEMPLATE_FORMAT.format(investor=INVESTOR)

    out_dir = template_dir / "_debug_outputs"
    out_dir.mkdir(parents=True, exist_ok=True)
    out_path = out_dir / "DEBUG_COM_output.pptx"

    print("POWERPNT processes BEFORE:")
    print(_tasklist_powerpnt())
    print(f"Count: {_count_powerpnt_processes()}")

    app = win32com.client.DispatchEx("PowerPoint.Application")
    app.Visible = True

    base = None
    try:
        print(f"Opening base: {base_path}")
        base = app.Presentations.Open(str(base_path), WithWindow=True)

        insert_index = base.Slides.Count
        print(f"Inserting standard slides at end. Base slides: {insert_index}")
        base.Slides.InsertFromFile(str(standard_path), insert_index)

        if out_path.exists():
            out_path.unlink()

        print(f"Saving: {out_path}")
        base.SaveAs(str(out_path))

        print("Open presentations (BEFORE close):")
        for i in range(1, app.Presentations.Count + 1):
            pres = app.Presentations.Item(i)
            print(f"  [{i}] Name='{pres.Name}' FullName='{pres.FullName}'")

    finally:
        try:
            if base is not None:
                base.Close()
        except Exception as e:
            print(f"Base close exception: {e}")

        try:
            print("Presentations count (AFTER base close):", app.Presentations.Count)
            for i in range(1, app.Presentations.Count + 1):
                pres = app.Presentations.Item(i)
                print(f"  Remaining [{i}] Name='{pres.Name}' FullName='{pres.FullName}'")
        except Exception as e:
            print(f"Enumerate after close exception: {e}")

        try:
            app.Quit()
        except Exception as e:
            print(f"Quit exception: {e}")

        app = None

    time.sleep(2)

    print("POWERPNT processes AFTER Quit:")
    print(_tasklist_powerpnt())
    print(f"Count: {_count_powerpnt_processes()}")

    print(f"Debug output saved: {out_path}")


if __name__ == "__main__":
    main()
