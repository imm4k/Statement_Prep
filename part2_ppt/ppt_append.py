from __future__ import annotations

import gc
import time

import pythoncom
import win32com.client

def combine_presentations(base_pptx_path: str, standard_pptx_path: str, out_pptx_path: str) -> None:
    import shutil
    import time
    from pathlib import Path

    import pythoncom
    import win32com.client

    def _now() -> float:
        return time.perf_counter()

    def _fmt(s: float) -> str:
        if s < 1:
            return f"{s * 1000:.0f} ms"
        return f"{s:.2f} s"

    base_src = Path(base_pptx_path)
    std_src = Path(standard_pptx_path)
    out_dst = Path(out_pptx_path)

    tmp_dir = out_dst.parent / "__tmp_combine_local"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    base_local = tmp_dir / f"__base__{base_src.name}"
    std_local = tmp_dir / f"__standard__{std_src.name}"
    out_local = tmp_dir / f"__out__{out_dst.name}"

    print("")
    print("ppt_append.combine_presentations diagnostics")
    print(f"  base_src: {base_src}")
    print(f"  std_src:  {std_src}")
    print(f"  out_dst:  {out_dst}")
    print(f"  staging:  {tmp_dir}")

    t0 = _now()
    shutil.copy2(base_src, base_local)
    t_copy_base = _now()
    shutil.copy2(std_src, std_local)
    t_copy_std = _now()

    print(f"  copy base to local: {_fmt(t_copy_base - t0)}")
    print(f"  copy std  to local: {_fmt(t_copy_std - t_copy_base)}")

    pythoncom.CoInitialize()

    ppt = None
    pres = None
    std_pres = None

    try:
        t1 = _now()
        ppt = win32com.client.DispatchEx("PowerPoint.Application")
        t2 = _now()

        try:
            ppt.DisplayAlerts = 0
        except Exception:
            pass

        print(f"  DispatchEx PowerPoint: {_fmt(t2 - t1)}")

        t3 = _now()
        pres = ppt.Presentations.Open(str(base_local), WithWindow=False)
        t4 = _now()
        print(f"  Open base (local): {_fmt(t4 - t3)}")

        t_std0 = _now()
        std_pres = ppt.Presentations.Open(str(std_local), WithWindow=False, ReadOnly=True)
        std_slide_count = std_pres.Slides.Count
        std_pres.Close()
        std_pres = None
        t_std1 = _now()
        print(f"  Open std for count (local): {_fmt(t_std1 - t_std0)} slides: {std_slide_count}")

        t5 = _now()
        insert_index = pres.Slides.Count
        pres.Slides.InsertFromFile(str(std_local), insert_index, 1, std_slide_count)
        t6 = _now()
        print(f"  InsertFromFile std (local): {_fmt(t6 - t5)}")

        if out_local.exists():
            try:
                out_local.unlink()
            except Exception:
                pass

        t7 = _now()
        pres.SaveAs(str(out_local))
        t8 = _now()
        print(f"  SaveAs out (local): {_fmt(t8 - t7)}")

    finally:
        try:
            if std_pres is not None:
                std_pres.Close()
        except Exception:
            pass

        try:
            if pres is not None:
                pres.Close()
        except Exception:
            pass

        try:
            if ppt is not None:
                ppt.Quit()
        except Exception:
            pass

        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    t9 = _now()
    out_dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(out_local, out_dst)
    t10 = _now()
    print(f"  copy out to destination: {_fmt(t10 - t9)}")
    print(f"  total combine time: {_fmt(t10 - t0)}")

    try:
        shutil.rmtree(tmp_dir)
        print("  cleaned up staging directory")
    except Exception:
        print("  staging directory cleanup failed")




