from __future__ import annotations

import gc
import time

import pythoncom
import win32com.client


def combine_presentations(base_pptx_path, standard_pptx_path, out_pptx_path) -> None:
    pythoncom.CoInitialize()
    app = None
    base = None

    try:
        app = win32com.client.DispatchEx("PowerPoint.Application")
        app.Visible = True
        app.DisplayAlerts = 0

        base = app.Presentations.Open(str(base_pptx_path), WithWindow=False)

        insert_index = base.Slides.Count
        print(f"Base slides: {insert_index}. Inserting standard slides at end.")

        base.Slides.InsertFromFile(str(standard_pptx_path), insert_index)

        base.SaveAs(str(out_pptx_path))
        base.Save()

    finally:
        try:
            if base is not None:
                base.Close()
        except Exception:
            pass

        try:
            if app is not None:
                app.Quit()
        except Exception:
            pass

        base = None
        app = None
        gc.collect()
        pythoncom.CoUninitialize()
        time.sleep(1.5)
