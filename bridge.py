"""
PANG & PPT Studio - Direct Desktop Bridge (No sockets required)
Exposes Python document engine directly into Javascript through window.pywebview.api!
Instant, secure, zero network latency, and works directly in any sandboxed desktop environment.
"""

import os
import sys
import json
import base64
import subprocess
import engine

class StudioBridge:
    def __init__(self):
        self.temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".studio_temp")
        os.makedirs(self.temp_dir, exist_ok=True)

    def ping(self):
        return {"status": "ok", "version": "1.0.0"}

    def inspect_file(self, file_path):
        try:
            res = engine.DocumentInspector.inspect_pdf(file_path)
            return {"success": True, "data": res}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def get_preview_data_url(self, file_path, page=1, dpi=150):
        try:
            png_bytes = engine.DocumentPreviewEngine.render_page(file_path, page, dpi)
            b64 = base64.b64encode(png_bytes).decode("ascii")
            return {"success": True, "data_url": f"data:image/png;base64,{b64}"}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def optimize_pdf(self, file_path, output_path=None, settings=None):
        try:
            if not output_path:
                base, ext = os.path.splitext(file_path)
                output_path = f"{base}_optimized{ext}"
            res = engine.PDFOptimizer.optimize(file_path, output_path, settings)
            return {"success": True, "result": res}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def repair_pdf(self, file_path, output_path=None, settings=None):
        try:
            if not output_path:
                base, ext = os.path.splitext(file_path)
                output_path = f"{base}_repaired{ext}"
            res = engine.PDFRepairEngine.repair_and_flatten(file_path, output_path, settings)
            return {"success": True, "result": res}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def pdf_to_ppt(self, file_path, output_path=None, mode="editable"):
        try:
            if not output_path:
                base, _ = os.path.splitext(file_path)
                output_path = f"{base}.pptx"
            res = engine.PDFToPowerPointConverter.convert(file_path, output_path, mode=mode)
            return {"success": True, "result": res}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def ppt_to_pdf(self, file_path, output_path=None):
        try:
            if not output_path:
                base, _ = os.path.splitext(file_path)
                output_path = f"{base}.pdf"
            res = engine.PowerPointToPDFConverter.convert(file_path, output_path)
            return {"success": True, "result": res}
        except Exception as e:
            return {"success": False, "error": str(e)}

    def open_location(self, file_path):
        try:
            norm = os.path.normpath(os.path.abspath(file_path))
            if sys.platform == "win32":
                subprocess.Popen(["explorer", "/select,", norm])
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", norm])
            else:
                subprocess.Popen(["xdg-open", os.path.dirname(norm)])
            return {"success": True, "path": norm}
        except Exception as e:
            return {"success": False, "error": str(e)}
