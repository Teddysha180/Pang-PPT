"""
PANG & PPT Studio - Local API & Desktop Server
Provides ultra-fast async REST endpoints for the desktop frontend:
- File inspection
- Optimization with live progress & presets
- Compatibility repair & selective flattening
- PDF to PPTX conversion (editable & high-fidelity)
- PPTX to PDF conversion
- Page previews (high-res PNG stream / data-url)
- Batch processing queue
"""

import os
import sys
import json
import uuid
import subprocess
import threading
import time
from http.server import HTTPServer, BaseHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
import engine

TEMP_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".studio_temp")
os.makedirs(TEMP_DIR, exist_ok=True)

# Store job status for background progress tracking
JOBS = {}

class StudioAPIHandler(BaseHTTPRequestHandler):
    def _set_headers(self, status=200, content_type="application/json"):
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, X-Filename")
        self.send_header("Cache-Control", "no-store, no-cache, must-revalidate")
        self.end_headers()

    def do_OPTIONS(self):
        self._set_headers(200)

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)

        # Serve static frontend files (index.html, style.css, app.js)
        if path == "/" or path == "/index.html":
            file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "index.html")
            with open(file_path, "rb") as f:
                self._set_headers(200, "text/html; charset=utf-8")
                self.wfile.write(f.read())
            return

        if path in ["/style.css", "/app.js"]:
            mime = "text/css" if path.endswith(".css") else "application/javascript"
            file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), path.lstrip("/"))
            if os.path.exists(file_path):
                with open(file_path, "rb") as f:
                    self._set_headers(200, mime)
                    self.wfile.write(f.read())
                return

        if path == "/api/health":
            self._set_headers(200)
            self.wfile.write(json.dumps({"status": "running", "engine": "PANG & PPT Studio 1.0"}).encode())
            return

        if path == "/api/preview":
            file_path = query.get("file", [None])[0]
            page = int(query.get("page", [1])[0])
            dpi = int(query.get("dpi", [150])[0])
            if not file_path or not os.path.exists(file_path):
                self._set_headers(404)
                self.wfile.write(json.dumps({"error": "File not found"}).encode())
                return
            try:
                png_bytes = engine.DocumentPreviewEngine.render_page(file_path, page, dpi)
                self._set_headers(200, "image/png")
                self.wfile.write(png_bytes)
            except Exception as e:
                self._set_headers(500)
                self.wfile.write(json.dumps({"error": str(e)}).encode())
            return

        if path == "/api/download":
            file_path = query.get("file", [None])[0]
            if not file_path or not os.path.exists(file_path):
                self._set_headers(404)
                self.wfile.write(json.dumps({"error": "File not found"}).encode())
                return
            filename = os.path.basename(file_path)
            self.send_response(200)
            self.send_header("Content-Type", "application/octet-stream")
            self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
            self.send_header("Content-Length", str(os.path.getsize(file_path)))
            self.send_header("Access-Control-Allow-Origin", "*")
            self.end_headers()
            with open(file_path, "rb") as f:
                while True:
                    chunk = f.read(64 * 1024)
                    if not chunk:
                        break
                    self.wfile.write(chunk)
            return

        if path == "/api/job-status":
            job_id = query.get("id", [None])[0]
            if not job_id or job_id not in JOBS:
                self._set_headers(404)
                self.wfile.write(json.dumps({"error": "Job not found"}).encode())
                return
            self._set_headers(200)
            self.wfile.write(json.dumps(JOBS[job_id]).encode())
            return

        self._set_headers(404)
        self.wfile.write(json.dumps({"error": "Not Found"}).encode())

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        content_length = int(self.headers.get("Content-Length", 0))

        # Direct binary stream upload for large documents (e.g. 90MB+)
        if path == "/api/upload":
            raw_filename = self.headers.get("X-Filename", "document.pdf")
            from urllib.parse import unquote
            filename = unquote(raw_filename)
            saved_path = os.path.join(TEMP_DIR, filename)

            # Stream chunks directly to disk without loading entire 90MB into memory buffer
            remaining = content_length
            chunk_size = 64 * 1024 # 64KB chunks
            with open(saved_path, "wb") as f_out:
                while remaining > 0:
                    read_len = min(remaining, chunk_size)
                    chunk = self.rfile.read(read_len)
                    if not chunk:
                        break
                    f_out.write(chunk)
                    remaining -= len(chunk)

            if os.path.exists(saved_path) and os.path.getsize(saved_path) > 0:
                self._set_headers(200)
                self.wfile.write(json.dumps({
                    "success": True,
                    "file_path": saved_path,
                    "file_name": filename,
                    "file_size": os.path.getsize(saved_path)
                }).encode())
                return
            else:
                self._set_headers(400)
                self.wfile.write(json.dumps({"error": "Failed to stream upload"}).encode())
                return

        post_data = self.rfile.read(content_length)
        body = {}
        if post_data:
            try:
                body = json.loads(post_data.decode("utf-8"))
            except Exception:
                pass

        if path == "/api/open-location":
            file_path = body.get("file_path")
            if not file_path or not os.path.exists(file_path):
                self._set_headers(400)
                self.wfile.write(json.dumps({"error": "Valid file_path is required"}).encode())
                return
            try:
                norm = os.path.normpath(os.path.abspath(file_path))
                if sys.platform == "win32":
                    subprocess.Popen(f'explorer.exe /select,"{norm}"', shell=True)
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", "-R", norm])
                else:
                    subprocess.Popen(["xdg-open", os.path.dirname(norm)])
                self._set_headers(200)
                self.wfile.write(json.dumps({"success": True, "path": norm}).encode())
            except Exception as e:
                self._set_headers(500)
                self.wfile.write(json.dumps({"error": str(e)}).encode())
            return

        if path == "/api/inspect":
            file_path = body.get("file_path")
            if not file_path or not os.path.exists(file_path):
                self._set_headers(400)
                self.wfile.write(json.dumps({"error": "Valid file_path is required"}).encode())
                return
            try:
                info = engine.DocumentInspector.inspect_pdf(file_path)
                self._set_headers(200)
                self.wfile.write(json.dumps({"success": True, "data": info}).encode())
            except Exception as e:
                self._set_headers(500)
                self.wfile.write(json.dumps({"error": str(e)}).encode())
            return

        if path in ["/api/optimize", "/api/repair", "/api/pdf2ppt", "/api/ppt2pdf"]:
            file_path = body.get("file_path")
            output_path = body.get("output_path")
            settings = body.get("settings", {})

            if not file_path or not os.path.exists(file_path):
                self._set_headers(400)
                self.wfile.write(json.dumps({"error": "Valid file_path is required"}).encode())
                return

            job_id = str(uuid.uuid4())
            JOBS[job_id] = {
                "id": job_id,
                "status": "processing",
                "progress": 0,
                "message": "Initializing...",
                "result": None,
                "error": None
            }

            def run_job(j_id, action, in_p, out_p, opts):
                def progress_cb(pct, msg):
                    JOBS[j_id]["progress"] = pct
                    JOBS[j_id]["message"] = msg

                try:
                    if action == "/api/optimize":
                        if not out_p:
                            b, ext = os.path.splitext(in_p)
                            out_p = f"{b}_optimized{ext}"
                        res = engine.PDFOptimizer.optimize(in_p, out_p, opts, progress_cb)
                    elif action == "/api/repair":
                        if not out_p:
                            b, ext = os.path.splitext(in_p)
                            out_p = f"{b}_repaired{ext}"
                        res = engine.PDFRepairEngine.repair_and_flatten(in_p, out_p, opts, progress_cb)
                    elif action == "/api/pdf2ppt":
                        if not out_p:
                            b, _ = os.path.splitext(in_p)
                            out_p = f"{b}.pptx"
                        res = engine.PDFToPowerPointConverter.convert(in_p, out_p, opts.get("mode", "editable"), progress_cb)
                    elif action == "/api/ppt2pdf":
                        if not out_p:
                            b, _ = os.path.splitext(in_p)
                            out_p = f"{b}.pdf"
                        res = engine.PowerPointToPDFConverter.convert(in_p, out_p, progress_cb)

                    JOBS[j_id]["status"] = "completed"
                    JOBS[j_id]["progress"] = 100
                    JOBS[j_id]["message"] = "Finished successfully"
                    JOBS[j_id]["result"] = res
                except Exception as ex:
                    JOBS[j_id]["status"] = "failed"
                    JOBS[j_id]["error"] = str(ex)
                    JOBS[j_id]["message"] = f"Failed: {str(ex)}"

            t = threading.Thread(target=run_job, args=(job_id, path, file_path, output_path, settings), daemon=True)
            t.start()

            self._set_headers(200)
            self.wfile.write(json.dumps({"job_id": job_id, "status": "processing"}).encode())
            return

        self._set_headers(404)
        self.wfile.write(json.dumps({"error": "Endpoint not found"}).encode())


def run_server(port=47891):
    server = HTTPServer(("127.0.0.1", port), StudioAPIHandler)
    print(f"PANG & PPT Studio Engine Server running on http://127.0.0.1:{port}")
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        server.server_close()


if __name__ == "__main__":
    p = 47891
    if len(sys.argv) > 1:
        try:
            p = int(sys.argv[1])
        except ValueError:
            pass
    run_server(p)
