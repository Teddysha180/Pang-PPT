"""
PANG & PPT Studio - Windows Desktop Application Entrypoint
Uses native pywebview with Chromium Edge WebView2 backend and direct Python Bridge.
Completely offline, zero network sockets needed, maximum execution security and performance.
"""

import os
import sys
import webview
from bridge import StudioBridge

def main():
    current_dir = os.path.dirname(os.path.abspath(__file__))
    index_path = os.path.join(current_dir, "index.html")

    api_bridge = StudioBridge()

    window = webview.create_window(
        title="PANG & PPT Studio — Professional PDF, PowerPoint & Document Studio",
        url=f"file://{index_path}",
        js_api=api_bridge,
        width=1380,
        height=880,
        min_size=(1024, 700),
        background_color="#090909"
    )

    webview.start(debug=False)

if __name__ == "__main__":
    main()
