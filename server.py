"""Servidor estático BR Drive — sem cache."""
import http.server, sys
from pathlib import Path

PORT = int(sys.argv[1]) if len(sys.argv) > 1 else 8502
DIR  = Path(sys.argv[2]) if len(sys.argv) > 2 else Path(__file__).parent / "brdrive_output"

class Handler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *a, **kw):
        super().__init__(*a, directory=str(DIR), **kw)
    def log_message(self, *a):
        pass
    def end_headers(self):
        self.send_header("Cache-Control", "no-cache, no-store, must-revalidate")
        self.send_header("Pragma", "no-cache")
        self.send_header("Expires", "0")
        super().end_headers()

http.server.HTTPServer(("0.0.0.0", PORT), Handler).serve_forever()
