import http.server, os, socketserver

PASTA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "brdrive_output")
PORTA = 8080

os.chdir(PASTA)

class Handler(http.server.SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        print(f"  [{self.address_string()}] {format % args}")

with socketserver.ThreadingTCPServer(("", PORTA), Handler) as httpd:
    httpd.serve_forever()
