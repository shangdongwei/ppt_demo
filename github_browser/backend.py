import argparse
from http.server import SimpleHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path


STATIC_DIR = Path(__file__).resolve().parent / "static"


class StaticHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(STATIC_DIR), **kwargs)

    def do_GET(self):
        if self.path == "/health":
            body = b"ok"
            self.send_response(200)
            self.send_header("Content-Type", "text/plain; charset=utf-8")
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)
            return

        if self.path == "/":
            self.path = "/index.html"

        return super().do_GET()

    def end_headers(self):
        self.send_header("Cache-Control", "no-store")
        self.send_header("X-Content-Type-Options", "nosniff")
        self.send_header("Referrer-Policy", "no-referrer")
        super().end_headers()

    def log_message(self, format, *args):
        return


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--host", default="127.0.0.1")
    parser.add_argument("--port", default=8000, type=int)
    args = parser.parse_args()

    if not STATIC_DIR.exists():
        raise SystemExit(f"static directory not found: {STATIC_DIR}")

    httpd = ThreadingHTTPServer((args.host, args.port), StaticHandler)
    print(f"Serving {STATIC_DIR} at http://{args.host}:{args.port}/")
    httpd.serve_forever()


if __name__ == "__main__":
    main()
