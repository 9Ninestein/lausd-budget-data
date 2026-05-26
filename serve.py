#!/usr/bin/env python3
"""
serve.py
========
Private HTTP server for the LAUSD board meeting materials site.
Requires a shared password (set via SITE_PASSWORD env var or --password).

Usage:
    # Set password once (or pass --password each time)
    set SITE_PASSWORD=your_password_here      # Windows
    export SITE_PASSWORD=your_password_here   # Mac/Linux

    python serve.py                           # listens on 0.0.0.0:8080
    python serve.py --port 9000
    python serve.py --password mysecret --port 8080

Then share http://<your-ip>:8080 with coworkers.
For remote access outside your network, use ngrok: https://ngrok.com
    ngrok http 8080

Requirements: none (uses Python standard library only)
"""

import argparse
import base64
import os
import sys
from functools import partial
from http.server import HTTPServer, SimpleHTTPRequestHandler

REALM = "LAUSD Board Materials"


class AuthHandler(SimpleHTTPRequestHandler):
    password: str = ""

    def do_HEAD(self):
        if not self._check_auth():
            return
        super().do_HEAD()

    def do_GET(self):
        if not self._check_auth():
            return
        super().do_GET()

    def _check_auth(self) -> bool:
        auth_header = self.headers.get("Authorization", "")
        if auth_header.startswith("Basic "):
            try:
                decoded = base64.b64decode(auth_header[6:]).decode("utf-8", errors="replace")
                _, _, provided = decoded.partition(":")
                if provided == self.__class__.password:
                    return True
            except Exception:
                pass
        # Send 401
        self.send_response(401)
        self.send_header("WWW-Authenticate", f'Basic realm="{REALM}"')
        self.send_header("Content-Type", "text/plain")
        self.send_header("Content-Length", "12")
        self.end_headers()
        self.wfile.write(b"Unauthorized")
        return False

    def log_message(self, fmt, *args):
        # Suppress credential details from logs
        if "401" not in str(args):
            super().log_message(fmt, *args)


def main():
    parser = argparse.ArgumentParser(description="Serve LAUSD board materials privately.")
    parser.add_argument("--port",     type=int, default=8080)
    parser.add_argument("--host",     default="0.0.0.0")
    parser.add_argument("--password", default=os.environ.get("SITE_PASSWORD", ""))
    args = parser.parse_args()

    if not args.password:
        print("ERROR: No password set.")
        print("  Set SITE_PASSWORD env var:  set SITE_PASSWORD=yourpassword")
        print("  Or pass --password yourpassword")
        sys.exit(1)

    AuthHandler.password = args.password

    # Serve from repo root (where index.html and boepdfs_tabs/ live)
    os.chdir(os.path.dirname(os.path.abspath(__file__)))

    server = HTTPServer((args.host, args.port), AuthHandler)

    local_ip = _local_ip()
    print(f"\nLAUSD Board Materials server running")
    print(f"  Local:   http://localhost:{args.port}")
    print(f"  Network: http://{local_ip}:{args.port}")
    print(f"  Password: {'*' * len(args.password)}")
    print(f"\nPress Ctrl+C to stop.\n")

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.")


def _local_ip() -> str:
    import socket
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("8.8.8.8", 80))
            return s.getsockname()[0]
    except Exception:
        return "127.0.0.1"


if __name__ == "__main__":
    main()
