import json
import secrets
import subprocess
import sys
import os
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlsplit

from .artifacts import Artifacts
from . import control


def make_server(root, port=7331):
    root = Path(root).resolve()
    token = secrets.token_urlsafe(32)
    worker = {'process': None}
    worker_lock = threading.Lock()
    class Handler(BaseHTTPRequestHandler):
        def valid_host(self):
            return self.headers.get('Host') in ('localhost:' + str(self.server.server_port),
                                               '127.0.0.1:' + str(self.server.server_port))

        def do_POST(self):
            if not self.valid_host() or self.headers.get('X-Resume-Token') != token:
                self.send_error(403)
                return
            origin = self.headers.get('Origin')
            if origin and origin != 'http://' + self.headers.get('Host', ''):
                self.send_error(403)
                return
            try:
                size = int(self.headers.get('Content-Length', '0'))
                if not 0 < size <= 4096:
                    raise ValueError('Invalid request size')
                body = json.loads(self.rfile.read(size))
                if self.path == '/api/control':
                    control.write(root, body['state'])
                    message = 'Requested ' + body['state']
                elif self.path == '/api/run':
                    if not Artifacts(root).state.get('run_id'):
                        raise ValueError('Initialize a source DOCX with the init command first')
                    with worker_lock:
                        if worker['process'] is not None and worker['process'].poll() is None:
                            raise ValueError('The monitor already has an active worker')
                        control.write(root, 'continue')
                        logs = root / 'workspace'
                        logs.mkdir(parents=True, exist_ok=True)
                        with (logs / 'worker.log').open('ab') as output:
                            worker['process'] = subprocess.Popen(
                                [sys.executable, '-m', 'src', '--root', str(root), 'resume'],
                                cwd=str(Path(__file__).resolve().parent.parent),
                                stdout=output, stderr=output,
                                creationflags=subprocess.CREATE_NO_WINDOW if os.name == 'nt' else 0)
                    message = 'Worker started; see status and audit events for results'
                else:
                    self.send_error(404)
                    return
            except (ValueError, KeyError, TypeError, OSError) as exc:
                self.send_error(400, str(exc))
                return
            data = json.dumps({'message': message}).encode()
            self.send_response(200)
            self.send_header('Content-Type', 'application/json')
            self.send_header('Content-Length', str(len(data)))
            self.end_headers()
            self.wfile.write(data)

        def do_GET(self):
            if not self.valid_host():
                self.send_error(403)
                return
            path = urlsplit(self.path).path
            if path == '/docs' or path.startswith('/docs/'):
                from .docs_viewer import page
                try:
                    status, data = page(root, self.path)
                except (OSError, ValueError, UnicodeError):
                    status, data = 404, b'Document unavailable'
                self.send_response(status)
                self.send_header('Content-Type', 'text/html; charset=utf-8')
                self.send_header('Content-Length', str(len(data)))
                self.send_header('Cache-Control', 'no-store')
                self.send_header('X-Content-Type-Options', 'nosniff')
                self.send_header('Referrer-Policy', 'no-referrer')
                self.send_header('Content-Security-Policy', "default-src 'none'; style-src 'unsafe-inline'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'")
                self.end_headers()
                self.wfile.write(data)
                return
            store = Artifacts(root)
            if path == '/':
                data = Path(__file__).with_name('monitor.html').read_text(encoding='utf-8').replace('__CONTROL_TOKEN__', token).encode('utf-8')
                mime = 'text/html; charset=utf-8'
            elif path == '/preview/status.json':
                data = json.dumps(store.status()).encode()
                mime = 'application/json'
            elif path == '/audit/events.jsonl':
                file = root / 'audit/events.jsonl'
                data = file.read_bytes() if file.exists() else b''
                mime = 'application/x-ndjson'
            else:
                allowed = {entry['path']: entry for name, entry in store.state.get('handles', {}).items()
                           if name in ('resume.layout.preview', 'resume.render.preview',
                                       'resume.layout.published', 'resume.render.published')}
                # Previously viewed immutable builds remain available without exposing workspace/source.
                if path.startswith('/artifacts/') and path.endswith(('.pdf', '.docx')):
                    name = path.removeprefix('/')
                    events = root / 'audit/events.jsonl'
                    if events.exists():
                        for line in events.read_text(encoding='utf-8').splitlines():
                            try:
                                ev = json.loads(line)
                                build = ev.get('build_id', '')
                                if ev.get('operation') == 'preview.publish' and name in (
                                    'artifacts/resume-' + build + '.pdf', 'artifacts/resume-' + build + '.docx'):
                                    allowed[name] = {}
                            except ValueError:
                                pass
                if path == '/preview/current.pdf':
                    current = store.state.get('preview')
                    if current:
                        self.send_response(302)
                        self.send_header('Location', '/' + current['pdf'])
                        self.send_header('Cache-Control', 'no-store')
                        self.end_headers()
                        return
                name = path.removeprefix('/')
                file = (root / name).resolve()
                if name not in allowed or root not in file.parents or not file.is_file():
                    self.send_error(404)
                    return
                data = file.read_bytes()
                mime = 'application/pdf' if file.suffix == '.pdf' else 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
            self.send_response(200)
            self.send_header('Content-Type', mime)
            self.send_header('Content-Length', str(len(data)))
            self.send_header('Cache-Control', 'no-store')
            self.send_header('X-Content-Type-Options', 'nosniff')
            self.end_headers()
            self.wfile.write(data)

    return ThreadingHTTPServer(('127.0.0.1', port), Handler)


def serve(root, port=7331):
    with make_server(root, port) as server:
        print('Preview monitor: http://localhost:' + str(server.server_port), flush=True)
        server.serve_forever()
