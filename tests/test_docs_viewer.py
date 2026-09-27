import http.client
import os
from pathlib import Path
import subprocess
import tempfile
import threading
import unittest
from urllib.parse import quote
from unittest.mock import patch

from src.docs_viewer import documents, page
from src.server import make_server


class DocsViewerTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)
        (self.root / 'docs/plans').mkdir(parents=True)
        (self.root / 'docs/plans/設計.md').write_text(
            '# Design\n\n## Scope\n\n## Scope\n\n**Backend**\n\n'
            '- [x] Done\n- [ ] Pending\n\n> Quote\n\n```py\nprint(1)\n```\n\n'
            '| A | B |\n| --- | --- |\n| one | two |\n\n'
            '[Other](../OTHER.md#target)\n\n![remote](https://example.com/image.png)\n\n'
            '<script>alert(1)</script>\n\n[bad](javascript:alert(1))\n\n'
            '[private](../../secret.md)\n\n[web](https://example.com)\n\n'
            '[file](file:///secret.md)\n\n@@article@@', encoding='utf-8')
        (self.root / 'docs/OTHER.md').write_text('# Target', encoding='utf-8')
        (self.root / 'docs/private.txt').write_text('not exposed')
        (self.root / 'secret.md').write_text('SECRET')

    def test_render_formatting_links_and_inert_content(self):
        status, body = page(self.root, '/docs/plans/' + quote('設計.md'))
        html = body.decode()
        self.assertEqual(200, status)
        for expected in ('<strong>Backend</strong>', '<table>', '<blockquote>',
                         'type="checkbox"', 'disabled', '<pre><code',
                         'href="/docs/OTHER.md#target"', 'id="scope-1"',
                         'href="https://example.com"', '&lt;script&gt;', '@@article@@'):
            self.assertIn(expected, html)
        for forbidden in ('<script>', '<img', 'href="javascript:', 'href="file:',
                          'href="../../secret.md"', 'SECRET'):
            self.assertNotIn(forbidden, html)

    def test_index_filter_missing_empty_and_refresh(self):
        self.assertEqual(['OTHER.md', 'plans/設計.md'], documents(self.root))
        _, html = page(self.root, '/docs/?q=other')
        self.assertIn(b'OTHER.md', html)
        self.assertNotIn(quote('設計.md').encode(), html)
        self.assertIn(b'No matching documents', page(self.root, '/docs/?q=absent')[1])
        (self.root / 'docs/empty.md').write_text('')
        self.assertIn(b'This document is empty', page(self.root, '/docs/empty.md')[1])
        (self.root / 'docs/empty.md').write_text('# Fresh')
        self.assertIn(b'>Fresh</h1>', page(self.root, '/docs/empty.md')[1])
        with tempfile.TemporaryDirectory() as empty:
            self.assertIn(b'No Markdown documents', page(empty, '/docs/')[1])

    def test_http_boundary_and_no_store_access(self):
        server = make_server(self.root, 0)
        thread = threading.Thread(target=server.serve_forever, daemon=True)
        thread.start()
        try:
            with patch('src.server.Artifacts', side_effect=AssertionError('No store access')):
                for path in ('/docs/', '/docs/OTHER.md', '/docs/plans/' + quote('設計.md')):
                    status, headers, body = self.request(server, path)
                    self.assertEqual(200, status)
                    self.assertIn("default-src 'none'", headers['Content-Security-Policy'])
                    self.assertEqual('no-store', headers['Cache-Control'])
                for path in ('/docs/../secret.md', '/docs/%2e%2e/secret.md',
                             '/docs/%252e%252e/secret.md', '/docs/%2fsecret.md',
                             '/docs/C:%5csecret.md', '/docs/..%5csecret.md',
                             '/docs/OTHER.md:$DATA', '/docs/private.txt',
                             '/docs/missing.md', '/docs/%00.md'):
                    status, _, body = self.request(server, path)
                    self.assertEqual(404, status, path)
                    self.assertNotIn(b'SECRET', body)
                    self.assertNotIn(str(self.root).encode(), body)
            self.assertFalse((self.root / 'manifest.json').exists())
            self.assertFalse((self.root / 'audit').exists())
            self.assertEqual(403, self.request(server, '/docs/', host='evil.example')[0])
        finally:
            server.shutdown()
            server.server_close()
            thread.join(2)

    def request(self, server, path, host=None):
        conn = http.client.HTTPConnection('127.0.0.1', server.server_port)
        try:
            conn.request('GET', path, headers={'Host': host} if host else {})
            response = conn.getresponse()
            return response.status, dict(response.getheaders()), response.read()
        finally:
            conn.close()

    def test_symlink_is_not_listed_or_read(self):
        alias = self.root / 'docs/alias.md'
        try:
            alias.symlink_to(self.root / 'secret.md')
        except OSError:
            self.skipTest('Symlink creation unavailable')
        self.assertNotIn('alias.md', documents(self.root))
        self.assertEqual(404, page(self.root, '/docs/alias.md')[0])

    @unittest.skipUnless(os.name == 'nt', 'Windows junction boundary')
    def test_junction_is_not_listed_or_read(self):
        outside = self.root / 'outside'
        outside.mkdir()
        (outside / 'secret.md').write_text('SECRET')
        alias = self.root / 'docs/alias'
        # Creation only: cleanup stays in Python and removes the junction itself.
        result = subprocess.run(['cmd', '/c', 'mklink', '/J', str(alias), str(outside)],
                                capture_output=True)
        self.assertEqual(0, result.returncode)
        try:
            self.assertNotIn('alias/secret.md', documents(self.root))
            self.assertEqual(404, page(self.root, '/docs/alias/secret.md')[0])
        finally:
            alias.rmdir()


if __name__ == '__main__':
    unittest.main()
