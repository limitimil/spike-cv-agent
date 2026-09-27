import hashlib
from datetime import datetime, timezone
from pathlib import Path
from uuid import uuid4

from . import audit
from .storage import atomic_bytes, read_json, write_json


def build_id():
    return datetime.now(timezone.utc).strftime('%Y%m%d-%H%M%S-%f') + '-' + uuid4().hex[:8]


class Artifacts:
    def __init__(self, root):
        self.root = Path(root).resolve()
        self.state = read_json(self.root / 'manifest.json', {})

    def resolve(self, handle):
        entry = self.state['handles'][handle]
        path = (self.root / entry['path']).resolve()
        if self.root not in path.parents:
            raise ValueError('Artifact escapes workflow root')
        if hashlib.sha256(path.read_bytes()).hexdigest() != entry['sha256']:
            raise ValueError('Artifact integrity check failed: ' + handle)
        return path

    def create(self, data, suffix, build=None):
        build = build or build_id()
        relative = 'artifacts/resume-' + build + suffix
        path = self.root / relative
        path.parent.mkdir(parents=True, exist_ok=True)
        # Exclusive creation; partial orphan files are never registered or reused.
        with path.open('xb') as stream:
            stream.write(data)
            stream.flush()
            import os
            os.fsync(stream.fileno())
        return dict(build_id=build, path=relative, sha256=hashlib.sha256(data).hexdigest())

    def register(self, handle, entry):
        self.state.setdefault('handles', {})[handle] = dict(entry)

    def commit(self, record):
        self.state['last_operation'] = record['operation']
        write_json(self.root / '.transaction.json', dict(state=self.state, event=record))
        self.recover()

    def recover(self):
        pending = self.root / '.transaction.json'
        tx = read_json(pending)
        if tx:
            audit.append(self.root, tx['event'])
            write_json(self.root / 'manifest.json', tx['state'])
            self.state = tx['state']
        self.project()
        if tx:
            pending.unlink()

    def project(self):
        preview = self.state.get('handles', {}).get('resume.render.preview')
        if preview:
            atomic_bytes(self.root / 'preview/current.pdf', self.resolve('resume.render.preview').read_bytes())
        write_json(self.root / 'preview/status.json', self.status())

    def status(self):
        return {key: self.state.get(key) for key in
                ('run_id', 'phase', 'last_operation', 'preview', 'summary', 'next_step', 'review', 'render_error')}
