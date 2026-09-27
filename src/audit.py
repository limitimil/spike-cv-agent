import json
import os
from datetime import datetime, timezone
from uuid import uuid4


def event(operation, run_id, **fields):
    return dict(timestamp=datetime.now(timezone.utc).isoformat(), event_id=uuid4().hex,
                operation=operation, run_id=run_id, **fields)


def append(root, record):
    path = root / 'audit/events.jsonl'
    path.parent.mkdir(parents=True, exist_ok=True)
    # A torn final line is retained as evidence; recovery appends a fresh full record.
    raw = path.read_bytes() if path.exists() else b''
    for line in raw.splitlines():
        try:
            if json.loads(line).get('event_id') == record['event_id']:
                return
        except (ValueError, UnicodeDecodeError):
            pass
    with path.open('ab') as stream:
        if raw and not raw.endswith(b'\n'):
            stream.write(b'\n')
        stream.write((json.dumps(record, ensure_ascii=False) + '\n').encode('utf-8'))
        stream.flush()
        os.fsync(stream.fileno())
