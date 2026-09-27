from .storage import read_json, write_json

STATES = ('continue', 'pause', 'stop_after_checkpoint', 'abort')


def read(root):
    value = read_json(root / 'control/agent-control.json', {'state': 'continue'})
    if value.get('state') not in STATES:
        raise ValueError('Unknown control state; refusing to continue')
    return value['state']


def write(root, state):
    if state not in STATES:
        raise ValueError('Unknown control state')
    write_json(root / 'control/agent-control.json', {'state': state})
