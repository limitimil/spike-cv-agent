import argparse
import json
import subprocess
from pathlib import Path

from . import control
from .artifacts import Artifacts
from .operations import REGISTRY
from .server import serve
from .storage import lock, read_json
from .workflow import Workflow
from .rendering import RendererFactory, RendererError


def main():
    parser = argparse.ArgumentParser(description='Resumable resume layout workflow')
    parser.add_argument('--root', type=Path, default=Path.cwd())
    sub = parser.add_subparsers(dest='command', required=True)
    init = sub.add_parser('init')
    init.add_argument('--source', type=Path, required=True)
    init.add_argument('--design', type=Path, default=Path('design.json'))
    for name in ('run', 'resume', 'checkpoint', 'status', 'operations', 'doctor'):
        sub.add_parser(name)
    change = sub.add_parser('control')
    change.add_argument('state', choices=control.STATES)
    for name in ('start', 'serve'):
        web = sub.add_parser(name)
        web.add_argument('--port', type=int, default=7331)
    op = sub.add_parser('operate')
    op.add_argument('name', choices=REGISTRY)
    parameters = op.add_mutually_exclusive_group()
    parameters.add_argument('--params', default='{}', help='JSON operation parameters, or @path to a JSON file')
    parameters.add_argument('--design', type=Path, help='Design JSON file for layout.apply')
    op.add_argument('--reason', required=True)
    for name in ('feedback', 'review'):
        command = sub.add_parser(name)
        command.add_argument('build')
        command.add_argument('message')
        if name == 'review':
            command.add_argument('--accept', action='store_true')
    publish = sub.add_parser('publish')
    publish.add_argument('build')
    args = parser.parse_args()
    root = args.root.resolve()
    try:
        if args.command in ('start', 'serve'):
            serve(root, args.port)
            return
        if args.command == 'doctor':
            import sys
            report = {'python': sys.version.split()[0], 'root': str(root),
                      'initialized': bool(Artifacts(root).state.get('run_id'))}
            try:
                renderer = RendererFactory.create()
                renderer.validate_environment()
                report.update(ready=True, renderer=renderer.metadata(), next_action='Ready to render')
            except RendererError as exc:
                report.update(ready=False, error_code=exc.code, next_action=exc.message)
            print(json.dumps(report, indent=2))
            if not report['ready']:
                parser.exit(1)
            return
        if args.command == 'control':
            # Separate channel intentionally does not wait for the workflow writer lock.
            control.write(root, args.state)
            print('Requested ' + args.state)
            return
        if args.command == 'status':
            print(json.dumps(Artifacts(root).status(), indent=2))
            return
        if args.command == 'operations':
            print(json.dumps({k: {'inputs': v.inputs, 'outputs': v.outputs, 'parameters': v.parameters}
                              for k, v in REGISTRY.items()}, indent=2))
            return
        with lock(root):
            workflow = Workflow(root)
            if args.command == 'init':
                workflow.init(args.source, read_json(args.design if args.design.is_absolute() else root / args.design))
            elif args.command in ('run', 'resume'):
                workflow.run()
            elif args.command == 'checkpoint':
                workflow.checkpoint()
            elif args.command == 'operate':
                if not args.reason.strip():
                    raise ValueError('A design reason is required')
                if args.design and args.name != 'layout.apply':
                    raise ValueError('--design is only supported for layout.apply')
                if args.design:
                    params = {'design': read_json(args.design)}
                elif args.params.startswith('@'):
                    params_path = Path(args.params[1:])
                    params = read_json(params_path if params_path.is_absolute() else root / params_path)
                    if not isinstance(params, dict):
                        raise ValueError('Parameter file must contain a JSON object')
                else:
                    params = json.loads(args.params)
                workflow.execute(args.name, params, args.reason)
            elif args.command == 'feedback':
                workflow.feedback(args.build, args.message)
            elif args.command == 'review':
                workflow.feedback(args.build, args.message, args.accept)
            elif args.command == 'publish':
                workflow.publish(args.build)
            print(json.dumps(workflow.store.status(), indent=2))
    except (ValueError, RuntimeError, OSError, KeyError, subprocess.SubprocessError) as exc:
        parser.exit(1, 'Error: ' + str(exc) + '\n')


if __name__ == '__main__':
    main()
