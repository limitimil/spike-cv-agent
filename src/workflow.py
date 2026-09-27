import copy
import json
from pathlib import Path

from . import audit, control
from .artifacts import Artifacts, build_id
from .operations import REGISTRY, validate_design
from .rendering import DocxToPdfRenderer, RendererFactory, RendererError
from .storage import atomic_bytes, read_json


class Workflow:
    """Caller holds the writer lock for all mutations."""
    def __init__(self, root, renderer: DocxToPdfRenderer = None):
        self.store = Artifacts(root)
        self.root = self.store.root
        self.renderer = renderer
        self.store.recover()

    @property
    def state(self):
        return self.store.state

    def record(self, operation, **fields):
        if getattr(self, 'operation_reason', None):
            fields['requested_reason'] = self.operation_reason
        self.store.commit(audit.event(operation, self.state.get('run_id'), **fields))

    def init(self, source, design):
        if self.state.get('run_id'):
            raise ValueError('Workflow already initialized')
        validate_design({'design': design})
        data = Path(source).read_bytes()
        REGISTRY['source.inspect'].execute(data, {})
        entry = self.store.create(data, '.docx')
        self.store.state = dict(run_id='run-' + build_id(), handles={}, phase='ready',
                                next_step=0, generation=0, design=design, summary=[])
        self.store.register('resume.source', entry)
        self.store.register('resume.layout.workspace', entry)
        self.record('source.import', artifact='resume.source', after=entry, reason='Initialize workflow from explicit source')
        self.working_copy()

    def working_copy(self):
        atomic_bytes(self.root / 'workspace/resume-working.docx', self.store.resolve('resume.layout.workspace').read_bytes())

    def execute(self, name, params, reason, advance=False):
        op = REGISTRY[name]
        op.validate(params)
        if op.lifecycle:
            self.operation_reason = reason
            try:
                return op.execute(self, params)
            finally:
                self.operation_reason = None
        before = copy.deepcopy(self.state['handles']['resume.layout.workspace'])
        data = self.store.resolve(op.inputs[0]).read_bytes()
        if name in ('layout.inspect', 'layout.compose', 'content.revise') and advance:
            raise ValueError('Composition operations are explicit; they do not advance the default cursor')
        if name in ('layout.compose', 'content.revise'):
            if self.state.get('next_step', 0) < 3:
                raise ValueError('Complete the existing layout step before composing')
            spec = params['specification']
            if before['sha256'] != spec['expected_workspace_sha256']:
                raise ValueError('Stale composition workspace checksum')
            inspection = self.state['handles'].get('resume.layout.inspection')
            if not inspection or inspection['sha256'] != spec['expected_inspection_sha256']:
                raise ValueError('Stale or missing composition inspection checksum')
            report = json.loads(self.store.resolve('resume.layout.inspection').read_text(encoding='utf-8'))
            if report.get('schema_version') != 1 or report.get('workspace_sha256') != before['sha256']:
                raise ValueError('Inspection belongs to a different workspace revision')
        output, result = op.execute(data, params)
        if name == 'content.revise':
            result['source'] = copy.deepcopy(self.state['handles']['resume.source'])
            report_entry = self.store.create(
                json.dumps(result, ensure_ascii=False, indent=2).encode('utf-8'), '.revision.json')
            self.store.register('resume.content.revision', report_entry)
        if output is not None:
            entry = self.store.create(output, '.docx')
            self.store.register(op.outputs[0], entry)
            self.state['phase'] = 'editing'
            self.state['review'] = None
            if name == 'content.revise':
                self.state['handles'].pop('resume.layout.inspection', None)
        else:
            entry = self.store.create(json.dumps(result, ensure_ascii=False, indent=2).encode('utf-8'), '.inspection.json')
            self.store.register(op.outputs[0], entry)
        if 'design' in params:
            old_design = self.state['design']
            self.state['design'] = params['design']
            before = old_design
        self.state['summary'] = (self.state['summary'] + [name + ': ' + reason])[-20:]
        if advance:
            self.state['next_step'] += 1
        self.record(name, artifact=op.outputs[0], build_id=entry['build_id'],
                    **op.audit_serializer(before, result, reason))
        if name != 'layout.inspect':
            self.working_copy()

    def render(self):
        workspace = self.state['handles']['resume.layout.workspace']
        pending = self.state.get('rendered')
        build = build_id()
        docx = None
        staged_pdf = self.root / 'workspace' / ('render-' + build + '.pdf')
        try:
            if self.renderer is None:
                self.renderer = RendererFactory.create()
            self.renderer.validate_environment()
            identity = self.renderer.metadata()
            if (pending and pending['workspace_sha256'] == workspace['sha256']
                    and pending['provenance'] == identity):
                return
            docx = self.store.create(self.store.resolve('resume.layout.workspace').read_bytes(), '.docx', build)
            self.store.register('resume.layout.checkpoint', docx)
            self.record('layout.checkpoint', build_id=build, artifact='resume.layout.checkpoint',
                        reason='Preserve immutable DOCX before rendering')
            staged_pdf.parent.mkdir(parents=True, exist_ok=True)
            self.renderer.render(self.root / docx['path'], staged_pdf)
            pdf = staged_pdf.read_bytes()
            if not pdf.startswith(b'%PDF-'):
                raise RendererError('RENDER_FAILED', 'Renderer output is not a PDF.')
            provenance = self.renderer.metadata()
            pdf_entry = self.store.create(pdf, '.pdf', build)
        except Exception as exc:
            error = exc if isinstance(exc, RendererError) else RendererError('RENDER_FAILED', str(exc))
            if docx is None:
                docx = self.store.create(self.store.resolve('resume.layout.workspace').read_bytes(), '.docx', build)
                self.store.register('resume.layout.checkpoint', docx)
            metadata = self.renderer.metadata() if self.renderer else {}
            self.state['phase'] = 'error'
            self.state['render_error'] = dict(code=error.code, message=error.message, build_id=build,
                                               checkpoint=docx['path'])
            self.record('render.docx_to_pdf', build_id=build, input_handle='resume.layout.checkpoint',
                        input_path=docx['path'], renderer=metadata.get('type'),
                        platform=metadata.get('platform'), renderer_metadata=metadata,
                        status='failed', error_code=error.code, reason=error.message)
            raise error from exc
        finally:
            staged_pdf.unlink(missing_ok=True)
        self.state['render_error'] = None
        self.state['rendered'] = dict(docx=docx, pdf=pdf_entry, provenance=provenance,
                                      workspace_sha256=workspace['sha256'])
        self.store.register('resume.render.candidate', pdf_entry)
        self.record('render.docx_to_pdf', build_id=build, input_handle='resume.layout.checkpoint',
                    input_path=docx['path'], artifact='resume.render.candidate',
                    renderer=provenance.get('type'), platform=provenance.get('platform'),
                    renderer_metadata=provenance, status='success',
                    reason='Render immutable document checkpoint')

    def checkpoint(self):
        self.render()
        pending = self.state['rendered']
        if self.state.get('preview', {}).get('build_id') == pending['docx']['build_id']:
            return
        self.state['generation'] += 1
        generation = self.state['generation']
        for handle, key in [('resume.layout.preview', 'docx'), ('resume.render.preview', 'pdf')]:
            self.store.register(handle, dict(pending[key], generation=generation))
        self.state['preview'] = dict(build_id=pending['docx']['build_id'], generation=generation,
                                     pdf=pending['pdf']['path'], docx=pending['docx']['path'],
                                     provenance=pending['provenance'])
        self.state['review'] = None
        self.record('preview.publish', build_id=pending['docx']['build_id'], generation=generation,
                    reason='Expose complete immutable checkpoint')

    def run(self):
        if not self.state.get('run_id'):
            raise ValueError('Initialize the workflow first')
        plan = [('source.inspect', {}, 'Inventory source structure and paragraph indices'),
                ('content.normalize', {}, 'Normalize nonbreaking spaces while preserving runs'),
                ('layout.apply', {'design': self.state['design']}, 'Apply declared resume design')]
        self.state['phase'] = 'running'
        self.record('workflow.resume', reason='Resume from persisted operation cursor')
        try:
            while True:
                command = control.read(self.root)
                if command == 'abort':
                    self.state['phase'] = 'aborted'
                    self.record('control.abort', reason='Stop at atomic operation boundary; retain last preview')
                    return
                if self.renderer is None:
                    try:
                        self.renderer = RendererFactory.create()
                    except RendererError:
                        self.render()  # Audit and retain DOCX before editing on an unsupported OS.
                if command in ('pause', 'stop_after_checkpoint'):
                    self.checkpoint()
                    self.state['phase'] = 'paused' if command == 'pause' else 'stopped'
                    self.record('control.' + command, reason='Checkpoint saved; process may exit safely')
                    return
                index = self.state['next_step']
                if index >= len(plan):
                    self.checkpoint()
                    if control.read(self.root) != 'continue':
                        continue
                    self.state['phase'] = 'awaiting_review'
                    self.record('workflow.complete', reason='Layout complete; awaiting build-specific visual review')
                    return
                name, params, reason = plan[index]
                self.execute(name, params, reason, advance=True)
        except Exception as exc:
            # Reload committed state so failed in-memory changes are never committed.
            self.store = Artifacts(self.root)
            self.store.recover()
            self.state['phase'] = 'error'
            self.record('workflow.error', reason=str(exc))
            raise

    def feedback(self, build, message, accepted=None):
        if accepted is not None and not isinstance(accepted, bool):
            raise ValueError('accepted must be a boolean')
        preview = self.state.get('preview')
        if not preview or preview['build_id'] != build:
            raise ValueError('Feedback target must match the current preview build')
        if not message.strip():
            raise ValueError('Feedback must include a message')
        if accepted is not None:
            self.state['review'] = dict(build_id=build, accepted=accepted, message=message)
        self.record('visual.review' if accepted is not None else 'human_feedback',
                    target_build=build, generation=preview['generation'], message=message,
                    accepted=accepted, reason='Human reviewed exact preview generation')

    def publish(self, build):
        preview = self.state.get('preview', {})
        review = self.state.get('review') or {}
        if preview.get('build_id') != build or review.get('build_id') != build or not review.get('accepted'):
            raise ValueError('Publication requires acceptance of this exact preview build')
        if self.state['handles']['resume.layout.workspace']['sha256'] != self.state['rendered']['workspace_sha256']:
            raise ValueError('Workspace changed after preview; create and review a new checkpoint')
        for source, target in [('resume.layout.preview', 'resume.layout.published'),
                               ('resume.render.preview', 'resume.render.published')]:
            self.store.resolve(source)
            self.store.register(target, self.state['handles'][source])
        self.state['phase'] = 'published'
        self.record('artifact.publish', build_id=build, provenance=preview['provenance'],
                    reason='Human-accepted local build published')
