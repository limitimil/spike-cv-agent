"""Windows Word adapter. The separate worker owns all Office calls."""
import importlib.util
import json
import platform
import os
import re
import subprocess
import sys
import tempfile
from pathlib import Path

from .base import DocxToPdfRenderer, RendererError


class WordAutomationRenderer(DocxToPdfRenderer):
    def __init__(self, timeout_seconds=120):
        self.timeout_seconds = timeout_seconds
        self._metadata = {'type': 'word_automation', 'platform': 'windows', 'word_version': None}
        self._validated = False
        self.diagnostics = ''  # Local troubleshooting only; never copied into artifact metadata.

    def metadata(self):
        return dict(self._metadata)

    def validate_environment(self):
        if self._validated:
            return
        if platform.system() != 'Windows':
            raise RendererError('UNSUPPORTED_PLATFORM', 'Microsoft Word Automation requires Windows.')
        if importlib.util.find_spec('win32com') is None:
            raise RendererError('RENDERER_UNAVAILABLE', 'Install the pinned Windows dependencies from requirements.txt.')
        import winreg
        try:
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, r'Word.Application\CLSID') as key:
                clsid = winreg.QueryValue(key, None)
            with winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, 'CLSID\\' + clsid + r'\LocalServer32') as key:
                command = winreg.QueryValue(key, None).strip()
            match = re.match(r'^"([^"]+)"|^(.+?\.exe)(?:\s|$)', command, re.IGNORECASE)
            executable = os.path.expandvars(next(group for group in match.groups() if group)) if match else ''
            if not executable or not Path(executable).is_file():
                raise FileNotFoundError('Registered Word executable is missing')
        except OSError as exc:
            raise RendererError('RENDERER_UNAVAILABLE',
                                'Microsoft Word desktop is not installed or its automation registration is missing.') from exc
        self._metadata = self._invoke('probe')
        self._validated = True

    def _invoke(self, mode, source=None, output=None):
        with tempfile.TemporaryDirectory(prefix='resume-word-') as directory:
            state = Path(directory) / 'owned-process.json'
            args = [sys.executable, '-m', 'src.rendering.word_worker', mode, str(state)]
            if source is not None:
                args.extend([str(source), str(output)])
            process = subprocess.Popen(args, cwd=str(Path(__file__).resolve().parents[2]),
                                       stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                                       creationflags=getattr(subprocess, 'CREATE_NO_WINDOW', 0))
            cleanup_error = None
            try:
                stdout, stderr = process.communicate(timeout=self.timeout_seconds)
                self.diagnostics = stderr.decode('utf-8', errors='replace')
            except subprocess.TimeoutExpired as exc:
                process.kill()
                process.communicate()
                raise RendererError('RENDERER_INIT_FAILED' if mode == 'probe' else 'RENDER_FAILED',
                                    'Word exceeded the ' + str(self.timeout_seconds) + '-second timeout.') from exc
            finally:
                if process.poll() is None:
                    process.kill()
                    process.communicate()
                cleanup_error = self._cleanup_owned_process(state)
            try:
                result = json.loads(stdout.decode('utf-8'))
            except (ValueError, UnicodeDecodeError) as exc:
                raise RendererError('RENDERER_INIT_FAILED' if mode == 'probe' else 'RENDER_FAILED',
                                    'Word worker exited without a valid result.') from exc
            if not result.get('ok'):
                raise RendererError(result.get('code', 'RENDER_FAILED'), result.get('message', 'Word rendering failed.'))
            if process.returncode or cleanup_error:
                raise RendererError('RENDER_FAILED', cleanup_error or 'Word worker failed.')
            return result['metadata']

    @staticmethod
    def _cleanup_owned_process(state):
        """Never terminate preexisting processes or a reused PID."""
        if not state.exists():
            return None  # Activation may fail before an owned PID can be identified.
        try:
            record = json.loads(state.read_text(encoding='utf-8'))
            if not record.get('owned'):
                return None
            import win32api
            import win32event
            import win32process
            try:
                handle = win32api.OpenProcess(0x100000 | 0x1000 | 0x0001, False, record['pid'])
            except Exception as exc:
                if getattr(exc, 'winerror', None) == 87:  # Process already exited.
                    return None
                raise
            try:
                created = win32process.GetProcessTimes(handle)['CreationTime'].isoformat()
                if created != record['created']:
                    return None
                if win32event.WaitForSingleObject(handle, 3000) == 258:
                    win32api.TerminateProcess(handle, 1)
                    if win32event.WaitForSingleObject(handle, 3000) == 258:
                        return 'The owned Word process did not exit after cleanup.'
            finally:
                handle.Close()
        except Exception:
            return 'Could not verify cleanup of the owned Word process; inspect Word before retrying.'
        return None

    def render(self, input_docx, output_pdf):
        source, output = Path(input_docx).resolve(), Path(output_pdf).resolve()
        if not source.is_file() or source.suffix.lower() != '.docx':
            raise RendererError('DOCUMENT_OPEN_FAILED', 'Input must be an existing DOCX file.')
        if output.exists():
            raise RendererError('RENDER_FAILED', 'Output already exists; immutable PDFs cannot be overwritten.')
        self.validate_environment()
        output.parent.mkdir(parents=True, exist_ok=True)
        with tempfile.TemporaryDirectory(prefix='resume-pdf-', dir=str(output.parent)) as directory:
            temporary = Path(directory) / 'export.pdf'
            metadata = self._invoke('render', source, temporary)
            if not temporary.exists() or not temporary.read_bytes().startswith(b'%PDF-'):
                raise RendererError('RENDER_FAILED', 'Word did not produce a valid PDF.')
            # Exclusive creation keeps an existing output safe even if it appeared during export.
            created = False
            try:
                with output.open('xb') as stream:
                    created = True
                    stream.write(temporary.read_bytes())
                    stream.flush()
                    import os
                    os.fsync(stream.fileno())
            except OSError as exc:
                if created:
                    output.unlink(missing_ok=True)
                raise RendererError('RENDER_FAILED', 'Could not save the rendered PDF.') from exc
            self._metadata = metadata
