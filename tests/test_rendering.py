from datetime import datetime, timezone
import json
from pathlib import Path
import subprocess
import tempfile
from types import ModuleType
import unittest
from unittest.mock import MagicMock, patch

from src.rendering import RendererFactory, RendererError
from src.rendering.word_windows import WordAutomationRenderer
from src.rendering.word_worker import automate


def fake_modules():
    modules = {name: ModuleType(name) for name in
               ('pythoncom', 'win32api', 'win32process', 'win32event', 'win32com', 'win32com.client')}
    modules['win32com'].client = modules['win32com.client']
    modules['pythoncom'].COINIT_APARTMENTTHREADED = 2
    modules['pythoncom'].CoInitializeEx = MagicMock()
    modules['pythoncom'].CoUninitialize = MagicMock()
    modules['win32com.client'].DispatchEx = MagicMock()
    modules['win32api'].OpenProcess = MagicMock()
    modules['win32api'].TerminateProcess = MagicMock()
    modules['win32process'].EnumProcesses = MagicMock(return_value=[100])
    modules['win32process'].GetWindowThreadProcessId = MagicMock(return_value=(1, 200))
    modules['win32process'].GetProcessTimes = MagicMock(return_value={
        'CreationTime': datetime(2026, 1, 1, tzinfo=timezone.utc)})
    modules['win32event'].WaitForSingleObject = MagicMock(return_value=0)
    return modules


class RendererTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.root = Path(self.temp.name)

    def test_factory_supports_only_windows(self):
        with patch('src.rendering.factory.platform.system', return_value='Windows'):
            self.assertIsInstance(RendererFactory.create(), WordAutomationRenderer)
        for system in ('Linux', 'Darwin', 'Other'):
            with self.subTest(system=system), patch('src.rendering.factory.platform.system', return_value=system):
                with self.assertRaises(RendererError) as error:
                    RendererFactory.create()
                self.assertEqual('UNSUPPORTED_PLATFORM', error.exception.code)

    def test_missing_dependency_has_explicit_code(self):
        with patch('src.rendering.word_windows.platform.system', return_value='Windows'), \
             patch('src.rendering.word_windows.importlib.util.find_spec', return_value=None):
            with self.assertRaises(RendererError) as error:
                WordAutomationRenderer().validate_environment()
            self.assertEqual('RENDERER_UNAVAILABLE', error.exception.code)

    def test_missing_word_registration(self):
        registry = ModuleType('winreg')
        registry.HKEY_CLASSES_ROOT = 1
        registry.OpenKey = MagicMock(side_effect=FileNotFoundError())
        with patch.dict('sys.modules', {'winreg': registry}), \
             patch('src.rendering.word_windows.platform.system', return_value='Windows'), \
             patch('src.rendering.word_windows.importlib.util.find_spec', return_value=object()):
            with self.assertRaises(RendererError) as error:
                WordAutomationRenderer().validate_environment()
            self.assertEqual('RENDERER_UNAVAILABLE', error.exception.code)

    def test_timeout_kills_worker_and_attempts_owned_cleanup(self):
        process = MagicMock()
        process.communicate.side_effect = [subprocess.TimeoutExpired('worker', 1), (b'', b'')]
        process.poll.return_value = 1
        renderer = WordAutomationRenderer(timeout_seconds=1)
        with patch('src.rendering.word_windows.subprocess.Popen', return_value=process), \
             patch.object(renderer, '_cleanup_owned_process') as cleanup:
            with self.assertRaises(RendererError) as error:
                renderer._invoke('render', self.root / 'in.docx', self.root / 'out.pdf')
            self.assertEqual('RENDER_FAILED', error.exception.code)
            process.kill.assert_called_once()
            cleanup.assert_called_once()

    def test_adapter_preserves_existing_output(self):
        source, output = self.root / 'in.docx', self.root / 'out.pdf'
        source.write_bytes(b'docx')
        output.write_bytes(b'existing')
        with self.assertRaises(RendererError):
            WordAutomationRenderer().render(source, output)
        self.assertEqual(b'existing', output.read_bytes())

    def test_adapter_rejects_invalid_pdf_without_publishing(self):
        source, output = self.root / 'in.docx', self.root / 'out.pdf'
        source.write_bytes(b'docx')
        renderer = WordAutomationRenderer()
        def invoke(mode, source, temporary):
            temporary.write_bytes(b'not a PDF')
            return renderer.metadata()
        with patch.object(renderer, 'validate_environment'), patch.object(renderer, '_invoke', side_effect=invoke):
            with self.assertRaises(RendererError):
                renderer.render(source, output)
        self.assertFalse(output.exists())

    def test_owned_cleanup_does_not_kill_reused_pid(self):
        modules = fake_modules()
        state = self.root / 'owned.json'
        state.write_text(json.dumps({'owned': True, 'pid': 200, 'created': 'different creation time'}))
        with patch.dict('sys.modules', modules):
            self.assertIsNone(WordAutomationRenderer._cleanup_owned_process(state))
        modules['win32api'].TerminateProcess.assert_not_called()

    def test_owned_cleanup_terminates_only_tracked_instance(self):
        modules = fake_modules()
        modules['win32event'].WaitForSingleObject.side_effect = [258, 0]
        state = self.root / 'owned.json'
        state.write_text(json.dumps({'owned': True, 'pid': 200, 'created':
            datetime(2026, 1, 1, tzinfo=timezone.utc).isoformat()}))
        with patch.dict('sys.modules', modules):
            self.assertIsNone(WordAutomationRenderer._cleanup_owned_process(state))
        modules['win32api'].TerminateProcess.assert_called_once_with(
            modules['win32api'].OpenProcess.return_value, 1)

    def test_word_lifecycle_and_native_export(self):
        modules = fake_modules()
        app = modules['win32com.client'].DispatchEx.return_value
        app.Version, app.Build = '16.0', '12345'
        app.AutomationSecurity = 1
        app.Options.UpdateLinksAtOpen = True
        app.Options.SaveNormalPrompt = True
        with patch.dict('sys.modules', modules):
            result = automate('render', self.root / 'owned.json', self.root / 'in.docx', self.root / 'out.pdf')
        self.assertTrue(result['ok'])
        self.assertEqual('16.0', result['metadata']['word_version'])
        self.assertFalse(app.Visible)
        self.assertEqual(1, app.AutomationSecurity)  # User settings restored before quit.
        self.assertTrue(app.Options.UpdateLinksAtOpen)
        self.assertTrue(app.Options.SaveNormalPrompt)
        self.assertTrue(app.Documents.Open.call_args.kwargs['ReadOnly'])
        document = app.Documents.Open.return_value
        self.assertEqual(17, document.ExportAsFixedFormat.call_args.kwargs['ExportFormat'])
        document.Close.assert_called_once_with(SaveChanges=0)
        app.Quit.assert_called_once_with(SaveChanges=0)
        modules['pythoncom'].CoUninitialize.assert_called_once()
        app.Documents.Add.assert_called_once_with(Visible=False)
        app.Documents.Add.return_value.Close.assert_called_once_with(SaveChanges=0)

    def test_word_failures_close_owned_app(self):
        for failure in ('init', 'open', 'export'):
            with self.subTest(failure=failure):
                modules = fake_modules()
                app = modules['win32com.client'].DispatchEx.return_value
                if failure == 'init':
                    modules['win32com.client'].DispatchEx.side_effect = RuntimeError('activation')
                    expected = 'RENDERER_INIT_FAILED'
                elif failure == 'open':
                    app.Documents.Open.side_effect = RuntimeError('open')
                    expected = 'DOCUMENT_OPEN_FAILED'
                else:
                    app.Documents.Open.return_value.ExportAsFixedFormat.side_effect = RuntimeError('export')
                    expected = 'RENDER_FAILED'
                with patch.dict('sys.modules', modules):
                    result = automate('render', self.root / 'owned.json', self.root / 'in.docx', self.root / 'out.pdf')
                self.assertEqual(expected, result['code'])
                if failure != 'init':
                    app.Quit.assert_called_once()
                modules['pythoncom'].CoUninitialize.assert_called_once()

    def test_preexisting_word_instance_is_not_modified_or_quit(self):
        modules = fake_modules()
        modules['win32process'].GetWindowThreadProcessId.return_value = (1, 100)
        app = modules['win32com.client'].DispatchEx.return_value
        with patch.dict('sys.modules', modules):
            result = automate('render', self.root / 'owned.json', self.root / 'in.docx', self.root / 'out.pdf')
        self.assertEqual('RENDERER_INIT_FAILED', result['code'])
        app.Documents.Open.assert_not_called()
        app.Quit.assert_not_called()
        self.assertFalse((self.root / 'owned.json').exists())
        app.Documents.Add.return_value.Close.assert_called_once_with(SaveChanges=0)
