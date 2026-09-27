"""Private subprocess entry point. Never called by workflow/CLI directly."""
import json
import sys
from pathlib import Path

from ..storage import write_json


def automate(mode, state_path, source=None, output=None):
    import pythoncom
    import win32api
    import win32com.client
    import win32process

    app = document = probe_document = None
    initialized = False
    owned = False
    result = None
    preferences = None
    stage = 'RENDERER_INIT_FAILED'
    try:
        pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
        initialized = True
        previous_pids = set(win32process.EnumProcesses())
        app = win32com.client.DispatchEx('Word.Application')
        # Word exposes HWND on Window, not Application. A private, invisible blank
        # document gives us the exact process without inspecting user windows.
        probe_document = app.Documents.Add(Visible=False)
        _, pid = win32process.GetWindowThreadProcessId(probe_document.ActiveWindow.Hwnd)
        owned = pid not in previous_pids
        if not owned:
            raise RuntimeError('Word did not create a dedicated process.')
        handle = win32api.OpenProcess(0x1000, False, pid)
        try:
            created = win32process.GetProcessTimes(handle)['CreationTime'].isoformat()
        finally:
            handle.Close()
        write_json(state_path, {'pid': pid, 'created': created, 'owned': True})
        probe_document.Close(SaveChanges=0)
        probe_document = None
        preferences = (app.DisplayAlerts, app.AutomationSecurity,
                       app.Options.SaveNormalPrompt, app.Options.UpdateLinksAtOpen)
        app.Visible = False
        app.DisplayAlerts = 0
        app.AutomationSecurity = 3  # msoAutomationSecurityForceDisable
        app.Options.SaveNormalPrompt = False
        app.Options.UpdateLinksAtOpen = False
        metadata = {'type': 'word_automation', 'platform': 'windows',
                    'word_version': str(app.Version), 'word_build': str(app.Build)}
        if mode == 'render':
            stage = 'DOCUMENT_OPEN_FAILED'
            document = app.Documents.Open(FileName=str(source), ConfirmConversions=False,
                ReadOnly=True, AddToRecentFiles=False, Visible=False,
                PasswordDocument='', WritePasswordDocument='', NoEncodingDialog=True,
                OpenAndRepair=False)
            stage = 'RENDER_FAILED'
            document.ExportAsFixedFormat(OutputFileName=str(output), ExportFormat=17,
                OpenAfterExport=False, OptimizeFor=0, Range=0, Item=0,
                IncludeDocProps=True, CreateBookmarks=1, DocStructureTags=True,
                BitmapMissingFonts=True, UseISO19005_1=False)
        result = {'ok': True, 'metadata': metadata}
    except Exception:
        import traceback
        traceback.print_exc(file=sys.stderr)
        messages = {'RENDERER_INIT_FAILED': 'Microsoft Word automation could not start a dedicated hidden instance.',
                    'DOCUMENT_OPEN_FAILED': 'Microsoft Word could not open the DOCX checkpoint.',
                    'RENDER_FAILED': 'Microsoft Word could not export the document to PDF.'}
        result = {'ok': False, 'code': stage, 'message': messages[stage]}
    finally:
        cleanup_failed = False
        if probe_document is not None:
            try:
                probe_document.Close(SaveChanges=0)
            except Exception:
                cleanup_failed = True
            probe_document = None
        if document is not None:
            try:
                document.Close(SaveChanges=0)
            except Exception:
                cleanup_failed = True
            document = None
        if app is not None and owned:
            if preferences is not None:
                try:
                    (app.DisplayAlerts, app.AutomationSecurity,
                     app.Options.SaveNormalPrompt, app.Options.UpdateLinksAtOpen) = preferences
                except Exception:
                    cleanup_failed = True
            try:
                app.Quit(SaveChanges=0)
            except Exception:
                cleanup_failed = True
        app = None
        if initialized:
            pythoncom.CoUninitialize()
        if cleanup_failed and result and result['ok']:
            result = {'ok': False, 'code': 'RENDER_FAILED', 'message': 'Word could not close cleanly after rendering.'}
    return result


def main():
    mode, state = sys.argv[1:3]
    paths = [Path(value).resolve() for value in sys.argv[3:]]
    try:
        result = automate(mode, Path(state), *paths)
    except Exception:
        result = {'ok': False, 'code': 'RENDERER_INIT_FAILED', 'message': 'The Word worker could not initialize.'}
    print(json.dumps(result), flush=True)
    return 0 if result['ok'] else 1


if __name__ == '__main__':
    sys.exit(main())
