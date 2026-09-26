# Resume layout workspace

Windows-first, resumable resume layout using Python and Microsoft Word. No Docker.
**AI agents: read [AGENTS.md](AGENTS.md) before operating this repository.**

## Open the product

Run `./resume.ps1` from this folder. Open **http://localhost:7331**.
The Python entry point is `.venv\Scripts\python resume.py`.
Production rendering currently supports Windows only; macOS/Linux fail explicitly.

The monitor shows phase, viewed build, recent audit events, and design changes.
Use **Run / resume**, **Pause**, **Stop at checkpoint**, and **Abort** to manage
execution. A newer preview loads only when you choose it, keeping feedback tied
to the exact document you inspected. Closing the monitor does not kill a worker;
use the control channel first if you want it to stop.

## One-time setup

Install Python 3.10+ and Microsoft Word desktop on Windows, then:

```powershell
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements.txt
.\resume.ps1 doctor
.\resume.ps1 init --source source/Chien_Li_Ming_2026_Reformatted.docx
.\resume.ps1
```

Run Word interactively once to complete activation and first-run setup. `doctor`
starts and closes a dedicated hidden Word instance and reports the runtime version.
It never opens the source resume. Initialize once;
if a manifest already exists, use `status` and resume it instead.

## Daily commands

| Command | Purpose |
| --- | --- |
| `doctor` | Check Python, initialization, and Word startup/version/shutdown |
| `status` | Read durable state and current preview identity |
| `operations` | List deterministic operation contracts |
| `run` / `resume` | Continue from the persisted operation cursor |
| `control pause` | Finish atomic operation, checkpoint, exit |
| `control stop_after_checkpoint` | Checkpoint and exit as stopped |
| `control abort` | Exit at the next operation boundary, preserve last preview |
| `control continue` | Permit a subsequent `resume` command |
| `checkpoint` | Snapshot, render, and expose an immutable preview |
| `feedback BUILD "message"` | Record feedback on current preview |
| `review BUILD "message" --accept` | Record explicit human acceptance |
| `publish BUILD` | Publish an accepted, current, unmodified build |

Prefix each command with `./resume.ps1`. `--help` lists commands.
`--root DIRECTORY` before the subcommand
selects an independent workspace.

CLI resume respects pause/abort; use `control continue` only when the human asks
to continue. The monitor's Run/resume button explicitly performs both actions.
Pause and stop require a successful PDF render to expose a complete checkpoint.
Missing Word leaves completed document operations saved and reports an
error; install/activate Word, then resume. Each Word worker times out after 120 seconds.
Abort waits for the current atomic operation, including any render in progress.

## Design changes and agent boundaries

The initial design is imported from `design.json`. Later edits do not silently
affect a run. Submit a complete proposed design file explicitly:

```powershell
.\resume.ps1 operate layout.apply --design proposed-design.json --reason "Increase role spacing for readability"
.\resume.ps1 checkpoint
```

See [docs/OPERATIONS.md](docs/OPERATIONS.md) for contracts and limits. Use the
inspection report's paragraph indices for semantic style mapping; no section
meaning is guessed. Unmapped paragraphs retain their styles, while font and page
settings apply globally. Tables and headers/footers are included; drawing/textbox
text is outside the initial implementation.

An agent asked to run the pipeline must not change tool code or directly mutate
documents as a workaround. Unsupported scenarios require a
[proposed plan](docs/PLAN_TEMPLATE.md) and the human's explicit decision before
implementation. Human acceptance and publication are never inferred from tests.

## Canonical Windows rendering

Microsoft Word Automation is the only production renderer. The deterministic
operation is `render.docx_to_pdf` (replacing the earlier `render.pdf` name).
`RendererFactory.create()` returns a `DocxToPdfRenderer`; callers use only
`validate_environment()`, `metadata()`, and `render(input_docx, output_pdf)`.
The Windows adapter owns hidden Word startup, native PDF export, and cleanup.
Other platforms fail with `UNSUPPORTED_PLATFORM`; there is no fallback.

Each render records the renderer type, platform, Word version, and Word build.
The DOCX/PDF pair shares one build ID. A changed renderer identity creates a new
build. Word opens the checkpoint read-only, disables document macros and link
updates, exports to a private staging file, then closes without saving.

An isolated worker bounds calls that may block. Cleanup closes the document and
quits its Word instance; if necessary, the parent terminates only the recorded
owned process after verifying its PID and creation time. It never kills all Word
processes. If activation hangs before the owned PID can be identified, targeted
cleanup cannot be guaranteed; inspect Word manually before retrying. Use this
from your signed-in Windows desktop session, not an unattended server service.

Failure codes are `RENDERER_UNAVAILABLE`, `RENDERER_INIT_FAILED`,
`DOCUMENT_OPEN_FAILED`, and `RENDER_FAILED`. A failed `render.docx_to_pdf` audit
event records the code and debug checkpoint. The previous preview and published
handles remain unchanged. `resume.layout.checkpoint` identifies the retained DOCX;
`status` exposes `render_error`. No failed PDF becomes a preview.

Word versions and fonts can affect pagination: inspect all pages and accept the
exact build before publication. Install the design font (default: Liberation
Sans), or explicitly choose an installed font. Future renderers require a new
approved plan and implement the same adapter interface; artifact handling,
auditing, and preview publication do not need renderer-specific changes.

## Durability and access

`manifest.json` is authoritative. Its `handles` map names to exact paths, build
IDs, checksums, and generations. `resume.layout.workspace` points to a committed
revision; `workspace/resume-working.docx` is an agent-only disposable projection.
Direct edits to the projection never enter a checkpoint. Source and completed
artifacts are never overwritten by pipeline code.

`resume.layout.preview` and `resume.render.preview` identify an immutable pair;
the corresponding `.published` handles identify a human-accepted pair. Audit
records describe intent, not shell commands. A transaction journal recovers
interruptions between audit append and atomic manifest replacement, deduplicating
event IDs. Torn audit tails remain as evidence and are skipped by readers. The OS
writer lock releases on process death. Orphan files are never chosen by filename.

Status and current-PDF files are separately atomic projections. The monitor reads
authoritative state and loads an immutable URL to avoid mismatched generations.
The separate control file contains desired state; its application is audited at
an operation boundary. A newer request may supersede an unapplied older request.
Use a local disk: recovery covers process interruption, not hardware failure.
Immutability is enforced by code and checksums, not against a filesystem owner.

The monitor binds to loopback and exposes checkpoint documents, status, and audit
events, never source or working files. Browser mutations require a session token
and matching host/origin. Human review and publication remain explicit CLI actions.

## Verify

```powershell
.venv\Scripts\python -m unittest discover -s tests -v
```

Tests cover recovery, controls, artifacts, review/publication, rendering identity,
and HTTP boundaries. Test PDFs use an explicit stub renderer; real Word
output still needs visual inspection. Modules under `src/` implement the registry,
artifacts, audit, rendering, controls, orchestration, and monitor.

References: [python-docx styles](https://python-docx.readthedocs.io/en/stable/user/styles-using.html)
and [Word PDF export](https://learn.microsoft.com/en-us/office/vba/api/word.document.exportasfixedformat).
