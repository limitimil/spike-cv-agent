# Resume pipeline: instructions for every AI agent

## Decide whether you are operating or developing

The default is **operator mode**. A request to run, resume, inspect, re-layout,
render, review, or export a resume authorizes use of the existing pipeline. It
does **not** authorize implementing or changing the tool. In operator mode:

- Read README.md and `docs/OPERATIONS.md` before running commands.
- Use `./resume.ps1` on Windows. Production rendering supports Windows only;
  macOS/Linux rendering must fail explicitly, with no fallback.
- Run `status`, `doctor`, and `operations` first. Resolve artifacts through
  manifest handles; never guess a filename, choose the newest file, or use a
  reference PDF as if this pipeline rendered it.
- Use only supported deterministic operations. Always give a concrete design
  reason. Inspect the source report before assigning semantic paragraph styles.
- Do not edit `src/`, tests, launchers, dependencies, documentation, design files,
  manifests, audit logs, checkpoint files, or working DOCX files directly as a
  workaround. Submit a complete design through `operate layout.apply`; record
  the reason. Updating tool code requires explicit implementation authorization.
  Preparing a separate proposed-design JSON file for that operation is allowed;
  do not edit the active manifest or initial design file to apply changes.
- Never call python-docx, Word automation, or any renderer ad hoc to bypass the operation registry.
  Never fabricate a PDF, review, acceptance, provenance, or audit event.
- Use `checkpoint` to expose a completed immutable DOCX/PDF pair. Preserve
  existing sources and artifacts. Respect the separate control channel; do not
  clear pause/abort unless the human asks to continue.
- Attribute feedback to the exact build the human inspected. Do not infer human
  acceptance from successful tests or a request to improve layout. `review
  BUILD --accept` requires the human's explicit acceptance of that exact build.
  `publish BUILD` also requires a request to publish or standing authorization.
- Report actual results and blockers. Missing Microsoft Word means rendering is
  blocked; do not install software, substitute a renderer, or loosen validation
  silently. Explain the required environment change and obtain authorization
  when needed.

## If an operation does not cover the scenario

Stop the affected work and propose a plan using `docs/PLAN_TEMPLATE.md`.
Explain the requested outcome, the exact missing capability, existing operation
options, proposed stable operation name, inputs/outputs, validation, audit
semantics, and tests. Mention any effect on existing artifacts or dependencies.

**The human decides each plan. Wait for explicit approval before implementing
the plan, changing the tool, or executing an unsupported workaround.** A proposal
is not approval. Silence, a timeout, previous approval of a different plan, or a
general request to run the pipeline is not approval. If scope changes materially,
present the revised plan for a new decision. Continue independent, already
authorized work where possible. Do not repeatedly ask for an already approved
action within the same plan.

## Developer mode

Enter only when the human explicitly asks to implement/change the tool or
approves a concrete implementation plan. Keep changes within that authorization.
Preserve the registry contracts, immutable checkpoints, single-writer lock,
durable cursor, audit recovery, and build-specific human acceptance. Add focused
tests for behavioral changes and run `python -m unittest discover -s tests -v`.
This stage is Windows-first: python-docx plus Microsoft Word Automation. Every
caller uses `DocxToPdfRenderer` through `RendererFactory`. Keep Word-specific APIs
inside `src/rendering/`. Do not introduce LibreOffice, macOS/Linux renderers,
Docker, cloud services, or automatic rewrite models without a new approved plan.
Document any new operation in `docs/OPERATIONS.md`.

## Communication preference

The user requests `/say` progress updates for each implementation step and when
human input is needed. Use the installed say skill when available. Speech through
the configured say host is authorized **without remote saving**: always pass
`--no-save`. Follow the skill's failure handling and report unavailable audio;
continue authorized implementation work. Never synthesize resume content unless
separately requested. Text updates must still state the result and blockers.
