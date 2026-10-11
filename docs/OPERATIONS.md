# Deterministic operation contracts

Start with `status`, `doctor`, and `operations`. Read `AGENTS.md` before acting.
Use the manifest's exact `handles` entries, never filename ordering.

Before visual layout work, read [the resume design guide](RESUME_DESIGN_GUIDE.md)
for human feedback, hierarchy, color, alignment, and visual review criteria.
The guide describes design intent; only the operation parameters below are
implemented capabilities. Missing visual controls require a proposed extension.

| Operation | Inputs | Outputs | Parameters / behavior |
| --- | --- | --- | --- |
| `source.inspect` | `resume.source` | `resume.source.inspection` | No parameters; JSON report of paragraph indices, styles, and sections. |
| `content.normalize` | Workspace | New workspace revision | No parameters; replace nonbreaking spaces, preserving emphasis. |
| `content.revise` | Workspace; current workspace inspection | New workspace revision; `resume.content.revision` | Explicit exact text replacements with checksum guards and a durable editorial change report. |
| `layout.apply` | Workspace | New workspace revision | Complete `design` object; font, margins, spacing, pagination, explicit style mapping. |
| `layout.inspect` | Workspace | `resume.layout.inspection` | No parameters; versioned structure report tied to the workspace checksum. DOCX, cursor, preview, and acceptance remain unchanged. |
| `layout.compose` | Workspace; current workspace inspection | New workspace revision | `specification`; explicit block order, table linearization, semantic theme, typography, visual paragraph rules, phrase highlights, compaction, and empty-paragraph removal. |
| `render.docx_to_pdf` | Workspace | `resume.layout.checkpoint`, `resume.render.candidate` | No parameters; immutable snapshot, native Word PDF, runtime provenance. Windows only. |
| `preview.publish` | Workspace / candidate | DOCX/PDF preview handles | No parameters; ensure matching candidate and expose a generation. |
| `visual.review` | Current preview | Audit review record | `build`, `message`, boolean `accepted`; human decision only. |
| `artifact.publish` | Accepted current preview | Published handles | `build`; requires acceptance and unchanged workspace. |

Workspace means `resume.layout.workspace`. Preview handles are
`resume.layout.preview` and `resume.render.preview`; published handles use the
`.published` suffix. The default pipeline is inspect → normalize → layout.apply →
checkpoint → await human review. Cursor advancement is committed with each step.
Resume skips committed edits. Checkpoint composes rendering and preview publication
and is idempotent for the same revision and renderer identity.

```powershell
.\resume.ps1 status
.\resume.ps1 doctor
.\resume.ps1 operations
.\resume.ps1 operate source.inspect --reason "Inspect paragraph indices before assigning styles"
.\resume.ps1 operate layout.apply --design proposed-design.json --reason "Separate role headings"
.\resume.ps1 checkpoint
.\resume.ps1 feedback EXACT_BUILD "Keep Experience unchanged"
# Only after human acceptance of this exact build:
.\resume.ps1 review EXACT_BUILD "Human accepted all pages" --accept
# Only when publication is requested or already authorized:
.\resume.ps1 publish EXACT_BUILD
```

Other operations accept `--params` as a JSON object or `@PATH` to a UTF-8 JSON
parameter file. Dedicated `feedback`,
`review`, and `publish` commands avoid shell JSON quoting. Every `operate` call
requires a nonempty reason. Rendering on macOS/Linux is explicitly unsupported.

Use `operate render.docx_to_pdf --reason "Render layout checkpoint"` to create a
candidate, then `operate preview.publish --reason "Expose checkpoint for review"`.
The `checkpoint` command composes both. The old `render.pdf` command is removed;
historical audit entries retain their original names. Existing manifests resume
without migration. A previous renderer's cached output is not reused for Word.

Rendering errors include `RENDERER_UNAVAILABLE`, `RENDERER_INIT_FAILED`,
`DOCUMENT_OPEN_FAILED`, `RENDER_FAILED`, and `UNSUPPORTED_PLATFORM`. Failures are
audited by `render.docx_to_pdf`, leave the prior preview/published handles intact,
and retain `resume.layout.checkpoint` for debugging. Use `doctor` to check Word
registration, startup, runtime version, and shutdown; never bypass a failure with
an alternative renderer or a hand-edited manifest.

| Design key | Validation / meaning |
| --- | --- |
| `body.font` | Nonempty installed family name |
| `body.size_pt` | 6–30 points |
| `section.space_before_pt` | 0–72 points on mapped Section paragraphs |
| `experience.role_spacing_pt` | 0–72 points on mapped Role paragraphs |
| `page.margin_inches` | 0.2–2 inches for all margins |
| `pagination.widow_control` | Boolean |
| `paragraph_styles` | Optional index-string map to Body, Section, Company, Role, Meta, Bullet |

Supply all numeric, font, and pagination keys. Unknown keys, invalid ranges, and
invalid paragraph indices are rejected before committing a revision. Font
installation is an environment responsibility, not confirmed by python-docx.

## Workspace inspection and composition

`layout.inspect` reports actual paragraph/table block order, nested cell containers,
paragraph runs and hyperlinks, numbering/formatting properties, explicit breaks,
and unsupported structures. Its report has `schema_version: 1`, `workspace_sha256`,
`blocks`, `containers`, and `sections`. IDs such as `body/b0` and
`body/b0/r0/c0/b0` describe that inspected revision; they do not assign semantic
roles and must not be reused after another edit. The source report and its flat
paragraph indices remain unchanged. Inspection creates an immutable JSON artifact
and an audit record, while preserving review, preview/published handles, phase,
workspace bytes, and the default cursor.

Complete the existing pipeline layout step before composing: `next_step` must be
at least 3. Both new operations are manual and never advance the default cursor.
Prepare a separate parameter file containing `{"specification": {...}}`, then use:

```powershell
.\resume.ps1 operate layout.inspect --reason "Inspect current blocks before composing"
# Read resume.layout.inspection through its exact manifest handle;
# fill the proposed file with that inspection's IDs and both current checksums.
.\resume.ps1 operate layout.compose --params '@proposed-composition.json' --reason "Put backend skills first and keep project details together"
.\resume.ps1 operate layout.inspect --reason "Verify the composed revision and obtain its new IDs"
.\resume.ps1 checkpoint
```

The required specification keys are integer `schema_version: 1`,
`expected_workspace_sha256`, and `expected_inspection_sha256`. Checksums are
64 lowercase hexadecimal characters. The CLI holds the existing writer lock;
composition checks both current handles, their artifact integrity, and that the
inspection describes the current workspace. Missing or stale inspection requires
reinspection. Unknown keys, invalid values, unsupported affected structures, and
preservation failures reject the operation before registering a revision.

All remaining specification keys are optional. Lists default to `[]` and
`typography` defaults to `{}`; omitted properties retain their current values.

| Specification key | Supported rule |
| --- | --- |
| `orders` | `[{"container": ID, "blocks": [ID, ...]}]`; one complete permutation of each selected body/cell container's direct blocks. Every required block occurs exactly once. No arbitrary cross-container moves. |
| `linearize_tables` | `[{"table": ID, "paragraph_order": [ID, ...]}]`; select a distinct body-level layout table and include every leaf paragraph exactly once, including empties and nested-table paragraphs. See restrictions below. |
| `typography` | Optional `font` (nonempty), `body_size_pt` (6-30), `margin_inches` (0.2-2), and `color` (`#RRGGBB`). Font/size/color apply to all supported paragraph runs, including hyperlink runs; margins apply to all sections. Color changes are explicit. Installed fonts are an environment responsibility. |
| `theme` | Complete semantic role map for `ink`, `teal`, `accent`, `surface`, `highlight`, `text`, `muted`, and `paper`; every value is `#RRGGBB`. Unknown or missing roles and unsafe required contrast pairs are rejected. Omit the whole map for legacy specifications without role references. |
| `paragraph_rules` | `[{"block": PARAGRAPH_ID, ...}]`; at most one rule per paragraph. Existing size, weight, spacing, indentation, and pagination fields remain. Optional `alignment` is `left`, `center`, `right`, or `justify`; `color_role` and `shading_role` reference `theme`; `borders` declares left/right/top/bottom role, `single`/`double`/`dotted`/`dashed` style, width, and spacing; `tab_stops` declares increasing positions in inches with left/center/right/decimal alignment and an optional leader. |
| `highlights` | `[{"block": PARAGRAPH_ID, "text": "API", "expected_matches": 1, "color_role": "teal", "shading_role": "highlight", "bold": true}]`; exact, case-sensitive phrase matching with a positive integer match count. `bold` defaults to `true` for backward compatibility. Role styling applies only to matched text across text runs/hyperlinks while retaining link context and unrelated run properties. A missing or different count fails; control/field runs are unsupported. |
| `compact` | `[{"blocks": [PARAGRAPH_ID, ...], "separator": "\t"}]`; at least two distinct, adjacent paragraphs in declared order and the same container. Allowed separators are one space, `" | "`, `"\n"`, or `"\t"`. A tab separator requires a declared tab stop on the first surviving paragraph. Retains that paragraph's layout and moves later content/links into it. |
| `whitespace` | `[{"block": PARAGRAPH_ID, "action": "remove_empty"}]`; explicitly removes a truly empty paragraph. No text, controls, numbering (including inherited numbering), or unsupported content. Cannot remove a table cell's only or required trailing paragraph. |

Compaction rejects overlapping groups, numbered paragraphs, differing inherited
paragraph styles, explicit breaks/tabs/page breaks, and unsupported wrapped content.
Paragraph run defaults must match, or every affected run must directly override
the differing properties; compaction must not silently change inherited appearance.
Whitespace targets cannot overlap compacted paragraphs.

Theme validation calculates and audits contrast ratios for body, secondary,
primary, emphasized, section-heading, phrase-highlight, and boundary role pairs.
Normal text requires at least 4.5:1; large text and non-text borders require at
least 3:1. Referenced foreground/background combinations are checked as well.
Tab positions must be unique, increasing, and within the narrowest section's
printable width after any declared margin change. Successful results audit the
complete specification, contrast evidence, visual rules, exact highlight counts,
and explicit inserted tab tokens.

Explicit font declarations override inherited ASCII/HAnsi/complex-script theme
font attributes. Explicit numeric indentation removes conflicting character-based
indentation attributes for that property. Other formatting and numbering remain.
Cached `lastRenderedPageBreak` markers are nonprinting metadata and are preserved
exactly once when a text run is split. Complete bookmark ranges contained within
one paragraph are retained with their text positions; cross-paragraph or invalid
bookmark ranges are rejected when affected by a transformation.

Table linearization reuses paragraph XML in the declared order at the table's body
position and removes its layout table shells. It preserves wording, run properties,
hyperlinks, and numbering. Horizontal `gridSpan` merges and simple nested tables
are supported. Vertical `vMerge`, legacy `hMerge`, fields, drawings/textboxes,
tracked changes, content controls, comment/section boundaries, and unsupported
wrapped table content are rejected. Cell widths, backgrounds, borders, and table
presentation are intentionally discarded; declare readable text color explicitly
when removing a dark contact-table background. Paragraph rules, highlights,
compaction, and whitespace removal may target the original inspected paragraph IDs
within the same composition. Otherwise merged/nested affected structures retain
their existing restrictions.

Composition applies table linearization, block orders, global typography,
paragraph rules, highlights, compaction, then whitespace removal. It validates
text/link/numbering/bookmark inventories and complete block membership before creating one
immutable DOCX revision. The journaled audit result includes the full specification,
input/output checksums, matched phrases, explicit whitespace changes, preservation
results, and linearized table IDs, original paragraph order, body positions before
compaction, and discarded shell XML hashes. Previous source/artifact bytes remain
intact; rejected operations retain the committed cursor, review, and handles.

A successful composition clears acceptance. Create a checkpoint, inspect every
rendered page, and obtain explicit human acceptance of that exact new build before
publication. The old complete preview remains available until the new checkpoint
succeeds; a changed workspace checksum requires a fresh render. These operations
do not replace the stored `design`, add default pipeline steps, migrate manifests,
add dependencies, or change renderer selection. Page count and visual quality
must be verified on the resulting PDF.

Current gaps: automatic content rewriting, unsupported structural transformations, automatic
page-count targets, ATS validation, and automatic visual review.
Never bypass these through direct edits or ad hoc scripts. Use
[PLAN_TEMPLATE.md](PLAN_TEMPLATE.md), explain the missing capability, and wait for
the human to decide each plan. Independent authorized operations may continue.

## Explicit wording revisions

`content.revise` applies human-reviewed wording supplied as data; it does not
generate wording. It is manual, works before or after layout, and never advances
the default cursor. Use `layout.inspect`
first and resolve its exact manifest handle. Prepare a separate UTF-8 JSON file:

```json
{
  "specification": {
    "schema_version": 1,
    "expected_workspace_sha256": "<current workspace SHA-256>",
    "expected_inspection_sha256": "<current inspection SHA-256>",
    "edits": [
      {
        "block": "body/b14",
        "old_text": "Precent",
        "new_text": "Present",
        "expected_matches": 1,
        "formatting": "preserve_uniform",
        "reason": "Correct spelling in the employment date"
      }
    ]
  }
}
```

The placeholders must be replaced with 64 lowercase hexadecimal characters.
The paragraph ID is illustrative; use the current inspection's IDs. Every key
shown is required and unknown keys are rejected. `edits` must be nonempty.
`old_text` must be nonblank; `new_text` may be empty for an explicit deletion but
must differ from the old text. Text cannot contain tabs, line breaks, control
characters, or invalid XML characters. A positive integer `expected_matches`
declares all exact, case-sensitive occurrences in that paragraph. All matches
are evaluated against the input revision; overlaps, including self-overlapping
matches, fail. There is no fuzzy match or implicit occurrence selection.

```powershell
.\resume.ps1 operate layout.inspect --reason "Locate exact wording and run formatting"
.\resume.ps1 operate content.revise --params '@workspace/wording-review/proposed-revision.json' --reason "Apply the settled grammar and concision revision"
.\resume.ps1 operate layout.inspect --reason "Inspect new paragraph text before keyword styling"
```

`preserve_uniform` is the only formatting policy. A replacement may cross text
nodes and runs only when their serialized run properties are identical; it
inherits those properties. Mixed formatting is rejected, even when it may render
similarly. Use smaller independently valid patches or resolve formatting in a
separately reviewed supported operation. No automatic emphasis removal occurs.

Affected paragraphs must contain plain text runs, optionally including nonprinting
cached page-break markers. Paragraph properties and numbering are retained.
Existing tabs and breaks outside edit spans are retained exactly, including their
XML properties and order. Their text positions change only by declared preceding
text edits. Edits cannot cross, insert, or remove these controls.
Word spelling/grammar `w:proofErr` markers are retained unchanged. Paired adjacent
empty `_GoBack` bookmarks are supported when no replacement crosses their
position; offsets shift only by declared preceding text changes. Invalid marker
types, unpaired/nonempty/other named bookmarks, and crossing edits are rejected.
The change report includes before/after marker and bookmark-position evidence.
Links, other bookmarks, comments, fields, wrappers, section boundaries,
tracked changes, drawings, and other unsupported structures in affected paragraphs
are rejected. Interior paragraphs inside cross-paragraph bookmark/comment ranges
or complex fields are rejected too. The existing merged/nested-table restrictions
apply. Ordinary table cells and existing headers/footers are addressable. Unaffected
links, bookmarks, numbering, and other content remain intact.

The operation changes only declared text nodes and their XML whitespace attribute.
Run/paragraph structure and other XML are checked for preservation; all saved
paragraph text is checked against the exact declared diff. Unaffected DOCX package
members retain their bytes. It stages a new DOCX and an immutable JSON report,
then registers both through the existing journaled transaction. The report handle
is `resume.content.revision`; it records full edits, input offsets, reasons,
formatting policy, source identity, checksums, and preservation results. The audit
event carries the same evidence. A failed staging attempt may leave unregistered
artifacts, which recovery never selects. A journaled commit recovers both handles
and deduplicates its audit event.

A successful revision clears acceptance and removes the obsolete inspection
handle. Source, older artifacts, preview, published handles, stored design, and
cursor are preserved. The changed workspace checksum requires a fresh render;
create a checkpoint and obtain build-specific review before publishing. This
operation adds no renderer, dependency, automatic model, bullet-glyph change, or
default pipeline step.

For wording-first runs, use explicit `operate` and `checkpoint` commands. The
default `run`/`resume` pipeline still follows its stored cursor and can apply the
initial layout; do not invoke it merely to render the revised wording.

## Read-only plans and documentation

Open `/docs/` on the local monitor server, or follow **Plans & docs** in the monitor.
The viewer lists `.md` files recursively beneath the workspace's `docs/` directory.
Use the filename filter, select a document, and use its heading navigation.
Direct URLs such as `/docs/plans/resume-visual-optimization.md` render Markdown as
HTML. Refresh to see saved edits. UTF-8 Markdown supports tables, task lists,
lists, code fences, quotes, and links. Embedded HTML is escaped and images display
their alt text without loading resources. Relative Markdown links must resolve to
existing permitted documents. Only HTTP(S) and mailto external links are enabled.

The viewer rejects paths outside `docs/`, non-Markdown files, and symlink/junction
aliases. It is read-only and works without a resume manifest or renderer. It does
not create an audit event, change control state, approve a plan, or accept a build.
Existing monitor/checkpoint routes are unchanged. Install pinned dependencies
from `requirements.txt` before starting the server.
