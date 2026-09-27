# Resume optimization guide

This is the canonical design instruction for planning, applying, and reviewing a
resume layout in this repository. Read it alongside [operation contracts](OPERATIONS.md).
The operation contract defines what the pipeline can do; this guide defines how
to make visual decisions with those capabilities.

The user's latest direction overrides examples and earlier feedback. A successful
render, preserved text, or preferred page count does not establish readability or
human acceptance.

## Establish the brief

Before changing layout, state the design contract in a proposal, composition
reason, or review note:

- intended audience and the information they should find first;
- requested section order and technical emphasis;
- preferred and maximum page counts;
- content-preservation and rewriting boundaries;
- the exact prior build whose feedback is being addressed;
- visual traits to preserve from the prior build or source;
- any uncertainty that must be resolved by rendering rather than assumption.

For the current resume, the standing brief is:

- optimize for HR reviewers and hiring managers scanning for relevant evidence;
- prefer two pages and never exceed three;
- prioritize Skills, About Me, Experience, then Education;
- present a full-stack engineer with stronger backend emphasis;
- preserve wording, claims, links, and full-stack breadth unless content editing is
  separately authorized.

Treat this list as a current brief, not a universal template. Reconcile it with
newer user instructions before each task.

## Inspect before assigning meaning

Read the current workspace inspection and the exact rendered reference before
assigning semantic roles. Identify actual section, employer, role, project,
bullet, link, table, and page-break boundaries. Paragraph indices or style names
alone do not prove a document's rendered reading order.

For each structural element, record what visual function it provides. A table may
group labels and values, associate dates with employers, organize contacts, or
create a navigation sidebar. If the structure is removed, the replacement must
restore those functions with supported headings, spacing, alignment, borders,
shading, or tab stops.

Never treat fewer tables, fewer columns, fewer pages, or a single reading stream
as an automatic improvement. Verify the result at whole-page and normal reading
scales.

## Choose the visual system before styling

Define one coherent system before applying individual formatting changes.

### Information hierarchy

Specify visible levels for:

1. name and contact header;
2. major sections;
3. employers or degrees;
4. roles and project groups;
5. body paragraphs and bullets;
6. dates and supporting metadata;
7. selected technical evidence.

Each level should differ through a small combination of size, weight, spacing,
color, border, shading, or alignment. Avoid a page where every line has similar
weight, and avoid using every available cue on every element.

### Semantic color roles

Choose colors by role rather than paragraph. Define at least:

| Role | Purpose |
| --- | --- |
| Paper | Main page background |
| Body | Long prose and ordinary bullet text |
| Primary | Name, section labels, and strongest landmarks |
| Secondary | Dates and supporting metadata |
| Accent | Small borders, project labels, or selected technical evidence |
| Surface | Gentle section or group background |
| Highlight | Restricted phrase-level emphasis |

Use a restrained palette with one principal accent family. Check the rendered
foreground/background combinations rather than assuming a named color is safe.
Normal text should reach at least 4.5:1 contrast; large text and non-text visual
boundaries should reach at least 3:1. Color must reinforce weight, spacing, or
placement instead of acting as the sole cue. Inspect a grayscale representation
when color is material to the hierarchy.

The approved generation 9 result is an example of an ink-and-teal system, not a
required palette for every resume. Its exact colors remain documented in the
[visual hierarchy plan](plans/resume-visual-hierarchy.md).

### Typography

Use an installed, legible font and a small, consistent type scale. The following
ranges are starting points for a dense professional resume, not fixed values:

| Element | Starting range |
| --- | --- |
| Name | 18–22 pt |
| Major section | 12–13.5 pt |
| Employer or degree | 11–12 pt |
| Role or project label | 10.5–11.5 pt |
| Body and bullets | 10.5–11 pt |
| Secondary metadata | 9–10 pt |

Prefer a readable body size and consistent leading over shrinking text to hit a
page target. Use weight and spacing to distinguish adjacent levels before adding
more font sizes.

## Compose for scanning

An HR reviewer should locate the main sections, employers, roles, dates, and
relevant skills without reading every paragraph. A hiring manager should be able
to follow a project from its overview to its technical evidence.

Keep these relationships visible:

- a company with its role and dates;
- a project label with its overview and supporting bullets;
- a skill category with its technologies;
- a degree with its institution, dates, and supporting details;
- a parent bullet with its nested evidence.

Use section gaps that are visibly stronger than gaps within one project. Use
project or employer gaps that are stronger than gaps between sibling bullets.
Keep labels and values on predictable edges. Do not fill a page with an
undifferentiated paragraph stream.

## Align by relationship

Use alignment to clarify relationships, not to decorate the page.

| Information | Preferred treatment |
| --- | --- |
| Long prose and bullets | Left aligned with a stable reading edge |
| Section headings | Consistently left aligned unless the whole page uses another clear system |
| Employer and role | Left aligned and visually grouped |
| Employment or education dates | Right aligned on the same row when the association remains obvious; otherwise placed nearby on the left |
| Name and contacts | Left aligned or deliberately centered as one compact group |
| Skill labels and values | Predictable aligned groups or supported columns |

Use declared tabs, paragraph alignment, or supported tables. Never align text with
repeated spaces. Avoid dates floating far from the employer or degree they describe.

## Treat bullets as hierarchy

Each bullet level needs one consistent marker, left indent, and hanging indent.
Wrapped lines must align with the first line's text rather than with the marker.

- Use the first level for a project, responsibility, or outcome.
- Use the second level for supporting implementation evidence.
- Use a third level only when the parent-child relationship is meaningful and easy
  to scan.
- Keep sibling spacing consistent. Use a larger gap for a new project or employer.
- Do not use deeper indentation to compensate for unclear grouping.

Inspect the actual glyphs after rendering. Different marker characters can have
different visual widths even when their numeric indents match.

## Emphasize selectively

Use an emphasis budget. Decide which evidence should attract attention and keep
ordinary prose quiet. For the current resume, backend technologies, data stores,
API work, reliability behavior, and testing evidence may receive selective
emphasis while frontend and end-to-end delivery remain visible.

Prefer a short existing phrase or outcome over highlighting an entire bullet.
Combine restrained bold, accent color, or gentle shading only when each cue helps.
Do not create new claims, metrics, ownership language, or stronger conclusions.
Too many highlighted phrases remove the distinction and make the page harder to
scan.

## Fit the page without crowding

Treat page count as a measured output. Two pages is preferred for the current
resume; three is the hard maximum. A readable three-page result is preferable to
an unreadable two-page result, but first use the available two-page canvas well.

Use this order of adjustment:

1. Remove accidental empty paragraphs or inherited spacing only through declared
   supported rules.
2. Establish reasonable margins and printable-width tab stops.
3. Set a readable type scale and consistent line spacing.
4. Choose an intentional page-two entry point at a section or project boundary.
5. Render and measure where content ends on each page.
6. Distribute remaining height through semantic gaps: sections, employers,
   projects, bullets, and education entries.
7. Re-render after every material adjustment and confirm the exact page count.

Use remaining space to improve rhythm. Do not add repeated blank paragraphs,
inflate every paragraph equally, create an oversized gap at one location, delete
content, or reduce text below a readable size. The practical target is content
reaching the lower text margin with balanced internal spacing, not ink touching
the physical page edge.

## Map the design to supported controls

Before composing, map every design choice to a documented field in
[operation contracts](OPERATIONS.md). `layout.compose` currently supports the
validated structural, typography, paragraph, theme, border, shading, tab, and
phrase-emphasis controls described there.

If a requested structure or transformation is outside the current contract, stop
that part and prepare a bounded proposal using the [plan template](PLAN_TEMPLATE.md).
Do not edit DOCX XML, invoke Word directly, change renderer behavior, or broaden a
validator as an operator workaround. A design instruction never authorizes a new
operation, content rewrite, renderer, acceptance, or publication.

## Iterate through exact artifacts

Use a deterministic loop:

1. Inspect the current workspace.
2. Prepare one complete composition specification for the intended revision.
3. Apply it through the registered operation with a concrete design reason.
4. Reinspect and verify structural and content preservation.
5. Create an immutable checkpoint.
6. Count pages in the exact rendered PDF.
7. Inspect every page at whole-page and normal reading scales.
8. Attribute findings to that exact build.
9. Apply further changes through another declared composition and repeat.

Resolve artifacts through manifest handles. Never pick a filename because it is
newest, edit a checkpoint, or use a reference PDF as if the pipeline rendered it.

## Review technical and visual quality separately

Technical validation should establish:

- the exact build and renderer provenance;
- preserved text, links, numbering, bookmarks, and unaffected package parts;
- supported structure with no unaddressed elements;
- matching manifest and artifact checksums;
- the exact PDF page count;
- no clipping, overlap, or broken output.

Visual review should answer:

- Can the reader find Skills, About Me, Experience, and Education at a glance?
- Are company, role, dates, project overview, and supporting bullets associated?
- Do font size, line spacing, paragraph spacing, and bullet indents form one system?
- Does each page use its practical text area without crowding or an obvious dead zone?
- Are page breaks intentional, with no stranded heading, employer header, project
  label, or education row?
- Do backend phrases stand out selectively while ordinary prose stays quiet?
- Do color, weight, spacing, and alignment reinforce the same hierarchy?
- Is the layout still understandable in grayscale?
- What improved over the prior build, and what became harder to scan?

Do not turn an informal glance into a measured reading-speed claim. No clipping is
necessary, but it does not establish readability or user satisfaction.

## Record feedback and acceptance precisely

For each feedback round, record:

- the exact build ID and generation inspected;
- the user's observation in their terms;
- what they explicitly liked;
- the reusable design lesson;
- the next hypothesis to test.

Keep build-specific evidence in the relevant plan or result document. Promote a
lesson into this guide only when it applies beyond that build. Later human feedback
can contradict earlier agent judgment and takes precedence.

Technical validation, agent visual review, human acceptance, and publication are
separate states. Never infer acceptance from a successful render, passing tests,
an optimization request, or informal praise of a different build.

## Final checklist

- The latest brief, section order, emphasis goal, and page target are explicit.
- Source structure and the exact prior build were inspected.
- The visual system was chosen before individual styling.
- Useful grouping was preserved or replaced with supported cues.
- Alignment communicates relationships and long prose keeps a stable left edge.
- The type scale, line rhythm, paragraph rhythm, and bullet hierarchy are coherent.
- Highlighting is selective, evidence-based, and readable in grayscale.
- Both pages use their practical text area without crowding, clipping, or orphans.
- Text, links, claims, numbering, and unaffected formatting are preserved.
- Every change maps to the current operation contract.
- Findings, feedback, acceptance, and publication refer to the exact build.

## Case study and provenance

The current method was distilled from the approved
[visual optimization plan](plans/resume-visual-optimization.md),
[table linearization extension](plans/resume-table-linearization.md),
[visual hierarchy extension](plans/resume-visual-hierarchy.md), and the
[generation 9 result](plans/resume-visual-optimization-result.md). Those documents
retain exact implementation details and build evidence; this guide retains the
reusable method.
