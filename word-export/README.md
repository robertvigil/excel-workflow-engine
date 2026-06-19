# Word export (`WorkflowWordExport.bas`)

A companion **Excel VBA macro** that exports a workflow into a **Word document** —
a clickable Table of Contents, one heading per step (nested by tree depth), each
step's `Notes` as the body text, and screenshots embedded inline.

This is a *separate add-in*, deliberately kept out of the core. The engine itself is a
single self-contained Office Script (`../workflow-engine.osts`) that runs on Excel for
the web or desktop. This macro is the opposite in every dimension that matters — it's
**VBA**, it's **Windows desktop only**, and it drives **Word** over COM automation — so
it lives in its own folder and never touches the core script or its README.

It isn't shipped/blessed; it's a working prototype of the "Word export" roadmap idea.
See `../ROADMAP.md` → *Word export* for why this path (Excel-drives-Word) was chosen over
the alternatives (pandoc, a Word-side macro, Power Automate).

## What it produces

- **Title** from the `Workflow Name` config value (falls back to the sheet name).
- **Table of Contents** — a real Word TOC *field*, so it's clickable (Ctrl+Click) and
  updates with the document.
- **Headings** — each step becomes a `Heading N` paragraph where *N* is its depth in the
  tree, so the TOC mirrors your outline. The step's `No.` is folded into the heading text
  (`1.2 Step name`); a `Label` value, if present, is used in place of the step name.
- **Body** — each step's `Notes` as paragraphs beneath its heading.
- **Images** — `![[key]]` tokens in `Notes` are replaced by the matching screenshot,
  inline, exactly where the token sits.
- **Captions + List of Figures** — `![[key|Caption text]]` adds a numbered Word caption
  under the image, and a **List of Figures** is built at the top (the figure equivalent of
  the TOC) whenever at least one captioned image resolves.
- **Page breaks** — each top-level (`Heading 1`) section starts on its own page (so the
  Title + TOC stand alone on page 1). Set via the Heading 1 style's "page break before"
  property, so it reflows cleanly. Toggle with the `PAGE_BREAK_BEFORE_H1` constant at the
  top of the macro.

## How to run it

Windows desktop **Excel + Word**, with macros enabled.

1. In Excel: **Alt+F11** (VBA editor) → **Insert ▸ Module** → paste in
   `WorkflowWordExport.bas` → close the editor.
2. Select the workflow **definition** sheet (not a Published sheet).
3. **Alt+F8** → **ExportWorkflowToWord** → **Run**. Word opens with the document.

It acts on the **active sheet**, same as the engine. A successful run opens Word; problems
are reported in a message box.

## Precondition — the tree must be numbered

The macro can't call the engine's TypeScript, so it reads structure straight from the
**`No.` column**: heading depth is the dot-count of the number (`2.3.1` → depth 3) and
order is the outline order of the numbers. So **run `Renumber` first.**

Export **refuses and instructs** — it never renumbers for you (that would mutate the source
as a side effect of a derived output). It stops with a specific message when:

- there's **no `No.` column** → "Add a 'No.' column and run Renumber";
- a step row is **blank in `No.`** → "Some steps aren't numbered — run Renumber";
- the numbers are **stale** (a child whose parent number is missing) → "Numbering is out of
  date — run Renumber".

## Images — the `![[key]]` scheme

Screenshots are referenced the same way the engine references steps: **by a stable name,
never by position.** No registry table — the image's *name is the key*.

**1. Put screenshots on their own sheet and declare it.** Add a config row:

| Property | Property Value |
|---|---|
| `Images Sheet` | `Screenshots` |

**2. Name each screenshot.** Give each image a stable key via the **Name Box**:

1. **Click the image once** so its selection handles show (you've selected the picture, not a
   cell behind it — the Name Box should read `Picture N`, not a cell address like `A1`).
2. Click into the **Name Box** — the little box at the far left of the formula bar, left of
   the `fx`, where cell addresses normally show.
3. **Type the key** (e.g. `login-page`) and press **Enter**. You must press Enter or the
   rename won't stick.

That name is the image's durable identity — it survives moving, resizing, and re-anchoring the
picture. Keep keys unique on the sheet, no spaces, not starting with a digit (`-`/`_` are fine).
To check it took, reselect the image and look at the Name Box.

> Renaming shapes via the Name Box needs **desktop Excel** — Excel for the web won't do it.
> That's fine: this export macro is Windows-desktop-only anyway.

**3. Reference it from any `Notes` cell:**

| Token | Result in Word |
|---|---|
| `![[build-settings]]` | the image named `build-settings`, embedded inline at that spot |
| `![[build-settings\|Build configuration screen]]` | the image **plus** a numbered caption, and an entry in the List of Figures |

Keys are matched case-insensitively. The `!` prefix means *embed* (vs. a `[[name]]` link that
resolves to an outline number — see below) — the same convention Obsidian and Markdown use.

**Resolution rule — resolve if found, else leave verbatim.** If the key matches a named
shape on the `Images Sheet`, the picture is embedded. If it doesn't (typo, deleted image, no
`Images Sheet`, or the named sheet doesn't exist), the literal `![[key]]` text is left in the
document as a visible "missing image" placeholder — never silently blanked. Same discipline
as the engine's other references.

### Example

`Images Sheet` config row → `Screenshots`. On the `Screenshots` sheet, a pasted screenshot
named `login-page` via the Name Box. Then in a step's `Notes`:

```
Open the admin console and sign in.
![[login-page|The login page]]
You should land on the dashboard.
```

→ in Word: the sentence, then the `login-page` image, then a "Figure N: The login page"
caption, then the closing sentence — and the figure listed in the List of Figures.

## Cross-references — the `[[step name]]` link

Cite another step from any `Notes` cell with its **exact step name** in double brackets:

```
Don't start this until [[Configure the build]] is done.
```

On export, `[[Configure the build]]` becomes a **clickable cross-reference** that jumps to that
step's heading and displays its **current number and name** — e.g. `2.3.3 Configure the build`
(matching the heading, so the reader knows both *where* and *what*). You reference by the step's
name — its stable identity — never by the number, because the number is the one thing that moves
when you restructure. Both the number and the jump target are recomputed on every export, so the
citation can never go stale: restructure, re-`Renumber`, re-export, and it updates automatically.

This is **Word-export only** — the source sheet keeps the durable `[[name]]` token; only the
generated document shows the live number. It does nothing in Excel (Organize / Renumber / Publish
never touch it), which is the point: a cross-reference is only useful in a rendered document.

**Resolution rule — resolve if found, else leave verbatim.** `[[name]]` → a hyperlink to the
matching step if one with that name exists on the sheet; otherwise the literal `[[name]]` is left
in the document as a visible "unresolved" placeholder (typo, deleted step, or a name that only
lives on another sheet). Matching is case-insensitive. Same discipline as the image tokens.

Under the hood, each heading is bookmarked as it's written and the link is a Word hyperlink to
that bookmark (Ctrl+Click to follow, like the TOC).

## Caveats / known rough edges

- **Windows only.** Excel-driving-Word via COM is unreliable on Mac and absent on Web.
- **English Word UI** — uses built-in style names `Title`, `Heading 1`…`9`, `Caption`.
- **Clipboard** — embedding images copies/pastes through the clipboard, so it overwrites
  whatever you had copied.
- **First open** — Word fields (TOC / List of Figures page numbers) may need a manual
  update the first time (Ctrl+A → F9) depending on your Word settings.
- **Prototype** — built but not yet round-trip tested in Excel/Word; expect to tweak the
  paste/sizing or style names for your environment.
