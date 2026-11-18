# docx-editor.dev

WYSIWYG editor and rendering engine for DOCX. Output must match MS Word: fonts, theme colors, styles, tables, headers/footers, section layout.

## Local instructions

Before starting work, read `local/AGENTS.md` if it exists. Follow its additional instructions throughout this checkout.

The `local/` directory is gitignored. Keep its contents private and do not copy them into committed files or pull requests. If the file is absent, continue with these shared instructions.

## Communication

Write all replies in ASD-STE100 Simplified Technical English.

- Use the active voice.
- Keep sentences to 20 words or fewer.
- Give one idea in each sentence.
- Use simple tenses: present, past, and future.
- Use the same word for the same idea each time.
- Do not use idioms, slang, or jargon.
- Keep paragraphs to 6 sentences or fewer.

Keep technical items exact. File paths, function names, flags, and numbers do not change. For example, write `packages/core/src/layout/semantic-layout.ts` and `w:contextualSpacing` in full.

In user-facing text, always use the name `EigenPal Pro License`. Do not rename the legal license files or `LicenseRef-EigenPal-Pro-Evaluation-1.0` metadata.

## Packages

One engine. Thin chrome on top.

| Package | What | Status |
| --- | --- | --- |
| `core` | **The engine.** `store/` (canonical tree, ops, OPC read/write), `layout/` (DOM-free), `output/` (paint), `editor/` (facade, surface, chrome registry), `contracts/`, `binding/`, `automation/` | published |
| `i18n` | Shared strings | published |
| `fonts` | Metric-compatible substitutes for Word's defaults | published |
| `vue` | The upstream Vue 3 adapter. Reference for the Angular port only; it leaves in a later step. | reference |

This fork removed the upstream `pro`, `editor-api`, `docx-to-pdf`, and `telemetry` packages (EigenPal Pro License), and `react`, `nuxt`, `docx-to-markdown`, and `fonts-cjk`. Never copy code from those upstream packages into this repository.

**The engine must resolve to ONE copy.** It holds module-level state — the HarfBuzz shaper and its cache budget, the grapheme boundary strategy, layout caches keyed by object identity. Two copies in a tree do not crash; they load the shaper twice and miss every identity-keyed cache, quietly. So `core` is a **peer** of every adapter. Never move `core` to a regular `dependency` of an adapter.

Inside `core`, each directory is a guarded lane with a declared dependency edge and environment (`store` and `layout` are DOM-free, `binding` is the only PM-aware one). The DAG is machine-readable in `packages/core/src/__tests__/core-lane-graph.ts` and documented in `docs/architecture/production-engine-packages.md`. A lane taking a new dependency edits that DAG.

Active production authority: `openspec/changes/typed-ooxml-paragraph-editor/`. Superseded proposals are not requirements.

## Architecture — one pipeline

```
bytes → readOoxmlPackage (bounded OPC/XML) → canonical OoxmlNode tree per part
→ TreeDocumentStore → semantic-layout → semantic-paint → serializeOoxmlPart
```

**The tree store is the only source of truth.** Painted pages ARE the editable surface — contenteditable, but the DOM is a picture: browser mutations are prevented and re-expressed as tree ops.

**ProseMirror is only a projection** of one tree revision. The reverse direction never reconstructs the tree from it: it diffs the edited doc against the tree it came from, emits the smallest `TreeDocOp`s that explain the difference, or refuses outright (`TreeBindingRejection`). A silently-dropped edit is worse than a refused one, because only the refusal can be reconciled. PM exists only in `core/src/binding/`; `store`, `layout`, `output` and `contracts` are PM-free, enforced by `store/__tests__/prosemirror-isolation.test.ts`.

- **Canonical tree** — typed kinds where layout needs them (paragraph/run/table); everything else is a lossless `generic` node. Invalid or misplaced known elements demote to generic. Unknown content never locks editing.
- **Fidelity** — structural, gated by two D9 oracles: `canonicalOoxmlFingerprint` and a save/reopen `semanticDigest`. Modeled XML parts re-emit normalized; byte identity applies to non-XML parts only.
- **Mutation** — `TreeDocumentStore.transact` over `TreeDocOp`s (node id + UTF-16 offset) is the only write path. The node index makes cell and nested paragraphs ordinary. Cross-cell joins refused (`not-adjacent-siblings`).
- **Layout** (`layout/semantic-layout.ts`) — DOM-free, injected `TextMeasurer`, points everywhere (twips convert at property-read boundaries). `storyBlocks` walks body/hdr/ftr roots and flattens block SDTs. Tables: row pagination, header-row repeats, vMerge, clamped gridSpan. Headers/footers laid out once per variant at flow height, attached per page. Incremental: per-block cache keys + flow checkpoints + convergence; a no-change pass returns previous pages by identity. Paragraph fidelity resolves through `layout/style-cascade.ts`: `w:spacing` line rules, first-line/hanging indents, `w:contextualSpacing`, `w:pBdr` on all edges, tab stops and leaders, `w:vanish` (not measured, not painted), list markers from `numbering.xml`, table styles via `basedOn` gated by `w:tblLook`.
- **Selection** — maps only through `data-paragraph-id`/`data-start` on text and `data-drawing-paragraph-id`/`data-drawing-start` on inline picture spacers. Page furniture is `contenteditable=false` + `[data-docx-hf]`, excluded from selection.
- **Caret** (`editor/surface-caret.ts`) — the engine paints its own and suppresses the native one (`caret-color: transparent` on the pages layer) for exactly as long as it does. Geometry comes from `caretAt` on the layout, never the DOM, so an empty paragraph gets a caret too. The element is furniture (`data-docx-marker`, `contenteditable=false`) on the page content box, not on a line. Fails soft: range selection, IME composition or an unplaced position hands the native caret back rather than leaving none.
- **Input** (`editor/surface-input.ts`) — one keymap for every host (Word's paragraph shortcuts, `beforeinput` dispatch, clipboard, IME readback) as factories over `PaginatedSurface`, so hosts cannot drift into separate hand-written keymaps.
- **`createDocxEditor`** (`editor/docx-editor.ts`) — the full `Editor` contract over the surface; unimplemented reads return typed empty values. `snapshot()` is version-cached: same reference until state moves, sub-objects reference-stable (`useSyncExternalStore` contract), `perf` outside it. `attach(el)`/`detach()` split creation from mounting; `detach()` parks the live surface and `attach(el)` moves it, keeping undo, selection and collaboration, even into another window.
- **Host window** (`editor/host-realm.ts`) — the editor can mount in a pop-out window or an iframe and move between windows. DOM code resolves its document and window from a node it owns (`hostDocumentOf`/`hostWindowOf`, `editorHostWindow(editor)`), at use time, never from script globals; `instanceof` a DOM constructor is false across windows. Long-lived document/window listeners go through `surface-host-listeners.ts` so a move rebinds them. `no-restricted-globals` and `docx/no-cross-realm-instanceof` enforce it; write `globalThis.x` only where any window will do.
- **Chrome registry** (`editor/chrome-controls.ts`) — `CHROME_GROUPS` is the toolbar taxonomy; `ChromeSlotId` (`text.bold`, `font.family`, …) is public API, so renames are breaking. `commandForSlot`/`commandForSlotValue` is the command table; `toolbarCommandState`/`runToolbarCommand` give shared can-before-exec. `ChromeControlState` says HOW a control dispatches (`command`/`value`/`save`), never whether it is enabled — enabled state has exactly one source, `toolbarCommandState`. Unwired slots render disabled with the engine's reason.

## Verify

```bash
bun run typecheck
bun run lint
bun run test
bun run api:check
bun run i18n:validate
openspec validate typed-ooxml-paragraph-editor --strict
```

- `bun run lint` runs oxlint with `.oxlintrc.json`. Its errors are the `max-lines` caps (1000 lines for most files, per-file caps up to 3500 for the handful already past it) and the repository rules in `scripts/oxlint/docx-rules.mjs` (HTML-from-strings sinks in `packages/*/src` non-test code, global `@keyframes` names, varargs array spread in `core/src/{store,layout}`, raw Word compatibility-mode comparisons outside `core/src/layout/compatibility/`) plus the framework-isolation `no-restricted-imports`. Nothing else catches the caps, so adding to a large file passes typecheck and the whole suite and fails CI. Extract; do not raise the cap — `scripts/check-max-lines-caps.mjs` also fails a cap that drifts more than 150 lines above its file. `scripts/__tests__/oxlint-rules.test.ts` keeps the repository rules firing; extend it when you add one. Lint covers `examples/*/src` as well as `packages/*/src` — the examples hit the same line caps.
- `bun run test` shards the suite one process per file across a worker pool (`scripts/test/run-parallel.mjs`, `--jobs N` to pin the width). That is also what CI runs. `bun test` still works and is the one to reach for when you want a single file, `-t`, or `--watch`; `bun run test:serial` is the whole suite the old way.
- `bun run test:changed` runs only the test files that import a changed file, directly or through other imports, compared with the merge base with `main`. `--list` prints the selection. It runs the whole suite when a change is outside what the import graph can follow, such as `bun.lock`, a `package.json`, a `tsconfig`, or the test runner. A change in `core` still selects most of the suite, because most tests import the `core` entry point.
- The pool is machine-aware, and must stay that way: the heavy OOXML-oracle files peak at multiple GiB each and slowest-first starts them together, which once froze a 48 GiB machine (fourteen workers, ~94 GiB footprint, WindowServer watchdog death). The runner caps width by installed memory, splits width and budget between concurrent runs on the machine (tmpdir registry), and gates dispatch on each file's measured peak footprint plus live memory pressure. `--jobs` pins the width but never lifts the gate. A single test file that needs more than a few GiB is a bug to file, not a budget to raise.
- A file that leaves state on `document` can only be caught by the serial run — per-file processes hide it. Scope DOM queries to the container you mounted.
- The pre-commit hook runs `lint-staged`, then its other gates side by side (`scripts/precommit/gates.mjs`). The API snapshot gate skips packages whose `dist/` is older than their sources, because a stale build fails correct commits; CI checks every package on a fresh build.
- `git commit --no-verify` is fine locally, but `bun run format` and `bun run lint` are not optional — the hook runs both, and they are the two gates nothing else covers. Run the other relevant scoped checks too, and report a bypassed failing gate instead of calling it passing.
- Compare the run against the non-clean baseline in the active change.
- `bun run format` before pushing.

## Styling

All editor chrome CSS and color tokens live in the core stylesheet; adapters only `@import` it. Never hardcode hex/rgba — use `--doc-*` tokens. The document canvas is not themed; it stays Word-faithful.

## Public API

API Extractor snapshots live in `docs/api/<pkg-slug>/<entry>.api.md`; CI runs `bun run api:check`. On drift: `bun run api:extract`, commit. Changing a `@public` symbol: tag it in TSDoc, rebuild, re-extract, commit.

Vue composables must declare a named `Use<Name>Return` interface and annotate the return type, or core's internal types leak into the snapshot.

## Security — untrusted input

**Every value from a DOCX, pasted HTML or embedded part is attacker-controlled.** A `.docx` is a zip of XML the sender fully controls: font names, hyperlink targets, shape attrs, image rels, run text. Sanitize at the bounded parse/trust boundary (XML read + typed/generic tree construction), never at render time, so every downstream sink receives a sanitized projection. Contract: `openspec/changes/typed-ooxml-paragraph-editor/specs/typed-ooxml-canonical-tree/`.

Audit these whenever you touch parsing, serialization, `output/semantic-paint*`, clipboard or print:

- **No HTML from strings** — no `innerHTML`/`outerHTML`/`insertAdjacentHTML`/ `document.write` on file-derived values. Use `createElement(NS)` + `setAttribute`/`textContent`.
- **URLs through `sanitizeHref`** — allowlist `http(s)/mailto/tel/ftp`, drop `javascript:`/`data:`/`vbscript:`/`file:`, strip embedded tab/LF/CR. Every `href`, image `hlinkHref`, and `window.open(...)` arg.
- **Escape strings interpolated into CSS** — `@font-face` family names, and any inline `style` built from file data.
- **XML** — no DTD/external entity resolution (XXE), no nested entity expansion (billion-laughs).
- **Zip** — decompression ratio/size cap; reject part/rel/media paths with `..` or a leading `/`.
- **No zero-click external fetch** — never auto-load `TargetMode="External"` rels, remote `src`, or CSS `url()`/`@import`. Gate remote loads behind an explicit user action.
- **Resource limits** — cap recursion depth (nested tables/shapes/SDT/groups) and element counts. Never feed a file-supplied number into allocation, `.repeat()` or a loop bound. No catastrophic-backtracking regex on file-derived strings.
- **Escape on save** — `escapeXml` every attacker-derived string; never template a raw value into markup.
- **Prototype pollution** — guard `JSON.parse` merges and any XML-attribute-name → object-key assignment against `__proto__`/`constructor`/`prototype`.
- **Field codes / OLE** — never execute or auto-resolve field instructions (DDE, `INCLUDE*`) or embedded OLE/macro content. Render inert.

```bash
grep -rnE "innerHTML|outerHTML|insertAdjacentHTML|v-html|document\\.write|window\\.open\\(|\\.href\\s*=|font-family:.*\\$\\{" packages --include="*.ts" --include="*.tsx" --include="*.vue" | grep -viE "test|\\.spec\\."
```

Fix sibling sinks when you fix one. The HTML sinks in that grep are also lint errors (`docx/no-html-sinks` in `scripts/oxlint/docx-rules.mjs`), so a new one fails `bun run lint`; the grep still covers what the AST selectors cannot (`v-html`, raw `href` assignment, CSS interpolation).

## i18n

`packages/i18n/en.json` is source of truth; other locales mirror its shape with `null` = fall back to English. A missing key fails CI.

The published catalogs cover only UI shipped by the library packages. Files under `examples/` must not add translation keys to `packages/i18n`. Keep example-only text in the example or `examples/shared/example-text.ts`. Examples may reuse existing keys for shipped editor controls. Examples use English unless the task explicitly requires local translations.

```ts
const { t } = useTranslation();
t('formattingBar.bold');
t('navigation.find.counter', { current: 3, total: 15 });
```

New library UI string: add to `en.json`, use `t('key')`, `bun run i18n:fix`. New language: `bun run i18n:new <code>`, fill nulls, `bun run i18n:status`. Never hardcode user-facing English in shipped library components. This restriction does not apply to example-only text.

## Documentation style

READMEs, guides, and developer docs follow the [Google developer documentation style guide](https://developers.google.com/style): second person, active voice, present tense, sentence-case headings, American spelling, short sentences, and one idea per sentence. Introduce each code sample with text. Use code font for filenames and identifiers. Diagrams are mermaid, not ASCII box drawing.

OOXML reference: `reference/quick-ref/wordprocessingml.md`, `themes-colors.md`; schemas in `reference/ecma-376/part1/schemas/`. PDFs are gitignored — run `bun run reference:fetch` once when needed.

## Conventions

- **PRs** — short factual title (conventional-commit prefix); body is the minimum the diff doesn't show, often one sentence. A PR that resolves an issue MUST end its BODY with `Fixes #N` — GitHub links and auto-closes only from closing keywords in the body or in commits, never from the title, so a title-only `(fixes #N)` leaves the issue open after merge. No `@`-mentions, unrelated issue numbers, file lists, tooling footers or emojis. Write the body in the [Google developer documentation style](https://developers.google.com/style), the same guide the docs site follows: second person, active voice, present tense, sentence-case headings, American spelling, short sentences, one idea each. Code font for identifiers and filenames; bold for UI elements only.
- **Bugs** — `gh issue view <N> --repo eigenpal/docx-editor`. Dev server `bun run dev` → `http://localhost:5173/`. Live demo `http://docx-editor.dev/editor`. Commit `fix: ... (fixes #N)`. Screenshots → `screenshots/`.
- **Issue labels** — every issue gets exactly ONE `area:*` and ONE `priority:*` label; the taxonomy table lives in `README.md` under "Issue conventions". Optional `component:*` names the package. The labels are the source of truth for the public roadmap board (org project 2): `.github/workflows/sync-labels-to-project.yml` mirrors them into the board's Area/Priority single-select fields. An option name must keep matching its label slug (lowercase, spaces→dashes) or the sync warns and skips; an issue with no `area:*` label is removed from the board; two labels of one prefix clear the field. ALWAYS file an issue with both labels AND a native issue type: `gh issue create --type Bug|Feature|Task --label "area:...,priority:..."` (`gh issue edit <N> --type ...` to fix one). Never leave an issue without a type or an area.
- **ESM only** — no `require()`.
- **TypeScript** — `typecheck` scripts run `bun check` (Bun 1.4.3 or later), which reports the same errors as TypeScript 7 and writes no files. Declaration builds run TypeScript 7, installed as `typescript7` (`npm:typescript@^7`). The `typescript` dependency stays on 6, because tsup, rollup-plugin-dts, Next.js, and other tools load it as a library, and TypeScript 7 has no library API. A package's own `tsc` is TypeScript 7. Where both are installed (the root, `examples/write-agent`), `tsc` is TypeScript 6, so scripts that need `tsc` there call `node node_modules/typescript7/bin/tsc`; never use bare `bunx tsc` at the root. Examples whose framework loads TypeScript as a library (Next.js, Astro, Nuxt, and the Vue example) install only TypeScript 6. Declarations come from `scripts/build-declarations.mjs` through `withDeclarations` in each tsup config; do not turn tsup's `dts` back on.
- **Tailwind** — scoped to `.docx-editor`; the scoping is baked into `dist/editor.css` at core build time (`scripts/build-core-styles.mjs` + `packages/core/tailwind.dist.config.cjs`), so the shipped file carries no raw `@tailwind` directive. Rendered output isn't always protected, so use inline styles on painted elements. Never put Tailwind utilities on the element that carries `docx-editor` itself — scoped utilities only match descendants.
- **Focus stealing** — painted pages are the editable surface, so any mousedown reaching them moves the caret.
- **Icons** — inline SVG (Material Symbol paths), not a font. A missing name renders raw text.
