# WordCraft VBA integration profile

Updated 2026-11-01. This is the WordCraft side of tiny-vba's M6 host-integration
milestone. The adapter belongs in `wordcraft-vba` (L5), above the document engine;
tiny-vba remains a separately versioned shared interpreter/editor. Project import is
inert. The opt-in editor has a per-document trust checkbox; without it the bridge
refuses every VBA host call. Successful text mutations use WordCraft commands and
are grouped into a single undo step.

Enable the optional editor with `cargo run -p wordcraft --features vba`; the browser
application exposes the same `vba` feature. Open View > Macros > Visual Basic or use
Alt+F11. Import authored `.bas`, `.cls` or `.tvform` source, then explicitly trust
the document before execution. Merely importing a project or opening the editor
does not execute its procedures. Imports are bounded to 4 MiB before decoding;
the shared IDE applies its narrower per-source/project/design budgets afterward.

## M6 task status

Updated 2026-11-01 after the M6.5–M6.8 completion work. Shared dependency pins
`9eb3e4111d8a78b88b26d22d26571e86efd93871` (the M6 resolution, `With`, implicit-member
and localization API), immutable Git revision, no path override.

| Task | Status | Evidence / remaining work |
|---|---|---|
| M6.1 Word object/metadata profile, live ranges and character positions | Selected profile implemented | Native Application, Document, Selection and Range objects; matching nominal `Word.*` type descriptors and completion catalogue; UTF-16 positions, optional/named `Document.Range` arguments, live-range edits, owner-bound `Me` forwarding during callbacks and stale-range refusal. Tables and inline objects refuse explicitly. This is a bounded plain-main-story profile. |
| M6.2 shared editor embedding and source routing | Implemented for the declared profile | The editor starts closed and opens from View > Macros > Visual Basic or Alt+F11. It uses host theme/fonts and shared editor shortcuts; a host Alt+F11 assignment takes precedence when opening. Native/browser save/import/export route through existing services. Imports are bounded, inert and tied to the requesting document; save acknowledgments are revision checked. Source, trust, host handles and pending lifecycle state follow each open document even while the editor is hidden. The shared renderer's chrome now goes through the host localization bridge (`Theme.localize` → WordCraft `i18n::t`), and the translation audit passes with the editor open; the audit exempts the authored default project title and entry name. Translated rows for the shared editor's strings land through craft-i18n separately. Browser picker/inbox logic is unit tested; an end-to-end browser runtime source-transfer observation remains open. Package persistence is M7. |
| M6.3 document mutations, waits, debugger and event lifecycle | Implemented for the declared profile | Real document mutation, undo grouping, partial effects on failure, fresh runs after user edits, explicit document callbacks, canonical automatic `Document_Close()`, WordCraft save/close cancellation extensions, retry/save ordering, reset/close reuse and shared modal-form wait/resume are covered. Ordinary macro-list runs of a `ThisDocument` procedure now deliver through the host-authorized, owner-bound document path (`WordVbaProject::run` routes non-standard entries), proven with an implicit-member mutation test. Standard Word Application event routing and automatic package Open/New delivery remain absent and are recorded as absent. There is no document calculation engine or asynchronous member in the selected profile. |
| M6.4 independent workflows and permitted reference observations | Authored workflows complete; observations NOT_RUN | Nineteen authored adapter workflows and three position-map tests use the actual engine/session and shared interpreter, including the M6.7 implicit-member and ordinary-run owner-binding cases. UI workflows cover editor access, document isolation and source routing. Microsoft Word reference observations remain NOT_RUN; no connected Word session was available. These are WordCraft contract tests, not Microsoft oracle results. |
| M6.5/M6.7 module-context conformance and document roots | Implemented | WordCraft declares the `ThisDocument` implicit members from its supported catalogue, its instance type and the host globals' result types (`configure_document_context`). Bare `Content`/`Range(...)` inside `ThisDocument` dispatch to the live owning document; bare names that are neither declared nor implicit refuse. The shared-side per-layer record is tiny-vba's `docs/m6-context-resolution-check.md`. |
| M6.8 host-object member completion | Implemented | Completion types receivers from the declared host import types (`Application.`, `ActiveDocument.`, `Selection.`), call shapes, `Me` and the `ThisDocument` code name, using the same catalogue. Code aids stay read-only and never invoke host members. |

## Initial main-story position contract

The position mapper indexes only a document's main story when it consists of
plain top-level paragraphs without inline objects. Word character positions are
zero-based, count paragraph marks and include the final paragraph mark. The adapter
maps supported boundaries to WordCraft's UTF-8 byte offsets, counts supplementary
Unicode scalars as two UTF-16 units, and refuses an offset inside a surrogate pair.
The final paragraph mark maps to a virtual final boundary. Table cells, floating
objects, fields and other inline-object positions are explicitly unsupported until
their widths and mutation semantics have dedicated coverage.

## Implemented host members

| Object | Members |
|---|---|
| Application | Name, Version, ActiveDocument, Selection |
| Document | Application, Name, Content, Range(Optional Start, Optional End) |
| Selection | Range, Text |
| Range | Start, End, Text (Get/Let), SetRange, InsertBefore, InsertAfter |

All other members return VBA error 438. Untrusted host access returns error 70;
stale/foreign document handles return 91; invalid positions return 5; invalid
argument types return 13; the bounded object table returns 7 when exhausted.
The host exposes the shared editor's API catalogue for these supported members.

The explicit `Word` reference snapshot registers Application, Document, Selection
and Range under adapter version `WordCraft.Word.v1`, not an installed Office typelib
version. Typed assignments query the same live object store used for invocation;
wrong nominal types raise 13 and closed/foreign handles raise 91. No library is loaded.
`Me` forwards undeclared document members to the actual owning Document during callbacks.

Arguments bind case-insensitively by parameter name before mutation. Unknown names
raise 448, missing required arguments 449, extra/duplicate arguments 450, and positional
arguments after named arguments 5. Positions accept integral VBA widths only and must
be non-negative and inside the supported story; text arguments require String and
reject unpaired UTF-16 surrogates. `Document.Range` defaults omitted Start to zero and
omitted End to the final story boundary, including positional omissions. See the
[documented Range signature](https://learn.microsoft.com/en-us/office/vba/api/word.document.range)
and [SetRange parameter names](https://learn.microsoft.com/en-us/office/vba/api/word.range.setrange).

## Lifecycle and calculation contract

Trusted `ThisDocument.Document_Close()` runs after the save decision and before
WordCraft removes the document. It has no Cancel parameter. This matches Word's
[document event signature](https://learn.microsoft.com/en-us/office/vba/word/concepts/objects-properties-methods/using-events-with-the-document-object).
WordCraft's explicit failure policy blocks removal on callback error and retains
prior edits in normal undo history. A successful callback that edits the document
causes a fresh save decision without replaying the callback. A modal form wait keeps
the document open; finish or reset, then retry Close. Reset abandons the pending event.

The earlier `Document_BeforeClose(Cancel As Boolean)` and
`Document_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)` hooks are **WordCraft
extensions**, not standard Word document events. The close extension is used only
when canonical `Document_Close()` is absent. A true ByRef Cancel prevents removal or
the write; successful save-hook edits reach the file, and errors/waits prevent writing.
Completed waiting hooks are consumed once on retry. In Word, cancellable
[DocumentBeforeClose](https://learn.microsoft.com/en-us/office/vba/api/word.application.documentbeforeclose)
and [DocumentBeforeSave](https://learn.microsoft.com/en-us/office/vba/api/word.application.documentbeforesave)
belong to Application and include a Document argument. Their WithEvents routing and
Application selection-change delivery are not implemented here.

Only trusted, active in-memory projects receive automatic events. Source import
is inert; opening a package does not yet attach or run source. Explicit host-authorized
`Document_Open` execution is tested, but automatic Open/New package delivery is absent.

WordCraft has no document formula/recalculation engine. VBA calculation callbacks
are unsupported and no recalculation is performed by the VBA bridge. Document
mutations use ordinary WordCraft editing commands and undo groups; a failing macro
does not roll back successful earlier edits.

The mapping is rebuilt for each host call while object handles persist for the
document lifetime. VBA-originated text edits update captured ranges in UTF-16 units.
External text edits while paused invalidate subsequent calls with error 91. Formatting
changes are not detected by the current text snapshot, and the host only maps plain
main-story paragraphs.
At a fresh invocation or Reset, ordinary external edits refresh the root snapshot
and invalidate old range handles; captured ranges never acquire guessed positions.
InsertBefore/InsertAfter expand the receiver to include its original text and the
insertion, as described in the [InsertAfter contract](https://learn.microsoft.com/en-us/office/vba/api/word.range.insertafter).
Replacing that expanded range later in the invocation joins the same undo step,
including when changing the selection is necessary. Opening another file keeps
an otherwise untouched document when it owns authored VBA source.

## Verification evidence (2026-11-01 M6 completion run)

Shared dependency pin `9eb3e4111d8a78b88b26d22d26571e86efd93871` (immutable Git
revision; remote availability verified by the cargo update that locked it).

- cargo test -p wordcraft-vba --locked --offline: 3 unit tests and 19 authored integration
  workflows passed, including the new M6 cases: ordinary-run `ThisDocument.FillTitle`
  binds the live owner and its implicit `Content` member mutates the document as one
  undo step, and `Document_Open` resolves bare `Range(0, 3)` to the owning document
  (the `abcde world` assertion covers both the implicit call and the explicit
  `Application.ActiveDocument.Range(3, 5)` follow-up). The earlier coverage holds:
  Unicode mapping/refusal boundaries, live text mutation, undo grouping, stale handles,
  frame recreation, failed-macro partial effects and paused external edits, explicit
  document events, reset/close host reuse and a modal shared-form wait/resume workflow.
- cargo test -p wordcraft-ui-egui --features vba --locked --offline --lib: 1004 passed,
  0 failed, 2 ignored — including `i18n::audit::untranslated_interface_text`, which now
  passes with the editor open. The shared renderer's chrome goes through the host
  localization bridge (`Theme.localize` → WordCraft `i18n::t`); the shared tiny-vba
  render regression pins the same contract; and the audit exempts the authored default
  project title and entry name (`Untitled Project`, `Main`) per the documented
  authored-data boundary. The 8 before_close/before_save/vba/browser-import extension
  regressions pass.
- `cargo check -p wordcraft --features vba --locked --offline`: native application passed.
- `cargo clippy -p wordcraft-ui-egui --features vba --all-targets --locked --offline -- -D warnings`: passed.
- Microsoft Word reference observations: NOT_RUN — no connected Word session was
  available on 2026-10-30 or 2026-11-01; recorded rather than inferred. The browser
  source-transfer observation also remains open (picker/inbox logic is unit tested and
  the browser application compiles). Cross-session/package persistence and package
  round trips remain M7. GridCraft's fork-side evidence is recorded in its
  `docs/m6-acceptance.md`; the shared per-layer M6.5 record is tiny-vba's
  `docs/m6-context-resolution-check.md`. Standard Word Application event routing and
  automatic Open/New package delivery remain absent from the declared profile.

The 2026-10-30 evidence is retained in the git history of this file; its failing
translation-audit gate is resolved by this update.
