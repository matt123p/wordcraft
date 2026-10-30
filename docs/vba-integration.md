# WordCraft VBA integration profile

Updated 2026-10-30. This is the WordCraft side of tiny-vba's M6 host-integration
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

| Task | Status | Evidence / remaining work |
|---|---|---|
| M6.1 Word object/metadata profile, live ranges and character positions | Selected profile implemented | Native Application, Document, Selection and Range objects; matching nominal `Word.*` type descriptors and completion catalogue; UTF-16 positions, optional/named `Document.Range` arguments, live-range edits, owner-bound `Me` forwarding during callbacks and stale-range refusal. Tables and inline objects refuse explicitly. This is a bounded plain-main-story profile. |
| M6.2 shared editor embedding and source routing | Partial | The editor starts closed and opens from View > Macros > Visual Basic or Alt+F11. It uses host theme/fonts and shared editor shortcuts; a host Alt+F11 assignment takes precedence when opening. Native/browser save/import/export route through existing services. Imports are bounded, inert and tied to the requesting document; save acknowledgments are revision checked. Source, trust, host handles and pending lifecycle state follow each open document even while the editor is hidden. Shared-label localization and an actual browser runtime source-transfer observation remain open. Package persistence is M7. |
| M6.3 document mutations, waits, debugger and event lifecycle | Partial | Real document mutation, undo grouping, partial effects on failure, pause/edit/resume refusal, fresh runs after user edits, explicit document callbacks, canonical automatic `Document_Close()`, WordCraft save/close cancellation extensions, retry/save ordering, reset/close reuse and shared modal-form wait/resume are covered. Standard Word Application event routing and automatic package Open/New delivery remain absent. There is no document calculation engine or asynchronous member in the selected profile. |
| M6.4 independent workflows and permitted reference observations | Partial | Seventeen authored adapter workflows and three position-map tests use the actual engine/session and shared interpreter. UI workflows cover editor access, document isolation and source routing. Microsoft Word reference observations remain NOT_RUN; these are WordCraft contract tests. |

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

## Verification evidence (2026-10-30)

- cargo test -p wordcraft-vba --locked --offline: 3 unit tests and 17 authored integration
  workflows passed. Covered Unicode mapping/refusal boundaries, live text mutation,
  undo grouping, stale handles, frame recreation, failed-macro partial effects and
  paused external edits, explicit document events, reset/close host reuse and a
  modal shared-form wait/resume workflow.
- cargo test -p wordcraft-ui-egui --features vba browser_vba_import_consumes_inbox_source_without_running_it --offline: passed.
- cargo test -p wordcraft-ui-egui --features vba vba_source_import_save_export_and_return --offline: passed.
- `cargo test -p wordcraft-ui-egui --features vba before_close --offline`: extension regressions pass, including Cancel writeback, abnormal stop and retry/save ordering. Canonical `Document_Close()` has separate live owner/mutation/failure coverage.
- cargo test -p wordcraft-ui-egui --features vba before_close_mutations_get_a_save_decision_without_replaying_the_event --offline: passed; cancel permits a fresh event attempt, then Save and Close do not replay it.
- `cargo test -p wordcraft-ui-egui --features vba before_save --offline`: three focused BeforeSave workflows passed against the M6 shared source checkout; handler edits reach the DOCX, a failing handler blocks the write while retaining partial edits, and `Cancel = SaveAsUI` cancels Save As without changing the existing DOCX.
- cargo test -p wordcraft-vba document_module_state_survives_successive_lifecycle_callbacks --offline: passed; the document-module instance retains state across repeated callbacks.
- cargo test -p wordcraft-ui-egui --features vba vba_projects_follow_open_documents_when_the_editor_switches --offline: passed; source modules and trust stay with the open document.
- `cargo check -p wordcraft --features vba --locked --offline`: native application passed.
- `cargo check -p wordcraft-web --features vba --target wasm32-unknown-unknown --locked --offline`: actual browser application passed, including the shared VBA source picker filters.
- `cargo clippy -p wordcraft-ui-egui --features vba --all-targets --locked --offline -- -D warnings`: passed.
- cargo xtask layers: passed; 21 crates, no layering violations.
- `CARGO_NET_OFFLINE=true cargo xtask ci`: all eight stages passed against the final immutable Git dependency pin (fmt, strict Clippy, workspace tests, i18n report, doctests, assets, layers and wasm). The parallel engine test exposed a hard-coded assertion that every session's document ID is `1`; the regression now compares with the ID captured from that session.
- `cargo test -p wordcraft-ui-egui --features vba --locked --offline --lib -- --skip i18n::audit::untranslated_interface_text`: 1003 passed, 0 failed, 2 ignored, 1 audit filtered. The earlier 143 feature failures were integration regressions caused by the editor opening by default; closing it by default and adding an explicit accessible entry point resolved them.
- `cargo test -p wordcraft-ui-egui --features vba --locked --offline`: 1003 passed, 1 failed, 2 ignored in the library suite; integration/doc tests were skipped after the failure. The unfiltered translation audit opens the editor and reported 43 unmarked English labels. The WordCraft-owned window title and trust text now use its translation path, but the shared renderer still needs an injected localization bridge. This gate is not green. Authored code and source identifiers must stay intact; the audit should distinguish those data names from renderer-owned interface labels.
- Microsoft Word reference observations: NOT_RUN.
- `cargo run -p wordcraft-ui-egui --features vba --example ui_shot --locked --offline`: original light/dark host/editor renders captured under `/tmp/m6-vba-editor-{light,dark}.png` and inspected. These are WordCraft output, not Microsoft reference screenshots.

The next embedding task is the shared renderer localization bridge, retaining the
full translation audit and original authored code. M6 also remains open for standard
Application event delivery, automatic Open/New project
attachment/delivery and permitted Microsoft Word reference observations. Ordinary-run
`ThisDocument` owner binding outside callbacks also remains unverified. No connected
Word reference session was available during this run. Browser picker/inbox logic is
unit tested and the browser application compiles, but an end-to-end browser runtime
source-transfer observation has not been recorded. Cross-session/package persistence
and package round trips remain M7. GridCraft's fork-side evidence is recorded in
`docs/m6-acceptance.md`. Shared dependencies pin `3f077dcb378536530b48ff54b962bab706740735`;
the exact Git revision can be verified offline from the local Cargo cache, while remote
availability is a separate publication check. The prerequisite feature branch must
be pushed before clean remote checkouts can resolve this pin.
