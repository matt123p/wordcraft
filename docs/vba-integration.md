# WordCraft VBA integration profile

Updated 2026-11-04: M7 adds inert packaged-source attachment when creating the
shared editor project. Standard/class modules and the selected `ThisDocument`
identity decode through the already pinned tiny-vba importer. Malformed or
unsupported-owner projects refuse before publishing a replacement project;
original package data remains intact. Every new invocation requires explicit
host authorization before VBA entry. No automatic Open/New handler execution
is introduced. Original project/direct signature/data parts survive `.docm`
save/reopen; source export remains separate from unsupported edited binary
writeback. Package-wide OPC signatures are dropped/logged on import into the
rewrite model. Signature bytes do not establish validity or execution trust.
[M7 acceptance](../../tiny-vba/docs/m7-acceptance.md) records the declared profile,
source identity, event lifecycle, exact effects and native/WASM evidence.

Updated 2026-11-01. This is the WordCraft side of tiny-vba's M6 host-integration
milestone. The adapter belongs in `wordcraft-vba` (L5), above the document engine;
tiny-vba remains a separately versioned shared interpreter/editor. Project import is
inert. The opt-in editor has a per-document trust checkbox; without it the bridge
refuses every VBA host call. Successful text mutations use WordCraft commands and
are grouped into a single undo step.

Enable the optional editor with `cargo run -p wordcraft --features vba`; the browser
application exposes the same `vba` feature. Open View > Macros > Visual Basic or use
Alt+F11. Import authored `.bas`/`.cls` source or a selected native `.frm`/`.frx`
form pair (legacy `.tvform` removal is scheduled separately), then explicitly trust
the document before execution. Merely importing a project or opening the editor
does not execute its procedures. Native paired browser imports are bounded to 8 MiB in aggregate; native resource
CFB data is bounded to 4 MiB and export source uses the narrower shared source limit.
Native file selection reads only the validated sibling companion. Other imports
retain their existing file bounds;
the shared IDE applies its narrower per-source/project/design budgets afterward.

## Editor error visibility (2026-11-02)

The editor displays a notice when document trust is disabled, and keeps the latest
message visible above the shared toolbar regardless of the debugger tab selected.
Host batch, command and execution failures enter this visible message log as well
as the application log. Host binding and project initialization failures display in
the VBA window even when the shared editor cannot initialize. Enable the trust
checkbox before running a macro that accesses the document.

Verification: the 12 VBA UI tests pass, including an invalid-entry command error,
visible runtime trust rejection, trust enablement, Step Into and successful Run.
Formatting and whitespace checks pass. These are authored WordCraft tests, not
Microsoft Word reference observations.

## M6 task status

Updated 2026-11-01 after the M6 event and attribute/code-view work. Both forks
pin published tiny-vba revision `fe86257c80c178baaa5701b5cf87597e4bc72daa`;
manifest and lockfile use immutable Git sources, with no path override.

| Task | Status | Evidence / remaining work |
|---|---|---|
| M6.1 Word object/metadata profile, live ranges and character positions | Selected profile implemented | Native Application, Document, Selection and Range objects; matching nominal `Word.*` type descriptors and completion catalogue; UTF-16 positions, optional/named `Document.Range` arguments, live-range edits, owner-bound `Me` forwarding during callbacks and stale-range refusal. Tables and inline objects refuse explicitly. This is a bounded plain-main-story profile. |
| M6.2 shared editor embedding and source routing | Implemented for the declared profile | The editor starts closed and opens from View > Macros > Visual Basic or Alt+F11. It uses host theme/fonts and shared editor shortcuts; a host Alt+F11 assignment takes precedence when opening. Native/browser save/import/export route through existing services. Imports are bounded, inert and tied to the requesting document; save acknowledgments are revision checked. Source, trust, host handles and pending lifecycle state follow each open document even while the editor is hidden. The shared renderer's chrome now goes through the host localization bridge (`Theme.localize` → WordCraft `i18n::t`), and the translation audit passes with the editor open; the audit exempts the authored default project title and entry name. Translated rows for the shared editor's strings land through craft-i18n separately. Actual Chromium import/export preserves exact CRLF/Unicode bytes without executing imported source. Asynchronous picker saves acknowledge only successful, revision- and project-instance-matching completion; cancellation/failure retains modified state. Package persistence is M7. |
| M6.3 document mutations, waits, debugger and event lifecycle | Implemented for the declared profile | Real document mutation, undo grouping, partial effects on failure, fresh runs after user edits, explicit document callbacks, canonical automatic `Document_Close()`, WordCraft save/close cancellation extensions, retry/save ordering, reset/close reuse and shared modal-form wait/resume are covered. Ordinary macro-list runs of a `ThisDocument` procedure now deliver through the host-authorized, owner-bound document path (`WordVbaProject::run` routes non-standard entries), proven with an implicit-member mutation test. Trusted Application BeforeClose/BeforeSave/WindowSelectionChange delivery is implemented through actual live WithEvents subscriptions, with signature validation, ordered ByRef propagation, pause/reset handling and lifecycle tails. Automatic package Open/New delivery remains absent. There is no document calculation engine or asynchronous member in the selected profile. |
| M6.4 independent workflows and permitted reference observations | Authored workflows complete; observations NOT_RUN | Twenty-five authored adapter workflows and three position-map tests use the actual engine/session and shared interpreter, including the M6.7 implicit-member and ordinary-run owner-binding cases. UI workflows cover editor access, document isolation and source routing. Microsoft Word reference observations remain NOT_RUN; no connected Word session was available. These are WordCraft contract tests, not Microsoft oracle results. |
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
Application selection-change delivery are implemented for explicitly connected,
trusted in-memory projects. Every handler signature is checked before effects;
handlers run in connection order and completed ByRef values feed the next handler.
Pauses/waits keep save/close pending, and Reset or trust revocation clears delivery.
Application handlers precede the document lifecycle tail; Cancel suppresses it.
Selection changes are coalesced while busy and do not recursively redeliver changes
made by the same selection callback, including edits after a debugger pause.

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

Validation against the implemented shared profile (final immutable pin above):

- `cargo test -p wordcraft-vba --locked --offline` (initial local-patch run, then repeated against the published Git pin):
  3 position-map unit tests and 25 actual document workflows pass. New cases cover
  Application events, signature/ByRef/connection behavior, trust, pause/reset,
  exact recorded-member metadata and single-step undo across multiple VM slices.
- Full `cargo xtask ci`: all eight stages PASS, including the 17-crate WASM
  check. Optional VBA UI suite: 1035 PASS, 2 existing ignored. Strict optional-
  feature Clippy passes against the final Git pin; the VBA-enabled browser app
  builds. The new UI regressions cover asynchronous source-save
  success/cancel/failure/stale completion, actual Alt+F11 frame routing,
  Application save/close decisions and selection callbacks resumed after a pause.
- Chromium exercised the actual WordCraft WASM app: source import was inert and
  fallback export bytes matched the original CRLF/Unicode source exactly. A rebuilt
  browser run also imported blank ThisDocument metadata and exported the exact
  recorded-member source without exposing metadata in its normal code view.
- Microsoft Word reference observations remain **NOT_RUN**. The owner has no Word
  installation and no connected session exists. Authored tests are not an Office
  oracle, and the fresh-comparison gate is deferred to tiny-vba M8.5 by owner direction.

## Attribute/code-view integration

The shared editor parses the exact recorded `VB_Invoke_Func` metadata shape and
omits it from normal code while retaining raw source for export. Attribute-only
`ThisDocument` source opens as blank code; explorer/Properties show the code name
separately from the source filename. Descriptions and Instancing are read-only;
View exposes a read-only raw metadata inspector. WordCraft refuses document
code-name rename requests explicitly. ThisDocument is host-owned and imported
project identities must not change without package writeback support.

Cross-session/package persistence and binary project writeback remain M7; no M7
implementation was started. GridCraft's evidence is in its `docs/m6-acceptance.md`.

Rebuilt-browser follow-up: Chromium confirmed inert attribute-only ThisDocument
import, exact Recorded.bas CRLF/Unicode export, and visible procedure rename
updating the hidden attribute target. With only the OS picker/writable boundary
mocked, pending and cancelled saves retained Modified; successful write/close
received exact raw bytes and then marked Saved. No page errors occurred. The
rich sample document's attempted Range mutation returned the declared error 5
plain-story refusal; native authored workflows establish successful host effects
on supported plain documents.

2026-11-01 language fix: the shared pin now includes implicit/explicit scalar
Variant For counters. This removes the reported CopyIngredients counter-type
diagnostic in the shared compiler; existing positive integral literal-Step
limits and separate Office host-member coverage still apply.

M6 sign-off update (2026-11-01): the owner accepted the tested M6 integration
profile and moved fresh Microsoft Excel/Word comparison to tiny-vba M8.5. These
observations remain NOT_RUN; they no longer block M7 unchanged-document/package
work. This scheduling decision does not change the supported host profile or
claim a new Office observation. See tiny-vba ADR 0032.

## M7A.1/M7A.2 native UserForms (2026-11-04)

The selected `native-userform-import-v1` profile now attaches embedded UserForms
from DOCM and native `.frm`/`.frx` pairs through the shared designer/code/Properties/
debugger/runtime. Immediate Label/TextBox/CommandButton children are supported;
unknown controls, containers, bindings and unsupported property/resource variants
refuse before publication. The shell export profile is Windows-1252; embedded
projects use their declared code page. Browser file picking supplies the explicit
pair together rather than publishing a source before its resources arrive.

Import is inert and never grants host authorization or fires Initialize. Authored
package tests prove explicit authorization, modal handler return, Reset and original
embedded project retention on save/reopen. Native edited source/resource saving
remains unsupported and cannot fall back to the custom archive. Full lifecycle/
control parity, edited persistence and legacy-format removal remain later M7A tasks.
See [native contracts](../../tiny-vba/docs/m7a-control-contracts.md) and
[acceptance](../../tiny-vba/docs/m7a-import-acceptance.md). No fresh Office
observation or signature validity is claimed.

## M7A.4 shared runtime increment (2026-11-04)

The host/editor now pin shared revision `fe86257c80c178baaa5701b5cf87597e4bc72daa`.
The flat native profile includes Label, TextBox, CommandButton, CheckBox,
OptionButton and ToggleButton. Explicit instances have independent fields,
windows and callbacks. Selected Value/Change/Click/group/Locked, tab ordering,
Default/Cancel and text-limit semantics and Controls/UserForms lookup are shared.
Unload separates displayed state/default ownership from retained code references;
control access reloads initial values and normal last-reference cleanup terminates.
Original embedded project resources remain unchanged on package save/reopen.

Instance input and form launch use the existing document execution-authorization
policy before dispatch, including handlers that only execute pure VBA. Editing,
compiling and exporting source remain available without granting execution trust.

See [the shared runtime record](../../tiny-vba/docs/m7a4-native-runtime.md) and ADRs
0043/0044 for scope/evidence and remaining requirements. M7A.4 remains in progress;
list/container/scroll/image/binding/platform packets, full focus/event/type parity,
native edited persistence and format removal are not implied. Fresh Office
comparisons remain NOT_RUN.

Validation for this increment: full workspace CI PASS; strict optional VBA/UI
Clippy PASS; both native package tests PASS natively and in actual Node 22
JavaScript/WASM; execution-consent regressions PASS through the real editor command
dispatcher; optional VBA web build PASS. Fixtures are original Rust builders under
MIT OR Apache-2.0. Package bytes remain unchanged, independently of native edited
saving. No fresh Office observation or vendor asset is included.
WordCraft full CI was rerun with `RUST_TEST_THREADS=1` after a PDF timing
assertion failed during concurrent checks; the isolated PDF case and all eight
full CI steps pass without a test/threshold change.

M7A.4 collection/lazy-instance follow-up (2026-11-04): the pinned shared runtime
adds Controls/UserForms For Each, retained enumerator/snapshot ownership, and
scalar project-class/native-form As New. Constructor reads use ordinary debugger
frames; declarations, Set targets and inspection remain inert. The native boolean
package case now verifies zero loaded forms before the first As New reference,
canonical child enumeration and UserForms identity, then actual input and exact
original package preservation. Expanded native families, focus/typed events and
MSForms type bindings remain open; this is not full M7A.4 acceptance.

Validation for this follow-up: full repository CI PASS (GridCraft 10 gates with
the licensed craft-fonts input; WordCraft 8 gates with serial tests), both native
package cases PASS natively and in actual Node 22 JavaScript/WASM, and strict
optional VBA/UI Clippy PASS. These are authored integration results.

M7A.4 Forms typing follow-up (2026-11-04): immutable shared revision `2ce3f205b1d6cce9a7bc7a6935cee1026c448d18`
adds native MSForms.UserForm/Control/Controls and leaf type identity, constrained
WithEvents interfaces and qualified VBA.UserForms access. The authored native
package case now enumerates into MSForms.Control/MSForms.UserForm variables
through VBA.UserForms while retaining lazy construction, canonical identity,
trust refusal, real boolean input and exact original project bytes on reopen.
Full CI, native/actual Node 22 WASM package/trust/input tests, strict optional
VBA/UI Clippy and the optional VBA web build PASS. Broader M7A.4
control/focus/platform semantics remain open.

M7A.4 native focus follow-up (2026-11-04): shared revision `2ce3f205b1d6cce9a7bc7a6935cee1026c448d18`
adds canonical ActiveControl/SetFocus, native container Enter/Exit and TextBox
BeforeUpdate/AfterUpdate with mutable typed cancellation and ordinary callback
frames. UI focus/input commands retain macro authorization and modal ownership;
Default/Cancel keyboard activation retains control focus. The native package
regression checks initial TextBox focus and checkbox focus after actual input,
with unchanged trust refusal, typed enumeration and original VBA preservation.
Full CI, native/actual Node 22 WASM package/trust/input tests, strict optional
VBA/UI Clippy and optional VBA web build PASS. The broader M7A.4
control/property/platform scope remains open in the shared implementation plan.

M7A.4 TextBox follow-up (2026-11-04): shared revision `2ce3f205b1d6cce9a7bc7a6935cee1026c448d18`
adds native UTF-16 selection/replacement, multiline/wrap and Enter/Tab/field
behavior, custom password masks and atomic text/selection input. The authored
macro package regression checks native initialization, MSForms.TextBox operations,
trusted input and exact original VBA bytes after reopen. Full CI, native/actual
Node 22 WASM package tests, strict optional VBA/UI Clippy and optional VBA web
build PASS. Full control/binding/property/platform scope remains open.

M7A.4 numeric follow-up (2026-11-04): shared revision `2ce3f205b1d6cce9a7bc7a6935cee1026c448d18`
adds selected native SpinButton/ScrollBar records, Long properties, orientation
and Change/spin/drag-only Scroll. Canonical input uses ordinary authorized focus/
validation frames, with pause/resume and reversed ranges. The fourth authored
native package test covers typed control access, trusted small/large steps and
original VBA-byte preservation. All eight full-CI gates, four native and actual
Node 22 WASM package tests, strict optional VBA/UI Clippy and optional VBA web
build PASS. The unchanged proofing timing case passed on a serial retry after
concurrent builds finished. Additional native control/property/binding/platform
scope remains open.

M7A.4 unbound ListBox follow-up (2026-11-04): shared revision `c578131f76c17228c6c5890a8ae5fdf9086d442c`
adds selected native ListBox records, ordinary indexed ten-column row data,
AddItem/RemoveItem/Clear, selection/Value/Text and column properties. Pointer/
keyboard input uses authorized focus/validation and ordinary Change/Click frames;
Stop/resume preserves committed state without replay. The fifth authored native
package test checks MSForms.ListBox access, trusted selection and exact original
VBA project bytes on reopen. All eight full-CI gates, five native and actual
Node 22 WASM cases, strict optional VBA/UI Clippy and optional VBA web build
PASS. Whole-array list APIs,
multiple selection, bindings, ComboBox and broader M7A.4 scope remain open.

M7A.4 ComboBox follow-up (2026-11-04): shared revision `6534f4fcbd70dfc54e4e332836f45912549fc37b`
adds selected native editable/drop-list ComboBox records, ordinary shared rows,
text/selection/matching and VM-owned drop state. Authorized text, navigation and
drop input preserves focus/validation and Change/Click/DropButtonClick debugger
continuations. The sixth authored package case checks MSForms.ComboBox access,
trusted row selection and UTF-16 free text with matching disabled, and exact
original project bytes after reopen. All eight full-CI gates, six native and
actual Node 22 WASM package cases, strict optional VBA/UI Clippy and optional
VBA web build PASS. Locale limits,
further list APIs/bindings/resources, containers and broader M7A.4 remain open.

M7A.4 list array follow-up (2026-11-04): shared revision `3688113be57d5413d7e96c29f0798d08c0d3506a`
adds native ListBox/ComboBox two-dimensional List loading, transposed Column
loading and indexed/current-row Column access (ADR 0053). The existing fifth/
sixth package cases now exercise nonzero source bounds, source-copy isolation
and current-row Column reads before authorized row/text input and exact original
project-byte save/reopen. All eight full-CI gates, six native and six actual
Node 22 WASM package cases, strict optional VBA/UI Clippy and the optional VBA
web build PASS. Whole-array reads,
additional loading/list/binding semantics and broader M7A.4 remain open; these
are authored expectations, with fresh Office comparisons NOT_RUN under M8.5.

M7A.4 readback/metadata follow-up (2026-11-04): shared revision `5cda307d4f9233175c915f6296b25896b12b0ffe`
adds detached zero-based List/Column array snapshots for the current unbound
ten-column packet, documented named parameters, native Tag/ControlTipText/
HelpContextID and runtime control tips (ADR 0054/0055). The list/ComboBox package
cases round-trip snapshots and mutate detached copies; a seventh package case
checks decoded metadata, ordinary MSForms.Control access, runtime writes,
authorized text input and exact original project/resource save/reopen. These
are authored expectations. All eight full-CI gates, seven native and seven actual
Node 22 WASM package cases, strict optional VBA/UI Clippy and the optional VBA
web build PASS. Further list/binding and
container/image/platform work, actual Help/F1 routing, native edited saving and
.tvform removal remain open. Fresh Office comparison remains NOT_RUN under M8.5.

M7A.4 native Frame follow-up (2026-11-05): shared revision `e246411ff33bf04d43a8c0e8d159c3679db257e9` supports
selected nested native Frame storages, immediate child Controls/Parent identity,
scoped tab/option groups, nested ActiveControl and cancellable ancestor focus.
Shared widgets use live relative geometry, clipping and accessible group ancestry.
The package regression exercises typed Frame access and trusted child text input,
then verifies exact original project resources after save/reopen. This does not
implement native edited-source/resource saving. MultiPage/TabStrip/Image/RefEdit,
further properties/bindings/platform/lifecycle semantics and .tvform removal
remain open; fresh Office observations remain NOT_RUN under M8.5.

Frame validation: all eight native and eight actual Node/WASM package cases PASS.
Full CI (eight gates), strict optional VBA/UI Clippy and the optional VBA web
build PASS. The Frame case also verifies authorization revocation; runtime text
and container state never replace original DOCM project/source/resource bytes.

M7A.4 native TabStrip follow-up (2026-11-05): shared revision `6bb45c6f607ecd2ae881ccdb314ace6e3073b29c` adds selected
native TabStrip records and distinct Tab/Tabs objects, indexed/named lookup,
enumeration, selection/default properties and authorized pointer/key callbacks.
Tabs remain separate from form Controls; source/resources remain independent.
Full M7A.4 remains open for MultiPage, Image/RefEdit, tab structural mutation and
writable Index/Name, further lifecycle/property/event/binding/platform cases.
Native edited persistence and .tvform removal remain M7A.5/.6 requirements.

The final TabStrip pin `fe86257c80c178baaa5701b5cf87597e4bc72daa` also rejects oversized
uncompressed tab character counts before they can alias the compression flag.

TabStrip validation: all nine native and nine actual Node/WASM package cases
PASS. Full CI (eight gates), strict optional VBA/UI Clippy and the optional VBA
web build PASS. The new case checks consent, distinct tab identity/defaults,
disabled input refusal, authorized selection, unchanged source/resources and
exact DOCM project bytes after reopen, plus authorization revocation.
