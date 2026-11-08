# WordCraft VBA integration profile

M7A.6 native-only files (2026-11-08): the shared custom form codec/APIs and
`.tvform` file routes/examples are removed under
[ADR 0099](../../tiny-vba/docs/decisions/0099-native-only-form-workflows.md). New forms use
inert native templates. Native `.frm`/`.frx` Save/Export and edited macro-enabled
copy export retain the M7A.5 snapshot/signature/IO contract. Useful legacy-era
host callback tests now import native pairs; there is no legacy importer or
converter. Shared revision `d9431ce4992ee9580acfe2fe0f7276de229bedaf`. All eight CI steps
pass in isolation with one test thread, along with all 29 workflow cases,
29 native/29 Node 22 WASM package cases, strict optional VBA/UI Clippy, 13 retained
VBA UI cases plus hidden-editor/canvas/copy-command checks and VBA-enabled web.
The migrated native import retains its real document callback assertions. The
host canvas was inspected; 13 shared lockfile packages changed by revision only. M7A.7 acceptance and
M8.5 fresh Office comparisons remain separate; older codec/support entries below
are historical and superseded.

M7A.5 persistence (2026-11-08): native `.frm`/`.frx` Save/Export and
**Export Edited DOCM Copy** (`ui.vbaExportNativePackage`) are implemented. The copy
compiles idle/reset authored source without execution, encodes edited native form
resources, removes invalidated VBA signatures and reopens the macro-enabled output
to check exact project bytes. File selection revalidates project/document/source/
design identity; native output commits atomically and refuses the current document
path. Copy export/download leaves the original editor modified and preserves its
original project bytes. Ordinary document Save retains those original VBA bytes;
open the edited copy to use the persisted changes. Native `.frm` Save acknowledges
only successful current-snapshot IO; browser download requests never do.
Shared revision `8800fcb07e628d4719a5c5459969503ba7e70f7c`.
See the [supported workflow and acceptance limits](../../tiny-vba/docs/m7a5-acceptance.md).
All eight CI steps pass in isolation with one test thread, along with strict
optional VBA/UI Clippy, 13 retained VBA UI cases, hidden-editor/canvas and new
copy-command cases, the VBA-enabled web build and 29 native/29 Node 22 WASM package
cases. The legacy/agile/v3 signature-removal regression and incomplete manifest
refusals pass. Thirteen shared lockfile packages changed by revision only. M7A.4 is complete for the common profile under
ADR 0097; uncommon extensions belong to P9, `.tvform` removal to M7A.6 and final
integration to M7A.7. Fresh Office observations remain M8.5 NOT_RUN. Earlier
incremental entries below retain historical status and are superseded by this
current scope and persistence contract.

M7A.4 button click focus and input identity (2026-11-08): shared ADR 0096
implements CommandButton TakeFocusOnClick, native True defaults and False clicks
that preserve text focus/selection and deferred validation. Normal Tab/explicit
focus, authorized callbacks and Reset remain intact. Input checks canonical object
membership after ordinary writes settle and after peer index shifts; smaller
active-session slices are respected without raising limits. Package regressions
retain typed False at Stop and normal AutoTab to the button; host canvases retain
text focus. Shared revision `bcf3dc859095df7a331760d1469b8047c00cacf4`.
All eight CI steps pass in isolation with one test thread, along with strict
optional Clippy, 13 VBA UI cases plus hidden-editor/canvas cases, the VBA-enabled
web build and all 28 native/28 Node 22 WASM package cases. The dark host canvas
was inspected. Thirteen shared lockfile packages changed by revision only.
Full M7A.4 remains open; native edited saving/removal and fresh Office traces remain
M7A.5/.6 and M8.5 NOT_RUN.

M7A.4 hidden TextBox caret viewports (2026-11-08): shared ADR 0095 separates
scrolling from visible bars. None/room-hidden axes reveal canonical carets while
fixed authored bounds and translated pointer geometry remain intact. Caller-themed
accessible bars keep the existing authorized callback/default path. The host
canvas fixture now shows long text scrolled to its final row with ScrollBars=None.
No API, format, host authority or budget change. Shared revision `25a35a403e862bb8510bfe7117d95f379aeea1fc`. All eight CI steps
pass in isolation with one test thread, along with strict optional Clippy, 13 VBA
UI cases plus hidden-editor/canvas cases, the VBA-enabled web build and all 28
native/28 Node 22 WASM package cases. The dark host canvas was inspected.
Thirteen shared lockfile packages changed by revision only. Full M7A.4, AutoSize, repeat/wheel/platform and geometry/font-change
edge cases remain open; fresh Office comparisons remain M8.5 NOT_RUN.

M7A.4 TextBox ScrollBars (2026-11-08): shared ADR 0094 implements native
ScrollBars import, ordinary Long None defaults/constants/errors and actual
bounded text viewports. Measured room, MultiLine and WordWrap determine axes;
track/drag waits for callbacks and preserves canonical text/ranges. Tagged UI
offsets reset with execution/control identity, follow caret changes and retain
scrolled pointer geometry. Native ListBox stored Both is accepted without a
public member. Package regressions check typed state at Stop and original
project preservation. Shared revision `05b8862abcf2d3204d26e1226acda372b89b6bb3`. All eight CI
steps pass in isolation with one test thread, along with strict optional Clippy,
13 VBA UI cases plus hidden-editor/canvas cases, the VBA-enabled web build and
all 28 native/28 Node 22 WASM package cases. The corrected dark host canvas was
inspected. Thirteen shared lockfile packages changed by revision only. Full M7A.4,
AutoSize/hidden-bar/repeat/platform and container scrolling remain open.

M7A.4 SelectionMargin (2026-11-08): shared ADR 0093 implements native
TextBox/ComboBox bit-26 import, ordinary Boolean True defaults and a theme-local
left selection strip. Actual line selection/drag uses canonical UTF-16/CRLF
ranges after mouse callbacks finish, including Stop/Continue and changed policy.
Empty native edit fields render blank with accessible names. Package regressions
check typed state/ranges at Stop and original VBA project preservation. Canvas
fixtures compare the default TextBox margin with an editable ComboBox margin off.
Shared revision `295ae7adde51765e42698d392b5140a7a33c341f`. All eight CI steps pass in isolation with one test
thread, along with strict optional Clippy, 13 VBA UI cases plus hidden-editor/
canvas cases, the VBA-enabled web build and all 28 native/28 Node 22 WASM package
cases. The dark host canvas was inspected. Thirteen shared lockfile packages
changed by revision only. Full M7A.4 remains open; edited native
saving and fresh Office observations remain M7A.5 and M8.5 NOT_RUN.

M7A.4 AutoTab (2026-11-08): shared ADR 0092 implements native TextBox/ComboBox
bit-30 import, ordinary Boolean False defaults and accepted user-edit tab
traversal through existing validation/focus continuations. Actual edit cause
survives autocomplete with unchanged text; caret-only and programmatic changes
stay passive. Package coverage checks imported typed state, Stop/Continue,
authorized text edits, focus transfer and original project preservation.
Shared revision `e5a62681f074dab000894118fa71977e438e26df`. All eight CI steps pass in isolation with one test
thread, along with strict optional Clippy, 13 VBA UI cases plus hidden-editor/
canvas cases, the VBA-enabled web build and all 28 native/28 Node 22 WASM package
cases. The dark host canvas was inspected. Thirteen shared lockfile packages
changed by revision only. Full M7A.4 remains open; edited native
saving and fresh Office observations remain M7A.5 and M8.5 NOT_RUN.

M7A.4 HideSelection (2026-11-07): shared ADR 0091 implements native
TextBox/ComboBox bit-29 flags, ordinary Boolean state and True file/dynamic
defaults. Inactive selected text stays highlighted when False; both ComboBox
presentations use canonical ranges without changing edit permission or drop
state. Text/Value, selected ranges, events, original resources and all budgets
remain independent. Highlight painting uses local clipped galleys outside font
locks. Package regressions check trust refusal, typed state and nonempty ranges
at Stop/resume, and unchanged original VBA bytes. Canvas fixtures include a
runtime-created ComboBox with persistent highlighting.
Shared revision `74f89f2bfb80c461d9791175291b48d892fa6006`.
Full CI (all eight steps, isolated with one test thread), strict optional
VBA/UI Clippy, 13 VBA UI cases plus hidden-VBE/canvas cases, the VBA-enabled
web build and all 28 native/28 Node 22 WASM package cases pass. The dark host
canvas was inspected. The 13 shared lockfile packages changed by revision only.
Remaining selection, properties,
AutoSize/scrolling, lifecycle/input, bindings and image/RefEdit/ActiveX/platform
cases stay open. Full M7A.4 remains IN PROGRESS; edited native saving is M7A.5
and fresh Office observations remain M8.5 NOT_RUN.

M7A.4 caption Alignment (2026-11-07): shared ADR 0090 implements native
CheckBox/OptionButton flag translation, ordinary Long Alignment state, Right
defaults and published constants. ToggleButton's documented disabled member is
not enabled. Real indicators and their bevels follow the selected side; caption
text stays left aligned and clipped in its own rectangle. Empty captions retain
accessible runtime names. Value/Null/groups/WordWrap/events, source/resources and
all budgets remain unchanged. Package regressions cover trust refusal, typed
properties, inert stopped inspection, one-time resume and original VBA bytes.
Canvas fixtures include a runtime-created left-caption CheckBox.
Shared revision `b8d203dd00bef56acce74deeb0d25d13811fbe1e`.
Full CI (all eight steps, isolated with one test thread), strict optional
VBA/UI Clippy, 13 VBA UI cases plus hidden-VBE/canvas cases, the VBA-enabled
web build and all 28 native/28 Node 22 WASM package cases pass. The dark host
canvas was rendered and inspected. The 13 shared lockfile packages changed by
revision only. Remaining applicable members,
AutoSize/scrolling, lifecycle/input, binding and image/RefEdit/ActiveX/platform
cases remain open. Full M7A.4 stays IN PROGRESS; edited native saving is M7A.5
and fresh Office observations remain M8.5 NOT_RUN.

M7A.4 Boolean captions (2026-11-07): shared ADR 0088 implements native
CheckBox/OptionButton/ToggleButton WordWrap flags, True defaults and ordinary
Boolean state. Value, Null/TripleState, option groups, events and original
resources remain independent. Local caption layout retains widget input, full
accessible labels/state and authored bounds; ToggleButton captions stay centred.
Package regressions verify trust refusal, typed wrapping, stopped inspection,
one-time resume and unchanged original VBA bytes. Canvas fixtures include a
runtime-created wrapped CheckBox. ADR 0089 stabilizes the restricted WASI test
runner through value-preserving callback forwarding, with no stack/VM/capability
budget increase or added arguments, environment or preopened directories.
Shared revision `b7010bd2738bbcac87a018b94bd261c2d8533f24`.
Full CI (all eight steps, isolated with one test thread), strict optional
VBA/UI Clippy, 13 VBA UI cases plus hidden-VBE/canvas cases, the VBA-enabled
web build and all 28 native/28 Node 22 WASM package cases pass. The dark host
canvas was rendered and inspected. All 13 shared packages changed by revision
only in the lockfile. Alignment/TextAlign and remaining
properties/defaults, AutoSize/scrolling, lifecycle/input, bindings and image/
RefEdit/ActiveX/platform cases remain open. Full M7A.4 stays IN PROGRESS; native
edited saving is M7A.5 and fresh Office observations remain M8.5 NOT_RUN.

M7A.4 caption WordWrap (2026-11-07): shared ADR 0087 decodes applicable
Label/CommandButton flags and exposes ordinary Boolean state with distinct file/
runtime defaults. Dynamic TextBox starts with WordWrap=True. Local caption
wrapping/clipping preserves fixed bounds, fonts and accelerators; unwrapped
multiline TextBox allocation retains its horizontal interaction/accessibility
bounds and complete UTF-16 text/cursor rows. Package cases cover typed wrapping
through trusted Stop/Continue, original resources, document edits/undo and denied
execution. Canvas fixtures exercise the same canonical properties.
Shared revision `2be3f2dcbf96b3798adb348649a131fe73436061`.
Full CI (all eight steps with isolated serial tests), strict optional VBA/UI
Clippy, 13 VBA UI cases, hidden-editor and canvas-prefix/Reset flows, VBA web
and 28 native/28 actual Node 22-WASM package cases pass. The updated document/
form render was inspected. Boolean-control wrapping, AutoSize,
scrolling and the remaining property/lifecycle/input/binding/image/RefEdit/ActiveX
and platform cases stay open. Full M7A.4 remains IN PROGRESS; native edited saving
is M7A.5 and fresh Office observations remain M8.5 NOT_RUN. VM/host protocols,
global font/style policy and execution/memory/permission budgets are unchanged.

M7A.4 TextAlign (2026-11-07): shared ADR 0086 implements selected native
Label/TextBox/ListBox/ComboBox alignment with ordinary Long state and published
fmTextAlign constants. Native ParagraphAlign and public centre/right values are
translated explicitly. Font assignment, text/selection and events stay independent;
local glyph/caret positions agree, including UTF-16 pointer selection. ComboBox
popup rows keep their own alignment. Package regressions exercise typed alignment
through trusted Stop/Continue, retain document edits/undo and original VBA resources,
and keep denied execution inert. Canvas fixtures include aligned native controls.
Shared revision `77465858e83c5e59790296b5d2fb1ed3c3b39374`.
Full CI (all eight steps with isolated serial tests), strict optional VBA/UI
Clippy, 13 VBA UI cases, hidden-editor and canvas-prefix/Reset flows, VBA web,
and 28 native/28 actual Node 22-WASM package cases pass. The updated document/
form render was inspected. Boolean/RefEdit and remaining property/default,
wrap/AutoSize, input/lifecycle/binding and ActiveX/platform cases stay open.
Full M7A.4 remains IN PROGRESS; native edited saving remains M7A.5 and fresh
Office observations remain M8.5 NOT_RUN. No VM/protocol/global style/font or
execution/memory/permission budget changes are included.

M7A.4 appearance audit (2026-11-07): shared ADRs 0083–0085 remove the
non-Office Page Font/colour exposure; Page children and client painting use the
owning MultiPage. Native records retain applicable BorderColor, BorderStyle,
SpecialEffect and BackStyle. Ordinary typed properties preserve documented
border/effect coupling, transparency and errors. Published Font aliases share the
canonical Font with exact Currency/Integer/String/Boolean values. Local painting
uses those fields and clips fixed-size controls to authored bounds. Package
cases cover Page ownership, typed appearance, trust refusal, Stop/Continue and
unchanged original VBA resources; canvas fixtures use aliases and appearance.
The follow-up list renderer removes its unconditional outside outline; canonical
BorderStyle/SpecialEffect own ListBox edges, while ComboBox popups keep their
separate frame. Canvas fixtures include runtime-created flat lists.
Shared revision `0bc169c5dd38f57634a77515a45a58370b1220b3`.
Full CI (eight steps, isolated serial tests), strict optional VBA/UI
Clippy, 13 VBA UI cases, the hidden-editor form, the canvas-prefix/Reset flow,
VBA web and all 28 native/28 actual Node-WASM package cases pass. The host
render was inspected. Remaining member/default/selection, lifecycle/focus, binding,
keyboard/mouse, image/RefEdit and ActiveX/platform cases stay open. Full M7A.4
remains IN PROGRESS; native edited saving is M7A.5. Fresh Office observations
remain M8.5 NOT_RUN. Budgets and the existing macro bridge are unchanged.

M7A.4 Font-assignment update (2026-11-07): shared ADR 0082 adds ordinary typed
Font setters on native UserForms and Font-capable controls. Explicit assignment
shares an object reference; old aliases survive replacement and unload, and new
controls still copy independent container defaults. Root setters retain lazy
load/reload and unrelated objects fail nominal validation before mutation.
Package cases cover assignment across Stop/Continue, trust and original-resource
preservation; canvas input tests now use an assigned Font. Shared revision
`38dacbd178dd8f25c98b4da8338f575e85df8171` passes full CI (eight steps, isolated
serial tests), strict optional VBA/UI Clippy, 13 VBA UI cases, the hidden-editor
form and canvas-prefix/Reset flow, VBA web and 26 native/26 actual Node-WASM
package cases. The host render was inspected. Nothing/external/other-assignment, ambient/platform, appearance and the
full remaining M7A.4 checklist stay open; fresh Office comparisons are NOT_RUN.

M7A.4 Font update (2026-11-07): shared ADR 0081 retains native StdFont/TextProps
and VBFrame Font blocks with exact fractional sizes and original resources.
Ordinary Font objects expose typed attributes, independent instance state and
container defaults for runtime-added controls and Pages. Local rendering uses
canonical values and existing host fonts; optional face selection belongs to the
caller. Extreme display sizes are bounded with a visible notice while retaining
the canonical value. Host package cases cover trust refusal, typed aliases,
Stop/Continue and unchanged original VBA resources; canvas tests exercise authored
typography with existing activation/input ordering.
Pinned shared product `f899c9925de3e2d4ef838f8ac78287ec9b2f4c2b` passes full CI
(all eight steps, isolated serial tests), strict optional VBA/UI Clippy, 13 VBA
UI cases, the hidden-editor form, the canvas-prefix/Reset flow, VBA web and all
25 native/25 actual Node-WASM package cases. The host document render was inspected. Font replacement, ambient/platform realization, borders/effects,
transparency and the remaining full M7A.4 checklist stay open. Fresh Office
comparisons remain M8.5 NOT_RUN.

M7A.4 colour update (2026-11-07): shared ADR 0080 retains native BackColor/ForeColor
bits and file defaults, exposes ordinary typed properties and paints authored RGB
values through the shared form renderer. Optional platform/client palette lookup
belongs to the caller's UI theme; missing client entries keep their canonical value
and show an unavailable-palette message. A runtime-created MultiPage Page-parent
lookup regression is fixed. Native package tests cover trust refusal, Stop/Continue,
one-shot document effects and unchanged original VBA resources. Host canvas tests
exercise authored form/control colours with existing activation/input ordering.
Pinned shared product `d3c99baab86fad9e33333fd4204066757baec63d` passes full CI
(8 steps, isolated serial tests), strict optional VBA/UI Clippy, 13 VBA UI cases,
the hidden-editor form, the canvas-prefix/Reset flow, VBA web and all 24 native/
24 actual Node-WASM package cases. The host document view was inspected.
Fonts, borders/effects, transparency, remaining ambient/platform/default/selection
and collection/lifecycle requirements stay open. Full M7A.4 remains incomplete;
fresh Office comparisons remain M8.5 NOT_RUN.

M7A.4 input-order update (2026-11-07): shared ADR 0079 orders earlier native
key/text/paste/IME input before document Deactivate. Existing keyboard receipts,
widget defaults and Change callbacks finish first; spanning composition stays
with the form through Stop. Reset/context/trust invalidation discards unapplied
input. Modal redirection cancels the later document gesture while preserving the
original keyboard receipt. No host API, authorization or budget is changed.
Canvas regressions now also pause in the preceding form Change callback.
Pinned shared product `231818a4179a6b096bcc66b1d8922ad0282ac7ce` passes full CI
(8 steps, isolated serial tests), strict optional VBA/UI Clippy, 13 VBA UI cases,
the hidden-editor form, the stopped canvas-prefix Continue/Reset flow, VBA web
and 23 native/23 actual Node-WASM package cases. The host view was inspected.
Full M7A.4 remains incomplete; fresh Office comparisons stay M8.5 NOT_RUN.

M7A.4 document-focus update (2026-11-07): ADR 0078 adds a bounded raw-input
barrier before document canvas input. Native modeless Deactivate runs through
existing authorization; Stop retains the original gesture, Continue restores and
repaints the document before replay, and Reset/source/document/trust/geometry
changes discard stale input. Reentrant activation also cancels a partially
replayed pointer capture. RefEdit retains its explicit worksheet-selection path.
Pinned shared product `ce6482640332cb67ccfe3372886a4039c832437e` passes full CI (8 steps, isolated serial
tests), strict optional VBA/UI Clippy, 13 VBA UI cases, the hidden-editor form,
the new document-focus/Reset flow, VBA web and 23 native/23 actual Node-WASM
package cases. Broader M7A.4 remains incomplete; fresh
Office observations stay M8.5 NOT_RUN.

M7A.4 activation update (2026-11-06): shared ADR 0077 adds canonical native
Show/Hide/modeless activation, reentrant lifecycle guards and cross-window key
continuations. Window focus commands execute VBA through existing authorization;
external application focus and passive hover do not emit lifecycle callbacks.
The new package regression checks a stopped Deactivate handler before a one-shot
host edit and preserves original VBA resources. Automatic document-canvas raw
input buffering/replay remains unimplemented; L01/L02 and full M7A.4 stay open.
Pinned product and final validation are recorded in ROADMAP.md.

Updated 2026-11-05: M7 adds inert packaged-source attachment when creating the
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
pin published tiny-vba revision `21898dced2dfed4649142a5cc84f8792acda896b`;
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

The host/editor now pin shared revision `4455df4082430d3d0d56412df73defd92c26c1df`.
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

M7A.4 native MultiPage/Page follow-up — 2026-11-05: shared revision
`18e412caa8d9dfd1a1ba984ea49c42bc490ef8da` adds bounded native f/o/x joins,
canonical Page/Pages identity and child Controls/Parent, indexed/named lookup,
enumeration, Value/SelectedItem and selected-page input/focus/rendering. Internal
tabs remain private. Physical selection respects dirty-field cancellation and
Stop/resume. TabStrip/MultiPage Click now carries its Long index and selecting
the current page/tab raises no selection event. Native source/resources remain
inert and separate; original project bytes survive package save/reopen. Dynamic
Pages/Tabs mutation, further page/Frame/property/resource/event/lifecycle cases,
Image/RefEdit, bindings and applicable platform controls remain M7A.4 requirements.
Native edited saving/.tvform removal remain M7A.5/.6; no full milestone sign-off
or fresh Office observations are claimed.

Validation for the MultiPage pin: full WordCraft CI (8 steps), ten native and ten
actual Node/WASM package cases, strict optional VBA/UI Clippy and optional VBA
web check PASS. The new DOCM case verifies inert import, native page selection,
selected-page child input/refusal, trust revocation and exact original project
resources after save/reopen. These are authored expectations, not fresh Word traces.

M7A.4 native Image/picture follow-up — 2026-11-05: shared revision
`4455df4082430d3d0d56412df73defd92c26c1df` adds bounded inert BMP/GIF/JPEG/ICO
resources, ordinary Image/stdole picture identity and typed assignment/clearing,
clip/stretch/zoom, alignment/tiling and selected AutoSize/border/effect state.
Image clicks preserve dirty text focus and ordinary callback/Stop continuations.
Canonical state stays in VM instances, with immutable shared pixel assets and
bounded derivative texture caches. Original source/resources remain separate.
Vector/platform picture services and the full image event/presentation matrix,
RefEdit, bindings, dynamic collections, platform controls and remaining lifecycle/
property semantics stay in M7A.4. Native edited saving/.tvform removal remain
M7A.5/.6; fresh Office observations remain NOT_RUN under M8.5.

Image validation: full WordCraft CI (8 steps), eleven native and eleven actual
Node/WASM package cases, strict optional VBA/UI Clippy and optional VBA web
check PASS. The new DOCM case verifies inert import, nominal Image/stdole
picture access, shared immutable resource identity, nonfocusable authorized
clicks, trust revocation and exact original project bytes after save/reopen.
These are authored expectations, not fresh Word observations.

M7A.4 native RefEdit follow-up — 2026-11-05 (shared revision `21898dced2dfed4649142a5cc84f8792acda896b`): the shared decoder/runtime adds
bounded inert native class-table resolution and ordinary RefEdit text/drop
objects, including nominal metadata and selected input/continuations. WordCraft
does not advertise an Excel worksheet range-picker service. The package case
checks inert import, consent, typed text, explicit modeless-picker refusal and
exact original DOCM VBA bytes; it does not establish Word/RefEdit platform parity.
Broader RefEdit/platform/property/event/binding/lifecycle semantics and dynamic
collections remain M7A.4 work; native edited saving/removal stay M7A.5/.6. Fresh
Office traces remain NOT_RUN under M8.5.

RefEdit validation: full WordCraft CI (8 steps), twelve native and twelve actual
Node/WASM package cases, strict optional VBA/UI Clippy and optional VBA web
check PASS. Thirteen existing VBA UI cases and the new native-form widget case
PASS; the latter checks hidden VBE windows, a real authored button callback and
Stop/F5 continuation without replay. Its two variants preserve ordinary form
state and reveal VBE only when a callback pauses/fails. DOCM cases retain exact
original project bytes and verify execution-consent revocation.

Native UserForm windows and callback advancement are independent of VBE window
visibility. The shared form renderer is used while VBE is hidden, and a paused or
failed callback brings the debugger forward. Document trust checks, source requests
and project/document ownership still use the existing Word bridge.

M7A.4 native Tabs/Pages mutation — 2026-11-05 (shared revision
`422c5c2338a05276e353b27018a531bf23231af4`): native collections support
Add/Remove/Clear, writable Integer Index and runtime-created member Name.
Reordering preserves aliases and selected identity; collection changes commit
before ordinary Change/Stop continuations. Added empty Pages and reordered
native children use the shared renderer; removed Page children reject widget
input. Original project/source/resource bytes remain separate. Fresh Office
comparison remains NOT_RUN under M8.5; Controls.Add/Remove, multiple selection,
bindings and the remaining M7A.4 semantic ledger stay open. Native edited saving
and .tvform removal remain M7A.5/.6. Integration validation PASS: all 13 native and 13 actual Node/WASM package
cases, full eight-step CI, strict optional VBA/UI Clippy, thirteen VBA UI cases,
the hidden-VBE native-form widget case and the optional VBA web check. The new
package case checks both Tabs and Pages, consent, retained aliases, unchanged
document text and exact original DOCM VBA bytes.


M7A.4 runtime controls and selection — 2026-11-05 (shared revision
`91b0c7870b76c2c71dbcc5964563195f6f39498d`): the shared runtime adds
selected native Controls.Add/Remove/Clear, ListBox Multi/Extended selection and
IR 12 canonical TypeName. A DOCM regression exercises dynamic typed controls,
selection, authorization and exact package-resource retention together. Empty
ControlSource/RowSource remain available; nonempty worksheet sources are refused
because this embedding supplies no Excel Range capability. No spreadsheet model
or external object activation is introduced. Selected Excel cell/row bindings live
in the shared Forms provider and GridCraft's existing workbook bridge. Validation PASS: full WordCraft CI (8 steps), strict optional-VBA/UI Clippy,
all 14 native and 14 actual Node/WASM package cases, 13 VBA UI tests, the hidden-VBE
native-form widget case and optional VBA web check. Native edited saving/.tvform removal remain
M7A.5/.6, the remaining M7A.4 ledger is open, and fresh Office evidence is M8.5
NOT_RUN.


M7A.4 runtime object names — 2026-11-06 (shared revision
`c48e109d95074153b37bb342d3206b50a33ee5fa`): the existing Word object store
provides validated Application/Document/Selection/Range type spellings through
the shared read-only resolver. TypeName retains actual identity through typed,
Object and Variant aliases without reading document text or invoking properties.
Names and work/allocation remain bounded; stale/foreign/closed handles refuse.
TypeName/VarType/IsArray accept their documented varname named argument. VarType
object-default evaluation remains a separate gap. Validation PASS: full CI
(8 steps), strict optional-VBA/UI Clippy, 13 VBA UI tests, the hidden-editor native
form case, optional web check and both new actual Node/WASM object-store/alias
cases. Fresh Office comparisons remain M8.5 NOT_RUN and broader M7A.4 stays open.


M7A.4 VarType default values — 2026-11-06 (ADR 0067/IR 13): the existing
Word store supplies pure validated default availability; actual Range/Selection
reads use the normal authorized Text getter. Object/Variant aliases, Stop/resume
and CStr(default) retain canonical text including the final paragraph mark.
Application/Document have no zero-argument default advertised by this adapter.
Project/native control defaults run through ordinary frames, preserving errors
and waits. Read-only regressions assert unchanged document and undo history.
Fresh Office comparisons remain M8.5 NOT_RUN; broader M7A.4 remains open.

Pinned shared revision: `5f93e8140b457de6645327b885b263ab6b472e13`.
Validation PASS: full CI (8 steps), 28 native workflows, two selected Node/WASM
workflows, strict optional VBA/UI Clippy, 13 VBA UI tests, the hidden-editor form
case and optional VBA web build. Source/resources and existing budgets are unchanged.


M7A.4 As New arrays — 2026-11-06, shared ADR 0068/IR 14: fixed/dynamic
project-class/native-form elements use ordinary lazy constructor frames. Identity,
Nothing recreation, Preserve/Erase, private/static/restored storage, canonical
ByRef and Stop/For Each continuations retain existing limits. A new authored
DOCM test checks per-window callback receivers, document writes with one undo
step, trust loss and exact original VBA resources. Pinned validation PASS at
`d3f00127a47f08996bd89968ebb7ec019ed13c8b`: full CI (8 steps, serial isolated
tests), strict optional VBA/UI Clippy, 13 VBA UI tests, the hidden-editor form
case, optional web and all 15 native/15 Node-WASM package cases. Broader M7A.4
and fresh M8.5 Office comparisons remain open.


UserForms without form components — 2026-11-06 (shared ADR 0069):
the existing built-in collection supports zero Count, empty iteration, aliases
and runtime type names, Item bounds errors and error 424 for Add of a missing
form. Normal semantic shadowing and qualified VBA.UserForms remain available.
No control library or form is loaded by the no-component provider. Authored host
cases check authorization, Stop/resume and unchanged document/undo state.
Pinned validation PASS at `0e7a3fea4f306a1aac8207e00b0a2deb206f6737`: full CI
(8 steps, isolated serial tests), 29 native workflows, three selected Node/WASM
workflows, strict optional VBA/UI Clippy, 13 VBA UI tests, the hidden-editor form
case and optional web check. Broader M7A.4 remains open.

M7A.4 keyboard update (2026-11-06): shared revision
`9fde04ce9ed8eab15b7f0609160617c774c2fc1b` adds selected native typed
KeyDown/KeyPress/KeyUp callbacks, mutable ByVal MSForms.ReturnInteger, Shift flags
and ordinary WithEvents delivery. The shared raw-input barrier waits for handlers
before default edits, preserves Stop/Continue and cancels stale Reset input.
All 16 native and 16 Node-WASM package cases pass, including keyboard trust,
debugger continuation, document text, one undo step and exact original resources.
Full CI (8 steps, isolated serial tests), strict optional VBA/UI Clippy, 13 VBA
UI cases, the hidden-editor form case and optional web compilation also pass.
Full physical key/platform, shortcut, root-window focus, IME/locale/focus behavior,
accelerators, mouse and the remaining M7A.4 checklist rows stay open. Fresh Office
comparisons remain M8.5 NOT_RUN. See shared ADR 0070.

M7A.4 keyboard-focus update (2026-11-06): shared revision
`38578095ba1fc1ae3695f92a6c2e89cdb29d42f6` adds selected root input focus,
canonical Tab/SetFocus/Space ordering, stopped character/release continuation,
editor shortcut isolation and matching web text/key normalization (ADR 0071).
The root eligibility guard distinguishes TabStop=False from disabled controls.
Full shared gates and WordCraft full CI (8 steps, isolated serial tests), strict
optional VBA/UI Clippy, 13 VBA UI cases, hidden-editor form, web and all 17 native/
17 Node-WASM package cases pass. Root eligibility, Stop/Continue, authority,
text mutation, one undo step and exact original resources are verified. Full M7A.4 and fresh Office comparisons remain open.

M7A.4 accelerator update (2026-11-06): shared ADR 0072 at `7d43f082de4b7827da1e1d93c257ebeb50eca4ae`
adds persisted native accelerators, canonical Accelerator properties and selected
Label/button/boolean/Tab/Page Alt actions. Focus validation, Stop/Continue and
existing authorization apply. Caption underlines use the host theme and fonts.
Duplicate routing explicitly refuses; non-ASCII/platform ordering, mouse,
lifecycle and other checklist rows remain open. M7A.4 remains incomplete; fresh
Office observations remain M8.5 NOT_RUN. Full CI (8 steps, isolated serial tests), strict optional VBA/UI Clippy, 13 VBA
UI tests, the hidden-editor form case, web and all 18 native/18 Node-WASM package
cases pass. The accelerator case preserves authorization, one-shot debugger
continuation, document text, one undo step and original VBA resources.

M7A.4 mouse update (2026-11-06): shared ADR 0073 adds typed native
MouseDown/MouseMove/MouseUp/DblClick callbacks, indexed arguments, mutable cancel
and selected simple-control capture/defaults. Stop/Continue, outside release and
Reset preserve input identity and source. Event signatures omit unused class-field
metadata while preserving typed arguments, executable method fields and all limits.
Text/list/numeric/indexed-tab pointer defaults, lifecycle and other checklist rows
remain open; M7A.4 is incomplete and Office observations remain M8.5 NOT_RUN.
Pinned shared revision `c9f829824c56397c101ac9228c89db289b788458` passes full
CI (8 steps, isolated serial tests), strict optional VBA/UI Clippy, 13 VBA UI
cases, hidden-editor form, web and all 19 native/19 Node-WASM package cases.
Stopped mouse callbacks preserve authority, one-shot text effects, one undo
step and exact original VBA resources.

The final mouse pin treats form commands as executable unless explicitly known
inert. The Word UI trust-revocation regression rejects InstanceMouse before
dispatch. Full CI (8 steps, isolated serial tests), strict optional Clippy, 13
VBA UI cases, hidden-editor form, web and 19 native/19 Node-WASM cases pass.

M7A.4 TextBox pointer update (2026-11-06): shared ADR 0074 adds selected
caret, captured drag, double-click cancellation/word selection and Shift extension
after ordinary mouse callbacks. Native AutoWordSelect is decoded with its inverse
file bit; new dynamic text controls retain their separate default. Canonical
UTF-16/CRLF, focus redirects, host authorization, Stop/Continue and Reset/source
identity remain enforced. Modal bounds stay stable while input waits. Editable
ComboBox/RefEdit, list/numeric/tab pointer defaults, further platform/selection and
lifecycle cases remain open; M7A.4 is incomplete and fresh Office comparisons
remain M8.5 NOT_RUN. Validation receipts follow below.

Accepted shared pin `3e8676cc754f59e06df64d2677c1f64af0216432`: full CI (8 steps, isolated serial tests), strict
optional VBA/UI Clippy, 20 native and 20 actual Node-WASM package cases, 13 VBA UI
cases, the hidden-editor form case and web checks pass. The new selection case
preserves authority, stopped continuation, one undo step and exact original
document/VBA resources. M7A.4 remains incomplete.

M7A.4 composite-pointer update (2026-11-06): shared ADR 0075 adds selected
ComboBox/RefEdit text/drop/picker, list/popup row, numeric step/repeat/drag and
indexed tab/page-header defaults after ordinary callbacks. Retained modal display
metadata preserves window presentation while callbacks run, without authorizing
input. Tabs retain identity through reorder and discard removed targets. Existing
host consent, snapshots, ordinary VBA state and resource preservation remain in
force; ComboBox AutoWordSelect preserves its inverse native bit. M7A.4 and the
full runtime checklist remain incomplete. Fresh Office observations are M8.5
NOT_RUN; native edited persistence/removal remain separate.
Pinned shared product `f027c5d14f0a57c6a0ebfa0f1cc2fb393cb9c2b7` passes full CI (8 steps, isolated
serial tests), strict optional VBA/UI Clippy, 13 VBA UI cases, the hidden-editor
form case, the VBA web build and all 21 native/21 actual Node-WASM package cases.
The new modal pointer regression preserves stopped authority, one-shot document
effects, one undo step and original resources.

M7A.4 nested-input update (2026-11-06): shared ADR 0076 services child mouse
callbacks while a parent mouse/key callback is parked. Raw input scopes share the
existing 64-event queue; held releases retain their original capture, modal areas
follow Show order, and open ComboBox popups stay above their owning form. Reset,
parent unload and owner-local keyboard barriers suppress stale/replayed defaults.
Canonical VBA state and existing host authority remain unchanged. Selected L02
cases are covered; M7A.4 and the full lifecycle/control checklist remain partial.
Fresh Office observations remain M8.5 NOT_RUN.
Pinned shared product `c99996351a80a234b23f2e81b54e210dae0bd530` passes full CI (8 steps, isolated
serial tests), strict optional VBA/UI Clippy, 13 VBA UI cases, the hidden-editor
form, VBA web and all 22 native/22 actual Node-WASM package cases. The nested
regression validates separate child/parent receipts, one-shot document effects,
one undo step and original resources.

ADR 0076 covers selected nested modal raw-input continuations and stale release
suppression. Full Activate/Deactivate, document/modeless focus, reentrant binding
and other remaining lifecycle/platform cases still require implementation and
acceptance. This does not finish L02 or M7A.4.
