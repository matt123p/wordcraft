# VBA documentation

This folder is the entry point for WordCraft's optional VBA integration. It
contains setup, host behavior, design decisions, support limits and historical
validation evidence. VBA documentation belongs here; the repository's main
README and general roadmap are maintained independently.

Start with the [integration guide](integration.md), which includes setup,
supported host behavior, native forms, package handling and remaining limits.

Shared runtime references use permanent links to the `v0.0.1` runtime commit;
reading this documentation does not require a sibling `tiny-vba` checkout.
Build and test commands run from this repository's root. Dated validation
receipts below describe their recorded revisions, rather than the current merge.
See the [main synchronization validation](validation.md) for current checks and
the remaining full-CI limitations.

## Setup and workflow

Run `cargo run -p wordcraft --features vba` from the repository root. The browser
application exposes the same `vba` feature. Open View > Macros > Visual Basic,
or use Alt+F11. Import `.bas`/`.cls` source or a native `.frm`/`.frx` pair;
imports are inert. Enable the per-document trust checkbox before running code
that accesses the document. Reset clears runtime instances. The removed `.tvform`
format is not supported.

The host supports the bounded plain-main-story Application, Document, Selection
and Range profile described in the [integration guide](integration.md#implemented-host-members).
Tables, inline objects and other unimplemented host members refuse explicitly.
Ordinary document saves preserve the original VBA project; persisting edited
source and native forms uses the explicit edited macro-enabled copy workflow.

## Implementation and evidence

`wordcraft-vba` owns the typed host bridge above the document engine. The UI
embeds tiny-vba's shared editor and supplies trust, file operations and lifecycle
callbacks. `wordcraft-docx` delegates package parsing and writing to `craft-docx`
through its model bridge; the local `vba_writeback` module prepares edited project
copies and removes invalidated signatures. Package-wide OPC signatures are
removed on import, while unchanged VBA project signatures remain opaque bytes.
Signature bytes never grant execution trust.

Fresh Office comparison remains M8.5 NOT_RUN. The integration guide's dated
milestone entries are historical; the tagged runtime below and native-only
workflow above supersede older runtime pins and `.tvform` plans.

Cargo dependencies use tiny-vba tag `v0.0.3`, resolved in `Cargo.lock` to
`1a79080db6b67ac531f282e391c97959ba52f1ba`.

Previous tagged runtime (2026-11-09): tiny-vba `v0.0.1` resolved to
`ac10617b5ee9d9112e5accffd07d6f6d1d9fb76b`.
This adopts the local M8 runtime fixes. Historical revision-specific receipts
below retain their original scope; fresh Office comparisons remain M8.5 NOT_RUN.
Validation: all nine CI steps, 69 native VBA tests, 40 Node 22 WASM tests,
strict optional VBA/UI Clippy and the VBA-enabled web build pass. The full
VBA UI run passed 1,224 cases (three ignored); its remaining document-input
replay case passes a focused rerun after disabling layout animations in the
fixture and asserting that the click remains queued across the pause.
Production canvas-change cancellation stays intact.

See the [WordCraft integration profile](integration.md) for setup, supported host
behavior, native forms and packages, acceptance evidence and remaining limits.
