# VBA documentation

Tagged runtime (2026-11-09): Cargo dependencies use tiny-vba tag `v0.0.1`,
resolved in `Cargo.lock` to `ac10617b5ee9d9112e5accffd07d6f6d1d9fb76b`.
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
