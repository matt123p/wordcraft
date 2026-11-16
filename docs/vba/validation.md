# Main synchronization validation

The feature branch incorporates main at
`9088ed777d4a6f0daae87d5f254f1893d584adcb`. VBA documentation changes are confined
to `docs/vba/`; the main README and general documentation match main.

Passed checks:

- Native VBA host tests: 69 passed.
- Edited-package and signature save/read regressions: two passed.
- The VBA paused-document-input/Reset regression: passed.
- Strict Clippy for the VBA host and VBA-enabled UI, including all targets.
- Workspace formatting and strict Clippy.
- Asset attribution, dependency layering, the WASM check and workspace doctests.
- Relative documentation links and pinned tiny-vba reference targets.

The complete `cargo xtask ci` run is **not green**. Its workspace test step stops
in `wordcraft-engine`: 1,439 tests pass and 25 fail, comprising 24 fit-to-page
assertions and one German/Portuguese hyphenation assertion. The fit-to-page test
source, proofing command source and font implementation are unchanged from main.
The documented proofing-data input (`CRAFT_I18N_DIR`) is unavailable in this
environment; the fit-to-page failures have not been independently diagnosed.
Later workspace test binaries, the translation audit and the language gate were
not reached by that run. Historical milestone receipts in this folder retain
their original revision scope.

## Main framework alignment

DOCX parsing and writing use main's `craft-docx` implementation and model bridge;
no copy of the old local DOCX parser or writer is retained. The local writeback
adapter only supplies the VBA-specific edited-copy and signature policy.
VBA package acceptance fixtures use main's bounded `craft-zip` Archive and Writer
APIs. The extra direct ZIP 3 dependency and its unused dependencies are removed.

The migrated package acceptance suite passes all six cases using `craft-zip`.
