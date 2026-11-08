//! WordCraft-owned VBA host bridge primitives.
//!
//! The interpreter stays in tiny-vba. This crate owns Word document identity, character
//! positions, host object lifetimes and the mapping from a VBA-visible Range to WordCraft's
//! document model. Package loading and trust remain explicit host decisions.
#![forbid(unsafe_code)]
#![deny(clippy::unwrap_used, clippy::expect_used, clippy::panic, clippy::unimplemented, clippy::todo, clippy::unreachable)]

mod host;
mod main_story;

pub use host::{ApplicationEvent, NativeDocumentSave, WordHost, WordHostState, WordVbaProject};
pub use main_story::{MainStoryMap, PositionAffinity, PositionError};
pub use vba_ide::{Action, FormMouseEvent, FormMouseInput, HostRequest as SourceRequest, Mode, SourceId};
pub use vba_ide::{NativeFormResource, NativeFormSave, ProjectImportLimits, native_form_companion_name};
pub use vba_values::Value;

/// Bounded source/form transfer before decoding or modifying an in-memory project.
pub const MAX_SOURCE_IMPORT_BYTES: u64 = vba_ide::MAX_PROJECT_BYTES as u64;
