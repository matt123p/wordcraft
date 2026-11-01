use std::collections::BTreeMap;
use std::sync::{
    Arc, RwLock,
    atomic::{AtomicBool, Ordering},
};

use serde_json::json;
use vba_foreign::{
    Completion, ForeignError, Host, HostArgument, HostCall, InvokeKind, ObjectTypeResolver, ReferenceLibrary, ReferenceSet, ReferenceType,
    ReferenceTypeKind, Reply,
};
use vba_ide::{Action, ApiMember, HostRequest, Ide, Mode, SourceId};
use vba_values::{NominalType, ObjectRef, Value, VbaString};
use wordcraft_doc::StoryRef;
use wordcraft_engine::Session;

use crate::{MainStoryMap, PositionAffinity};

const MAX_HOST_OBJECTS: usize = 4096;
const MAX_RUN_TICKS: usize = 1024;
const WORD_TYPES_VERSION: &str = "WordCraft.Word.v1";

#[derive(Clone, Debug)]
enum Object {
    Application,
    Document,
    Selection,
    Range { start: usize, end: usize },
}

#[derive(Clone, Debug)]
struct ObjectEntry {
    document: u64,
    value: Object,
}

/// The host's authoritative object store also serves bounded read-only type queries.
#[derive(Debug, Default)]
struct WordObjects {
    identity: u64,
    document: u64,
    closed: AtomicBool,
    values: RwLock<BTreeMap<u64, ObjectEntry>>,
}

impl WordObjects {
    fn get(&self, reference: ObjectRef) -> Result<Object, ForeignError> {
        if self.closed.load(Ordering::Relaxed) || reference.host != self.identity || reference.generation != self.document {
            return Err(failure(91, "foreign or stale Word object reference"));
        }
        self.values
            .read()
            .map_err(|_| failure(91, "Word object store is unavailable"))?
            .get(&reference.object)
            .filter(|entry| entry.document == self.document)
            .map(|entry| entry.value.clone())
            .ok_or_else(|| failure(91, "Word object reference is no longer valid"))
    }
}

impl ObjectTypeResolver for WordObjects {
    fn identity(&self) -> u64 {
        self.identity
    }

    fn supports_type(&self, reference: ObjectRef, expected: &NominalType) -> Result<bool, ForeignError> {
        let name = match self.get(reference)? {
            Object::Application => "application",
            Object::Document => "document",
            Object::Selection => "selection",
            Object::Range { .. } => "range",
        };
        Ok(expected.library() == "word" && expected.version() == WORD_TYPES_VERSION && expected.name() == name)
    }
}

/// Synchronous, bounded WordCraft host for the selected main-story VBA profile.
///
/// A host is tied to one open document for its lifetime. Closing or switching that document
/// invalidates every object handle. Range handles are updated after VBA-originated edits; edits
/// made elsewhere while execution is paused are detected as stale at the next host call.
pub struct WordHost<'a> {
    session: &'a mut Session,
    state: &'a mut WordHostState,
}

#[derive(Default)]
pub struct WordHostState {
    identity: u64,
    document: u64,
    next_object: u64,
    objects: Arc<WordObjects>,
    root_refs: BTreeMap<String, ObjectRef>,
    mutation_count: usize,
    last_text: String,
    authorized: bool,
}

impl Drop for WordHostState {
    fn drop(&mut self) {
        self.objects.closed.store(true, Ordering::Relaxed);
    }
}

impl<'a> WordHost<'a> {
    pub fn new(session: &'a mut Session, state: &'a mut WordHostState) -> Result<Self, String> {
        let document = session.document_id();
        if document == 0 {
            return Err("Word host identity is unavailable".into());
        }
        // A project is bound to the current Word window/document. A new project is made when a
        // document is activated, so host identity and document lifetime share this unique token.
        if state.document != document {
            *state = WordHostState {
                identity: document,
                document,
                next_object: 1,
                objects: Arc::new(WordObjects { identity: document, document, ..WordObjects::default() }),
                root_refs: BTreeMap::new(),
                mutation_count: 0,
                last_text: session.doc.plain_text(StoryRef::Body),
                authorized: false,
            };
            for (name, value) in [("Application", Object::Application), ("Document", Object::Document), ("Selection", Object::Selection)] {
                let reference = state.allocate(value)?;
                state.root_refs.insert(name.into(), reference);
            }
        }
        Ok(Self { session, state })
    }

    fn allocate(&mut self, value: Object) -> Result<ObjectRef, String> {
        self.state.allocate(value)
    }
}

impl WordHostState {
    fn allocate(&mut self, value: Object) -> Result<ObjectRef, String> {
        let mut objects = self.objects.values.write().map_err(|_| "Word object store is unavailable")?;
        if objects.len() >= MAX_HOST_OBJECTS {
            return Err("Word VBA object budget exceeded".into());
        }
        let id = self.next_object;
        self.next_object = self.next_object.checked_add(1).ok_or("Word VBA object identity exhausted")?;
        objects.insert(id, ObjectEntry { document: self.document, value });
        Ok(ObjectRef { host: self.identity, object: id, generation: self.document })
    }
}

impl WordHost<'_> {
    fn document_is_current(&self) -> Result<(), ForeignError> {
        if self.session.document_id() != self.state.document {
            return Err(failure(91, "Word document reference is no longer current"));
        }
        if self.session.doc.plain_text(StoryRef::Body) != self.state.last_text {
            return Err(failure(91, "Word document changed outside the active VBA call"));
        }
        Ok(())
    }

    fn resolve(&self, reference: ObjectRef) -> Result<Object, ForeignError> {
        self.document_is_current()?;
        self.state.objects.get(reference)
    }

    fn range(&mut self, start: usize, end: usize) -> Result<Value, ForeignError> {
        let map = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?;
        if start > end || end > map.len() {
            return Err(failure(5, "Word Range position is outside the main story"));
        }
        map.position(start).map_err(|error| failure(5, error.to_string()))?;
        map.position(end).map_err(|error| failure(5, error.to_string()))?;
        let reference = self.allocate(Object::Range { start, end }).map_err(|message| failure(7, message))?;
        Ok(Value::Object(reference))
    }

    fn selection_range(&mut self) -> Result<Value, ForeignError> {
        let map = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?;
        let (start, end) = self.session.sel.ordered();
        let (start, end) = if self.session.sel.is_collapsed() {
            let at = map.index(&start, PositionAffinity::Start).map_err(|error| failure(5, error.to_string()))?;
            (at, at)
        } else {
            (
                map.index(&start, PositionAffinity::Start).map_err(|error| failure(5, error.to_string()))?,
                map.index(&end, PositionAffinity::End).map_err(|error| failure(5, error.to_string()))?,
            )
        };
        self.range(start.min(end), start.max(end))
    }

    fn range_text(&self, start: usize, end: usize) -> Result<String, ForeignError> {
        let map = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?;
        if start > end || end > map.len() {
            return Err(failure(5, "Word Range position is outside the main story"));
        }
        let first = map.position(start).map_err(|error| failure(5, error.to_string()))?;
        let last = map.position(end).map_err(|error| failure(5, error.to_string()))?;
        let mut text = self.session.doc.copy_range(&first, &last).plain_text().replace('\n', "\r");
        if end == map.len() && start < end {
            text.push('\r');
        }
        Ok(text)
    }

    fn string_argument(call: &HostCall, index: usize) -> Result<String, ForeignError> {
        let argument = call.arguments.get(index).ok_or_else(|| failure(450, "missing Word argument"))?;
        let Value::String(value) = &argument.value else {
            return Err(failure(13, "Word argument must be a String"));
        };
        String::from_utf16(value.units()).map_err(|_| failure(13, "WordCraft cannot store an unpaired UTF-16 surrogate"))
    }

    fn integer_argument(call: &HostCall, index: usize) -> Result<usize, ForeignError> {
        let argument = call.arguments.get(index).ok_or_else(|| failure(450, "missing Word argument"))?;
        let number = match argument.value {
            Value::Byte(value) => i64::from(value),
            Value::Integer(value) => i64::from(value),
            Value::Long(value) => i64::from(value),
            Value::LongLong(value) => value,
            _ => return Err(failure(13, "Word position must be an integer")),
        };
        usize::try_from(number).map_err(|_| failure(5, "Word position must be non-negative"))
    }

    fn set_range_text(&mut self, reference: ObjectRef, start: usize, end: usize, text: &str) -> Result<(), ForeignError> {
        let map = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?;
        if start > end || end > map.len() {
            return Err(failure(5, "Word Range position is outside the main story"));
        }
        let terminal_mark_selected = end == map.len();
        let content_end = if terminal_mark_selected { end.saturating_sub(1) } else { end };
        let start = if start == end && terminal_mark_selected && start > content_end { content_end } else { start };
        if start > content_end {
            return Err(failure(5, "Word Range cannot replace the final paragraph mark"));
        }
        let first = map.position(start).map_err(|error| failure(5, error.to_string()))?;
        let last = map.position(content_end).map_err(|error| failure(5, error.to_string()))?;
        let mut normalized = text.replace("\r\n", "\n").replace('\r', "\n");
        if terminal_mark_selected && normalized.ends_with('\n') {
            normalized.pop();
        }
        let replacement_units = normalized.encode_utf16().count();

        self.session.run("select.range", &json!({"anchor": first, "focus": last})).map_err(|error| failure(5, error.to_string()))?;
        if self.state.mutation_count > 0 {
            self.session.join_next_undo();
        }
        self.session.run("text.insert", &json!({"text": normalized, "raw": true})).map_err(|error| failure(5, error.to_string()))?;
        self.state.mutation_count = self.state.mutation_count.saturating_add(1);
        self.state.last_text = self.session.doc.plain_text(StoryRef::Body);

        let new_end = start.saturating_add(replacement_units);
        let target_end = new_end.saturating_add(usize::from(terminal_mark_selected));
        let mut objects = self.state.objects.values.write().map_err(|_| failure(91, "Word object store is unavailable"))?;
        for (id, entry) in objects.iter_mut() {
            if let Object::Range { start: range_start, end: range_end } = &mut entry.value {
                if *id == reference.object {
                    *range_start = start;
                    *range_end = target_end;
                } else {
                    *range_start = transform_position(*range_start, start, content_end, new_end, false);
                    *range_end = transform_position(*range_end, start, content_end, new_end, true);
                }
            }
        }
        Ok(())
    }

    pub fn roots(&self) -> BTreeMap<String, ObjectRef> {
        self.state.root_refs.clone()
    }

    /// Access the live editor session while the macro is paused. Any text change made through it
    /// invalidates this host before the next VBA dispatch, so captured ranges are never reused at
    /// guessed positions.
    pub fn session_mut(&mut self) -> &mut Session {
        self.session
    }

    pub fn set_authorized(&mut self, authorized: bool) {
        self.state.authorized = authorized;
    }

    fn begin_batch(&mut self) -> Result<(), String> {
        let text = self.session.doc.plain_text(StoryRef::Body);
        if text != self.state.last_text {
            // A new invocation can use the current document roots after ordinary user edits.
            // Captured ranges retain their old identity and must never acquire guessed positions.
            self.state
                .objects
                .values
                .write()
                .map_err(|_| "Word object store is unavailable")?
                .retain(|_, entry| !matches!(entry.value, Object::Range { .. }));
            self.state.last_text = text;
        }
        self.state.mutation_count = 0;
        Ok(())
    }

    fn invoke_call(&mut self, mut call: HostCall, object: Object) -> Result<Value, ForeignError> {
        if call.member.len() > 128 {
            return Err(failure(438, "Word member is outside the supported M6 profile"));
        }
        let member = call.member.to_ascii_lowercase();
        let parameters: &[&str] = match (&object, member.as_str(), call.kind) {
            (Object::Application, "name" | "version" | "activedocument" | "selection", InvokeKind::Get)
            | (Object::Document, "application" | "name" | "content", InvokeKind::Get)
            | (Object::Selection, "range" | "text", InvokeKind::Get)
            | (Object::Range { .. }, "start" | "end" | "text", InvokeKind::Get) => &[],
            (Object::Range { .. }, "text", InvokeKind::Let) => &["Value"],
            (Object::Range { .. }, "setrange", InvokeKind::Call) => &["Start", "End"],
            (Object::Document, "range", InvokeKind::Get | InvokeKind::Call) => &["Start", "End"],
            (Object::Range { .. }, "insertbefore" | "insertafter", InvokeKind::Call) => &["Text"],
            _ => return Err(failure(438, "Word member is outside the supported M6 profile")),
        };
        let optional = matches!((&object, member.as_str()), (Object::Document, "range"));
        bind_arguments(&mut call, parameters, optional)?;
        match (object, member.as_str(), call.kind) {
            (Object::Application, "name", InvokeKind::Get) => string_value("WordCraft"),
            (Object::Application, "version", InvokeKind::Get) => string_value(env!("CARGO_PKG_VERSION")),
            (Object::Application, "activedocument", InvokeKind::Get) => {
                let reference = self.state.root_refs.get("Document").copied().ok_or_else(|| failure(91, "active document is unavailable"))?;
                Ok(Value::Object(reference))
            }
            (Object::Application, "selection", InvokeKind::Get) => {
                let reference = self.state.root_refs.get("Selection").copied().ok_or_else(|| failure(91, "selection is unavailable"))?;
                Ok(Value::Object(reference))
            }
            (Object::Document, "application", InvokeKind::Get) => {
                let reference = self.state.root_refs.get("Application").copied().ok_or_else(|| failure(91, "application is unavailable"))?;
                Ok(Value::Object(reference))
            }
            (Object::Document, "name", InvokeKind::Get) => {
                let name = self
                    .session
                    .file_name
                    .as_deref()
                    .or_else(|| self.session.path.as_ref().and_then(|path| path.file_name()?.to_str()))
                    .unwrap_or("Document1");
                string_value(name)
            }
            (Object::Document, "content", InvokeKind::Get) => {
                let map = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?;
                self.range(0, map.len())
            }
            (Object::Document, "range", InvokeKind::Get | InvokeKind::Call) => {
                let map = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?;
                let start = if call.arguments.first().is_some_and(|argument| matches!(argument.value, Value::Missing)) {
                    0
                } else {
                    Self::integer_argument(&call, 0)?
                };
                let end = if call.arguments.get(1).is_some_and(|argument| matches!(argument.value, Value::Missing)) {
                    map.len()
                } else {
                    Self::integer_argument(&call, 1)?
                };
                self.range(start, end)
            }
            (Object::Selection, "range", InvokeKind::Get) => self.selection_range(),
            (Object::Selection, "text", InvokeKind::Get) => {
                let map = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?;
                let (start, end) = self.session.sel.ordered();
                let (start, end) = if self.session.sel.is_collapsed() {
                    let at = map.index(&start, PositionAffinity::Start).map_err(|error| failure(5, error.to_string()))?;
                    (at, at)
                } else {
                    (
                        map.index(&start, PositionAffinity::Start).map_err(|error| failure(5, error.to_string()))?,
                        map.index(&end, PositionAffinity::End).map_err(|error| failure(5, error.to_string()))?,
                    )
                };
                string_value(&self.range_text(start.min(end), start.max(end))?)
            }
            (Object::Range { start, end: _ }, "start", InvokeKind::Get) => {
                Ok(Value::Long(i32::try_from(start).map_err(|_| failure(5, "Word position exceeds Long"))?))
            }
            (Object::Range { start: _, end }, "end", InvokeKind::Get) => {
                Ok(Value::Long(i32::try_from(end).map_err(|_| failure(5, "Word position exceeds Long"))?))
            }
            (Object::Range { start, end }, "text", InvokeKind::Get) => string_value(&self.range_text(start, end)?),
            (Object::Range { start, end }, "text", InvokeKind::Let) => {
                let text = Self::string_argument(&call, 0)?;
                self.set_range_text(call.target, start, end, &text)?;
                Ok(Value::Empty)
            }
            (Object::Range { start: _, end: _ }, "setrange", InvokeKind::Call) => {
                let new_start = Self::integer_argument(&call, 0)?;
                let new_end = Self::integer_argument(&call, 1)?;
                let map = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?;
                if new_start > new_end || new_end > map.len() {
                    return Err(failure(5, "Word SetRange positions are invalid"));
                }
                map.position(new_start).map_err(|error| failure(5, error.to_string()))?;
                map.position(new_end).map_err(|error| failure(5, error.to_string()))?;
                let mut objects = self.state.objects.values.write().map_err(|_| failure(91, "Word object store is unavailable"))?;
                if let Some(entry) = objects.get_mut(&call.target.object)
                    && let Object::Range { start, end } = &mut entry.value
                {
                    *start = new_start;
                    *end = new_end;
                }
                Ok(Value::Empty)
            }
            (Object::Range { start, end }, "insertafter", InvokeKind::Call) | (Object::Range { start, end }, "insertbefore", InvokeKind::Call) => {
                let text = Self::string_argument(&call, 0)?;
                let original_len = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?.len();
                let at = if member == "insertafter" { end } else { start };
                self.set_range_text(call.target, at, at, &text)?;
                let new_len = MainStoryMap::new(&self.session.doc).map_err(|error| failure(5, error.to_string()))?.len();
                let inserted = new_len.saturating_sub(original_len);
                // InsertBefore/After expand the receiver to include its original text and the
                // insertion. Text assignment uses the narrower replacement range above.
                let new_start = if inserted > 0 && start == end && end == original_len { start.saturating_sub(1) } else { start };
                let mut objects = self.state.objects.values.write().map_err(|_| failure(91, "Word object store is unavailable"))?;
                if let Some(entry) = objects.get_mut(&call.target.object)
                    && let Object::Range { start, end: range_end } = &mut entry.value
                {
                    *start = new_start;
                    *range_end = end.saturating_add(inserted);
                }
                Ok(Value::Empty)
            }
            _ => Err(failure(438, "Word member is outside the supported M6 profile")),
        }
    }
}

impl Host for WordHost<'_> {
    fn identity(&self) -> u64 {
        self.state.identity
    }

    fn invoke(&mut self, call: HostCall) -> Reply {
        if !self.state.authorized {
            return Reply::Error(failure(70, "VBA host access requires explicit document trust"));
        }
        let result = self.resolve(call.target).and_then(|object| self.invoke_call(call, object));
        match result {
            Ok(value) => Reply::Complete(Completion { value, writes: Vec::new() }),
            Err(error) => Reply::Error(error),
        }
    }

    fn resume(&mut self, _: u64) -> Reply {
        Reply::Error(failure(5, "WordCraft does not expose asynchronous host waits in this profile"))
    }

    fn cancel(&mut self, _: u64) {}
}

/// Shared tiny-vba project configured with WordCraft's explicit roots and API catalogue.
pub struct WordVbaProject {
    ide: Ide,
    batch_active: bool,
    run_ticks: usize,
}

impl WordVbaProject {
    pub fn new(host: &WordHost<'_>) -> Result<Self, String> {
        let mut ide = Ide::new();
        let imports = [("Application", "Application"), ("ActiveDocument", "Document"), ("Selection", "Selection")]
            .into_iter()
            .map(|(name, root)| vba_foreign::Import { name: name.into(), root: root.into(), member: None, kind: InvokeKind::Get })
            .collect();
        ide.configure_host(imports, host.roots(), api_catalog()).map_err(|error| error.to_string())?;
        // M6.7/M6.8: declare the document profile's implicit `Me` members for the
        // `ThisDocument` module, its instance type, and each host global's result
        // type. Metadata only; code aids never invoke host members.
        let document_members = [(
            "thisdocument",
            api_catalog()
                .iter()
                .filter(|item| item.supported && item.owner == "Document")
                .map(|item| item.name.to_ascii_lowercase())
                .collect::<Vec<_>>(),
        )]
        .into_iter()
        .map(|(module, members)| (module.to_string(), members))
        .collect();
        let document_types = [("thisdocument", "Document")].into_iter().map(|(name, ty)| (name.to_string(), ty.to_string())).collect();
        let host_types = [("Application", "Application"), ("ActiveDocument", "Document"), ("Selection", "Selection"), ("ThisDocument", "Document")]
            .into_iter()
            .map(|(name, ty)| (name.to_string(), ty.to_string()))
            .collect();
        ide.configure_document_context(document_members, document_types, host_types).map_err(|error| error.to_string())?;
        ide.configure_reference_bindings(
            ReferenceSet::new(vec![ReferenceLibrary {
                name: "Word".into(),
                version: WORD_TYPES_VERSION.into(),
                available: true,
                types: ["Application", "Document", "Selection", "Range"]
                    .into_iter()
                    .map(|name| ReferenceType { name: name.into(), kind: ReferenceTypeKind::Object })
                    .collect(),
                constants: Vec::new(),
            }])
            .map_err(|error| error.to_string())?,
        )
        .map_err(|error| error.to_string())?;
        ide.configure_type_resolver(Some(host.state.objects.clone())).map_err(|error| error.to_string())?;
        Ok(Self { ide, batch_active: false, run_ticks: 0 })
    }

    pub fn ide(&self) -> &Ide {
        &self.ide
    }

    pub fn ide_mut(&mut self) -> &mut Ide {
        &mut self.ide
    }

    pub fn has_document_event(&self, code_name: &str, handler: &str) -> bool {
        self.ide.modules.iter().any(|module| {
            module.kind == vba_ide::ModuleKind::Document
                && module.code_name.as_deref().is_some_and(|name| name.eq_ignore_ascii_case(code_name))
                && module.procedures.iter().any(|procedure| procedure.name.eq_ignore_ascii_case(handler))
        })
    }

    fn has_standard_entry(&self, entry: &str) -> bool {
        self.ide.modules.iter().any(|module| {
            module.kind == vba_ide::ModuleKind::Standard && module.procedures.iter().any(|procedure| procedure.name.eq_ignore_ascii_case(entry))
        })
    }

    pub fn is_completed(&self) -> bool {
        matches!(self.ide.mode(), Mode::Completed)
    }

    pub fn is_waiting(&self) -> bool {
        matches!(self.ide.mode(), Mode::Waiting)
    }

    pub fn is_reset(&self) -> bool {
        matches!(self.ide.mode(), Mode::Design)
    }

    pub fn is_running(&self) -> bool {
        matches!(self.ide.mode(), Mode::Running)
    }

    pub fn import_source(&mut self, name: &str, source: &str) -> Result<(), String> {
        if name.to_ascii_lowercase().ends_with(".cls")
            && let Some(code_name) = document_code_name(source).filter(|name| name.eq_ignore_ascii_case("ThisDocument"))
        {
            return self.ide.add_document_module(name, &code_name, source).map(|_| ()).map_err(|error| error.to_string());
        }
        self.ide.import_source_file(name, source).map(|_| ()).map_err(|error| error.to_string())
    }

    pub fn take_source_requests(&mut self) -> Vec<HostRequest> {
        self.ide.take_requests()
    }

    pub fn mark_source_saved(&mut self, id: SourceId, revision: u64) -> Result<(), String> {
        self.ide.mark_saved(id, revision).map_err(|error| error.to_string())
    }

    pub fn import_document_source(&mut self, source_name: &str, code_name: &str, source: &str) -> Result<(), String> {
        self.ide.add_document_module(source_name, code_name, source).map(|_| ()).map_err(|error| error.to_string())
    }

    /// Deliver a host-selected Word document event. Importing an event module never invokes it.
    pub fn dispatch_document_event(&mut self, code_name: &str, handler: &str, host: &mut WordHost<'_>) -> Result<Mode, String> {
        self.dispatch_document_event_with_arguments(code_name, handler, Vec::new(), host)
    }

    /// Deliver an event with host-captured arguments. ByRef changes can be retrieved after
    /// completion with `take_callback_outputs`; imports remain inert.
    pub fn dispatch_document_event_with_arguments(
        &mut self,
        code_name: &str,
        handler: &str,
        arguments: Vec<Value>,
        host: &mut WordHost<'_>,
    ) -> Result<Mode, String> {
        if !matches!(self.ide.mode(), Mode::Design | Mode::Completed) {
            return Err("Word document events require an idle or completed project".into());
        }
        host.begin_batch()?;
        self.batch_active = true;
        self.ide.clear_callback_outputs();
        let owner = host.state.root_refs.get("Document").copied().ok_or("Word document root is unavailable")?;
        if let Err(error) = self.ide.dispatch_document_event_for_owner(code_name, handler, arguments, Some(owner), host) {
            self.batch_active = false;
            return Err(error.to_string());
        }
        for _ in 0..MAX_RUN_TICKS {
            let mode = self.ide.mode();
            if !matches!(mode, Mode::Running) {
                if matches!(mode, Mode::Completed | Mode::Failed | Mode::Design | Mode::Stale) {
                    self.batch_active = false;
                }
                return Ok(mode);
            }
            self.ide.tick(host).map_err(|error| error.to_string())?;
        }
        self.batch_active = false;
        Err("Word VBA event tick budget exceeded".into())
    }

    pub fn take_callback_outputs(&mut self) -> Option<Vec<(usize, Value)>> {
        self.ide.take_callback_outputs()
    }

    pub fn take_boolean_callback_output(&mut self, index: usize) -> Option<bool> {
        self.take_callback_outputs()?.into_iter().find(|(argument, _)| *argument == index).and_then(|(_, value)| match value {
            Value::Boolean(value) => Some(value),
            _ => None,
        })
    }

    pub fn run(&mut self, entry: &str, host: &mut WordHost<'_>) -> Result<Mode, String> {
        // Word's macro list can run a `ThisDocument` procedure directly. Deliver
        // it through the host-authorized, owner-bound document path so `Me`
        // binds to the live document exactly as during lifecycle events.
        if !self.has_standard_entry(entry) && self.has_document_event("ThisDocument", entry) {
            return self.dispatch_document_event("ThisDocument", entry, host);
        }
        self.ide.entry = entry.into();
        if !self.batch_active {
            host.begin_batch()?;
            self.batch_active = true;
            self.run_ticks = 0;
        }
        if let Err(error) = self.ide.dispatch(Action::Run, host) {
            self.batch_active = false;
            return Err(error.to_string());
        }
        for _ in 0..MAX_RUN_TICKS {
            let mode = self.ide.mode();
            if !matches!(mode, Mode::Running) {
                if matches!(mode, Mode::Completed | Mode::Failed | Mode::Design | Mode::Stale) {
                    self.batch_active = false;
                }
                return Ok(mode);
            }
            if let Err(error) = self.ide.tick(host) {
                self.batch_active = false;
                return Err(error.to_string());
            }
        }
        self.batch_active = false;
        Err("Word VBA execution tick budget exceeded".into())
    }

    pub fn begin(&mut self, entry: &str, host: &mut WordHost<'_>) -> Result<Mode, String> {
        self.ide.entry = entry.into();
        if !self.batch_active {
            host.begin_batch()?;
            self.batch_active = true;
            self.run_ticks = 0;
        }
        self.ide.dispatch(Action::Run, host).map_err(|error| error.to_string())?;
        Ok(self.ide.mode())
    }

    pub fn pause(&mut self, host: &mut WordHost<'_>) -> Result<Mode, String> {
        self.ide.dispatch(Action::Pause, host).map_err(|error| error.to_string())?;
        for _ in 0..64 {
            if !matches!(self.ide.mode(), Mode::Running) {
                break;
            }
            self.run_ticks = self.run_ticks.saturating_add(1);
            self.ide.tick(host).map_err(|error| error.to_string())?;
        }
        Ok(self.ide.mode())
    }

    pub fn continue_execution(&mut self, host: &mut WordHost<'_>) -> Result<Mode, String> {
        self.ide.dispatch(Action::Run, host).map_err(|error| error.to_string())?;
        Ok(self.ide.mode())
    }

    /// Advance an already-dispatched UI action with a fixed instruction-tick budget.
    pub fn advance(&mut self, host: &mut WordHost<'_>) -> Result<Mode, String> {
        if !self.batch_active {
            host.begin_batch()?;
            self.batch_active = true;
            self.run_ticks = 0;
        }
        for _ in 0..64 {
            let mode = self.ide.mode();
            if !matches!(mode, Mode::Running | Mode::Waiting) {
                if matches!(mode, Mode::Completed | Mode::Failed | Mode::Design | Mode::Stale) {
                    self.batch_active = false;
                }
                return Ok(mode);
            }
            self.run_ticks += 1;
            if self.run_ticks > MAX_RUN_TICKS {
                self.batch_active = false;
                return Err("Word VBA execution tick budget exceeded".into());
            }
            self.ide.tick(host).map_err(|error| error.to_string())?;
        }
        Ok(self.ide.mode())
    }

    pub fn reset(&mut self, host: &mut WordHost<'_>) -> Result<(), String> {
        self.ide.dispatch(Action::Reset, host).map_err(|error| error.to_string())?;
        self.batch_active = false;
        self.run_ticks = 0;
        host.begin_batch()?;
        Ok(())
    }

    pub fn close(&mut self, host: &mut WordHost<'_>, discard_modified: bool) -> Result<(), String> {
        self.ide.close(host, discard_modified).map_err(|error| error.to_string())?;
        self.batch_active = false;
        self.run_ticks = 0;
        host.begin_batch()?;
        Ok(())
    }
}

fn transform_position(position: usize, start: usize, end: usize, new_end: usize, end_affinity: bool) -> usize {
    if position < start {
        position
    } else if position > end {
        position.saturating_sub(end.saturating_sub(start)).saturating_add(new_end.saturating_sub(start))
    } else if position == end || (position == start && end_affinity) {
        new_end
    } else {
        start
    }
}

fn document_code_name(source: &str) -> Option<String> {
    source.lines().find_map(|line| {
        let line = line.trim();
        let (attribute, value) = line.split_once('=')?;
        if !attribute.trim().eq_ignore_ascii_case("Attribute VB_Name") {
            return None;
        }
        let value = value.trim();
        let quoted = value.strip_prefix('"')?.strip_suffix('"')?;
        (!quoted.is_empty()).then(|| quoted.to_string())
    })
}

fn string_value(value: &str) -> Result<Value, ForeignError> {
    VbaString::from_utf8(value).map(Value::String).map_err(|error| failure(7, error.to_string()))
}

fn failure(number: i32, description: impl Into<String>) -> ForeignError {
    ForeignError { number, description: description.into() }
}

fn bind_arguments(call: &mut HostCall, parameters: &[&str], optional: bool) -> Result<(), ForeignError> {
    if call.arguments.len() > parameters.len() {
        return Err(failure(450, "wrong number of Word arguments"));
    }
    let mut ordered = vec![None; parameters.len()];
    let mut named = false;
    for (position, argument) in std::mem::take(&mut call.arguments).into_iter().enumerate() {
        let index = if let Some(name) = &argument.name {
            named = true;
            parameters
                .iter()
                .position(|parameter| parameter.eq_ignore_ascii_case(name) || (*parameter == "Value" && name == "$value"))
                .ok_or_else(|| failure(448, "unknown named Word argument"))?
        } else if named {
            return Err(failure(5, "positional Word argument after a named argument"));
        } else {
            position
        };
        let slot = ordered.get_mut(index).ok_or_else(|| failure(450, "wrong number of Word arguments"))?;
        if slot.is_some() {
            return Err(failure(450, "duplicate Word argument"));
        }
        if !optional && matches!(argument.value, Value::Missing) {
            return Err(failure(449, "required Word argument is missing"));
        }
        *slot = Some(argument);
    }
    call.arguments = ordered
        .into_iter()
        .map(|argument| match argument {
            Some(argument) => Ok(argument),
            None if optional => Ok(HostArgument { name: None, value: Value::Missing, by_ref: false }),
            None => Err(failure(449, "required Word argument is missing")),
        })
        .collect::<Result<_, _>>()?;
    Ok(())
}

fn api_catalog() -> Vec<ApiMember> {
    [
        ("Application", "Name", "Property Get Name() As String"),
        ("Application", "Version", "Property Get Version() As String"),
        ("Application", "ActiveDocument", "Property Get ActiveDocument() As Document"),
        ("Application", "Selection", "Property Get Selection() As Selection"),
        ("Document", "Application", "Property Get Application() As Application"),
        ("Document", "Name", "Property Get Name() As String"),
        ("Document", "Content", "Property Get Content() As Range"),
        ("Document", "Range", "Function Range(Optional Start As Variant, Optional End As Variant) As Range"),
        ("Selection", "Range", "Property Get Range() As Range"),
        ("Selection", "Text", "Property Get Text() As String"),
        ("Range", "Start", "Property Get Start() As Long"),
        ("Range", "End", "Property Get End() As Long"),
        ("Range", "Text", "Property Get Text() As String; Property Let Text(ByVal value As String)"),
        ("Range", "SetRange", "Sub SetRange(ByVal Start As Long, ByVal End As Long)"),
        ("Range", "InsertAfter", "Sub InsertAfter(ByVal Text As String)"),
        ("Range", "InsertBefore", "Sub InsertBefore(ByVal Text As String)"),
    ]
    .into_iter()
    .map(|(owner, name, signature)| ApiMember {
        library: "Word".into(),
        owner: owner.into(),
        name: name.into(),
        signature: signature.into(),
        supported: true,
    })
    .collect()
}
