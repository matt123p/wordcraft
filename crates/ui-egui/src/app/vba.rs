use super::{WordApp, theme};

impl WordApp {
    #[cfg(feature = "vba")]
    pub(crate) fn forget_vba_project(&mut self, document: u64) {
        self.vba_projects.remove(&document);
        self.vba_source_saves.retain(|_, save| save.document != document);
        self.vba_close_event_pending = None;
        self.vba_save_event_pending = None;
        if self.vba_document == document {
            self.vba_project = None;
            self.vba_host_state = Default::default();
            self.vba_trusted = false;
            self.vba_document = 0;
        }
    }

    #[cfg(feature = "vba")]
    pub(crate) fn dispatch_vba_before_close(&mut self) -> Result<bool, String> {
        let document = self.session.document_id();
        if !self.vba_trusted {
            self.vba_close_event_pending = None;
            return Ok(false);
        }
        if let Some((pending_document, start_rev, notified)) = self.vba_close_event_pending
            && pending_document == document
        {
            let completed = self.vba_project.as_ref().is_some_and(wordcraft_vba::WordVbaProject::is_completed);
            if completed {
                if !notified {
                    self.vba_close_event_pending = Some((document, start_rev, true));
                    let cancelled = self.vba_project.as_mut().and_then(|project| project.take_boolean_callback_output(0)) == Some(true);
                    if cancelled {
                        self.vba_close_event_pending = None;
                        return Err("Word VBA BeforeClose cancelled closing the document.".into());
                    }
                    return Ok(self.session.rev() != start_rev);
                }
                self.vba_close_event_pending = None;
                return Ok(false);
            }
            if self.vba_project.as_ref().is_some_and(wordcraft_vba::WordVbaProject::is_reset) {
                self.vba_close_event_pending = None;
                return Ok(false);
            }
            if notified {
                self.vba_close_event_pending = None;
                return Ok(false);
            }
            self.vba_open = true;
            return Err("Word VBA BeforeClose is waiting; finish or reset the VBA procedure before closing.".into());
        }
        let handler = self.vba_project.as_ref().and_then(|project| {
            if project.has_document_event("ThisDocument", "Document_Close") {
                Some(("Document_Close", Vec::new()))
            } else if project.has_document_event("ThisDocument", "Document_BeforeClose") {
                // Retain the explicitly documented WordCraft extension. Word's cancellable
                // close event belongs to Application; ThisDocument has Document_Close().
                Some(("Document_BeforeClose", vec![wordcraft_vba::Value::Boolean(false)]))
            } else {
                None
            }
        });
        if self.vba_project.is_none() {
            return Ok(false);
        }
        let start_rev = self.session.rev();
        self.vba_open = true;
        let mut host = wordcraft_vba::WordHost::new(&mut self.session, &mut self.vba_host_state)?;
        host.set_authorized(true);
        let project = self.vba_project.as_mut().ok_or("Word VBA project closed during BeforeClose")?;
        let mode = if let Some(mode) = project.dispatch_application_event(wordcraft_vba::ApplicationEvent::DocumentBeforeClose, &mut host)? {
            mode
        } else if let Some((handler, arguments)) = handler {
            project.dispatch_document_event_with_arguments("ThisDocument", handler, arguments, &mut host)?
        } else {
            return Ok(false);
        };
        let project = self.vba_project.as_ref().ok_or("Word VBA project closed during BeforeClose")?;
        if project.is_completed() {
            self.vba_close_event_pending = Some((document, start_rev, true));
            let cancelled = self.vba_project.as_mut().and_then(|project| project.take_boolean_callback_output(0)) == Some(true);
            if cancelled {
                self.vba_close_event_pending = None;
                return Err("Word VBA BeforeClose cancelled closing the document.".into());
            }
            Ok(self.session.rev() != start_rev)
        } else if !project.is_reset() {
            self.vba_close_event_pending = Some((document, start_rev, false));
            Err(format!("Word VBA BeforeClose is waiting ({mode:?}); finish or reset the VBA procedure before closing."))
        } else {
            Err(format!("Word VBA BeforeClose ended in {mode:?}; the document remains open."))
        }
    }

    #[cfg(feature = "vba")]
    pub(crate) fn dispatch_vba_before_save(&mut self, save_as_ui: bool) -> Result<bool, String> {
        let document = self.session.document_id();
        if !self.vba_trusted {
            self.vba_save_event_pending = None;
            return Ok(false);
        }
        if let Some((pending_document, pending_save_as_ui, notified)) = self.vba_save_event_pending
            && pending_document == document
        {
            let completed = self.vba_project.as_ref().is_some_and(wordcraft_vba::WordVbaProject::is_completed);
            if completed {
                if !notified {
                    self.vba_save_event_pending = Some((document, pending_save_as_ui, true));
                    let cancelled = self.vba_project.as_mut().and_then(|project| project.take_boolean_callback_output(1)) == Some(true);
                    if cancelled {
                        self.vba_save_event_pending = None;
                    }
                    return Ok(cancelled);
                }
                self.vba_save_event_pending = None;
            } else if self.vba_project.as_ref().is_some_and(wordcraft_vba::WordVbaProject::is_reset) {
                self.vba_save_event_pending = None;
            } else {
                self.vba_open = true;
                return Err("Word VBA BeforeSave is waiting or has stopped; finish or reset the VBA procedure before saving.".into());
            }
        }
        let has_handler = self.vba_project.as_ref().is_some_and(|project| project.has_document_event("ThisDocument", "Document_BeforeSave"));
        if self.vba_project.is_none() {
            return Ok(false);
        }
        self.vba_open = true;
        let mut host = wordcraft_vba::WordHost::new(&mut self.session, &mut self.vba_host_state)?;
        host.set_authorized(true);
        let project = self.vba_project.as_mut().ok_or("Word VBA project closed during BeforeSave")?;
        let mode =
            if let Some(mode) = project.dispatch_application_event(wordcraft_vba::ApplicationEvent::DocumentBeforeSave { save_as_ui }, &mut host)? {
                mode
            } else if has_handler {
                project.dispatch_document_event_with_arguments(
                    "ThisDocument",
                    "Document_BeforeSave",
                    vec![wordcraft_vba::Value::Boolean(save_as_ui), wordcraft_vba::Value::Boolean(false)],
                    &mut host,
                )?
            } else {
                return Ok(false);
            };
        let project = self.vba_project.as_ref().ok_or("Word VBA project closed during BeforeSave")?;
        if project.is_completed() {
            self.vba_save_event_pending = Some((document, save_as_ui, true));
            let cancelled = self.vba_project.as_mut().and_then(|project| project.take_boolean_callback_output(1)) == Some(true);
            if cancelled {
                self.vba_save_event_pending = None;
            }
            Ok(cancelled)
        } else if !project.is_reset() {
            self.vba_save_event_pending = Some((document, save_as_ui, false));
            Err(format!("Word VBA BeforeSave is waiting ({mode:?}); finish or reset the VBA procedure before saving."))
        } else {
            Err(format!("Word VBA BeforeSave ended in {mode:?}; the document was not saved."))
        }
    }

    #[cfg(feature = "vba")]
    pub(crate) fn sync_vba_document(&mut self) {
        let document = self.session.document_id();
        if self.vba_document != document {
            if self.vba_document != 0
                && self.session.is_open(self.vba_document)
                && let Some(project) = self.vba_project.take()
            {
                self.vba_projects.insert(
                    self.vba_document,
                    crate::VbaDocumentProject {
                        host: std::mem::take(&mut self.vba_host_state),
                        project,
                        trusted: self.vba_trusted,
                        close_event_pending: self.vba_close_event_pending.take(),
                        save_event_pending: self.vba_save_event_pending.take(),
                    },
                );
            }
            if let Some(saved) = self.vba_projects.remove(&document) {
                self.vba_host_state = saved.host;
                self.vba_project = Some(saved.project);
                self.vba_trusted = saved.trusted;
                self.vba_close_event_pending = saved.close_event_pending;
                self.vba_save_event_pending = saved.save_event_pending;
            } else {
                self.vba_project = None;
                self.vba_host_state = Default::default();
                self.vba_trusted = false;
                self.vba_close_event_pending = None;
                self.vba_save_event_pending = None;
            }
            self.vba_document = document;
        }
        let session = &self.session;
        self.vba_projects.retain(|document, _| session.is_open(*document));
    }

    /// Coalesce selection changes while a macro is running; notify only an idle trusted project.
    #[cfg(feature = "vba")]
    pub(crate) fn notify_vba_selection(&mut self) {
        self.sync_vba_document();
        let current = (self.session.document_id(), self.session.sel.clone());
        let Some(previous) = self.vba_selection.as_ref() else {
            self.vba_selection = Some(current);
            return;
        };
        if previous == &current {
            return;
        }
        if previous.0 != current.0 || !self.vba_trusted {
            self.vba_selection = Some(current);
            return;
        }
        let Some(project) = self.vba_project.as_mut() else {
            self.vba_selection = Some(current);
            return;
        };
        if !project.is_completed() {
            return;
        }
        self.vba_selection = Some(current);
        let result = wordcraft_vba::WordHost::new(&mut self.session, &mut self.vba_host_state).and_then(|mut host| {
            host.set_authorized(true);
            project.dispatch_application_event(wordcraft_vba::ApplicationEvent::WindowSelectionChange, &mut host)
        });
        // A handler's own selection mutation belongs to this notification, not another event.
        self.vba_selection = Some((self.session.document_id(), self.session.sel.clone()));
        match result {
            Err(error) => self.status(error),
            Ok(Some(_)) if self.vba_project.as_ref().is_some_and(|project| !project.is_completed()) => self.vba_open = true,
            _ => {}
        }
    }

    /// Dispatch modeless deactivation before the document receives its raw gesture.
    #[cfg(feature = "vba")]
    pub(crate) fn vba_document_raw_input(&mut self, ctx: &egui::Context, raw: &mut egui::RawInput) {
        self.sync_vba_document();
        let Some(project) = self.vba_project.as_mut() else {
            self.vba_editor.cancel_document_input(ctx, raw);
            return;
        };
        let document = self.session.document_id();
        if self.vba_trusted && self.vba_editor.document_input_ready(project.ide()) {
            self.vba_open = false;
        }
        let visible = !self.vba_open;
        let area = (self.vba_trusted && !self.ui.backstage).then_some(self.canvas.viewport).flatten();
        let Some(command) = self.vba_editor.document_raw_input(ctx, project.ide(), document, area, visible, raw) else { return };
        let result = wordcraft_vba::WordHost::new(&mut self.session, &mut self.vba_host_state).and_then(|mut host| {
            host.set_authorized(self.vba_trusted);
            project.prepare_ui_action(&mut host)?;
            dispatch_editor_command(project, &mut host, self.vba_trusted, command)
        });
        if let Err(error) = result {
            self.vba_editor.cancel_document_input(ctx, raw);
            project.ide_mut().message(error);
        }
    }

    /// Advance document-owned VBA and render forms independently of VBE visibility.
    #[cfg(feature = "vba")]
    pub(crate) fn vba_window(&mut self, ctx: &egui::Context, tokens: theme::Tokens) {
        self.sync_vba_document();
        if !self.has_document() {
            return;
        }
        if !self.vba_open
            && self.vba_project.as_ref().is_none_or(|p| {
                !p.is_running()
                    && !p.is_waiting()
                    && !matches!(p.ide().mode(), wordcraft_vba::Mode::Break | wordcraft_vba::Mode::Failed)
                    && !p.ide().form_instances().is_ok_and(|instances| instances.into_iter().any(|f| p.ide().form_instance_visible(f)))
            })
        {
            return;
        }
        let mut host = match wordcraft_vba::WordHost::new(&mut self.session, &mut self.vba_host_state) {
            Ok(host) => host,
            Err(error) => {
                log::warn!("VBA host could not bind to the active document: {error}");
                egui::Window::new(crate::tl!("Visual Basic")).open(&mut self.vba_open).show(ctx, |ui| {
                    ui.label(format!("VBA: {error}"));
                });
                return;
            }
        };
        if self.vba_project.is_none() {
            match wordcraft_vba::WordVbaProject::new(&host) {
                Ok(project) => self.vba_project = Some(project),
                Err(error) => {
                    log::warn!("VBA project could not initialize: {error}");
                    egui::Window::new(crate::tl!("Visual Basic")).open(&mut self.vba_open).show(ctx, |ui| {
                        ui.label(format!("VBA: {error}"));
                    });
                    return;
                }
            }
        }
        let project = match self.vba_project.as_mut() {
            Some(project) => project,
            None => return,
        };
        let selection_event = project.selection_event_pending();
        let editor = &mut self.vba_editor;
        let ui_theme = vba_ui_egui::Theme {
            // The shared VBA editor's chrome goes through WordCraft's interface
            // localization; authored source, identifiers and form captions never do.
            localize: |english| std::borrow::Cow::Borrowed(crate::i18n::t(english)),
            dark: tokens.dark,
            chrome: tokens.panel,
            editor: tokens.input,
            text: tokens.text,
            muted: tokens.text_dim,
            border: tokens.border,
            accent: tokens.accent,
            ..vba_ui_egui::Theme::light()
        };
        let mut open = self.vba_open;
        let mut trusted = self.vba_trusted;
        let mut source_requests = Vec::new();
        let previous_mode = project.ide().mode();
        let mut export_package = false;
        let commands = if open {
            let mut commands = Vec::new();
            egui::Window::new(crate::tl!("Visual Basic")).open(&mut open).default_size([960.0, 680.0]).vscroll(false).show(ctx, |ui| {
                if ui.button(crate::tl!("Export Edited DOCM Copy")).clicked() {
                    export_package = true;
                }
                ui.checkbox(&mut trusted, crate::tl!("Trust this document's VBA and allow it to edit this document"));
                if !trusted {
                    ui.label(crate::tl!("VBA document access is blocked. Enable document trust above to allow macros to access this document."));
                }
                if let Some(message) = project.ide().messages.last() {
                    ui.horizontal_wrapped(|ui| {
                        ui.label(format!("VBA: {message}"));
                    });
                }
                commands = editor.show(ui, project.ide(), ui_theme);
            });
            commands
        } else {
            editor.show_forms(ctx, project.ide(), ui_theme)
        };
        host.set_authorized(trusted);
        if !trusted
            && (project.is_running() || project.is_waiting() || project.ide().mode() == wordcraft_vba::Mode::Break)
            && let Err(error) = project.reset(&mut host)
        {
            project.ide_mut().message(error);
        }
        if trusted && let Err(error) = project.prepare_ui_action(&mut host) {
            log::warn!("VBA host batch could not start: {error}");
            project.ide_mut().message(error);
            ctx.request_repaint();
            self.vba_open = open;
            self.vba_trusted = trusted;
            return;
        }
        for command in commands {
            if let Err(error) = dispatch_editor_command(project, &mut host, trusted, command) {
                log::warn!("VBA editor command failed: {error}");
                project.ide_mut().message(error.to_string());
                ctx.request_repaint();
            }
        }
        source_requests.extend(project.take_source_requests());
        if trusted && let Err(error) = project.advance(&mut host) {
            log::warn!("VBA execution stopped: {error}");
            project.ide_mut().message(error);
            ctx.request_repaint();
        }
        source_requests.extend(project.take_source_requests());
        if project.is_running() {
            ctx.request_repaint();
        }
        if project.ide().mode() != previous_mode && matches!(project.ide().mode(), wordcraft_vba::Mode::Break | wordcraft_vba::Mode::Failed) {
            open = true;
        }
        if selection_event && !project.selection_event_pending() {
            self.vba_selection = Some((self.session.document_id(), self.session.sel.clone()));
        }
        self.vba_open = open;
        self.vba_trusted = trusted;
        for request in source_requests {
            self.route_vba_source_request(request);
        }
        if export_package && let Err(error) = self.run("ui.vbaExportNativePackage", serde_json::json!({})) {
            self.status(error);
        }
    }

    #[cfg(feature = "vba")]
    #[cfg(not(target_arch = "wasm32"))]
    pub(crate) fn route_vba_source_request(&mut self, request: wordcraft_vba::SourceRequest) {
        use crate::file_dialogs::{AfterPick, FileDialogRequest};
        let document = self.session.document_id();
        let request = match request {
            wordcraft_vba::SourceRequest::NativeForm { snapshot, save } => {
                let project = self.vba_project.as_ref().ok_or("native save project closed").map(|project| project.instance());
                let Ok(project) = project else {
                    self.status("Native save project closed");
                    return;
                };
                let Some(file) = snapshot.files().first() else {
                    return;
                };
                let name = file.name.clone();
                let after = AfterPick::VbaNativeWrite { document, project, snapshot, save };
                if matches!(self.ask_file(FileDialogRequest::Save { name, dir: None }, after), crate::file_dialogs::Asked::Busy) {
                    self.status("Finish the open file dialog before saving the native form");
                }
                return;
            }
            wordcraft_vba::SourceRequest::RenameDocument { .. } => {
                self.status("Word document code names are host-owned; renaming ThisDocument is unsupported.");
                return;
            }
            wordcraft_vba::SourceRequest::Import => {
                let after = AfterPick::VbaImport { document };
                if matches!(self.ask_file(FileDialogRequest::Open { purpose: "vba_source".into() }, after), crate::file_dialogs::Asked::Busy) {
                    self.status("Finish the open file dialog before importing VBA source.");
                }
                return;
            }
            wordcraft_vba::SourceRequest::Save { id, name, text, revision } => {
                let source = Some((id, revision));
                (FileDialogRequest::Save { name: name.clone(), dir: None }, AfterPick::VbaWrite { document, name, text, source })
            }
            wordcraft_vba::SourceRequest::Export { name, text } => {
                (FileDialogRequest::Save { name: name.clone(), dir: None }, AfterPick::VbaWrite { document, name, text, source: None })
            }
            wordcraft_vba::SourceRequest::ReturnToHost => {
                self.vba_open = false;
                return;
            }
            wordcraft_vba::SourceRequest::ConfirmClose => {
                self.status("Save or discard modified VBA source before closing the project.");
                return;
            }
        };
        let (picker, after) = request;
        if matches!(self.ask_file(picker, after), crate::file_dialogs::Asked::Busy) {
            self.status("Finish the open file dialog before working with VBA source.");
        }
    }

    #[cfg(feature = "vba")]
    #[cfg(target_arch = "wasm32")]
    pub(crate) fn route_vba_source_request(&mut self, request: wordcraft_vba::SourceRequest) {
        match request {
            wordcraft_vba::SourceRequest::NativeForm { snapshot, .. } => {
                if let Some(download) = &self.services.download {
                    for file in snapshot.files() {
                        if let Err(error) = download(&file.name, &file.bytes) {
                            self.status(format!("Native download failed: {error}"));
                            return;
                        }
                    }
                    self.status("Native files were offered for download; source remains modified until committed.");
                } else {
                    self.status("Native form downloads are unavailable");
                }
            }
            wordcraft_vba::SourceRequest::RenameDocument { .. } => {
                self.status("Word document code names are host-owned; renaming ThisDocument is unsupported.");
                return;
            }
            wordcraft_vba::SourceRequest::Import => {
                if let Some(open) = &self.services.open_async {
                    open("vba_source");
                    self.vba_import_pending = Some(self.session.document_id());
                } else {
                    self.status("VBA source picking is unavailable in this browser host.");
                }
            }
            wordcraft_vba::SourceRequest::Save { id, name, text, revision } => self.download_vba_source(&name, &text, Some((id, revision))),
            wordcraft_vba::SourceRequest::Export { name, text } => self.download_vba_source(&name, &text, None),
            wordcraft_vba::SourceRequest::ReturnToHost => self.vba_open = false,
            wordcraft_vba::SourceRequest::ConfirmClose => {
                self.status("Save or discard modified VBA source before closing the project.");
            }
        }
    }

    #[cfg(feature = "vba")]
    #[cfg(any(target_arch = "wasm32", test))]
    pub(crate) fn download_vba_source(&mut self, name: &str, text: &str, source: Option<(wordcraft_vba::SourceId, u64)>) {
        if self.vba_source_saves.len() >= 64 {
            self.status("Finish a pending VBA source save before starting another.");
            return;
        }
        let Some(project) = self.vba_project.as_ref() else {
            return;
        };
        let save = crate::VbaSourceSave { document: self.session.document_id(), project: project.instance(), source, name: name.to_owned() };
        let saves = self.services.async_saves.clone();
        if let Some(saves) = &saves {
            let _ = saves.take_started();
        }
        let result = self
            .services
            .download
            .as_ref()
            .ok_or_else(|| "source downloads are unavailable".to_owned())
            .and_then(|download| download(name, text.as_bytes()));
        match result {
            Ok(()) => {
                if let Some(ticket) = saves.as_ref().and_then(crate::async_saves::AsyncSaves::take_started) {
                    self.vba_source_saves.insert(ticket, save);
                    self.status(format!("Saving VBA source {name}"));
                } else {
                    self.finish_vba_source_save(save, &crate::async_saves::SaveOutcome::Saved);
                }
            }
            Err(error) => self.status(format!("Could not download VBA source: {error}")),
        }
    }
    #[cfg(feature = "vba")]
    pub(crate) fn settle_vba_source_save(&mut self, ticket: u64, outcome: &crate::async_saves::SaveOutcome) -> bool {
        let Some(save) = self.vba_source_saves.remove(&ticket) else {
            return false;
        };
        self.finish_vba_source_save(save, outcome);
        true
    }

    #[cfg(feature = "vba")]
    fn finish_vba_source_save(&mut self, save: crate::VbaSourceSave, outcome: &crate::async_saves::SaveOutcome) {
        use crate::async_saves::SaveOutcome;
        match outcome {
            SaveOutcome::Cancelled => self.status(format!("VBA source save cancelled: {}", save.name)),
            SaveOutcome::Failed(error) => self.status(format!("Could not save VBA source {}: {error}", save.name)),
            SaveOutcome::Saved => {
                if let Some((id, revision)) = save.source {
                    let project = if self.vba_document == save.document {
                        self.vba_project.as_mut()
                    } else {
                        self.vba_projects.get_mut(&save.document).map(|saved| &mut saved.project)
                    };
                    let result = project
                        .filter(|project| project.instance() == save.project)
                        .ok_or_else(|| "the originating project has closed or changed".to_owned())
                        .and_then(|project| project.mark_source_saved(id, revision));
                    if let Err(error) = result {
                        self.status(format!("Source was saved but remains modified: {error}"));
                        return;
                    }
                }
                self.status(format!("Saved VBA source {}", save.name));
            }
        }
    }
}

/// Keep source/compile/design/export commands available while document execution is untrusted.
#[cfg(feature = "vba")]
fn dispatch_editor_command(
    project: &mut wordcraft_vba::WordVbaProject,
    host: &mut wordcraft_vba::WordHost<'_>,
    trusted: bool,
    command: vba_ui_egui::Command,
) -> Result<(), String> {
    if !trusted && command.executes_code() {
        return Err("Word VBA requires explicit document authorization".into());
    }
    vba_ui_egui::apply(project.ide_mut(), command, host).map_err(|error| error.to_string())
}

#[cfg(all(test, feature = "vba"))]
mod native_form_authorization_tests {
    use super::dispatch_editor_command;
    #[test]
    fn instance_input_cannot_execute_pure_vba_after_document_trust_is_revoked() -> Result<(), Box<dyn std::error::Error>> {
        let mut session = wordcraft_engine::Session::new(wordcraft_doc::Document::from_text("unchanged"));
        let mut state = wordcraft_vba::WordHostState::default();
        let mut host = wordcraft_vba::WordHost::new(&mut session, &mut state)?;
        let mut project = wordcraft_vba::WordVbaProject::new(&host)?;
        let id = project.ide_mut().add_form("GuardForm")?;
        project
            .ide_mut()
            .set_text("Private Sub UserForm_QueryClose(Cancel As Integer,CloseMode As Integer)\nDebug.Print \"executed\"\nEnd Sub".into())?;
        let run = || vba_ui_egui::Command::Form(vba_ui_egui::FormCommand::Run(id, false));
        assert!(dispatch_editor_command(&mut project, &mut host, false, run()).is_err());
        assert!(project.ide().output().is_empty());
        host.set_authorized(true);
        project.prepare_ui_action(&mut host)?;
        dispatch_editor_command(&mut project, &mut host, true, run())?;
        for _ in 0..100 {
            project.advance(&mut host)?;
            if project.ide().mode() == wordcraft_vba::Mode::Completed {
                break;
            }
        }
        let instance = *project.ide().form_instances()?.first().ok_or("instance")?;
        host.set_authorized(false);
        assert!(
            dispatch_editor_command(
                &mut project,
                &mut host,
                false,
                vba_ui_egui::Command::Form(vba_ui_egui::FormCommand::InstanceEvent(instance, None, None))
            )
            .is_err()
        );
        let mouse = vba_ui_egui::Command::Form(vba_ui_egui::FormCommand::InstanceMouse(
            instance,
            None,
            wordcraft_vba::FormMouseInput { event: wordcraft_vba::FormMouseEvent::Down, button: 1, shift: 0, x: 0., y: 0., index: -1 },
        ));
        let refused = dispatch_editor_command(&mut project, &mut host, false, mouse);
        assert!(refused.is_err_and(|message| message.contains("requires explicit document authorization")));
        assert!(project.ide().output().is_empty());
        assert!(project.ide().form_instance_visible(instance));
        Ok(())
    }
}

#[cfg(feature = "vba")]
impl WordApp {
    pub(crate) fn export_native_package_copy(&mut self) -> Result<(), String> {
        let project = self.vba_project.as_ref().ok_or("VBA project is closed")?;
        let host = wordcraft_vba::WordHost::new(&mut self.session, &mut self.vba_host_state)?;
        let snapshot = project.prepare_native_project_export(&host)?;
        let bytes =
            wordcraft_engine::io::encode_edited_vba_package_copy(snapshot.document(), snapshot.original_project(), snapshot.encoded_project())?;
        project.validate_native_project_export(&host, &snapshot)?;
        #[cfg(not(target_arch = "wasm32"))]
        {
            let document = self.session.document_id();
            let project = project.instance();
            let request = crate::file_dialogs::FileDialogRequest::Save { name: "EditedDocument.docm".into(), dir: None };
            let after = crate::file_dialogs::AfterPick::VbaPackageWrite { document, project, snapshot: Box::new(snapshot), bytes };
            match self.ask_file(request, after) {
                crate::file_dialogs::Asked::Busy => Err("Finish the open file dialog before exporting the native document".into()),
                crate::file_dialogs::Asked::Done(result) => result.map(|_| ()),
                crate::file_dialogs::Asked::Pending => Ok(()),
            }
        }
        #[cfg(target_arch = "wasm32")]
        {
            let download = self.services.download.as_ref().ok_or("Document downloads are unavailable")?;
            download("EditedDocument.docm", &bytes)?;
            self.status("Edited DOCM copy offered for download");
            Ok(())
        }
    }
}
