use super::{WordApp, theme};

impl WordApp {
    #[cfg(feature = "vba")]
    pub(crate) fn forget_vba_project(&mut self, document: u64) {
        self.vba_projects.remove(&document);
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
        let Some((handler, arguments)) = handler else {
            return Ok(false);
        };
        let start_rev = self.session.rev();
        self.vba_open = true;
        let mut host = wordcraft_vba::WordHost::new(&mut self.session, &mut self.vba_host_state)?;
        host.set_authorized(true);
        let mode = self.vba_project.as_mut().ok_or("Word VBA project closed during BeforeClose")?.dispatch_document_event_with_arguments(
            "ThisDocument",
            handler,
            arguments,
            &mut host,
        )?;
        let project = self.vba_project.as_ref().ok_or("Word VBA project closed during BeforeClose")?;
        if project.is_completed() {
            self.vba_close_event_pending = Some((document, start_rev, true));
            let cancelled = self.vba_project.as_mut().and_then(|project| project.take_boolean_callback_output(0)) == Some(true);
            if cancelled {
                self.vba_close_event_pending = None;
                return Err("Word VBA BeforeClose cancelled closing the document.".into());
            }
            Ok(self.session.rev() != start_rev)
        } else if project.is_waiting() {
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
        if !has_handler {
            return Ok(false);
        }
        self.vba_open = true;
        let mut host = wordcraft_vba::WordHost::new(&mut self.session, &mut self.vba_host_state)?;
        host.set_authorized(true);
        let mode = self.vba_project.as_mut().ok_or("Word VBA project closed during BeforeSave")?.dispatch_document_event_with_arguments(
            "ThisDocument",
            "Document_BeforeSave",
            vec![wordcraft_vba::Value::Boolean(save_as_ui), wordcraft_vba::Value::Boolean(false)],
            &mut host,
        )?;
        let project = self.vba_project.as_ref().ok_or("Word VBA project closed during BeforeSave")?;
        if project.is_completed() {
            self.vba_save_event_pending = Some((document, save_as_ui, true));
            let cancelled = self.vba_project.as_mut().and_then(|project| project.take_boolean_callback_output(1)) == Some(true);
            if cancelled {
                self.vba_save_event_pending = None;
            }
            Ok(cancelled)
        } else if project.is_waiting() || project.is_running() {
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

    /// The shared VBA editor's floating window, when it is open and a document is showing.
    #[cfg(feature = "vba")]
    pub(crate) fn vba_window(&mut self, ctx: &egui::Context, tokens: theme::Tokens) {
        self.sync_vba_document();
        if !self.vba_open || !self.has_document() {
            return;
        }
        let mut host = match wordcraft_vba::WordHost::new(&mut self.session, &mut self.vba_host_state) {
            Ok(host) => host,
            Err(error) => {
                log::warn!("VBA host could not bind to the active document: {error}");
                return;
            }
        };
        if self.vba_project.is_none() {
            match wordcraft_vba::WordVbaProject::new(&host) {
                Ok(project) => self.vba_project = Some(project),
                Err(error) => {
                    log::warn!("VBA project could not initialize: {error}");
                    return;
                }
            }
        }
        let project = match self.vba_project.as_mut() {
            Some(project) => project,
            None => return,
        };
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
        egui::Window::new(crate::tl!("Visual Basic")).open(&mut open).default_size([960.0, 680.0]).vscroll(false).show(ctx, |ui| {
            ui.checkbox(&mut trusted, crate::tl!("Trust this document's VBA and allow it to edit this document"));
            host.set_authorized(trusted);
            let commands = editor.show(ui, project.ide(), ui_theme);
            for command in commands {
                if let Err(error) = vba_ui_egui::apply(project.ide_mut(), command, &mut host) {
                    log::warn!("VBA editor command failed: {error}");
                }
            }
            source_requests.extend(project.take_source_requests());
            if let Err(error) = project.advance(&mut host) {
                log::warn!("VBA execution stopped: {error}");
            }
            source_requests.extend(project.take_source_requests());
            if project.is_running() {
                ui.ctx().request_repaint();
            }
        });
        self.vba_open = open;
        self.vba_trusted = trusted;
        for request in source_requests {
            self.route_vba_source_request(request);
        }
    }

    #[cfg(feature = "vba")]
    #[cfg(not(target_arch = "wasm32"))]
    pub(crate) fn route_vba_source_request(&mut self, request: wordcraft_vba::SourceRequest) {
        use crate::file_dialogs::{AfterPick, FileDialogRequest};
        let document = self.session.document_id();
        let request = match request {
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
    #[cfg(target_arch = "wasm32")]
    fn download_vba_source(&mut self, name: &str, text: &str, source: Option<(wordcraft_vba::SourceId, u64)>) {
        let result = self
            .services
            .download
            .as_ref()
            .ok_or_else(|| "source downloads are unavailable".to_owned())
            .and_then(|download| download(name, text.as_bytes()));
        match result {
            Ok(()) => {
                if let Some((id, revision)) = source
                    && let Some(project) = self.vba_project.as_mut()
                    && let Err(error) = project.mark_source_saved(id, revision)
                {
                    self.status(format!("Source was downloaded but remains modified: {error}"));
                    return;
                }
                self.status(format!("Downloaded VBA source {name}"));
            }
            Err(error) => self.status(format!("Could not download VBA source: {error}")),
        }
    }
}
