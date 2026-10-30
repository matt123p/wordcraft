use vba_foreign::{Host, HostCall, InvokeKind};
use vba_ide::Mode;
use wordcraft_doc::{Document, StoryRef};
use wordcraft_engine::Session;
use wordcraft_vba::{WordHost, WordHostState, WordVbaProject};

#[test]
fn authored_macro_mutates_the_live_document_as_one_undo_step() {
    let mut session = Session::new(Document::from_text(""));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project
        .import_source(
            "Module1.bas",
            "Attribute VB_Name = \"Module1\"\nPublic Sub Main()\n    ActiveDocument.Content.Text = \"Hello\"\n    Selection.Range.InsertAfter \" 😀\"\nEnd Sub",
        )
        .expect("import authored source");

    let mode = project.run("Main", &mut host).expect("macro run");
    assert_eq!(mode, Mode::Completed);
    assert_eq!(session.doc.plain_text(StoryRef::Body), "Hello 😀");
    assert_eq!(session.undo_depth(), 1, "the macro's writes form one undo step");
    assert!(session.undo());
    assert_eq!(session.doc.plain_text(StoryRef::Body), "");
}

#[test]
fn typed_word_application_document_selection_and_range_mutate_the_live_document() {
    let mut session = Session::new(Document::from_text(""));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project.import_source("Module1.bas", "Public Sub Main()\nDim app As Word.Application\nDim doc As Word.Document\nDim selected As Word.Selection\nDim content As Word.Range\nSet app = Application\nSet doc = app.ActiveDocument\nSet selected = app.Selection\nSet content = doc.Content\ncontent.Text = \"typed Word\"\nselected.Range.InsertAfter \"!\"\nEnd Sub").expect("typed source");
    assert_eq!(project.run("Main", &mut host).expect("typed run"), Mode::Completed);
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "typed Word!");
    assert_eq!(host.session_mut().undo_depth(), 1);
}

#[test]
fn word_nominal_type_mismatch_reports_vba_error_13() {
    let mut session = Session::new(Document::from_text(""));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project.import_source("Module1.bas", "Public Sub Main()\nOn Error Resume Next\nDim content As Word.Range\nSet content = ActiveDocument\nActiveDocument.Content.Text = CStr(Err.Number)\nEnd Sub").expect("typed mismatch source");
    assert_eq!(project.run("Main", &mut host).expect("handled mismatch"), Mode::Completed);
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "13");
}

#[test]
fn named_word_range_arguments_bind_by_parameter_name() {
    let mut session = Session::new(Document::from_text("abcdef"));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project.import_source("Module1.bas", "Public Sub Main()\nDim content As Word.Range\nSet content = ActiveDocument.Content\ncontent.SetRange eNd:=4, START:=1\ncontent.Text = \"!\"\nEnd Sub").expect("named source");
    assert_eq!(project.run("Main", &mut host).expect("named run"), Mode::Completed);
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "a!ef");
}

#[test]
fn document_range_optional_and_named_positions_select_the_requested_text() {
    for (expression, expected) in [
        ("ActiveDocument.Range()", "!"),
        ("ActiveDocument.Range(1,)", "a!"),
        ("ActiveDocument.Range(,3)", "!def"),
        ("ActiveDocument.Range(End:=3)", "!def"),
        ("ActiveDocument.Range(End:=4, Start:=1)", "a!ef"),
    ] {
        let mut session = Session::new(Document::from_text("abcdef"));
        let mut state = WordHostState::default();
        let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
        host.set_authorized(true);
        let mut project = WordVbaProject::new(&host).expect("project");
        project
            .import_source(
                "Module1.bas",
                &format!("Public Sub Main()\nDim content As Word.Range\nSet content = {expression}\ncontent.Text = \"!\"\nEnd Sub"),
            )
            .expect("range source");
        assert_eq!(project.run("Main", &mut host).expect("range run"), Mode::Completed, "{expression}");
        assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), expected, "{expression}");
    }
}

#[test]
fn insertion_expands_the_captured_range_to_include_its_original_text() {
    for (range, method, expected) in [
        ("ActiveDocument.Range(1, 3)", "InsertBefore", "aXd"),
        ("ActiveDocument.Range(1, 3)", "InsertAfter", "aXd"),
        ("ActiveDocument.Content", "InsertAfter", "X"),
    ] {
        let mut session = Session::new(Document::from_text("abcd"));
        let mut state = WordHostState::default();
        let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
        host.set_authorized(true);
        let mut project = WordVbaProject::new(&host).expect("project");
        project
            .import_source(
                "Module1.bas",
                &format!(
                    "Public Sub Main()\nDim content As Word.Range\nSet content = {range}\ncontent.{method} \"!\"\ncontent.Text = \"X\"\nEnd Sub"
                ),
            )
            .expect("live range source");
        assert_eq!(project.run("Main", &mut host).expect("live range run"), Mode::Completed);
        assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), expected, "{range}.{method} retains the original range extent");
        assert_eq!(host.session_mut().undo_depth(), 1);
    }
}

#[test]
fn invalid_word_arguments_report_exact_errors_without_mutation() {
    for (invocation, error, initial) in [
        ("content.SetRange Start:=0", 449, "keep"),
        ("content.SetRange Unknown:=0, End:=1", 448, "keep"),
        ("content.InsertAfter \"bad\", \"extra\"", 450, "keep"),
        ("content.SetRange -1, 1", 5, "keep"),
        ("content.InsertAfter 42", 13, "keep"),
        ("content.SetRange 1, 2", 5, "😀"),
        ("Set content = ActiveDocument.Range(1, 2)", 5, "😀"),
    ] {
        let mut session = Session::new(Document::from_text(initial));
        let mut state = WordHostState::default();
        let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
        host.set_authorized(true);
        let mut project = WordVbaProject::new(&host).expect("project");
        project.import_source("Module1.bas", &format!("Public Sub Main()\nDim content As Word.Range\nSet content = ActiveDocument.Content\nOn Error Resume Next\n{invocation}\nIf Err.Number <> {error} Then ActiveDocument.Content.Text = \"wrong error\"\nEnd Sub")).expect("negative source");
        assert_eq!(project.run("Main", &mut host).expect("handled argument error"), Mode::Completed, "{invocation}");
        assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), initial, "{invocation}: exact error {error} with no partial write");
        assert_eq!(host.session_mut().undo_depth(), 0, "argument refusal is side-effect-free");
    }
}

#[test]
fn completed_project_can_run_again_after_an_ordinary_document_edit() {
    let mut session = Session::new(Document::from_text("initial"));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project.import_source("Module1.bas", "Public Sub Main()\nActiveDocument.Content.Text = \"macro\"\nEnd Sub").expect("source");
    assert_eq!(project.run("Main", &mut host).expect("first run"), Mode::Completed);
    let document = host.roots().get("Document").copied().expect("document root");
    let vba_foreign::Reply::Complete(vba_foreign::Completion { value: vba_values::Value::Object(captured), .. }) =
        host.invoke(HostCall { kind: InvokeKind::Get, target: document, member: "Content".into(), arguments: Vec::new() })
    else {
        panic!("capture a live range");
    };
    host.session_mut().run("text.insert", &serde_json::json!({"text": " user edit", "raw": true})).expect("ordinary edit");
    assert_eq!(project.run("Main", &mut host).expect("second run"), Mode::Completed);
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "macro");
    let stale = host.invoke(HostCall { kind: InvokeKind::Get, target: captured, member: "Text".into(), arguments: Vec::new() });
    assert!(
        matches!(stale, vba_foreign::Reply::Error(error) if error.number == 91),
        "a fresh run must not revive a range captured before the user edit"
    );
}

#[test]
fn host_handles_refuse_after_an_external_document_edit() {
    let mut session = Session::new(Document::from_text("before"));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let application = host.roots().get("Application").copied().expect("application root");

    host.session_mut().run("text.insert", &serde_json::json!({"text":" external", "raw":true})).expect("external edit");
    let reply = host.invoke(HostCall { kind: InvokeKind::Get, target: application, member: "Name".into(), arguments: Vec::new() });

    assert!(matches!(reply, vba_foreign::Reply::Error(ref error) if error.number == 91));
}

#[test]
fn host_identity_and_root_handles_survive_ui_frame_recreation() {
    let mut session = Session::new(Document::from_text("persistent"));
    let mut state = WordHostState::default();
    let application = {
        let mut host = WordHost::new(&mut session, &mut state).expect("first frame host");
        host.set_authorized(true);
        host.roots().get("Application").copied().expect("application")
    };
    let mut host = WordHost::new(&mut session, &mut state).expect("next frame host");
    host.set_authorized(true);
    let reply = host.invoke(HostCall { kind: InvokeKind::Get, target: application, member: "Name".into(), arguments: Vec::new() });
    assert!(matches!(reply, vba_foreign::Reply::Complete(_)));
}

#[test]
fn document_host_calls_require_explicit_trust() {
    let mut session = Session::new(Document::from_text("trusted only"));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    let application = host.roots().get("Application").copied().expect("application");
    let reply = host.invoke(HostCall { kind: InvokeKind::Get, target: application, member: "Name".into(), arguments: Vec::new() });
    assert!(matches!(reply, vba_foreign::Reply::Error(ref error) if error.number == 70));
}

#[test]
fn imported_word_event_runs_only_when_host_dispatches_it() {
    let mut session = Session::new(Document::from_text(""));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project
        .import_source(
            "ThisDocument.cls",
            "Attribute VB_Name = \"ThisDocument\"\nPrivate Sub Document_Open()\n    Me.Content.Text = \"opened\"\nEnd Sub",
        )
        .expect("document module");
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "");
    let mode = project.dispatch_document_event("ThisDocument", "Document_Open", &mut host).expect("dispatch event");
    assert_eq!(mode, Mode::Completed);
    assert_eq!(session.doc.plain_text(StoryRef::Body), "opened");
}

#[test]
fn document_module_state_survives_successive_lifecycle_callbacks() {
    let mut session = Session::new(Document::from_text(""));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project
        .import_source(
            "ThisDocument.cls",
            "Attribute VB_Name = \"ThisDocument\"\nPrivate attempts As Integer\nPrivate Sub Document_BeforeClose(Cancel As Boolean)\n    attempts = attempts + 1\n    If attempts = 1 Then\n        ActiveDocument.Content.Text = \"first\"\n    Else\n        ActiveDocument.Content.Text = \"second\"\n    End If\nEnd Sub",
        )
        .expect("document module");

    assert_eq!(
        project
            .dispatch_document_event_with_arguments("ThisDocument", "Document_BeforeClose", vec![vba_values::Value::Boolean(false)], &mut host)
            .expect("first close event"),
        Mode::Completed
    );
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "first");
    assert_eq!(
        project
            .dispatch_document_event_with_arguments("ThisDocument", "Document_BeforeClose", vec![vba_values::Value::Boolean(false)], &mut host)
            .expect("second close event"),
        Mode::Completed
    );
    assert_eq!(session.doc.plain_text(StoryRef::Body), "second", "the ThisDocument instance preserves module state between callbacks");
}

#[test]
fn reset_and_close_leave_the_document_host_usable() {
    let mut session = Session::new(Document::from_text("available"));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project.import_source("Module1.bas", "Public Sub Main()\nEnd Sub").expect("source");
    assert_eq!(project.run("Main", &mut host).expect("run"), Mode::Completed);
    project.reset(&mut host).expect("reset");
    project.close(&mut host, true).expect("close");

    let application = host.roots().get("Application").copied().expect("application");
    let reply = host.invoke(HostCall { kind: InvokeKind::Get, target: application, member: "Name".into(), arguments: Vec::new() });
    assert!(matches!(reply, vba_foreign::Reply::Complete(_)));
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "available");
}

#[test]
fn modal_shared_form_waits_resumes_and_mutates_the_word_document() {
    use vba_ide::forms::{Bounds, Control, ControlKind, Form};

    let form = Form {
        name: "EditorForm".into(),
        caption: "Editor Form".into(),
        width: 300.0,
        height: 130.0,
        controls: vec![
            Control {
                id: 1,
                name: "TextBox1".into(),
                kind: ControlKind::TextBox,
                bounds: Bounds { left: 10.0, top: 10.0, width: 160.0, height: 24.0 },
                caption: String::new(),
                text: String::new(),
                enabled: true,
                visible: true,
                tab_stop: true,
                group: None,
            },
            Control {
                id: 2,
                name: "CommandButton1".into(),
                kind: ControlKind::CommandButton,
                bounds: Bounds { left: 10.0, top: 50.0, width: 100.0, height: 24.0 },
                caption: "Done".into(),
                text: String::new(),
                enabled: true,
                visible: true,
                tab_stop: true,
                group: None,
            },
        ],
        tab_order: vec![1, 2],
    };
    let archive = vba_ide::form_archive::encode(
        &form,
        "Private Sub TextBox1_Change()\nActiveDocument.Content.Text = TextBox1.Text\nEnd Sub\nPrivate Sub CommandButton1_Click()\nMe.Hide\nEnd Sub",
    )
    .expect("form archive");
    let mut session = Session::new(Document::from_text(""));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project.import_source("EditorForm.tvform", &archive).expect("form import");
    let form_id =
        project.ide().modules.iter().find(|module| project.ide().form(module.id).is_some()).map(|module| module.id).expect("imported form module");
    project.ide_mut().run_form(form_id, true, &mut host).expect("show modal form");
    assert_eq!(project.advance(&mut host).expect("start form"), Mode::Waiting);
    assert!(project.ide().form_visible(form_id));

    project.ide_mut().form_event(form_id, Some(1), Some("typed in form".into()), &mut host).expect("text event");
    assert_eq!(project.advance(&mut host).expect("resume text event"), Mode::Waiting);
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "typed in form");

    project.ide_mut().form_event(form_id, Some(2), None, &mut host).expect("button event");
    assert_eq!(project.advance(&mut host).expect("resume button event"), Mode::Completed);
    assert!(!project.ide().form_visible(form_id));
}

#[test]
fn failed_macro_keeps_prior_mutations_undoable_as_one_batch() {
    let mut session = Session::new(Document::from_text(""));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project
        .import_source("Module1.bas", "Public Sub Main()\n    ActiveDocument.Content.Text = \"Kept\"\n    Selection.Range.NotInProfile\nEnd Sub")
        .expect("source");

    let mode = project.run("Main", &mut host).expect("macro dispatch");
    assert_eq!(mode, Mode::Failed);
    assert_eq!(session.doc.plain_text(StoryRef::Body), "Kept");
    assert_eq!(session.undo_depth(), 1);
    assert!(session.undo());
    assert_eq!(session.doc.plain_text(StoryRef::Body), "");
}

#[test]
fn paused_macro_refuses_to_resume_after_an_external_edit() {
    let mut session = Session::new(Document::from_text("start"));
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).expect("Word host");
    host.set_authorized(true);
    let mut project = WordVbaProject::new(&host).expect("project");
    project
        .import_source(
            "Module1.bas",
            "Public Sub Main()\n    ActiveDocument.Content.Text = \"macro\"\n    Selection.Range.InsertAfter \" later\"\nEnd Sub",
        )
        .expect("source");

    assert_eq!(project.begin("Main", &mut host).expect("start"), Mode::Running);
    assert_eq!(project.pause(&mut host).expect("pause"), Mode::Break);
    host.session_mut().run("text.insert", &serde_json::json!({"text":" user", "raw":true})).expect("user edit");
    project.continue_execution(&mut host).expect("continue");

    assert_eq!(project.advance(&mut host).expect("resume tick"), Mode::Failed);
    assert_eq!(session.doc.plain_text(StoryRef::Body), " userstart");
}
