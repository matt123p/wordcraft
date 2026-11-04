//! Authored native form attachment; source/resources never grant execution permission.
#[path = "support/native_forms.rs"]
mod fixture;
mod support;
use std::sync::Arc;
use vba_ide::{Action, Mode};
use wordcraft_doc::Document;
use wordcraft_docx::{Flavor, read, write_as};
use wordcraft_engine::Session;
use wordcraft_vba::{WordHost, WordHostState, WordVbaProject};
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_forms_attach_from_docm_inertly_run_after_trust_and_retain_original_resources() {
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", fixture::SOURCE),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nEnd Sub"),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let package = write_as(&doc, Flavor::MacroDocument).unwrap();
    let mut session = Session::new(read(&package).unwrap());
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    let id = project.ide().modules.iter().find(|m| m.name == "NativeForm.frm").unwrap().id;
    assert!(project.ide().form(id).is_some());
    assert!(project.ide().output().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Waiting, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["initialized".to_string()]);
    project.ide_mut().form_event(id, Some(13), None, &mut host).unwrap();
    for _ in 0..500 {
        project.ide_mut().tick(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(project.ide().mode(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["initialized".to_string(), "initial".to_string()]);
    project.ide_mut().dispatch(Action::Reset, &mut host).unwrap();
    assert_eq!(project.ide().modules.iter().find(|m| m.id == id).unwrap().text, fixture::SOURCE);
    let output = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&output).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

fn boolean_project(project: Vec<u8>) -> Vec<u8> {
    form_project(project, fixture::boolean_companion())
}
fn form_project(project: Vec<u8>, resource: Vec<u8>) -> Vec<u8> {
    use std::io::{Cursor, Read, Write};
    let mut file = cfb::CompoundFile::open(Cursor::new(fixture::add_designer(project))).unwrap();
    let mut resource = cfb::CompoundFile::open(Cursor::new(resource[4..].to_vec())).unwrap();
    for name in ["f", "o"] {
        let mut bytes = Vec::new();
        resource.open_stream(format!("/{name}")).unwrap().read_to_end(&mut bytes).unwrap();
        file.create_stream(format!("/NativeForm/{name}")).unwrap().write_all(&bytes).unwrap();
    }
    file.into_inner().into_inner()
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_boolean_instances_run_through_the_word_bridge_and_keep_package_bytes() {
    let original = boolean_project(support::project(&[
        ("NativeForm", "BaseClass", fixture::SOURCE),
        (
            "Launch",
            "Module",
            "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,c As MSForms.Control,u As MSForms.UserForm,n As Long\nDebug.Print VBA.UserForms.Count\nf.Check1.Value=True\nf.ChoiceA.Value=True\nf.ChoiceB.Value=True\nDebug.Print f.ChoiceA.Value\nDebug.Print f.ChoiceB.Value\nf.Show 0\nDebug.Print f.ActiveControl Is f.Entry\nFor Each c In f.Controls\nn=n+1\nNext c\nDebug.Print n\nn=0\nFor Each u In VBA.UserForms\nIf u Is f Then n=n+1\nNext u\nDebug.Print n\nEnd Sub",
        ),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().output().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["0", "initialized", "False", "True", "True", "7", "1"]);
    let instance = *project.ide().form_instances().unwrap().first().unwrap();
    assert_eq!(project.ide().form_instance_field(instance, Some(14), "Value").unwrap(), wordcraft_vba::Value::Boolean(true));
    project.prepare_ui_action(&mut host).unwrap();
    project.ide_mut().form_instance_event(instance, Some(14), None, &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_field(instance, Some(14), "Value").unwrap(), wordcraft_vba::Value::Boolean(false));
    assert_eq!(
        project.ide().form_instance_field(instance, None, "VbaActiveControl").unwrap(),
        project.ide().form_instance_field(instance, None, "VbaCheck1").unwrap()
    );
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_text_selection_and_input_preserve_docm_project_bytes() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,t As MSForms.TextBox\nSet t=f.Entry\nt.Text=\"abcde\"\nt.SelStart=1\nt.SelLength=2\nt.SelText=\"X\"\nf.Show 0\nDebug.Print t.MultiLine\nDebug.Print t.PasswordChar\nDebug.Print t.Text\nDebug.Print t.SelStart\nEnd Sub",
            ),
        ]),
        fixture::text_companion(0xacf0481b, 42),
    );
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().output().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed);
    assert_eq!(project.ide().output(), vec!["initialized", "True", "*", "aXde", "2"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    project.prepare_ui_action(&mut host).unwrap();
    project.ide_mut().form_instance_text(f, 12, vba_values::VbaString::from_utf8("z").unwrap(), (1, 0), &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_field(f, Some(12), "VbaSelStart").unwrap(), wordcraft_vba::Value::Long(1));
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_numeric_controls_keep_docm_identity_and_canonical_input() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,s As MSForms.SpinButton\nSet s=f.Spin1\ns.Value=7\nf.Show 0\nDebug.Print s.Value\nDebug.Print f.Scroll1.Min\nEnd Sub",
            ),
        ]),
        fixture::numeric_companion(),
    );
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().output().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed);
    assert_eq!(project.ide().output(), vec!["initialized", "7", "10"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    for (id, large) in [(14, false), (15, true)] {
        project.prepare_ui_action(&mut host).unwrap();
        project.ide_mut().form_instance_numeric(f, id, vba_ide::FormNumericInput::Step { up: true, large }, &mut host).unwrap();
        for _ in 0..100 {
            project.advance(&mut host).unwrap();
            if project.ide().mode() == Mode::Completed {
                break;
            }
        }
    }
    assert_eq!(project.ide().form_instance_field(f, Some(14), "Value").unwrap(), wordcraft_vba::Value::Long(9));
    assert_eq!(project.ide().form_instance_field(f, Some(15), "Value").unwrap(), wordcraft_vba::Value::Long(-5));
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_listbox_rows_and_authorized_input_preserve_docm_project_bytes() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,l As MSForms.ListBox\nSet l=f.List1\nl.AddItem \"alpha\"\nl.AddItem \"beta\"\nl.List(0,1)=\"A\"\nl.List(1,1)=\"B\"\nl.ListIndex=0\nf.Show 0\nDebug.Print l.ListCount\nDebug.Print l.Value\nEnd Sub",
            ),
        ]),
        fixture::list_companion(),
    );
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().output().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed);
    assert_eq!(project.ide().output(), vec!["initialized", "2", "A"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    project.prepare_ui_action(&mut host).unwrap();
    project.ide_mut().form_instance_list(f, 14, 1, &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_field(f, Some(14), "ListIndex").unwrap(), wordcraft_vba::Value::Long(1));
    assert_eq!(
        project.ide().form_instance_field(f, Some(14), "Value").unwrap(),
        wordcraft_vba::Value::String(vba_values::VbaString::from_utf8("B").unwrap())
    );
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_combobox_text_selection_and_input_preserve_docm_project_bytes() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,l As MSForms.ComboBox\nSet l=f.Combo1\nl.AddItem \"alpha\"\nl.AddItem \"beta\"\nl.List(0,1)=\"A\"\nl.List(1,1)=\"B\"\nl.ListIndex=0\nf.Show 0\nDebug.Print l.ListCount\nDebug.Print l.Value\nEnd Sub",
            ),
        ]),
        fixture::combo_companion(3, 2),
    );
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().output().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed);
    assert_eq!(project.ide().output(), vec!["initialized", "2", "A"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    project.prepare_ui_action(&mut host).unwrap();
    project.ide_mut().form_instance_list(f, 14, 1, &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_field(f, Some(14), "ListIndex").unwrap(), wordcraft_vba::Value::Long(1));
    assert_eq!(
        project.ide().form_instance_field(f, Some(14), "Value").unwrap(),
        wordcraft_vba::Value::String(vba_values::VbaString::from_utf8("B").unwrap())
    );
    project.prepare_ui_action(&mut host).unwrap();
    project.ide_mut().form_instance_text(f, 14, vba_values::VbaString::from_utf8("free😀").unwrap(), (6, 0), &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(project.ide().mode(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().form_instance_field(f, Some(14), "ListIndex").unwrap(), wordcraft_vba::Value::Long(-1));
    assert_eq!(
        project.ide().form_instance_field(f, Some(14), "Value").unwrap(),
        wordcraft_vba::Value::String(vba_values::VbaString::from_utf8("free😀").unwrap())
    );
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}
