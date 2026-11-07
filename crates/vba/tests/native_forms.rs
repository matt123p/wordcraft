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
fn settle_native(project: &mut WordVbaProject, host: &mut WordHost<'_>) -> Result<(), Box<dyn std::error::Error>> {
    for _ in 0..500 {
        project.advance(host)?;
        if project.ide().mode() != Mode::Running {
            return Ok(());
        }
    }
    Err("native form timeout".into())
}
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
    let streams = resource.walk().filter(|entry| entry.is_stream()).map(|entry| entry.path().to_path_buf()).collect::<Vec<_>>();
    for name in streams {
        let mut bytes = Vec::new();
        resource.open_stream(&name).unwrap().read_to_end(&mut bytes).unwrap();
        let target = format!("/NativeForm{}", name.display());
        let parent = std::path::Path::new(&target).parent().unwrap();
        file.create_storage_all(parent).unwrap();
        file.create_stream(&target).unwrap().write_all(&bytes).unwrap();
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
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,l As MSForms.ListBox\nSet l=f.List1\nDim rows(-1 To 0,3 To 4) As String\nrows(-1,3)=\"alpha\"\nrows(-1,4)=\"A\"\nrows(0,3)=\"beta\"\nrows(0,4)=\"B\"\nl.List=rows\nrows(-1,4)=\"changed source\"\nDim snapshot As Variant\nsnapshot=l.List\nl.List=snapshot\nsnapshot(0,1)=\"detached snapshot\"\nl.ListIndex=0\nf.Show 0\nDebug.Print l.ListCount\nDebug.Print l.Column(1)\nEnd Sub",
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
                "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,l As MSForms.ComboBox\nSet l=f.Combo1\nDim columns(3 To 4,-1 To 0) As String\ncolumns(3,-1)=\"alpha\"\ncolumns(4,-1)=\"A\"\ncolumns(3,0)=\"beta\"\ncolumns(4,0)=\"B\"\nl.Column=columns\ncolumns(4,-1)=\"changed source\"\nDim snapshot As Variant\nsnapshot=l.Column\nl.Column=snapshot\nsnapshot(1,0)=\"detached snapshot\"\nl.ListIndex=0\nf.Show 0\nDebug.Print l.ListCount\nDebug.Print l.Column(1)\nEnd Sub",
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

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_metadata_and_control_tips_keep_original_docm_resources() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,c As MSForms.Control
Set c=f.Entry
Debug.Print c.Tag,c.ControlTipText,c.HelpContextID
c.Tag="runtime tag"
c.ControlTipText="Runtime help"
c.HelpContextID=42
f.Tag="form tag"
f.HelpContextID=43
Debug.Print f.Tag,f.HelpContextID,c.Tag,c.ControlTipText,c.HelpContextID
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::metadata_companion(),
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
    assert_eq!(project.ide().output(), vec!["initialized", "tag-12\ttip-12\t112", "form tag\t43\truntime tag\tRuntime help\t42"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    assert_eq!(project.ide().form_instance_field(f, None, "HelpContextID").unwrap(), wordcraft_vba::Value::Long(43));
    assert_eq!(
        project.ide().form_instance_field(f, Some(12), "ControlTipText").unwrap(),
        wordcraft_vba::Value::String(vba_values::VbaString::from_utf8("Runtime help").unwrap())
    );
    let native = project.ide().native_form(f.source).unwrap().clone();
    project.prepare_ui_action(&mut host).unwrap();
    project.ide_mut().form_instance_text(f, 12, vba_values::VbaString::from_utf8("edited").unwrap(), (6, 0), &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(project.ide().native_form(f.source), Some(&native));
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_nested_frame_collections_and_input_preserve_docm_resources() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,g As MSForms.Frame
Set g=f.Frame1
Debug.Print f.Controls.Count,g.Controls.Count,f.Frame2.Controls.Count
Debug.Print f.Entry.Parent Is f.Frame2
f.Show 0
Debug.Print f.ActiveControl.Name,g.ActiveControl.Name
End Sub"#,
            ),
        ]),
        fixture::frame_companion(2),
    );
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().output().is_empty());
    assert!(project.ide().form_instances().unwrap().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["initialized", "3\t1\t1", "True", "Frame1\tFrame2"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    let native = project.ide().native_form(f.source).unwrap().clone();
    project.prepare_ui_action(&mut host).unwrap();
    project.ide_mut().form_instance_text(f, 12, vba_values::VbaString::from_utf8("nested😀").unwrap(), (8, 0), &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(
        project.ide().form_instance_field(f, Some(12), "VbaText").unwrap(),
        wordcraft_vba::Value::String(vba_values::VbaString::from_utf8("nested😀").unwrap())
    );
    assert_eq!(project.ide().native_form(f.source), Some(&native));
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_tabstrip_objects_and_input_preserve_docm_resources() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,s As MSForms.TabStrip
Set s=f.Tabs1
Debug.Print f.Controls.Count,s.Tabs.Count,s.SelectedItem Is s.FirstTab
s.Value=1
Debug.Print s.SelectedItem.Caption
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::tab_companion(0),
    );
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().output().is_empty());
    assert!(project.ide().form_instances().unwrap().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["initialized", "4\t3\tTrue", "Second😀"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    let native = project.ide().native_form(f.source).unwrap().clone();
    project.prepare_ui_action(&mut host).unwrap();
    assert!(project.ide_mut().form_instance_tab(f, 14, 2, &mut host).is_err());
    project.ide_mut().form_instance_tab(f, 14, 0, &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_field(f, Some(14), "Value").unwrap(), wordcraft_vba::Value::Long(0));
    assert_eq!(project.ide().native_form(f.source), Some(&native));
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_multipage_objects_and_input_preserve_docm_resources() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,s As MSForms.MultiPage
Set s=f.Multi1
Debug.Print f.Controls.Count,s.Pages.Count,s.SelectedItem Is s.FirstTab
s.Value=1
Debug.Print s.SelectedItem.Caption
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::multipage_companion(0),
    );
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().output().is_empty());
    assert!(project.ide().form_instances().unwrap().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["initialized", "4\t3\tTrue", "Second😀"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    let native = project.ide().native_form(f.source).unwrap().clone();
    project.prepare_ui_action(&mut host).unwrap();
    assert!(project.ide_mut().form_instance_tab(f, 20, 2, &mut host).is_err());
    project.ide_mut().form_instance_tab(f, 20, 0, &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_field(f, Some(20), "Value").unwrap(), wordcraft_vba::Value::Long(0));
    assert!(project.ide_mut().form_instance_text(f, 41, vba_values::VbaString::from_utf8("hidden").unwrap(), (6, 0), &mut host).is_err());
    project.ide_mut().form_instance_text(f, 40, vba_values::VbaString::from_utf8("selected").unwrap(), (8, 0), &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(
        project.ide().form_instance_field(f, Some(40), "VbaText").unwrap(),
        wordcraft_vba::Value::String(vba_values::VbaString::from_utf8("selected").unwrap())
    );
    assert_eq!(project.ide().native_form(f.source), Some(&native));
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_image_objects_and_input_preserve_docm_resources() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,s As MSForms.Image
Set s=f.Image1
Dim p As stdole.IPictureDisp
Set p=s.Picture
Set f.Image2.Picture=p
Debug.Print f.Controls.Count,p.Width,p.Height,p.Type
Debug.Print f.Image2.Picture Is p
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::image_companion(),
    );
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().output().is_empty());
    assert!(project.ide().form_instances().unwrap().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["initialized", "5\t1200\t720\t1", "True"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    let native = project.ide().native_form(f.source).unwrap().clone();
    project.prepare_ui_action(&mut host).unwrap();
    project.ide_mut().form_instance_event(f, Some(14), None, &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(
        project.ide().form_instance_field(f, None, "VbaActiveControl").unwrap(),
        project.ide().form_instance_field(f, None, "VbaEntry").unwrap()
    );
    let first = project.ide().form_instance_picture(f, 14).unwrap().unwrap();
    let second = project.ide().form_instance_picture(f, 15).unwrap().unwrap();
    assert!(Arc::ptr_eq(&first.rgba, &second.rgba));
    assert_eq!(project.ide().native_form(f.source), Some(&native));
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_refedit_text_preserves_docm_resources_without_an_excel_picker() {
    let original = form_project(
        support::project(&[
            ("NativeForm", "BaseClass", fixture::SOURCE),
            (
                "Launch",
                "Module",
                r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,r As RefEdit.RefEdit
Set r=f.Entry
Debug.Print r.Value,r.Parent Is f
f.Show 0
End Sub"#,
            ),
        ]),
        fixture::refedit_companion(),
    );
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().form_instances().unwrap().is_empty());
    assert!(project.ide().output().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["initialized", "initial\tTrue"]);
    let f = *project.ide().form_instances().unwrap().first().unwrap();
    let native = project.ide().native_form(f.source).unwrap().clone();
    project.prepare_ui_action(&mut host).unwrap();
    assert!(project.ide_mut().form_instance_drop(f, 12, true, &mut host).is_err());
    project.ide_mut().form_instance_text(f, 12, vba_values::VbaString::from_utf8("typed reference").unwrap(), (0, 0), &mut host).unwrap();
    for _ in 0..100 {
        project.advance(&mut host).unwrap();
        if project.ide().mode() == Mode::Completed {
            break;
        }
    }
    assert_eq!(
        project.ide().form_instance_field(f, Some(12), "VbaText").unwrap(),
        wordcraft_vba::Value::String(vba_values::VbaString::from_utf8("typed reference").unwrap())
    );
    assert_eq!(project.ide().form_range_selection().unwrap(), None);
    assert_eq!(project.ide().native_form(f.source), Some(&native));
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_dynamic_tabs_and_pages_keep_aliases_trust_and_docm_resources() {
    for pages in [false, true] {
        let (control, collection, id) = if pages { ("Multi1", "Pages", 20) } else { ("Tabs1", "Tabs", 14) };
        let source = format!(
            "Attribute VB_Name = \"Launch\"\nSub Main()\nDim f As New NativeForm,s As Object,t As Object\nSet s=f.{control}\nSet t=s.{collection}(0)\nt.Index=1\ns.{collection}.Add Name:=\"Runtime\",Caption:=\"Added\",Index:=0\ns.{collection}.Remove \"BlockedTab\"\nDebug.Print s.Value,s.{collection}(2) Is t,s.{collection}.Count\nf.Show 0\nEnd Sub"
        );
        let original = form_project(
            support::project(&[("NativeForm", "BaseClass", fixture::SOURCE), ("Launch", "Module", &source)]),
            if pages { fixture::multipage_companion(0) } else { fixture::tab_companion(0) },
        );
        let mut doc = Document::from_text("unchanged");
        doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
        let mut session = Session::new(doc);
        let mut state = WordHostState::default();
        let mut host = WordHost::new(&mut session, &mut state).unwrap();
        let mut project = WordVbaProject::new(&host).unwrap();
        assert!(project.ide().form_instances().unwrap().is_empty());
        assert!(project.run("Main", &mut host).is_err());
        host.set_authorized(true);
        assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
        assert_eq!(project.ide().output(), vec!["initialized", "2\tTrue\t3"]);
        let f = *project.ide().form_instances().unwrap().first().unwrap();
        let native = project.ide().native_form(f.source).unwrap().clone();
        project.prepare_ui_action(&mut host).unwrap();
        project.ide_mut().form_instance_tab(f, id, 0, &mut host).unwrap();
        for _ in 0..500 {
            project.advance(&mut host).unwrap();
            if project.ide().mode() != Mode::Running {
                break;
            }
        }
        assert_eq!(project.ide().mode(), Mode::Completed, "{:?}", project.ide().diagnostics);
        assert_eq!(project.ide().form_instance_field(f, Some(id), "Value").unwrap(), wordcraft_vba::Value::Long(0));
        if pages {
            assert!(project.ide_mut().form_instance_text(f, 40, vba_values::VbaString::from_utf8("hidden").unwrap(), (0, 0), &mut host).is_err());
        }
        assert_eq!(project.ide().native_form(f.source), Some(&native));
        let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
        assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
        assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
        host.set_authorized(false);
        assert!(project.prepare_ui_action(&mut host).is_err());
    }
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_runtime_controls_selection_and_typenames_preserve_docm_and_refuse_excel_sources() {
    let source = r#"Attribute VB_Name = "Launch"
Sub Main()
Dim f As New NativeForm,t As MSForms.TextBox,l As MSForms.ListBox
Set t=f.Controls.Add("Forms.TextBox.1","Added")
t.Value="new"
Set l=f.List1
l.AddItem "alpha"
l.AddItem "beta"
l.MultiSelect=2
l.Selected(0)=True
l.Selected(1)=True
Debug.Print TypeName(f),TypeName(t),TypeName(l),f.Controls.Count
Debug.Print t.Value,l.Selected(0),l.Selected(1),IsNull(l.Value)
On Error Resume Next
f.Entry.ControlSource="A1"
Debug.Print Err.Number,f.Entry.ControlSource,f.Entry.Value
Err.Clear
l.RowSource="A1:A2"
Debug.Print Err.Number,l.RowSource,l.ListCount
f.Show 0
End Sub"#;
    let original =
        form_project(support::project(&[("NativeForm", "BaseClass", fixture::SOURCE), ("Launch", "Module", source)]), fixture::list_companion());
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.ide().form_instances().unwrap().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["initialized", "NativeForm\tTextBox\tListBox\t5", "new\tTrue\tTrue\tTrue", "380\t\tinitial", "380\t\t2"]);
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    assert_eq!(read(&saved).unwrap().passthrough.get("word/vbaProject.bin").unwrap().as_slice(), original);
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_form_arrays_keep_callback_receivers_trust_and_docm_resources() -> Result<(), Box<dyn std::error::Error>> {
    let launch = "Attribute VB_Name = \"Launch\"\nSub Main()\nDim forms(0 To 1) As New NativeForm,a As NativeForm\nDebug.Print UserForms.Count\nforms(0).Show 0\nforms(0).Entry.Text=\"first\"\nforms(1).Show 0\nforms(1).Entry.Text=\"second\"\nSet a=forms(0)\nSet forms(0)=Nothing\nforms(0).Show 0\nDebug.Print a.Entry.Text,forms(1).Entry.Text,forms(0).Entry.Text\nEnd Sub";
    let form = fixture::SOURCE.replace("Debug.Print Entry.Text", "ActiveDocument.Content.Text=Entry.Text\r\nDebug.Print Entry.Text");
    let original = fixture::add_designer(support::project(&[("NativeForm", "BaseClass", &form), ("Launch", "Module", launch)]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(read(&write_as(&doc, Flavor::MacroDocument)?)?);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.ide().output().is_empty());
    assert!(project.ide().form_instances()?.is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output(), vec!["0", "initialized", "initialized", "initialized", "first\tsecond\tinitial"]);
    let instances = project.ide().form_instances()?;
    assert_eq!(instances.len(), 3);
    let expected = vba_values::Value::String(vba_values::VbaString::from_utf8("second")?);
    let target = instances
        .into_iter()
        .find(|instance| project.ide().form_instance_field(*instance, Some(12), "VbaText").is_ok_and(|value| value == expected))
        .ok_or("second instance")?;
    project.prepare_ui_action(&mut host)?;
    project.ide_mut().form_instance_event(target, Some(13), None, &mut host)?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().mode(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(project.ide().output().last().map(String::as_str), Some("second"));
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "second");
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_keyboard_waits_for_debugger_and_preserves_document_authority_and_resources() -> Result<(), Box<dyn std::error::Error>> {
    let form = format!(
        "{}\nPrivate Sub Entry_KeyPress(ByVal KeyANSI As MSForms.ReturnInteger)\nDebug.Print KeyANSI.Value\nStop\nActiveDocument.Content.Text=Chr(KeyANSI.Value)\nKeyANSI.Value=0\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show 0\nNativeForm.Entry.SetFocus\nEnd Sub"),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Completed);
    let f = *project.ide().form_instances()?.first().ok_or("form")?;
    project.prepare_ui_action(&mut host)?;
    let epoch = project.ide_mut().form_instance_key(f, Some(12), vba_ide::FormKeyInput::Text(vba_values::VbaString::from_utf8("a")?), &mut host)?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().mode(), Mode::Break);
    assert!(project.ide().form_instance_key_result(f, epoch)?.is_none());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    project.ide_mut().dispatch(Action::Run, &mut host)?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_key_result(f, epoch)?.ok_or("receipt")?.code, 0);
    assert_eq!(project.ide().output(), vec!["initialized", "97"]);
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "a");
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_root_keys_distinguish_tabstop_from_disabled_controls_and_keep_docm_resources() -> Result<(), Box<dyn std::error::Error>> {
    for disabled in [false, true] {
        let form = format!(
            "{}\nPrivate Sub UserForm_KeyDown(ByVal k As MSForms.ReturnInteger,ByVal s As Integer)\nStop\nActiveDocument.Content.Text=\"root\"\nDebug.Print \"root\"\nk.Value=0\nEnd Sub",
            fixture::SOURCE
        );
        let launch = format!(
            "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Entry.TabStop=False\nNativeForm.Apply.TabStop=False\n{}\nNativeForm.Show 0\nEnd Sub",
            if disabled { "Dim c As MSForms.Control\nFor Each c In NativeForm.Controls\nc.Enabled=False\nNext c" } else { "" }
        );
        let original = fixture::add_designer(support::project(&[("NativeForm", "BaseClass", &form), ("Launch", "Module", &launch)]));
        let mut doc = Document::from_text("unchanged");
        doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
        let mut session = Session::new(doc);
        let mut state = WordHostState::default();
        let mut host = WordHost::new(&mut session, &mut state)?;
        let mut project = WordVbaProject::new(&host)?;
        host.set_authorized(true);
        assert_eq!(project.run("Main", &mut host)?, Mode::Completed);
        let f = *project.ide().form_instances()?.first().ok_or("form")?;
        project.prepare_ui_action(&mut host)?;
        let epoch = project.ide_mut().form_instance_key(
            f,
            None,
            vba_ide::FormKeyInput::Code { event: vba_ide::FormKeyEvent::Down, code: 65, shift: 0 },
            &mut host,
        )?;
        for _ in 0..500 {
            project.advance(&mut host)?;
            if project.ide().mode() != Mode::Running {
                break;
            }
        }
        assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
        if disabled {
            assert_eq!(project.ide().mode(), Mode::Break);
            assert!(project.ide().form_instance_key_result(f, epoch)?.is_none());
            project.ide_mut().dispatch(Action::Run, &mut host)?;
            for _ in 0..500 {
                project.advance(&mut host)?;
                if project.ide().mode() != Mode::Running {
                    break;
                }
            }
        }
        assert_eq!(project.ide().mode(), Mode::Completed, "{:?}", project.ide().diagnostics);
        assert_eq!(project.ide().form_instance_key_result(f, epoch)?.ok_or("receipt")?.code, 0);
        assert_eq!(project.ide().output(), if disabled { vec!["initialized", "root"] } else { vec!["initialized"] });
        if disabled {
            assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "root");
            assert_eq!(host.session_mut().undo_depth(), 1);
            assert!(host.session_mut().undo());
        }
        assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
        assert_eq!(
            read(&write_as(&host.session_mut().doc, Flavor::MacroDocument)?)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(),
            original
        );
        host.set_authorized(false);
        assert!(project.prepare_ui_action(&mut host).is_err());
    }
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_accelerator_click_waits_for_debugger_and_preserves_document_resources() -> Result<(), Box<dyn std::error::Error>> {
    let form = fixture::SOURCE.replace("Debug.Print Entry.Text", "Stop\r\nActiveDocument.Content.Text=\"accelerated\"\r\nDebug.Print Entry.Text");
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        (
            "Launch",
            "Module",
            "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show 0\nNativeForm.Apply.Accelerator=\"A\"\nNativeForm.Entry.SetFocus\nEnd Sub",
        ),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Completed);
    let f = *project.ide().form_instances()?.first().ok_or("form")?;
    project.prepare_ui_action(&mut host)?;
    let epoch = project.ide_mut().form_instance_key(
        f,
        Some(12),
        vba_ide::FormKeyInput::Code { event: vba_ide::FormKeyEvent::Down, code: 65, shift: 4 },
        &mut host,
    )?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().mode(), Mode::Break);
    assert!(project.ide().form_instance_key_result(f, epoch)?.is_none());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    project.ide_mut().dispatch(Action::Run, &mut host)?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_key_result(f, epoch)?.ok_or("receipt")?.code, 0);
    assert_eq!(project.ide().output(), vec!["initialized", "initial"]);
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "accelerated");
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_mouse_callback_waits_for_debugger_and_preserves_document_resources() -> Result<(), Box<dyn std::error::Error>> {
    let form = format!(
        "{}\nPrivate Sub Apply_MouseDown(ByVal Button As Integer,ByVal Shift As Integer,ByVal X As Single,ByVal Y As Single)\nStop\nActiveDocument.Content.Text=\"mouse\"\nDebug.Print \"down\"\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show 0\nNativeForm.Entry.SetFocus\nEnd Sub"),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Completed);
    let f = *project.ide().form_instances()?.first().ok_or("form")?;
    project.prepare_ui_action(&mut host)?;
    let epoch = project.ide_mut().form_instance_mouse(
        f,
        Some(13),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut host,
    )?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().mode(), Mode::Break);
    assert!(project.ide().form_instance_mouse_result(f, epoch)?.is_none());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    project.ide_mut().dispatch(Action::Run, &mut host)?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_mouse_result(f, epoch)?, Some(true));
    assert_eq!(project.ide().output(), vec!["initialized", "down"]);
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "mouse");
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_text_pointer_selection_after_mouse_preserves_document_resources() -> Result<(), Box<dyn std::error::Error>> {
    let form = format!(
        "{}\nPrivate Sub Entry_MouseDown(ByVal Button As Integer,ByVal Shift As Integer,ByVal X As Single,ByVal Y As Single)\nStop\nActiveDocument.Content.Text=\"mouse\"\nDebug.Print \"down\"\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show 0\nNativeForm.Entry.SetFocus\nEnd Sub"),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Completed);
    let f = *project.ide().form_instances()?.first().ok_or("form")?;
    project.prepare_ui_action(&mut host)?;
    let epoch = project.ide_mut().form_instance_mouse(
        f,
        Some(12),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut host,
    )?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().mode(), Mode::Break);
    assert!(project.ide().form_instance_mouse_result(f, epoch)?.is_none());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    project.ide_mut().dispatch(Action::Run, &mut host)?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_mouse_result(f, epoch)?, Some(true));
    project.prepare_ui_action(&mut host)?;
    project.ide_mut().form_instance_text(f, 12, vba_values::VbaString::from_utf8("initial")?, (1, 3), &mut host)?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_field(f, Some(12), "VbaSelStart")?, vba_values::Value::Long(1));
    assert_eq!(project.ide().form_instance_field(f, Some(12), "VbaSelLength")?, vba_values::Value::Long(3));
    assert_eq!(project.ide().form_instance_field(f, Some(12), "AutoWordSelect")?, vba_values::Value::Boolean(false));
    assert_eq!(project.ide().output(), vec!["initialized", "down"]);
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "mouse");
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_modal_pointer_callback_preserves_document_resources() -> Result<(), Box<dyn std::error::Error>> {
    let form = format!(
        "{}\nPrivate Sub Entry_MouseDown(ByVal Button As Integer,ByVal Shift As Integer,ByVal X As Single,ByVal Y As Single)\nStop\nActiveDocument.Content.Text=\"mouse\"\nDebug.Print \"down\"\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nNativeForm.Entry.SetFocus\nEnd Sub"),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Waiting);
    let f = *project.ide().form_instances()?.first().ok_or("form")?;
    project.prepare_ui_action(&mut host)?;
    let epoch = project.ide_mut().form_instance_mouse(
        f,
        Some(12),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut host,
    )?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().mode(), Mode::Break);
    assert!(project.ide().form_instance_has_modal_window(f));
    assert!(!project.ide().form_instance_can_deliver(f));
    assert!(project.ide().form_instance_mouse_result(f, epoch)?.is_none());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    project.ide_mut().dispatch(Action::Run, &mut host)?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_mouse_result(f, epoch)?, Some(true));
    assert!(project.ide().form_instance_has_modal_window(f));
    project.prepare_ui_action(&mut host)?;
    project.ide_mut().form_instance_text(f, 12, vba_values::VbaString::from_utf8("initial")?, (1, 3), &mut host)?;
    for _ in 0..500 {
        project.advance(&mut host)?;
        if project.ide().mode() != Mode::Running {
            break;
        }
    }
    assert_eq!(project.ide().form_instance_field(f, Some(12), "VbaSelStart")?, vba_values::Value::Long(1));
    assert_eq!(project.ide().form_instance_field(f, Some(12), "VbaSelLength")?, vba_values::Value::Long(3));
    assert_eq!(project.ide().form_instance_field(f, Some(12), "AutoWordSelect")?, vba_values::Value::Boolean(false));
    assert_eq!(project.ide().output(), vec!["initialized", "down"]);
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "mouse");
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_nested_modal_pointer_callback_preserves_document_resources() -> Result<(), Box<dyn std::error::Error>> {
    let form = format!(
        "{}\nPrivate Sub Entry_MouseDown(ByVal Button As Integer,ByVal Shift As Integer,ByVal X As Single,ByVal Y As Single)\nDim child As NativeForm\nIf Me.Caption=\"Parent\" Then\nSet child=New NativeForm\nchild.Caption=\"Child\"\nchild.Show\nElse\nStop\nActiveDocument.Content.Text=\"mouse\"\nDebug.Print \"down\"\nEnd If\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        (
            "Launch",
            "Module",
            "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Caption=\"Parent\"\nNativeForm.Show\nNativeForm.Entry.SetFocus\nEnd Sub",
        ),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Waiting);
    let f = *project.ide().form_instances()?.first().ok_or("form")?;
    project.prepare_ui_action(&mut host)?;
    let epoch = project.ide_mut().form_instance_mouse(
        f,
        Some(12),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut host,
    )?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().mode(), Mode::Waiting);
    let parent_epoch = epoch;
    let child = *project.ide().form_instances()?.last().ok_or("child")?;
    assert_ne!(child, f);
    assert_eq!(project.ide().form_instance_modal_rank(f), Some(0));
    assert_eq!(project.ide().form_instance_modal_rank(child), Some(1));
    assert!(!project.ide().form_instance_can_deliver(f));
    project.prepare_ui_action(&mut host)?;
    let epoch = project.ide_mut().form_instance_mouse(
        child,
        Some(12),
        vba_ide::FormMouseInput { event: vba_ide::FormMouseEvent::Down, button: 1, shift: 3, x: 72., y: 12., index: -1 },
        &mut host,
    )?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().mode(), Mode::Break);
    assert!(project.ide().form_instance_has_modal_window(f));
    assert!(!project.ide().form_instance_can_deliver(f));
    assert!(project.ide().form_instance_mouse_result(child, epoch)?.is_none());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    project.ide_mut().dispatch(Action::Run, &mut host)?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().form_instance_mouse_result(child, epoch)?, Some(true));
    assert!(project.ide().form_instance_mouse_result(f, parent_epoch)?.is_none());
    project.prepare_ui_action(&mut host)?;
    project.ide_mut().form_instance_event(child, None, None, &mut host)?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().form_instance_mouse_result(f, parent_epoch)?, Some(true));
    assert!(project.ide().form_instance_has_modal_window(f));
    project.prepare_ui_action(&mut host)?;
    project.ide_mut().form_instance_text(f, 12, vba_values::VbaString::from_utf8("initial")?, (1, 3), &mut host)?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().form_instance_field(f, Some(12), "VbaSelStart")?, vba_values::Value::Long(1));
    assert_eq!(project.ide().form_instance_field(f, Some(12), "VbaSelLength")?, vba_values::Value::Long(3));
    assert_eq!(project.ide().form_instance_field(f, Some(12), "AutoWordSelect")?, vba_values::Value::Boolean(false));
    assert_eq!(project.ide().output(), vec!["initialized", "initialized", "down"]);
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "mouse");
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_window_deactivation_preserves_trust_stop_and_document_resources() -> Result<(), Box<dyn std::error::Error>> {
    let form = format!(
        "{}\nPrivate Sub UserForm_Activate()\nDebug.Print Me.Caption & \" active\"\nEnd Sub\nPrivate Sub UserForm_Deactivate()\nDebug.Print Me.Caption & \" inactive\"\nIf Me.Caption=\"B\" Then\nStop\nActiveDocument.Content.Text=\"deactivated\"\nEnd If\nEnd Sub",
        fixture::SOURCE
    );
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", &form),
        (
            "Launch",
            "Module",
            "Attribute VB_Name = \"Launch\"\nPublic a As NativeForm,b As NativeForm\nSub Main()\nSet a=New NativeForm\nSet b=New NativeForm\na.Caption=\"A\"\nb.Caption=\"B\"\na.Show 0\nb.Show 0\nEnd Sub",
        ),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Completed);
    let active = project.ide().form_active_window().ok_or("active window")?;
    assert_eq!(project.ide().output(), vec!["initialized", "initialized", "A active", "A inactive", "B active"]);
    project.prepare_ui_action(&mut host)?;
    project.ide_mut().form_instance_window_focus(active, false, &mut host)?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().mode(), Mode::Break);
    assert!(project.ide().form_active_window().is_none());
    assert!(!project.ide().form_instance_can_deliver(active));
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    project.ide_mut().dispatch(Action::Run, &mut host)?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "deactivated");
    assert_eq!(project.ide().output().iter().filter(|line| *line == "B inactive").count(), 1);
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_colours_keep_canonical_bits_trusted_continuations_and_original_document_resources() -> Result<(), Box<dyn std::error::Error>> {
    let form = "Attribute VB_Name = \"NativeForm\"\nAttribute VB_PredeclaredId = True\nPrivate Sub UserForm_Initialize()\nMe.BackColor=RGB(230,240,250)\nEntry.ForeColor=RGB(80,30,20)\nEnd Sub\nPrivate Sub Apply_Click()\nEntry.BackColor=RGB(250,230,200)\nActiveDocument.Content.Text=CStr(Entry.BackColor)\nStop\nEntry.ForeColor=RGB(1,2,3)\nActiveDocument.Content.Text=ActiveDocument.Content.Text & \":\" & CStr(Entry.ForeColor)\nMe.Hide\nEnd Sub";
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nEnd Sub"),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.ide().output().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Waiting);
    let owner = project.ide().form_active_window().ok_or("form")?;
    assert_eq!(project.ide().form_instance_field(owner, None, "BackColor")?, vba_values::Value::Long(16445670));
    project.prepare_ui_action(&mut host)?;
    project.ide_mut().form_instance_event(owner, Some(13), None, &mut host)?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().mode(), Mode::Break);
    assert_eq!(project.ide().form_instance_field(owner, Some(12), "BackColor")?, vba_values::Value::Long(13166330));
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "13166330");
    project.ide_mut().dispatch(Action::Run, &mut host)?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().mode(), Mode::Completed);
    assert_eq!(project.ide().form_instance_field(owner, Some(12), "ForeColor")?, vba_values::Value::Long(197121));
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "13166330\n:197121");
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn native_fonts_keep_typed_aliases_trusted_continuations_and_original_document_resources() -> Result<(), Box<dyn std::error::Error>> {
    let form = "Attribute VB_Name = \"NativeForm\"\nAttribute VB_PredeclaredId = True\nPrivate Sub UserForm_Initialize()\nEntry.Font.Name=\"Authored face\"\nEntry.Font.Size=12.25\nEnd Sub\nPrivate Sub Apply_Click()\nDim font As MSForms.NewFont\nSet font=Entry.Font\nActiveDocument.Content.Text=CStr(font.Size)\nStop\nfont.Bold=True\nfont.Size=14.5\nActiveDocument.Content.Text=ActiveDocument.Content.Text & \":\" & CStr(font.Size)\nMe.Hide\nEnd Sub";
    let original = fixture::add_designer(support::project(&[
        ("NativeForm", "BaseClass", form),
        ("Launch", "Module", "Attribute VB_Name = \"Launch\"\nSub Main()\nNativeForm.Show\nEnd Sub"),
    ]));
    let mut doc = Document::from_text("unchanged");
    doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(original.clone()));
    let mut session = Session::new(doc);
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state)?;
    let mut project = WordVbaProject::new(&host)?;
    assert!(project.ide().output().is_empty());
    assert!(project.run("Main", &mut host).is_err());
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host)?, Mode::Waiting);
    let owner = project.ide().form_active_window().ok_or("form")?;
    assert_eq!(project.ide().form_instance_font(owner, Some(12))?.size, 122500);
    project.prepare_ui_action(&mut host)?;
    project.ide_mut().form_instance_event(owner, Some(13), None, &mut host)?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().mode(), Mode::Break);
    assert_eq!(project.ide().form_instance_font(owner, Some(12))?.name, "Authored face");
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "12.25");
    project.ide_mut().dispatch(Action::Run, &mut host)?;
    settle_native(&mut project, &mut host)?;
    assert_eq!(project.ide().mode(), Mode::Completed);
    assert_eq!((project.ide().form_instance_font(owner, Some(12))?.size, project.ide().form_instance_font(owner, Some(12))?.weight), (145000, 700));
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "12.25\n:14.5");
    assert_eq!(host.session_mut().undo_depth(), 1);
    assert!(host.session_mut().undo());
    assert_eq!(host.session_mut().doc.plain_text(wordcraft_doc::StoryRef::Body), "unchanged");
    let saved = write_as(&host.session_mut().doc, Flavor::MacroDocument)?;
    assert_eq!(read(&saved)?.passthrough.get("word/vbaProject.bin").ok_or("project")?.as_slice(), original);
    host.set_authorized(false);
    assert!(project.prepare_ui_action(&mut host).is_err());
    Ok(())
}
