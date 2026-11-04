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
