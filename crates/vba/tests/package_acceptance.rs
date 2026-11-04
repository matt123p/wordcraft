//! M7 authored DOCM package acceptance; no Office observation or signature validation claim.
mod support;
use std::io::{Cursor, Read};
use std::sync::Arc;
use vba_ide::{Action, Mode};
use wordcraft_doc::{Document, StoryRef};
use wordcraft_docx::{Flavor, read, write_as};
use wordcraft_engine::Session;
use wordcraft_vba::{WordHost, WordHostState, WordVbaProject};
const STANDARD: &str = include_str!("../../../tests/fixtures/m7/Package.bas");
const SINK: &str = include_str!("../../../tests/fixtures/m7/AppSink.cls");
const DOCUMENT: &str = include_str!("../../../tests/fixtures/m7/ThisDocument.cls");
fn package() -> Vec<u8> {
    let mut doc = Document::from_text("initial");
    doc.passthrough.insert(
        "word/vbaProject.bin".into(),
        Arc::new(support::project(&[("Package", "Module", STANDARD), ("ThisDocument", "Document", DOCUMENT), ("AppSink", "Class", SINK)])),
    );
    // Opaque sentinels prove byte/relationship transport, never cryptographic validity.
    doc.passthrough.insert("word/vbaProjectSignature.bin".into(), Arc::new(b"authored opaque signature sentinel".to_vec()));
    doc.passthrough.insert("word/vbaData.xml".into(), Arc::new(b"<authored/>".to_vec()));
    doc.passthrough.insert("wordcraft:vbaProject.related".into(), Arc::new(concat!(
        "http://schemas.microsoft.com/office/2006/relationships/vbaProjectSignature\tword/vbaProjectSignature.bin\tapplication/vnd.ms-office.vbaProjectSignature\n",
        "http://schemas.microsoft.com/office/2006/relationships/wordVbaData\tword/vbaData.xml\tapplication/vnd.ms-word.vbaData+xml\n"
    ).as_bytes().to_vec()));
    write_as(&doc, Flavor::MacroDocument).unwrap()
}
fn part(bytes: &[u8], name: &str) -> Vec<u8> {
    let mut zip = zip::ZipArchive::new(Cursor::new(bytes)).unwrap();
    let mut data = vec![];
    zip.by_name(name).unwrap().read_to_end(&mut data).unwrap();
    data
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn unchanged_docm_load_compile_trust_run_save_reopen() {
    let input = package();
    let mut session = Session::new(read(&input).unwrap());
    let original = session.doc.passthrough.get("word/vbaProject.bin").unwrap().clone();
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert_eq!(project.ide().mode(), Mode::Design);
    assert_eq!(project.ide().modules.iter().map(|m| m.text.as_str()).collect::<Vec<_>>(), [STANDARD, DOCUMENT, SINK]);
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "initial");
    assert_eq!(host.session_mut().undo_depth(), 0);
    project.ide_mut().entry = "Main".into();
    project.ide_mut().dispatch(Action::Compile, &mut host).unwrap();
    assert!(project.run("Main", &mut host).is_err(), "attachment/compile must not grant trust");
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "initial");
    host.set_authorized(true);
    assert_eq!(project.run("Main", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "package11");
    assert_eq!(host.session_mut().undo_depth(), 1);
    let output = write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap();
    for name in ["word/vbaProject.bin", "word/vbaProjectSignature.bin", "word/vbaData.xml"] {
        assert_eq!(part(&output, name), part(&input, name), "{name}");
    }
    let rels = String::from_utf8(part(&output, "word/_rels/vbaProject.bin.rels")).unwrap();
    assert!(rels.contains("vbaProjectSignature") && rels.contains("wordVbaData"));
    let back = read(&output).unwrap();
    assert_eq!(back.plain_text(StoryRef::Body), "package11");
    assert_eq!(back.passthrough.get("word/vbaProject.bin"), Some(&original));
    let mut reopened = Session::new(back);
    let mut fresh_state = WordHostState::default();
    let mut fresh_host = WordHost::new(&mut reopened, &mut fresh_state).unwrap();
    let mut fresh_project = WordVbaProject::new(&fresh_host).unwrap();
    assert_eq!(fresh_project.ide().modules.iter().map(|m| m.text.as_str()).collect::<Vec<_>>(), [STANDARD, DOCUMENT, SINK]);
    assert!(fresh_project.run("Main", &mut fresh_host).is_err(), "consent does not survive reopening");
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn packaged_thisdocument_event_is_explicit_and_owner_bound() {
    let mut session = Session::new(read(&package()).unwrap());
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    assert!(project.dispatch_document_event("ThisDocument", "Document_Open", &mut host).is_err());
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "initial");
    host.set_authorized(true);
    assert_eq!(project.dispatch_document_event("ThisDocument", "Document_Open", &mut host).unwrap(), Mode::Completed);
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "owner-bound");
    assert_eq!(project.ide().modules[1].code_name.as_deref(), Some("ThisDocument"));
    assert_eq!(read(&write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap()).unwrap().plain_text(StoryRef::Body), "owner-bound");
}
#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn malformed_or_unknown_document_project_refuses_attachment_without_effects() {
    for bytes in [b"invalid CFB".to_vec(), support::project(&[("OtherDocument", "Document", "Attribute VB_Name = \"OtherDocument\"\n")])] {
        let mut doc = Document::from_text("keep");
        doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(bytes.clone()));
        let mut session = Session::new(doc);
        let mut state = WordHostState::default();
        let mut host = WordHost::new(&mut session, &mut state).unwrap();
        assert!(WordVbaProject::new(&host).is_err());
        assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "keep");
        assert_eq!(host.session_mut().doc.passthrough["word/vbaProject.bin"].as_slice(), bytes);
        assert_eq!(host.session_mut().undo_depth(), 0);
    }
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn packaged_word_code_edits_keep_hidden_metadata_in_plain_source_only() {
    let mut session = Session::new(read(&package()).unwrap());
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    project.ide_mut().active = 0;
    let display = vba_ide::editor_text(STANDARD);
    assert!(!display.contains("Attribute"));
    let edited = project.ide().active_module().unwrap().edited_source(&display.replace("\"package\"", "\"edited\""));
    project.ide_mut().set_text(edited.clone()).unwrap();
    assert!(edited.contains("Attribute Main.VB_ProcData.VB_Invoke_Func"));
    project.ide_mut().dispatch(Action::Save, &mut host).unwrap();
    assert!(project.take_source_requests().iter().any(|r| matches!(r,vba_ide::HostRequest::Save {text,..} if text == &edited)));
    let m = project.ide().active_module().unwrap();
    let (id, revision) = (m.id, m.revision);
    project.ide_mut().set_text(format!("{edited}\n' intervening edit\n")).unwrap();
    assert!(project.mark_source_saved(id, revision).is_err());
    let mut reopened = Session::new(read(&write_as(&host.session_mut().doc, Flavor::MacroDocument).unwrap()).unwrap());
    let mut fresh_state = WordHostState::default();
    let fresh_host = WordHost::new(&mut reopened, &mut fresh_state).unwrap();
    let fresh_project = WordVbaProject::new(&fresh_host).unwrap();
    assert_eq!(fresh_project.ide().modules[0].text, STANDARD);
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn packaged_withevents_subscription_cancels_save_and_reset_disconnects_it() {
    use wordcraft_vba::ApplicationEvent;
    let mut session = Session::new(read(&package()).unwrap());
    let mut state = WordHostState::default();
    let mut host = WordHost::new(&mut session, &mut state).unwrap();
    let mut project = WordVbaProject::new(&host).unwrap();
    host.set_authorized(true);
    assert_eq!(project.run("Connect", &mut host).unwrap(), Mode::Completed, "{:?}", project.ide().diagnostics);
    assert_eq!(
        project.dispatch_application_event(ApplicationEvent::DocumentBeforeSave { save_as_ui: false }, &mut host).unwrap(),
        Some(Mode::Completed)
    );
    assert_eq!(project.take_boolean_callback_output(2), Some(true));
    assert_eq!(host.session_mut().doc.plain_text(StoryRef::Body), "before-save");
    project.reset(&mut host).unwrap();
    assert!(project.dispatch_application_event(ApplicationEvent::DocumentBeforeSave { save_as_ui: false }, &mut host).unwrap().is_none());
}

#[cfg_attr(all(target_arch = "wasm32", target_os = "unknown"), wasm_bindgen_test::wasm_bindgen_test)]
#[cfg_attr(not(all(target_arch = "wasm32", target_os = "unknown")), test)]
fn rewritten_word_package_invalidates_package_signature_but_keeps_original_vba_signature_bytes() {
    use std::io::Write;
    let input = package();
    let mut zip = zip::ZipArchive::new(Cursor::new(&input)).unwrap();
    let mut out = zip::ZipWriter::new(Cursor::new(Vec::new()));
    for i in 0..zip.len() {
        let mut f = zip.by_index(i).unwrap();
        let name = f.name().to_string();
        let mut data = Vec::new();
        f.read_to_end(&mut data).unwrap();
        if name == "_rels/.rels" {
            data = String::from_utf8(data).unwrap().replace("</Relationships>",
                "<Relationship Id=\"package-signature\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"_xmlsignatures/origin.xml\"/></Relationships>"
            ).into_bytes();
        }
        out.start_file(&name, zip::write::SimpleFileOptions::default()).unwrap();
        out.write_all(&data).unwrap();
    }
    out.start_file("_xmlsignatures/origin.xml", zip::write::SimpleFileOptions::default()).unwrap();
    out.write_all(b"<authored-package-signature/>").unwrap();
    let signed = out.finish().unwrap().into_inner();
    let doc = read(&signed).unwrap();
    let output = write_as(&doc, Flavor::MacroDocument).unwrap();
    let mut back = zip::ZipArchive::new(Cursor::new(&output)).unwrap();
    assert!(back.by_name("_xmlsignatures/origin.xml").is_err());
    assert!(!String::from_utf8(part(&output, "_rels/.rels")).unwrap().contains("digital-signature/origin"));
    assert_eq!(part(&output, "word/vbaProjectSignature.bin"), part(&input, "word/vbaProjectSignature.bin"));
}
