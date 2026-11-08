//! Explicit native project replacement; ordinary saves retain their original graph.
use crate::DocxError;
use std::sync::Arc;
use wordcraft_doc::Document;
/// Prepare a document copy with an already encoded project and no VBA signatures.
/// The caller owns source compilation/encoding, revision checks and atomic IO.
pub fn prepare_edited_vba_package(doc: &Document, expected: &[u8], edited: &[u8]) -> Result<Document, DocxError> {
    use crate::package::{MAX_VBA_RELATED, VBA_PROJECT_PART, VBA_RELATED};
    let fail = |s: &str| DocxError::Xml(s.into());
    if doc.passthrough.get(VBA_PROJECT_PART).map(|b| b.as_slice()) != Some(expected) {
        return Err(fail("VBA project changed during save"));
    }
    if edited.len() > 16 * 1024 * 1024 || !edited.starts_with(&[0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1]) {
        return Err(fail("invalid encoded VBA project"));
    }
    let mut candidate = doc.clone();
    let mut kept = String::new();
    if let Some(manifest) = doc.passthrough.get(VBA_RELATED) {
        let manifest = std::str::from_utf8(manifest).map_err(|_| fail("invalid VBA relation manifest"))?;
        if manifest.lines().count() > MAX_VBA_RELATED {
            return Err(fail("VBA relation limit"));
        }
        let mut records = Vec::new();
        let mut removed = std::collections::BTreeSet::new();
        for line in manifest.lines() {
            let mut parts = line.split('\t');
            let (Some(kind), Some(path), Some(content_type), None) = (parts.next(), parts.next(), parts.next(), parts.next()) else {
                return Err(fail("invalid VBA relation manifest"));
            };
            if kind.is_empty()
                || content_type.is_empty()
                || !path.starts_with("word/")
                || path.contains("..")
                || path.contains('\\')
                || path.to_ascii_lowercase().ends_with(".rels")
                || path.eq_ignore_ascii_case(VBA_PROJECT_PART)
                || path.eq_ignore_ascii_case("word/document.xml")
                || !doc.passthrough.contains_key(path)
            {
                return Err(fail("incomplete or unsupported VBA relation manifest"));
            }
            let signature = |s: &str| {
                let s = s.to_ascii_lowercase();
                s.contains("vbaprojectsignature") || s.contains("digital-signature")
            };
            if signature(kind) || signature(content_type) {
                removed.insert(path.to_ascii_lowercase());
            }
            records.push((line, path));
        }
        candidate.passthrough.retain(|name, _| !removed.contains(&name.to_ascii_lowercase()));
        for (line, path) in records {
            if !removed.contains(&path.to_ascii_lowercase()) {
                kept.push_str(line);
                kept.push('\n');
            }
        }
    }
    if kept.is_empty() {
        candidate.passthrough.remove(VBA_RELATED);
    } else {
        candidate.passthrough.insert(VBA_RELATED.into(), Arc::new(kept.into_bytes()));
    }
    candidate.passthrough.insert(VBA_PROJECT_PART.into(), Arc::new(edited.to_vec()));
    Ok(candidate)
}

#[cfg(test)]
#[allow(clippy::unwrap_used)]
mod tests {
    use super::*;
    #[test]
    fn replacement_drops_each_signature_variant_and_retains_vba_data() {
        let mut doc = Document::default();
        doc.passthrough.insert("word/vbaProject.bin".into(), Arc::new(b"original".to_vec()));
        let mut manifest = String::new();
        for (name, kind) in [("legacy", "vbaProjectSignature"), ("agile", "vbaProjectSignatureAgile"), ("v3", "vbaProjectSignatureV3")] {
            let path = format!("word/{name}.bin");
            manifest.push_str(&format!("http://schemas.microsoft.com/office/2006/relationships/{kind}\t{path}\tapplication/vnd.ms-office.{kind}\n"));
            doc.passthrough.insert(path, Arc::new(b"signature sentinel".to_vec()));
        }
        manifest
            .push_str("http://schemas.microsoft.com/office/2006/relationships/wordVbaData\tword/vbaData.xml\tapplication/vnd.ms-word.vbaData+xml\n");
        doc.passthrough.insert("word/vbaData.xml".into(), Arc::new(b"<data/>".to_vec()));
        doc.passthrough.insert("wordcraft:vbaProject.related".into(), Arc::new(manifest.into_bytes()));
        let edited = b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1encoded-source-sentinel";
        let candidate = prepare_edited_vba_package(&doc, b"original", edited).unwrap();
        assert!(!candidate.passthrough.values().any(|bytes| bytes.as_slice() == b"signature sentinel"));
        assert_eq!(candidate.passthrough.get("word/vbaData.xml"), doc.passthrough.get("word/vbaData.xml"));
        let bytes = crate::write_as(&candidate, crate::Flavor::MacroDocument).unwrap();
        let back = crate::read(&bytes).unwrap();
        assert_eq!(back.passthrough.get("word/vbaProject.bin").map(|b| b.as_slice()), Some(edited.as_slice()));
        assert_eq!(back.passthrough.get("word/vbaData.xml"), doc.passthrough.get("word/vbaData.xml"));
        assert!(prepare_edited_vba_package(&doc, b"stale", edited).is_err());
        let mut incomplete = doc.clone();
        incomplete.passthrough.remove("word/vbaData.xml");
        assert!(prepare_edited_vba_package(&incomplete, b"original", edited).is_err());
        let mut unsupported = doc.clone();
        unsupported.passthrough.insert("wordcraft:vbaProject.related".into(), Arc::new(b"data\t../document.xml\tapplication/xml\n".to_vec()));
        assert!(prepare_edited_vba_package(&unsupported, b"original", edited).is_err());
        assert_eq!(doc.passthrough.get("word/vbaProject.bin").map(|b| b.as_slice()), Some(b"original".as_slice()));
    }
}
