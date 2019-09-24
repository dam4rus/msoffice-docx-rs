extern crate msoffice_docx;

use msoffice_docx::package::Package;
use std::path::PathBuf;

#[test]
#[ignore]
fn test_package_load() {
    let test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_docx_file = test_dir.join("tests/sample.docx");

    let package = Package::from_file(&sample_docx_file).unwrap();
    assert!(package.app_info.is_some());
    assert!(package.core.is_some());
    assert!(package.main_document.is_some());
    assert_eq!(package.main_document_relationships.len(), 14);
    assert!(package.styles.is_some());
    assert_eq!(package.medias.len(), 4);
}
