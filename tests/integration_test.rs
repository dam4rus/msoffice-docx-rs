extern crate msoffice_docx;

use msoffice_docx::package::Package;
use std::path::PathBuf;

#[test]
#[ignore]
fn test_package_load() {
    let test_dir = PathBuf::from(env!("CARGO_MANIFEST_DIR"));
    let sample_docx_file = test_dir.join("tests/sample.docx");

    Package::from_file(&sample_docx_file).unwrap();
}