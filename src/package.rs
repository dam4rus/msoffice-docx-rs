use std::{
    fs::File,
    path::{Path, PathBuf},
    io::Read,
};
use msoffice_shared::{
    docprops::{AppInfo, Core},
    relationship::Relationship,
    xml::XmlNode,
};
use zip::ZipArchive;
use crate::wml::Document;

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Package {
    pub app_info: Option<AppInfo>,
    pub core: Option<Core>,
    pub main_document: Option<Document>,
    pub main_document_relationships: Vec<Relationship>,
    pub medias: Vec<PathBuf>,
}

impl Package {
    pub fn from_file(file_path: &Path) -> Result<Self, Box<dyn std::error::Error>> {
        let file = File::open(file_path)?;
        let mut zipper = ZipArchive::new(&file)?;

        let mut instance: Self = Default::default();
        if let Ok(app_info_file) = &mut zipper.by_name("docProps/app.xml") {
            instance.app_info = Some(AppInfo::from_zip_file(app_info_file)?);
        }

        if let Ok(core_file) = &mut zipper.by_name("docProps/core.xml") {
            instance.core = Some(Core::from_zip_file(core_file)?);
        }

        if let Ok(main_document_file) = &mut zipper.by_name("word/document.xml") {
            let mut xml_string = String::new();
            main_document_file.read_to_string(&mut xml_string)?;
            let root = XmlNode::from_str(xml_string)?;
            instance.main_document = Some(Document::from_xml_element(&root)?);
        }

        Ok(instance)
    }
}