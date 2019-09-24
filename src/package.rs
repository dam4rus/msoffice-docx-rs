use crate::wml::{document::Document, styles::Styles};
use msoffice_shared::{
    docprops::{AppInfo, Core},
    relationship::Relationship,
    xml::zip_file_to_xml_node,
};
use std::{
    fs::File,
    path::{Path, PathBuf},
};
use zip::ZipArchive;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Package {
    pub app_info: Option<AppInfo>,
    pub core: Option<Core>,
    pub main_document: Option<Document>,
    pub main_document_relationships: Vec<Relationship>,
    pub styles: Option<Styles>,
    pub medias: Vec<PathBuf>,
}

impl Package {
    pub fn from_file(file_path: &Path) -> Result<Self> {
        let file = File::open(file_path)?;
        let mut zipper = ZipArchive::new(&file)?;

        let mut instance: Self = Default::default();
        for idx in 0..zipper.len() {
            let mut zip_file = zipper.by_index(idx)?;

            match zip_file.name() {
                "docProps/app.xml" => instance.app_info = Some(AppInfo::from_zip_file(&mut zip_file)?),
                "docProps/core.xml" => instance.core = Some(Core::from_zip_file(&mut zip_file)?),
                "word/document.xml" => {
                    let xml_node = zip_file_to_xml_node(&mut zip_file)?;
                    instance.main_document = Some(Document::from_xml_element(&xml_node)?)
                }
                "word/_rels/document.xml.rels" => {
                    instance.main_document_relationships = zip_file_to_xml_node(&mut zip_file)?
                        .child_nodes
                        .iter()
                        .map(Relationship::from_xml_element)
                        .collect::<Result<Vec<_>>>()?
                }
                "word/styles.xml" => {
                    let xml_node = zip_file_to_xml_node(&mut zip_file)?;
                    instance.styles = Some(Styles::from_xml_element(&xml_node)?)
                }
                path if path.starts_with("word/media/") => instance.medias.push(PathBuf::from(file_path)),
                _ => (),
            }
        }

        Ok(instance)
    }
}

