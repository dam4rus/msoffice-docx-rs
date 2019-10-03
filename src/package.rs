use crate::{
    wml::{
        document::{
            Border, Color, Document, EastAsianLayout, Em, FitText, Fonts, HighlightColor, HpsMeasure, Language,
            PPrBase, RPrBase, Shd, SignedHpsMeasure, SignedTwipsMeasure, TextEffect, Underline, P, R, SectPrContents,
        },
        settings::Settings,
        simpletypes::TextScale,
        styles::{Style, Styles, StyleType},
    },
};
use log::error;
use msoffice_shared::{
    docprops::{AppInfo, Core},
    drawingml::sharedstylesheet::OfficeStyleSheet,
    relationship::{Relationship, THEME_RELATION_TYPE},
    sharedtypes::{OnOff, VerticalAlignRun},
    xml::zip_file_to_xml_node,
};
use std::{
    collections::HashMap,
    error::Error,
    ffi::OsStr,
    fs::File,
    path::{Path, PathBuf},
};
use zip::ZipArchive;

pub type ParagraphProperties = PPrBase;

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RunProperties {
    pub style: Option<String>,
    pub fonts: Option<Fonts>,
    pub bold: Option<OnOff>,
    pub complex_script_bold: Option<OnOff>,
    pub italic: Option<OnOff>,
    pub complex_script_italic: Option<OnOff>,
    pub all_capitals: Option<OnOff>,
    pub all_small_capitals: Option<OnOff>,
    pub strikethrough: Option<OnOff>,
    pub double_strikethrough: Option<OnOff>,
    pub outline: Option<OnOff>,
    pub shadow: Option<OnOff>,
    pub emboss: Option<OnOff>,
    pub imprint: Option<OnOff>,
    pub no_proofing: Option<OnOff>,
    pub snap_to_grid: Option<OnOff>,
    pub vanish: Option<OnOff>,
    pub web_hidden: Option<OnOff>,
    pub color: Option<Color>,
    pub spacing: Option<SignedTwipsMeasure>,
    pub width: Option<TextScale>,
    pub kerning: Option<HpsMeasure>,
    pub position: Option<SignedHpsMeasure>,
    pub font_size: Option<HpsMeasure>,
    pub complex_script_font_size: Option<HpsMeasure>,
    pub highlight: Option<HighlightColor>,
    pub underline: Option<Underline>,
    pub effect: Option<TextEffect>,
    pub border: Option<Border>,
    pub shading: Option<Shd>,
    pub fit_text: Option<FitText>,
    pub vertical_alignment: Option<VerticalAlignRun>,
    pub rtl: Option<OnOff>,
    pub complex_script: Option<OnOff>,
    pub emphasis_mark: Option<Em>,
    pub language: Option<Language>,
    pub east_asian_layout: Option<EastAsianLayout>,
    pub special_vanish: Option<OnOff>,
    pub o_math: Option<OnOff>,
}

impl RunProperties {
    pub fn from_vec(properties_vec: &[RPrBase]) -> Self {
        properties_vec
            .iter()
            .fold(Default::default(), |mut instance: Self, property| {
                match property {
                    RPrBase::RunStyle(style) => instance.style = Some(style.clone()),
                    RPrBase::RunFonts(fonts) => instance.fonts = Some(fonts.clone()),
                    RPrBase::Bold(b) => instance.bold = Some(*b),
                    RPrBase::ComplexScriptBold(b) => instance.complex_script_bold = Some(*b),
                    RPrBase::Italic(i) => instance.italic = Some(*i),
                    RPrBase::ComplexScriptItalic(i) => instance.complex_script_italic = Some(*i),
                    RPrBase::Capitals(caps) => instance.all_capitals = Some(*caps),
                    RPrBase::SmallCapitals(small_caps) => instance.all_small_capitals = Some(*small_caps),
                    RPrBase::Strikethrough(strike) => instance.strikethrough = Some(*strike),
                    RPrBase::DoubleStrikethrough(dbl_strike) => instance.double_strikethrough = Some(*dbl_strike),
                    RPrBase::Outline(outline) => instance.outline = Some(*outline),
                    RPrBase::Shadow(shadow) => instance.shadow = Some(*shadow),
                    RPrBase::Emboss(emboss) => instance.emboss = Some(*emboss),
                    RPrBase::Imprint(imprint) => instance.imprint = Some(*imprint),
                    RPrBase::NoProofing(no_proof) => instance.no_proofing = Some(*no_proof),
                    RPrBase::SnapToGrid(snap_to_grid) => instance.snap_to_grid = Some(*snap_to_grid),
                    RPrBase::Vanish(vanish) => instance.vanish = Some(*vanish),
                    RPrBase::WebHidden(web_hidden) => instance.web_hidden = Some(*web_hidden),
                    RPrBase::Color(color) => instance.color = Some(*color),
                    RPrBase::Spacing(spacing) => instance.spacing = Some(*spacing),
                    RPrBase::Width(width) => instance.width = Some(*width),
                    RPrBase::Kerning(kerning) => instance.kerning = Some(*kerning),
                    RPrBase::Position(pos) => instance.position = Some(*pos),
                    RPrBase::FontSize(size) => instance.font_size = Some(*size),
                    RPrBase::ComplexScriptFontSize(cs_size) => instance.complex_script_font_size = Some(*cs_size),
                    RPrBase::Highlight(color) => instance.highlight = Some(*color),
                    RPrBase::Underline(u) => instance.underline = Some(*u),
                    RPrBase::Effect(effect) => instance.effect = Some(*effect),
                    RPrBase::Border(border) => instance.border = Some(*border),
                    RPrBase::Shading(shd) => instance.shading = Some(*shd),
                    RPrBase::FitText(fit_text) => instance.fit_text = Some(*fit_text),
                    RPrBase::VerticalAlignment(align) => instance.vertical_alignment = Some(*align),
                    RPrBase::Rtl(rtl) => instance.rtl = Some(*rtl),
                    RPrBase::ComplexScript(cs) => instance.complex_script = Some(*cs),
                    RPrBase::EmphasisMark(em) => instance.emphasis_mark = Some(*em),
                    RPrBase::Language(lang) => instance.language = Some(lang.clone()),
                    RPrBase::EastAsianLayout(ea_layout) => instance.east_asian_layout = Some(*ea_layout),
                    RPrBase::SpecialVanish(vanish) => instance.special_vanish = Some(*vanish),
                    RPrBase::OMath(o_math) => instance.o_math = Some(*o_math),
                }

                instance
            })
    }

    pub fn update_with(mut self, other: Self) -> Self {
        self.style = other.style.or(self.style);
        self.fonts = match (self.fonts, other.fonts) {
            (Some(lhs), Some(rhs)) => Some(lhs.update_with(rhs)),
            (lhs, rhs) => rhs.or(lhs),
        };
        self.bold = other.bold.or(self.bold);
        self.complex_script_bold = other.complex_script_bold.or(self.complex_script_bold);
        self.italic = other.italic.or(self.italic);
        self.complex_script_italic = other.complex_script_italic.or(self.complex_script_italic);
        self.all_capitals = other.all_capitals.or(self.all_capitals);
        self.all_small_capitals = other.all_small_capitals.or(self.all_small_capitals);
        self.strikethrough = other.strikethrough.or(self.strikethrough);
        self.double_strikethrough = other.double_strikethrough.or(self.double_strikethrough);
        self.outline = other.outline.or(self.outline);
        self.shadow = other.shadow.or(self.shadow);
        self.emboss = other.emboss.or(self.emboss);
        self.imprint = other.imprint.or(self.imprint);
        self.no_proofing = other.no_proofing.or(self.no_proofing);
        self.snap_to_grid = other.snap_to_grid.or(self.snap_to_grid);
        self.vanish = other.vanish.or(self.vanish);
        self.web_hidden = other.web_hidden.or(self.web_hidden);
        self.color = other.color.or(self.color);
        self.spacing = other.spacing.or(self.spacing);
        self.width = other.width.or(self.width);
        self.kerning = other.kerning.or(self.kerning);
        self.position = other.position.or(self.position);
        self.font_size = other.font_size.or(self.font_size);
        self.complex_script_font_size = other.complex_script_font_size.or(self.complex_script_font_size);
        self.highlight = other.highlight.or(self.highlight);
        self.underline = other.underline.or(self.underline);
        self.effect = other.effect.or(self.effect);
        self.border = other.border.or(self.border);
        self.shading = other.shading.or(self.shading);
        self.fit_text = other.fit_text.or(self.fit_text);
        self.vertical_alignment = other.vertical_alignment.or(self.vertical_alignment);
        self.rtl = other.rtl.or(self.rtl);
        self.complex_script = other.complex_script.or(self.complex_script);
        self.emphasis_mark = other.emphasis_mark.or(self.emphasis_mark);
        self.language = other.language.or(self.language);
        self.east_asian_layout = other.east_asian_layout.or(self.east_asian_layout);
        self.special_vanish = other.special_vanish.or(self.special_vanish);
        self.o_math = other.o_math.or(self.o_math);
        self
    }

    pub fn update_with_style_on_another_level(mut self, other: Self) -> Self {
        self.style = other.style.or(self.style);
        self.fonts = match (self.fonts, other.fonts) {
            (Some(lhs), Some(rhs)) => Some(lhs.update_with(rhs)),
            (lhs, rhs) => rhs.or(lhs),
        };
        self.bold = update_or_toggle_on_off(self.bold, other.bold);
        self.complex_script_bold = update_or_toggle_on_off(self.complex_script_bold, other.complex_script_bold);
        self.italic = update_or_toggle_on_off(self.italic, other.italic);
        self.complex_script_italic = update_or_toggle_on_off(self.complex_script_italic, other.complex_script_italic);
        self.all_capitals = update_or_toggle_on_off(self.all_capitals, other.all_capitals);
        self.all_small_capitals = update_or_toggle_on_off(self.all_small_capitals, other.all_small_capitals);
        self.strikethrough = update_or_toggle_on_off(self.strikethrough, other.strikethrough);
        self.double_strikethrough = update_or_toggle_on_off(self.double_strikethrough, other.double_strikethrough);
        self.outline = update_or_toggle_on_off(self.outline, other.outline);
        self.shadow = update_or_toggle_on_off(self.shadow, other.shadow);
        self.emboss = update_or_toggle_on_off(self.emboss, other.emboss);
        self.imprint = update_or_toggle_on_off(self.imprint, other.imprint);
        self.no_proofing = update_or_toggle_on_off(self.no_proofing, other.no_proofing);
        self.snap_to_grid = update_or_toggle_on_off(self.snap_to_grid, other.snap_to_grid);
        self.vanish = update_or_toggle_on_off(self.vanish, other.vanish);
        self.web_hidden = update_or_toggle_on_off(self.web_hidden, other.web_hidden);
        self.color = other.color.or(self.color);
        self.spacing = other.spacing.or(self.spacing);
        self.width = other.width.or(self.width);
        self.kerning = other.kerning.or(self.kerning);
        self.position = other.position.or(self.position);
        self.font_size = other.font_size.or(self.font_size);
        self.complex_script_font_size = other.complex_script_font_size.or(self.complex_script_font_size);
        self.highlight = other.highlight.or(self.highlight);
        self.underline = other.underline.or(self.underline);
        self.effect = other.effect.or(self.effect);
        self.border = other.border.or(self.border);
        self.shading = other.shading.or(self.shading);
        self.fit_text = other.fit_text.or(self.fit_text);
        self.vertical_alignment = other.vertical_alignment.or(self.vertical_alignment);
        self.rtl = update_or_toggle_on_off(self.rtl, other.rtl);
        self.complex_script = update_or_toggle_on_off(self.complex_script, self.complex_script);
        self.emphasis_mark = other.emphasis_mark.or(self.emphasis_mark);
        self.language = other.language.or(self.language);
        self.east_asian_layout = other.east_asian_layout.or(self.east_asian_layout);
        self.special_vanish = update_or_toggle_on_off(self.special_vanish, other.special_vanish);
        self.o_math = update_or_toggle_on_off(self.o_math, other.o_math);
        self
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ResolvedStyle {
    pub paragraph_properties: Box<ParagraphProperties>,
    pub run_properties: Box<RunProperties>,
    // pub table_properties: TblPrBase,
    // pub table_row_properties: TrPr,
    // pub table_cell_properties: TcPr,
}

impl ResolvedStyle {
    pub fn from_wml_style(style: &Style) -> Self {
        let paragraph_properties = Box::new(style
            .paragraph_properties
            .as_ref()
            .map(|p_pr| p_pr.base.clone())
            .unwrap_or_default());

        let run_properties = Box::new(style
            .run_properties
            .as_ref()
            .map(|r_pr| RunProperties::from_vec(&r_pr.r_pr_bases))
            .unwrap_or_default());

        Self { paragraph_properties, run_properties }
    }

    pub fn update_with(mut self, other: Self) -> Self {
        *self.paragraph_properties = self.paragraph_properties.update_with(*other.paragraph_properties);
        *self.run_properties = self.run_properties.update_with(*other.run_properties);
        self
    }

    pub fn update_with_style_on_another_level(mut self, other: Self) -> Self {
        *self.paragraph_properties = self.paragraph_properties.update_with(*other.paragraph_properties);
        *self.run_properties = self
            .run_properties
            .update_with_style_on_another_level(*other.run_properties);
        self
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Package {
    pub app_info: Option<AppInfo>,
    pub core: Option<Core>,
    pub main_document: Option<Box<Document>>,
    pub main_document_relationships: Vec<Relationship>,
    pub styles: Option<Box<Styles>>,
    pub settings: Option<Box<Settings>>,
    pub medias: Vec<PathBuf>,
    pub themes: HashMap<String, OfficeStyleSheet>,
}

impl Package {
    pub fn from_file(file_path: &Path) -> Result<Self, Box<dyn Error>> {
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
                    instance.main_document = Some(Box::new(Document::from_xml_element(&xml_node)?));
                }
                "word/_rels/document.xml.rels" => {
                    instance.main_document_relationships = zip_file_to_xml_node(&mut zip_file)?
                        .child_nodes
                        .iter()
                        .map(Relationship::from_xml_element)
                        .collect::<Result<Vec<_>, Box<dyn Error>>>()?;
                }
                "word/styles.xml" => {
                    let xml_node = zip_file_to_xml_node(&mut zip_file)?;
                    instance.styles = Some(Box::new(Styles::from_xml_element(&xml_node)?));
                }
                "word/settings.xml" => {
                    let xml_node = zip_file_to_xml_node(&mut zip_file)?;
                    instance.settings = Some(Box::new(Settings::from_xml_element(&xml_node)?));
                }
                path if path.starts_with("word/media/") => instance.medias.push(PathBuf::from(file_path)),
                path if path.starts_with("word/theme/") => {
                    let file_stem = match Path::new(path).file_stem().and_then(OsStr::to_str).map(String::from) {
                        Some(name) => name,
                        None => {
                            error!("Couldn't get file name of theme");
                            continue;
                        }
                    };
                    let style_sheet = OfficeStyleSheet::from_xml_element(&zip_file_to_xml_node(&mut zip_file)?)?;
                    instance.themes.insert(file_stem, style_sheet);
                }
                _ => (),
            }
        }

        Ok(instance)
    }

    pub fn resolve_default_style(&self) -> Option<ResolvedStyle> {
        self.styles
            .as_ref()
            .and_then(|styles| styles.document_defaults.as_ref())
            .map(|doc_defaults| {
                let run_properties = Box::new(doc_defaults
                    .run_properties_default
                    .as_ref()
                    .and_then(|r_pr_default| r_pr_default.0.as_ref())
                    .map(|r_pr| RunProperties::from_vec(&r_pr.r_pr_bases))
                    .unwrap_or_default()
                );

                let paragraph_properties = Box::new(doc_defaults
                    .paragraph_properties_default
                    .as_ref()
                    .and_then(|p_pr_default| p_pr_default.0.as_ref())
                    .map(|p_pr| p_pr.base.clone())
                    .unwrap_or_default()
                );

                ResolvedStyle{ run_properties, paragraph_properties }
            })
    }

    pub fn resolve_default_paragraph_style(&self) -> Option<ResolvedStyle> {
        let styles = self.styles
            .as_ref()
            .map(|styles| &styles.styles)?;

        let default_style = styles
            .iter()
            .find(|style| match (&style.style_type, &style.is_default) {
                (Some(StyleType::Paragraph), Some(true)) => true,
                _ => false
            })?;
        
        Some(ResolvedStyle::from_wml_style(default_style))
    }

    pub fn resolve_paragraph_style(&self, paragraph: &P) -> Option<ResolvedStyle> {
        paragraph
            .properties
            .as_ref()
            .and_then(|props| props.base.style.as_ref())
            .and_then(|style_name| self.resolve_style(style_name))
    }

    fn resolve_style<T: AsRef<str>>(&self, style_id: T) -> Option<ResolvedStyle> {
        // TODO(kalmar.robert) Use caching
        let styles = self
            .styles
            .as_ref()
            .map(|styles| &styles.styles)?;

        let top_most_style = styles
            .iter()
            .find(|style| {
                style
                    .style_id
                    .as_ref()
                    .filter(|s_id| (*s_id).as_str() == style_id.as_ref())
                    .is_some()
            })?;

        let style_hierarchy: Vec<&Style> = std::iter::successors(Some(top_most_style), |child_style| {
            styles.iter().find(|style| style.style_id == child_style.based_on)
        })
        .collect();

        Some(style_hierarchy
            .iter()
            .rev()
            .fold(Default::default(), |mut resolved_style: ResolvedStyle, style| {
                if let Some(style_p_pr) = &style.paragraph_properties {
                    *resolved_style.paragraph_properties =
                        resolved_style.paragraph_properties.update_with(style_p_pr.base.clone());
                }

                if let Some(style_r_pr) = &style.run_properties {
                    let folded_style_r_pr = RunProperties::from_vec(&style_r_pr.r_pr_bases);
                    *resolved_style.run_properties = resolved_style.run_properties.update_with(folded_style_r_pr);
                }

                resolved_style
            }))
    }

    pub fn resolve_style_inheritance(&self, paragraph: &P, run: &R) -> Option<ResolvedStyle> {
        let default_style = self.resolve_default_style();
        let paragraph_style = self
            .resolve_paragraph_style(paragraph)
            .or_else(|| self.resolve_default_paragraph_style());
        let run_style = resolve_run_style(run);

        let calced_style = match (paragraph_style, run_style) {
            (Some(p_style), Some(r_style)) => Some(p_style.update_with_style_on_another_level(r_style)),
            (p_style, r_style) => p_style.or(r_style),
        };

        match (default_style, calced_style) {
            (Some(def_style), Some(calced_style)) => Some(def_style.update_with(calced_style)),
            (def_style, calced_style) => def_style.or(calced_style),
        }
    }

    pub fn get_main_document_theme(&self) -> Option<&OfficeStyleSheet> {
        let theme_relation = self
            .main_document_relationships
            .iter()
            .find(|rel| rel.rel_type == THEME_RELATION_TYPE)?;

        let rel_target_file = Path::new(theme_relation.target.as_str())
            .file_stem()
            .and_then(OsStr::to_str)?;

        self.themes.get(rel_target_file)
    }

    pub fn get_main_document_section_properties(&self) -> Option<&SectPrContents> {
        self
            .main_document
            .as_ref()
            .and_then(|main_document| main_document.body.as_ref())
            .and_then(|body| body.section_properties.as_ref())
            .and_then(|sect_pr| sect_pr.contents.as_ref())
    }
}

pub fn resolve_run_style(run: &R) -> Option<ResolvedStyle> {
    run.run_properties
        .as_ref()
        .map(|props| RunProperties::from_vec(&props.r_pr_bases))
        .map(|run_props| ResolvedStyle {
            run_properties: Box::new(run_props),
            ..Default::default()
        })
}

fn update_or_toggle_on_off(lhs: Option<OnOff>, rhs: Option<OnOff>) -> Option<OnOff> {
    match (lhs, rhs) {
        (Some(lhs), Some(rhs)) => Some(lhs ^ rhs),
        (lhs, rhs) => rhs.or(lhs),
    }
}

#[cfg(test)]
mod tests {
    use super::{resolve_run_style, Package, ParagraphProperties, RunProperties};
    use crate::wml::{
        document::{
            Document, PPr, PPrBase, PPrGeneral, ParaRPr, RPr, RPrBase, TextAlignment, Underline, UnderlineType, P, R,
        },
        settings::Settings,
        styles::{DocDefaults, PPrDefault, RPrDefault, Style, Styles},
    };
    use msoffice_shared::docprops::{AppInfo, Core};

    #[test]
    fn test_size_of() {
        use std::mem::size_of;

        println!("sizeof Package: {}", size_of::<Package>());
        println!("sizeof AppInfo: {}", size_of::<AppInfo>());
        println!("sizeof Core: {}", size_of::<Core>());
        println!("sizeof Document: {}", size_of::<Document>());
        println!("sizeof Styles: {}", size_of::<Styles>());
        println!("sizeof Settings: {}", size_of::<Settings>());
    }

    fn doc_defaults_for_test() -> DocDefaults {
        let default_p_pr = PPr {
            base: PPrBase {
                start_on_next_page: Some(false),
                ..Default::default()
            },
            ..Default::default()
        };

        let default_r_pr = RPr {
            r_pr_bases: vec![RPrBase::Bold(true), RPrBase::Italic(false)],
            ..Default::default()
        };

        DocDefaults {
            paragraph_properties_default: Some(PPrDefault(Some(default_p_pr))),
            run_properties_default: Some(RPrDefault(Some(default_r_pr))),
        }
    }

    fn styles_for_test() -> Vec<Style> {
        let normal_style = Style {
            name: Some(String::from("Normal")),
            style_id: Some(String::from("Normal")),
            paragraph_properties: Some(PPrGeneral {
                base: PPrBase {
                    start_on_next_page: Some(true),
                    ..Default::default()
                },
                ..Default::default()
            }),
            run_properties: Some(RPr {
                r_pr_bases: vec![RPrBase::Italic(true)],
                ..Default::default()
            }),
            ..Default::default()
        };

        let child_style = Style {
            name: Some(String::from("Child")),
            style_id: Some(String::from("Child")),
            based_on: Some(String::from("Normal")),
            paragraph_properties: Some(PPrGeneral {
                base: PPrBase {
                    text_alignment: Some(TextAlignment::Center),
                    ..Default::default()
                },
                ..Default::default()
            }),
            run_properties: Some(RPr {
                r_pr_bases: vec![RPrBase::Underline(Underline {
                    value: Some(UnderlineType::Single),
                    ..Default::default()
                })],
                ..Default::default()
            }),
            ..Default::default()
        };

        vec![normal_style, child_style]
    }

    fn paragraph_for_test() -> P {
        P {
            properties: Some(PPr {
                base: PPrBase {
                    style: Some(String::from("Child")),
                    keep_lines_on_one_page: Some(true),
                    ..Default::default()
                },
                run_properties: Some(ParaRPr {
                    bases: vec![RPrBase::Bold(true), RPrBase::Italic(true)],
                    ..Default::default()
                }),
                ..Default::default()
            }),
            ..Default::default()
        }
    }

    fn run_for_test() -> R {
        R {
            run_properties: Some(RPr {
                r_pr_bases: vec![RPrBase::Bold(true), RPrBase::Italic(true)],
                ..Default::default()
            }),
            ..Default::default()
        }
    }

    #[test]
    pub fn test_resolve_default_style() {
        let package = Package {
            styles: Some(Box::new(Styles {
                document_defaults: Some(doc_defaults_for_test()),
                latent_styles: None,
                styles: Vec::new(),
            })),
            ..Default::default()
        };

        let default_style = package.resolve_default_style().unwrap();
        assert_eq!(
            *default_style.paragraph_properties,
            ParagraphProperties {
                start_on_next_page: Some(false),
                ..Default::default()
            },
        );
        assert_eq!(
            *default_style.run_properties,
            RunProperties {
                bold: Some(true),
                italic: Some(false),
                ..Default::default()
            }
        );
    }

    #[test]
    pub fn test_resolve_paragraph_style() {
        let package = Package {
            styles: Some(Box::new(Styles {
                document_defaults: None,
                latent_styles: None,
                styles: styles_for_test(),
            })),
            ..Default::default()
        };

        let paragraph_style = package.resolve_paragraph_style(&paragraph_for_test()).unwrap();
        assert_eq!(
            *paragraph_style.paragraph_properties,
            ParagraphProperties {
                start_on_next_page: Some(true),
                text_alignment: Some(TextAlignment::Center),
                ..Default::default()
            }
        );

        assert_eq!(
            *paragraph_style.run_properties,
            RunProperties {
                italic: Some(true),
                underline: Some(Underline {
                    value: Some(UnderlineType::Single),
                    ..Default::default()
                }),
                ..Default::default()
            }
        );
    }

    #[test]
    pub fn test_resolve_run_style() {
        let run_properties = resolve_run_style(&run_for_test()).unwrap();
        assert_eq!(
            *run_properties.run_properties,
            RunProperties {
                bold: Some(true),
                italic: Some(true),
                ..Default::default()
            }
        );
    }

    #[test]
    pub fn test_resolve_style_inheritance() {
        let package = Package {
            styles: Some(Box::new(Styles {
                document_defaults: Some(doc_defaults_for_test()),
                latent_styles: None,
                styles: styles_for_test(),
            })),
            ..Default::default()
        };

        let style = package
            .resolve_style_inheritance(&paragraph_for_test(), &run_for_test())
            .unwrap();
        assert_eq!(
            *style.paragraph_properties,
            ParagraphProperties {
                start_on_next_page: Some(true),
                text_alignment: Some(TextAlignment::Center),
                ..Default::default()
            }
        );
        assert_eq!(
            *style.run_properties,
            RunProperties {
                bold: Some(true),
                italic: Some(false),
                underline: Some(Underline {
                    value: Some(UnderlineType::Single),
                    ..Default::default()
                }),
                ..Default::default()
            }
        );
    }
}
