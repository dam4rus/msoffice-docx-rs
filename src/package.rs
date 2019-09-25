use crate::{
    error::NoSuchStyleError,
    wml::{
        document::{
            Border, Color, Document, EastAsianLayout, Em, FitText, Fonts, HighlightColor, HpsMeasure, Language,
            PPrBase, RPrBase, Shd, SignedHpsMeasure, SignedTwipsMeasure, TextEffect, Underline, P, R,
        },
        simpletypes::TextScale,
        styles::{PPrDefault, RPrDefault, Style, Styles, TblStylePr},
        table::{TcPr, TrPr},
    },
};
use msoffice_shared::{
    docprops::{AppInfo, Core},
    relationship::Relationship,
    sharedtypes::{OnOff, VerticalAlignRun},
    xml::zip_file_to_xml_node,
};
use std::{
    fs::File,
    path::{Path, PathBuf},
};
use zip::ZipArchive;

type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

type ParagraphProperties = PPrBase;

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
    pub size: Option<HpsMeasure>,
    pub complex_script_size: Option<HpsMeasure>,
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
                    RPrBase::Size(size) => instance.size = Some(*size),
                    RPrBase::ComplexScriptSize(cs_size) => instance.complex_script_size = Some(*cs_size),
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

    pub fn update_with(mut self, other: &Self) -> Self {
        self.style = other.style.as_ref().cloned().or(self.style);
        self.fonts = other.fonts.as_ref().cloned().or(self.fonts);
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
        self.size = other.size.or(self.size);
        self.complex_script_size = other.complex_script_size.or(self.complex_script_size);
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
        self.language = other.language.as_ref().cloned().or(self.language);
        self.east_asian_layout = other.east_asian_layout.or(self.east_asian_layout);
        self.special_vanish = other.special_vanish.or(self.special_vanish);
        self.o_math = other.o_math.or(self.o_math);
        self
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ResolvedStyle {
    pub paragraph_properties: ParagraphProperties,
    pub run_properties: RunProperties,
    // pub table_properties: TblPrBase,
    // pub table_row_properties: TrPr,
    // pub table_cell_properties: TcPr,
}

impl ResolvedStyle {
    pub fn update_with(mut self, other: &Self) -> Self {
        self.paragraph_properties = self.paragraph_properties.update_with(&other.paragraph_properties);
        self.run_properties = self.run_properties.update_with(&other.run_properties);
        self
    }
}

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

    pub fn resolve_default_style(&self) -> Option<ResolvedStyle> {
        self.styles
            .as_ref()
            .and_then(|styles| styles.document_defaults.as_ref())
            .map(|doc_defaults| {
                let mut resolved_style: ResolvedStyle = Default::default();
                if let Some(RPrDefault(Some(ref r_pr))) = doc_defaults.run_properties_default {
                    resolved_style.run_properties = RunProperties::from_vec(&r_pr.r_pr_bases)
                }

                if let Some(PPrDefault(Some(ref p_pr))) = doc_defaults.paragraph_properties_default {
                    resolved_style.paragraph_properties = p_pr.base.clone();
                }

                resolved_style
            })
    }

    pub fn resolve_style<T: AsRef<str>>(&self, style_id: T) -> std::result::Result<ResolvedStyle, NoSuchStyleError> {
        // TODO(kalmar.robert) Use caching
        let styles = self
            .styles
            .as_ref()
            .map(|styles| &styles.styles)
            .ok_or(NoSuchStyleError {})?;

        let top_most_style = styles
            .iter()
            .find(|style| {
                style
                    .style_id
                    .as_ref()
                    .filter(|s_id| (*s_id).as_str() == style_id.as_ref())
                    .is_some()
            })
            .ok_or(NoSuchStyleError {})?;

        let style_hierarchy: Vec<&Style> = std::iter::successors(Some(top_most_style), |child_style| {
            styles.iter().find(|style| style.style_id == child_style.based_on)
        })
        .collect();

        Ok(style_hierarchy
            .iter()
            .rev()
            .fold(Default::default(), |mut resolved_style: ResolvedStyle, style| {
                if let Some(style_p_pr) = &style.paragraph_properties {
                    resolved_style.paragraph_properties =
                        resolved_style.paragraph_properties.update_with(&style_p_pr.base);
                }

                if let Some(style_r_pr) = &style.run_properties {
                    let folded_style_r_pr = RunProperties::from_vec(&style_r_pr.r_pr_bases);
                    resolved_style.run_properties = resolved_style.run_properties.update_with(&folded_style_r_pr);
                }

                resolved_style
            }))
    }

    pub fn resolve_paragraph_style(
        &self,
        paragraph: &P,
    ) -> std::result::Result<Option<ResolvedStyle>, NoSuchStyleError> {
        let default_style = self.resolve_default_style();
        let paragraph_style = paragraph
            .properties
            .as_ref()
            .and_then(|props| props.base.style.as_ref())
            .map(|style_name| self.resolve_style(style_name))
            .transpose()?;

        Ok(match (default_style, paragraph_style) {
            (Some(def_style), Some(par_style)) => Some(def_style.update_with(&par_style)),
            (def_style, par_style) => def_style.or(par_style),
        })
    }

    pub fn resolve_run_style(
        &self,
        paragraph: &P,
        run: &R,
    ) -> std::result::Result<Option<ResolvedStyle>, NoSuchStyleError> {
        let resolved_par_style = self.resolve_paragraph_style(paragraph)?;
        let run_properties = run
            .run_properties
            .as_ref()
            .map(|props| RunProperties::from_vec(&props.r_pr_bases));

        Ok(match (resolved_par_style, run_properties) {
            (Some(par_style), Some(run_props)) => Some(ResolvedStyle {
                run_properties: par_style.run_properties.update_with(&run_props),
                ..par_style
            }),
            (_, Some(run_props)) => Some(ResolvedStyle {
                run_properties: run_props,
                ..Default::default()
            }),
            (par_style, _) => par_style,
        })
    }
}

#[cfg(test)]
mod tests {
    #[test]
    pub fn test_size_of() {
        use std::mem::size_of;

        println!("{}", size_of::<super::ResolvedStyle>());
    }
}
