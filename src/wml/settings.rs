use msoffice_shared::sharedtypes::{OnOff, TwipsMeasure};
use super::{
    document::{Rel, Language},
    simpletypes::DecimalNumber,
};

/*
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Settings {
    // <xsd:element name="writeProtection" type="CT_WriteProtection" minOccurs="0"/>
    pub write_propection: Option<WriteProtection>,
    // <xsd:element name="view" type="CT_View" minOccurs="0"/>
    pub view: Option<View>,
    // <xsd:element name="zoom" type="CT_Zoom" minOccurs="0"/>
    pub zoom: Option<Zoom>,
    // <xsd:element name="removePersonalInformation" type="CT_OnOff" minOccurs="0"/>
    pub remove_personal_information: Option<OnOff>,
    // <xsd:element name="removeDateAndTime" type="CT_OnOff" minOccurs="0"/>
    pub remove_date_and_time: Option<OnOff>,
    // <xsd:element name="doNotDisplayPageBoundaries" type="CT_OnOff" minOccurs="0"/>
    pub do_not_display_page_boundaries: Option<OnOff>,
    // <xsd:element name="displayBackgroundShape" type="CT_OnOff" minOccurs="0"/>
    pub display_background_shape: Option<OnOff>,
    // <xsd:element name="printPostScriptOverText" type="CT_OnOff" minOccurs="0"/>
    pub print_post_script_over_text: Option<OnOff>,
    // <xsd:element name="printFractionalCharacterWidth" type="CT_OnOff" minOccurs="0"/>
    pub print_fractional_character_width: Option<OnOff>,
    // <xsd:element name="printFormsData" type="CT_OnOff" minOccurs="0"/>
    pub print_forms_data: Option<OnOff>,
    // <xsd:element name="embedTrueTypeFonts" type="CT_OnOff" minOccurs="0"/>
    pub embed_true_type_fonts: Option<OnOff>,
    // <xsd:element name="embedSystemFonts" type="CT_OnOff" minOccurs="0"/>
    pub embed_system_fonts: Option<OnOff>,
    // <xsd:element name="saveSubsetFonts" type="CT_OnOff" minOccurs="0"/>
    pub save_subset_fonts: Option<OnOff>,
    // <xsd:element name="saveFormsData" type="CT_OnOff" minOccurs="0"/>
    pub save_forms_data: Option<OnOff>,
    // <xsd:element name="mirrorMargins" type="CT_OnOff" minOccurs="0"/>
    pub mirror_margins: Option<OnOff>,
    // <xsd:element name="alignBordersAndEdges" type="CT_OnOff" minOccurs="0"/>
    pub align_borders_and_edges: Option<OnOff>,
    // <xsd:element name="bordersDoNotSurroundHeader" type="CT_OnOff" minOccurs="0"/>
    pub borders_do_not_surround_header: Option<OnOff>,
    // <xsd:element name="bordersDoNotSurroundFooter" type="CT_OnOff" minOccurs="0"/>
    pub borders_do_not_surround_footer: Option<OnOff>,
    // <xsd:element name="gutterAtTop" type="CT_OnOff" minOccurs="0"/>
    pub gutter_at_top: Option<OnOff>,
    // <xsd:element name="hideSpellingErrors" type="CT_OnOff" minOccurs="0"/>
    pub hide_spelling_errors: Option<OnOff>,
    // <xsd:element name="hideGrammaticalErrors" type="CT_OnOff" minOccurs="0"/>
    pub hide_grammatical_errors: Option<OnOff>,
    // <xsd:element name="activeWritingStyle" type="CT_WritingStyle" minOccurs="0" maxOccurs="unbounded"/>
    pub active_writing_styles: Vec<WritingStyle>,
    // <xsd:element name="proofState" type="CT_Proof" minOccurs="0"/>
    pub proof_state: Option<Proof>,
    // <xsd:element name="formsDesign" type="CT_OnOff" minOccurs="0"/>
    pub forms_design: Option<OnOff>,
    // <xsd:element name="attachedTemplate" type="CT_Rel" minOccurs="0"/>
    pub attached_template: Option<Rel>,
    // <xsd:element name="linkStyles" type="CT_OnOff" minOccurs="0"/>
    pub link_styles: Option<OnOff>,
    // <xsd:element name="stylePaneFormatFilter" type="CT_StylePaneFilter" minOccurs="0"/>
    pub style_pane_format_filter: Option<StylePaneFilter>,
    // <xsd:element name="stylePaneSortMethod" type="CT_StyleSort" minOccurs="0"/>
    pub style_pane_sort_method: Option<StyleSort>,
    // <xsd:element name="documentType" type="CT_DocType" minOccurs="0"/>
    pub document_type: Option<DocType>,
    // <xsd:element name="mailMerge" type="CT_MailMerge" minOccurs="0"/>
    pub mail_merge: Option<MailMerge>,
    // <xsd:element name="revisionView" type="CT_TrackChangesView" minOccurs="0"/>
    pub revision_view: Option<TrackChangesView>,
    // <xsd:element name="trackRevisions" type="CT_OnOff" minOccurs="0"/>
    pub track_revisions: Option<OnOff>,
    // <xsd:element name="doNotTrackMoves" type="CT_OnOff" minOccurs="0"/>
    pub do_not_track_moves: Option<OnOff>,
    // <xsd:element name="doNotTrackFormatting" type="CT_OnOff" minOccurs="0"/>
    pub do_not_track_formatting: Option<OnOff>,
    // <xsd:element name="documentProtection" type="CT_DocProtect" minOccurs="0"/>
    pub document_protection: Option<DocProtect>,
    // <xsd:element name="autoFormatOverride" type="CT_OnOff" minOccurs="0"/>
    pub auto_format_override: Option<OnOff>,
    // <xsd:element name="styleLockTheme" type="CT_OnOff" minOccurs="0"/>
    pub style_lock_theme: Option<OnOff>,
    // <xsd:element name="styleLockQFSet" type="CT_OnOff" minOccurs="0"/>
    pub style_lock_set: Option<OnOff>,
    // <xsd:element name="defaultTabStop" type="CT_TwipsMeasure" minOccurs="0"/>
    pub default_tab_stop: Option<TwipsMeasure>,
    // <xsd:element name="autoHyphenation" type="CT_OnOff" minOccurs="0"/>
    pub auto_hyphenation: Option<OnOff>,
    // <xsd:element name="consecutiveHyphenLimit" type="CT_DecimalNumber" minOccurs="0"/>
    pub consecutive_hyphen_limit: Option<DecimalNumber>,
    // <xsd:element name="hyphenationZone" type="CT_TwipsMeasure" minOccurs="0"/>
    pub hyphenation_zone: Option<TwipsMeasure>,
    // <xsd:element name="doNotHyphenateCaps" type="CT_OnOff" minOccurs="0"/>
    pub do_not_hyphenate_capitals: Option<OnOff>,
    // <xsd:element name="showEnvelope" type="CT_OnOff" minOccurs="0"/>
    pub show_envelope: Option<OnOff>,
    // <xsd:element name="summaryLength" type="CT_DecimalNumberOrPrecent" minOccurs="0"/>
    pub summary_length: Option<DecimalNumberOrPercent>,
    // <xsd:element name="clickAndTypeStyle" type="CT_String" minOccurs="0"/>
    pub click_and_type_style: Option<String>,
    // <xsd:element name="defaultTableStyle" type="CT_String" minOccurs="0"/>
    pub default_table_style: Option<String>,
    // <xsd:element name="evenAndOddHeaders" type="CT_OnOff" minOccurs="0"/>
    pub even_and_odd_headers: Option<OnOff>,
    // <xsd:element name="bookFoldRevPrinting" type="CT_OnOff" minOccurs="0"/>
    pub book_fold_revision_printing: Option<OnOff>,
    // <xsd:element name="bookFoldPrinting" type="CT_OnOff" minOccurs="0"/>
    pub book_fold_printing: Option<OnOff>,
    // <xsd:element name="bookFoldPrintingSheets" type="CT_DecimalNumber" minOccurs="0"/>
    pub book_fold_printing_sheets: Option<OnOff>,
    // <xsd:element name="drawingGridHorizontalSpacing" type="CT_TwipsMeasure" minOccurs="0"/>
    pub drawing_grid_horizontal_spacing: Option<TwipsMeasure>,
    // <xsd:element name="drawingGridVerticalSpacing" type="CT_TwipsMeasure" minOccurs="0"/>
    pub drawing_grid_vertical_spacing: Option<TwipsMeasure>,
    // <xsd:element name="displayHorizontalDrawingGridEvery" type="CT_DecimalNumber" minOccurs="0"/>
    pub display_horizontal_drawing_grid_every: Option<DecimalNumber>,
    // <xsd:element name="displayVerticalDrawingGridEvery" type="CT_DecimalNumber" minOccurs="0"/>
    pub display_vertical_drawing_grid_every: Option<DecimalNumber>,
    // <xsd:element name="doNotUseMarginsForDrawingGridOrigin" type="CT_OnOff" minOccurs="0"/>
    pub do_not_use_margins_for_drawing_grid_origin: Option<OnOff>,
    // <xsd:element name="drawingGridHorizontalOrigin" type="CT_TwipsMeasure" minOccurs="0"/>
    pub drawing_grid_horizontal_origin: Option<TwipsMeasure>,
    // <xsd:element name="drawingGridVerticalOrigin" type="CT_TwipsMeasure" minOccurs="0"/>
    pub drawing_grid_vertical_origin: Option<TwipsMeasure>,
//       <xsd:element name="doNotShadeFormData" type="CT_OnOff" minOccurs="0"/>
    pub do_not_shade_form_data: Option<OnOff>,
//       <xsd:element name="noPunctuationKerning" type="CT_OnOff" minOccurs="0"/>
    pub no_punctuation_kerning: Option<OnOff>,
//       <xsd:element name="characterSpacingControl" type="CT_CharacterSpacing" minOccurs="0"/>
    pub character_spacing_control: Option<CharacterSpacing>,
//       <xsd:element name="printTwoOnOne" type="CT_OnOff" minOccurs="0"/>
    pub print_two_on_one: Option<OnOff>,
//       <xsd:element name="strictFirstAndLastChars" type="CT_OnOff" minOccurs="0"/>
    pub strict_first_and_last_chars: Option<OnOff>,
//       <xsd:element name="noLineBreaksAfter" type="CT_Kinsoku" minOccurs="0"/>
    pub no_line_breaks_after: Option<Kinsoku>,
//       <xsd:element name="noLineBreaksBefore" type="CT_Kinsoku" minOccurs="0"/>
    pub no_line_breaks_before: Option<Kinsoku>,
//       <xsd:element name="savePreviewPicture" type="CT_OnOff" minOccurs="0"/>
    pub save_preview_picture: Option<OnOff>,
//       <xsd:element name="doNotValidateAgainstSchema" type="CT_OnOff" minOccurs="0"/>
    pub do_not_validate_against_schema: Option<OnOff>,
//       <xsd:element name="saveInvalidXml" type="CT_OnOff" minOccurs="0"/>
    pub save_invalid_xml: Option<OnOff>,
//       <xsd:element name="ignoreMixedContent" type="CT_OnOff" minOccurs="0"/>
    pub ignore_mixed_content: Option<OnOff>,
//       <xsd:element name="alwaysShowPlaceholderText" type="CT_OnOff" minOccurs="0"/>
    pub always_show_placeholder_text: Option<OnOff>,
//       <xsd:element name="doNotDemarcateInvalidXml" type="CT_OnOff" minOccurs="0"/>
    pub do_not_demarcate_invalid_xml: Option<OnOff>,
//       <xsd:element name="saveXmlDataOnly" type="CT_OnOff" minOccurs="0"/>
    pub save_xml_data_only: Option<OnOff>,
//       <xsd:element name="useXSLTWhenSaving" type="CT_OnOff" minOccurs="0"/>
    pub use_xslt_when_saving: Option<OnOff>,
//       <xsd:element name="saveThroughXslt" type="CT_SaveThroughXslt" minOccurs="0"/>
    pub save_through_xslt: Option<SaveThroughXslt>,
//       <xsd:element name="showXMLTags" type="CT_OnOff" minOccurs="0"/>
    pub show_xml_tags: Option<OnOff>,
//       <xsd:element name="alwaysMergeEmptyNamespace" type="CT_OnOff" minOccurs="0"/>
    pub always_merge_empty_namespace: Option<OnOff>,
//       <xsd:element name="updateFields" type="CT_OnOff" minOccurs="0"/>
    pub update_fields: Option<OnOff>,
//       <xsd:element name="footnotePr" type="CT_FtnDocProps" minOccurs="0"/>
    pub footnote_properties: Option<FtnDocProps>,
//       <xsd:element name="endnotePr" type="CT_EdnDocProps" minOccurs="0"/>
    pub endnote_properties: Option<EdnDocProps>,
//       <xsd:element name="compat" type="CT_Compat" minOccurs="0"/>
    pub compatibility: Option<Compat>,
//       <xsd:element name="docVars" type="CT_DocVars" minOccurs="0"/>
    pub document_variables: Option<DocVars>,
//       <xsd:element name="rsids" type="CT_DocRsids" minOccurs="0"/>
    pub revision_ids: Option<DocRsids>,
//       <xsd:element ref="m:mathPr" minOccurs="0" maxOccurs="1"/>
//       <xsd:element name="attachedSchema" type="CT_String" minOccurs="0" maxOccurs="unbounded"/>
    pub attached_schemas: Vec<String>,
//       <xsd:element name="themeFontLang" type="CT_Language" minOccurs="0" maxOccurs="1"/>
    pub theme_font_lang: Option<Language>,
//       <xsd:element name="clrSchemeMapping" type="CT_ColorSchemeMapping" minOccurs="0"/>
    pub color_scheme_mapping: Option<ColorSchemeMapping>,
//       <xsd:element name="doNotIncludeSubdocsInStats" type="CT_OnOff" minOccurs="0"/>
    pub do_not_include_subdocs_in_stats: Option<OnOff>,
//       <xsd:element name="doNotAutoCompressPictures" type="CT_OnOff" minOccurs="0"/>
    pub do_not_auto_compress_pictures: Option<OnOff>,
//       <xsd:element name="forceUpgrade" type="CT_Empty" minOccurs="0" maxOccurs="1"/>
    pub force_upgrade: bool,
//       <xsd:element name="captions" type="CT_Captions" minOccurs="0" maxOccurs="1"/>
    pub captions: Option<Captions>,
//       <xsd:element name="readModeInkLockDown" type="CT_ReadingModeInkLockDown" minOccurs="0"/>
    pub read_move_ink_lock_down: Option<ReadingModeInkLockDown>,
//       <xsd:element name="smartTagType" type="CT_SmartTagType" minOccurs="0" maxOccurs="unbounded"/>
    pub smart_tag_types: Vec<SmartTagType>,
//       <xsd:element ref="sl:schemaLibrary" minOccurs="0" maxOccurs="1"/>
//       <xsd:element name="doNotEmbedSmartTags" type="CT_OnOff" minOccurs="0"/>
    pub do_not_embed_smart_tags: Option<OnOff>,
//       <xsd:element name="decimalSymbol" type="CT_String" minOccurs="0" maxOccurs="1"/>
    pub decimal_symbol: Option<String>,
//       <xsd:element name="listSeparator" type="CT_String" minOccurs="0" maxOccurs="1"/>
    pub list_separator: Option<String>,
}
*/