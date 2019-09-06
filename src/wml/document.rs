use super::error::ParseHexColorError;
use msoffice_shared::{
    drawingml::{parse_hex_color_rgb, HexColorRGB},
    error::{
        MissingAttributeError, MissingChildNodeError, NotGroupMemberError, ParseBoolError, PatternRestrictionError,
    },
    relationship::RelationshipId,
    sharedtypes::{
        CalendarType, Lang, OnOff, PositiveUniversalMeasure, TwipsMeasure, UniversalMeasure, UniversalMeasureUnit,
        VerticalAlignRun,
    },
    xml::{parse_xml_bool, XmlNode},
};
use regex::Regex;
use std::str::FromStr;

pub type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

pub type UcharHexNumber = u8;
pub type LongHexNumber = String; // length=4
pub type ShortHexNumber = String; // length=2
pub type UnqualifiedPercentage = i32;
pub type DecimalNumber = i32;
pub type UnsignedDecimalNumber = u32;
pub type DateTime = String;
pub type MacroName = String; // maxLength=33
pub type EightPointMeasure = u64;
pub type PointMeasure = u64;
pub type TextScalePercent = f64; // pattern=0*(600|([0-5]?[0-9]?[0-9]))%
pub type TextScaleDecimal = i32; // 0 <= n <= 600
pub type TextScale = TextScalePercent;

fn parse_text_scale_percent(s: &str) -> Result<f64> {
    let re = Regex::new("^0*(600|([0-5]?[0-9]?[0-9]))%$").expect("valid regexp should be provided");
    let captures = re.captures(s).ok_or_else(|| PatternRestrictionError::NoMatch)?;
    Ok(captures[1].parse::<i32>()? as f64 / 100.0)
}

#[cfg(test)]
#[test]
pub fn test_parse_text_scale_percent() {
    assert_eq!(parse_text_scale_percent("100%").unwrap(), 1.0);
    assert_eq!(parse_text_scale_percent("600%").unwrap(), 6.0);
    assert_eq!(parse_text_scale_percent("333%").unwrap(), 3.33);
    assert_eq!(parse_text_scale_percent("0%").unwrap(), 0.0);
}

fn parse_on_off_xml_element(xml_node: &XmlNode) -> std::result::Result<Option<OnOff>, ParseBoolError> {
    xml_node
        .attributes
        .get("val")
        .map(|val| parse_xml_bool(val))
        .transpose()
}

#[derive(Default, Debug, Clone, PartialEq)]
pub struct Charset {
    pub value: Option<UcharHexNumber>,
    pub character_set: Option<String>,
}

impl Charset {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Charset = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => instance.value = Some(UcharHexNumber::from_str_radix(value, 16)?),
                "characterSet" => instance.character_set = Some(value.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum DecimalNumberOrPercent {
    Int(UnqualifiedPercentage),
    Percentage(String),
}

// pub enum TextScale {
//     Percent(TextScalePercent),
//     Decimal(TextScaleDecimal),
// }

#[derive(Debug, Clone, EnumString, PartialEq)]
pub enum ThemeColor {
    #[strum(serialize = "dark1")]
    Dark1,
    #[strum(serialize = "light1")]
    Light1,
    #[strum(serialize = "dark2")]
    Dark2,
    #[strum(serialize = "light2")]
    Light2,
    #[strum(serialize = "accent1")]
    Accent1,
    #[strum(serialize = "accent2")]
    Accent2,
    #[strum(serialize = "accent3")]
    Accent3,
    #[strum(serialize = "accent4")]
    Accent4,
    #[strum(serialize = "accent5")]
    Accent5,
    #[strum(serialize = "accent6")]
    Accent6,
    #[strum(serialize = "hyperlink")]
    Hyperlink,
    #[strum(serialize = "followedHyperlink")]
    FollowedHyperlink,
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "background1")]
    Background1,
    #[strum(serialize = "text1")]
    Text1,
    #[strum(serialize = "background2")]
    Background2,
    #[strum(serialize = "text2")]
    Text2,
}

#[derive(Debug, Clone, EnumString, PartialEq)]
pub enum HighlightColor {
    #[strum(serialize = "black")]
    Black,
    #[strum(serialize = "blue")]
    Blue,
    #[strum(serialize = "cyan")]
    Cyan,
    #[strum(serialize = "green")]
    Green,
    #[strum(serialize = "magenta")]
    Magenta,
    #[strum(serialize = "red")]
    Red,
    #[strum(serialize = "yellow")]
    Yellow,
    #[strum(serialize = "white")]
    White,
    #[strum(serialize = "darkBlue")]
    DarkBlue,
    #[strum(serialize = "darkCyan")]
    DarkCyan,
    #[strum(serialize = "darkGreen")]
    DarkGreen,
    #[strum(serialize = "darkMagenta")]
    DarkMagenta,
    #[strum(serialize = "darkRed")]
    DarkRed,
    #[strum(serialize = "darkYellow")]
    DarkYellow,
    #[strum(serialize = "darkGray")]
    DarkGray,
    #[strum(serialize = "lightGray")]
    LightGray,
    #[strum(serialize = "none")]
    None,
}

#[derive(Debug, Clone, PartialEq)]
pub enum HexColor {
    Auto,
    RGB(HexColorRGB),
}

impl FromStr for HexColor {
    type Err = ParseHexColorError;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        match s {
            "auto" => Ok(HexColor::Auto),
            _ => Ok(HexColor::RGB(parse_hex_color_rgb(s)?)),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum SignedTwipsMeasure {
    Decimal(i32),
    UniversalMeasure(UniversalMeasure),
}

impl FromStr for SignedTwipsMeasure {
    // TODO custom error type
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        // TODO maybe use TryFrom instead?
        if let Ok(value) = s.parse::<i32>() {
            Ok(SignedTwipsMeasure::Decimal(value))
        } else {
            Ok(SignedTwipsMeasure::UniversalMeasure(s.parse()?))
        }
    }
}

impl SignedTwipsMeasure {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let val = xml_node
            .attributes
            .get("val")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
        Ok(val.parse()?)
    }
}

#[cfg(test)]
impl SignedTwipsMeasure {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} val="123.456mm"></{node_name}>"#, node_name = node_name)
    }

    pub fn test_instance() -> Self {
        SignedTwipsMeasure::UniversalMeasure(UniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter))
    }
}

#[cfg(test)]
#[test]
pub fn test_signed_twips_measure_from_str() {
    use msoffice_shared::sharedtypes::UniversalMeasureUnit;

    assert_eq!(
        SignedTwipsMeasure::from_str("-123").unwrap(),
        SignedTwipsMeasure::Decimal(-123),
    );

    assert_eq!(
        SignedTwipsMeasure::from_str("123").unwrap(),
        SignedTwipsMeasure::Decimal(123),
    );

    assert_eq!(
        SignedTwipsMeasure::from_str("123mm").unwrap(),
        SignedTwipsMeasure::UniversalMeasure(UniversalMeasure::new(123.0, UniversalMeasureUnit::Millimeter)),
    );
}

#[cfg(test)]
#[test]
pub fn test_signed_twips_measure_from_xml() {
    let xml = SignedTwipsMeasure::test_xml("signedTwipsMeasure");
    let signed_twips_measure = SignedTwipsMeasure::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(signed_twips_measure, SignedTwipsMeasure::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub enum HpsMeasure {
    Decimal(u64),
    UniversalMeasure(PositiveUniversalMeasure),
}

impl FromStr for HpsMeasure {
    type Err = Box<dyn ::std::error::Error>;

    fn from_str(s: &str) -> Result<Self> {
        if let Ok(value) = s.parse::<u64>() {
            Ok(HpsMeasure::Decimal(value))
        } else {
            Ok(HpsMeasure::UniversalMeasure(s.parse()?))
        }
    }
}

impl HpsMeasure {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let val = xml_node
            .attributes
            .get("val")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
        Ok(val.parse()?)
    }
}

#[cfg(test)]
impl HpsMeasure {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} val="123.456mm"></{node_name}>"#, node_name = node_name)
    }

    pub fn test_instance() -> Self {
        HpsMeasure::UniversalMeasure(PositiveUniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter))
    }
}

#[cfg(test)]
#[test]
pub fn test_hps_measure_from_str() {
    use msoffice_shared::sharedtypes::UniversalMeasureUnit;

    assert_eq!("123".parse::<HpsMeasure>().unwrap(), HpsMeasure::Decimal(123));
    assert_eq!(
        "123.456mm".parse::<HpsMeasure>().unwrap(),
        HpsMeasure::UniversalMeasure(PositiveUniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter)),
    );
}

#[cfg(test)]
#[test]
pub fn test_hps_measure_from_xml() {
    let xml = HpsMeasure::test_xml("hpsMeasure");
    let hps_measure = HpsMeasure::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(hps_measure, HpsMeasure::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub enum SignedHpsMeasure {
    Decimal(i32),
    UniversalMeasure(UniversalMeasure),
}

impl FromStr for SignedHpsMeasure {
    // TODO custom error type
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> ::std::result::Result<Self, Self::Err> {
        // TODO maybe use TryFrom instead?
        if let Ok(value) = s.parse::<i32>() {
            Ok(SignedHpsMeasure::Decimal(value))
        } else {
            Ok(SignedHpsMeasure::UniversalMeasure(s.parse()?))
        }
    }
}

impl SignedHpsMeasure {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let val = xml_node
            .attributes
            .get("val")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
        Ok(val.parse()?)
    }
}

#[cfg(test)]
impl SignedHpsMeasure {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} val="123.456mm"></{node_name}>"#, node_name = node_name)
    }

    pub fn test_instance() -> Self {
        SignedHpsMeasure::UniversalMeasure(UniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter))
    }
}

#[cfg(test)]
#[test]
pub fn test_signed_hps_measure_from_str() {
    use msoffice_shared::sharedtypes::UniversalMeasureUnit;

    assert_eq!(
        SignedHpsMeasure::from_str("-123").unwrap(),
        SignedHpsMeasure::Decimal(-123),
    );

    assert_eq!(
        SignedHpsMeasure::from_str("123").unwrap(),
        SignedHpsMeasure::Decimal(123),
    );

    assert_eq!(
        SignedHpsMeasure::from_str("123mm").unwrap(),
        SignedHpsMeasure::UniversalMeasure(UniversalMeasure::new(123.0, UniversalMeasureUnit::Millimeter)),
    );
}

#[cfg(test)]
#[test]
pub fn test_signed_hps_measure_from_xml() {
    let xml = SignedHpsMeasure::test_xml("signedHpsMeasure");
    let hps_measure = SignedHpsMeasure::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(hps_measure, SignedHpsMeasure::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct Color {
    pub value: HexColor,
    pub theme_color: Option<ThemeColor>,
    pub theme_tint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,
}

impl Color {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut value = None;
        let mut theme_color = None;
        let mut theme_tint = None;
        let mut theme_shade = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => value = Some(attr_value.parse()?),
                "themeColor" => theme_color = Some(attr_value.parse()?),
                "themeTint" => theme_tint = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                "themeShade" => theme_shade = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                _ => (),
            }
        }

        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;

        Ok(Self {
            value,
            theme_color,
            theme_tint,
            theme_shade,
        })
    }
}

#[cfg(test)]
impl Color {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} val="ffffff" themeColor="accent1" themeTint="ff" themeShade="ff">
        </{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            value: HexColor::RGB([0xff, 0xff, 0xff]),
            theme_color: Some(ThemeColor::Accent1),
            theme_tint: Some(0xff),
            theme_shade: Some(0xff),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_color_from_xml() {
    let xml = Color::test_xml("color");
    let color = Color::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(color, Color::test_instance());
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum ProofErrType {
    #[strum(serialize = "spellStart")]
    SpellingStart,
    #[strum(serialize = "spellEnd")]
    SpellingEnd,
    #[strum(serialize = "gramStart")]
    GrammarStart,
    #[strum(serialize = "gramEnd")]
    GrammarEnd,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ProofErr {
    pub error_type: ProofErrType,
}

impl ProofErr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let type_attr = xml_node
            .attributes
            .get("type")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "type"))?;

        Ok(Self {
            error_type: type_attr.parse()?,
        })
    }
}

#[cfg(test)]
impl ProofErr {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} type="spellStart"></{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            error_type: ProofErrType::SpellingStart,
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_proof_err_from_xml() {
    let xml = ProofErr::test_xml("proofErr");
    let proof_err = ProofErr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(proof_err, ProofErr::test_instance());
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum EdGrp {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "everyone")]
    Everyone,
    #[strum(serialize = "administrators")]
    Administrators,
    #[strum(serialize = "contributors")]
    Contributors,
    #[strum(serialize = "editors")]
    Editors,
    #[strum(serialize = "owners")]
    Owners,
    #[strum(serialize = "current")]
    Current,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum DisplacedByCustomXml {
    #[strum(serialize = "next")]
    Next,
    #[strum(serialize = "prev")]
    Prev,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Perm {
    pub id: String,
    pub displaced_by_custom_xml: Option<DisplacedByCustomXml>,
}

impl Perm {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut id = None;
        let mut displaced_by_custom_xml = None;
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "id" => id = Some(value.clone()),
                "displacedByCustomXml" => displaced_by_custom_xml = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            id: id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?,
            displaced_by_custom_xml,
        })
    }
}

#[cfg(test)]
impl Perm {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} id="Some id", displacedByCustomXml="next"></{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            id: String::from("Some id"),
            displaced_by_custom_xml: Some(DisplacedByCustomXml::Next),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_perm_from_xml() {
    let xml = Perm::test_xml("perm");
    let perm = Perm::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(perm, Perm::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct PermStart {
    pub permission: Perm,
    pub editor_group: Option<EdGrp>,
    pub editor: Option<String>,
    pub first_column: Option<DecimalNumber>,
    pub last_column: Option<DecimalNumber>,
}

impl PermStart {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let permission = Perm::from_xml_element(xml_node)?;
        let mut editor_group = None;
        let mut editor = None;
        let mut first_column = None;
        let mut last_column = None;
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "edGrp" => editor_group = Some(value.parse()?),
                "ed" => editor = Some(value.clone()),
                "colFirst" => first_column = Some(value.parse()?),
                "colLast" => last_column = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            permission,
            editor_group,
            editor,
            first_column,
            last_column,
        })
    }
}

#[cfg(test)]
impl PermStart {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} id="Some id" displacedByCustomXml="next" edGrp="everyone" ed="rfrostkalmar@gmail.com" colFirst="0" colLast="1">
        </{node_name}>"#, node_name=node_name)
    }

    pub fn test_instance() -> Self {
        Self {
            permission: Perm::test_instance(),
            editor_group: Some(EdGrp::Everyone),
            editor: Some(String::from("rfrostkalmar@gmail.com")),
            first_column: Some(0),
            last_column: Some(1),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_perm_start_from_xml() {
    let xml = PermStart::test_xml("permStart");
    let perm_start = PermStart::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(perm_start, PermStart::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct Markup {
    pub id: DecimalNumber,
}

impl Markup {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let id_attr = xml_node
            .attributes
            .get("id")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?;

        Ok(Self { id: id_attr.parse()? })
    }
}

#[cfg(test)]
impl Markup {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} id="0"></{node_name}>"#, node_name = node_name)
    }

    pub fn test_instance() -> Self {
        Self { id: 0 }
    }
}

#[cfg(test)]
#[test]
pub fn test_markup_from_xml() {
    let xml = Markup::test_xml("markup");
    let markup = Markup::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(markup, Markup::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct MarkupRange {
    pub base: Markup,
    pub displaced_by_custom_xml: Option<DisplacedByCustomXml>,
}

impl MarkupRange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = Markup::from_xml_element(xml_node)?;
        let displaced_by_custom_xml = xml_node
            .attributes
            .get("displacedByCustomXml")
            .map(|value| value.parse())
            .transpose()?;

        Ok(Self {
            base,
            displaced_by_custom_xml,
        })
    }
}

#[cfg(test)]
impl MarkupRange {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} id="0" displacedByCustomXml="next"></{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            base: Markup::test_instance(),
            displaced_by_custom_xml: Some(DisplacedByCustomXml::Next),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_markup_range_from_xml() {
    let xml = MarkupRange::test_xml("markupRange");
    let markup_range = MarkupRange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(markup_range, MarkupRange::test_instance());
}
#[derive(Debug, Clone, PartialEq)]
pub struct BookmarkRange {
    pub base: MarkupRange,
    pub first_column: Option<DecimalNumber>,
    pub last_column: Option<DecimalNumber>,
}

impl BookmarkRange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = MarkupRange::from_xml_element(xml_node)?;
        let first_column = xml_node
            .attributes
            .get("colFirst")
            .map(|value| value.parse())
            .transpose()?;

        let last_column = xml_node
            .attributes
            .get("colLast")
            .map(|value| value.parse())
            .transpose()?;

        Ok(Self {
            base,
            first_column,
            last_column,
        })
    }
}

#[cfg(test)]
impl BookmarkRange {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} id="0" displacedByCustomXml="next" colFirst="0" colLast="1">
        </{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            base: MarkupRange::test_instance(),
            first_column: Some(0),
            last_column: Some(1),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_bookmark_range_from_xml() {
    let xml = BookmarkRange::test_xml("bookmarkRange");
    let bookmark_range = BookmarkRange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(bookmark_range, BookmarkRange::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct Bookmark {
    pub base: BookmarkRange,
    pub name: String,
}

impl Bookmark {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = BookmarkRange::from_xml_element(xml_node)?;
        let name = xml_node
            .attributes
            .get("name")
            .cloned()
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?;

        Ok(Self { base, name })
    }
}

#[cfg(test)]
impl Bookmark {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} id="0" displacedByCustomXml="next" colFirst="0" colLast="1" name="Some name">
        </{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            base: BookmarkRange::test_instance(),
            name: String::from("Some name"),
        }
    }
}

#[cfg(test)]
#[test]
fn test_bookmark_from_xml() {
    let xml = Bookmark::test_xml("bookmark");
    let bookmark = Bookmark::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(bookmark, Bookmark::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct MoveBookmark {
    pub base: Bookmark,
    pub author: String,
    pub date: DateTime,
}

impl MoveBookmark {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = Bookmark::from_xml_element(xml_node)?;
        let author = xml_node
            .attributes
            .get("author")
            .cloned()
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "author"))?;

        let date = xml_node
            .attributes
            .get("date")
            .cloned()
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "date"))?;

        Ok(Self { base, author, date })
    }
}

#[cfg(test)]
impl MoveBookmark {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} id="0" displacedByCustomXml="next" colFirst="0" colLast="1" name="Some name" author="John Smith" date="2001-10-26T21:32:52">
        </{node_name}>"#, node_name=node_name)
    }

    pub fn test_instance() -> Self {
        Self {
            base: Bookmark::test_instance(),
            author: String::from("John Smith"),
            date: DateTime::from("2001-10-26T21:32:52"),
        }
    }
}

#[cfg(test)]
#[test]
fn test_move_bookmark_from_xml() {
    let xml = MoveBookmark::test_xml("moveBookmark");
    let move_bookmark = MoveBookmark::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(move_bookmark, MoveBookmark::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct TrackChange {
    pub base: Markup,
    pub author: String,
    pub date: Option<DateTime>,
}

impl TrackChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = Markup::from_xml_element(xml_node)?;
        let author = xml_node
            .attributes
            .get("author")
            .cloned()
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "author"))?;

        let date = xml_node.attributes.get("date").cloned();

        Ok(Self { base, author, date })
    }
}

#[cfg(test)]
impl TrackChange {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} id="0" author="John Smith" date="2001-10-26T21:32:52"></{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            base: Markup::test_instance(),
            author: String::from("John Smith"),
            date: Some(DateTime::from("2001-10-26T21:32:52")),
        }
    }
}

#[cfg(test)]
#[test]
fn test_track_change_from_xml() {
    let xml = TrackChange::test_xml("trackChange");
    let track_change = TrackChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(track_change, TrackChange::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct Attr {
    pub uri: String,
    pub name: String,
    pub value: String,
}

impl Attr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut uri = None;
        let mut name = None;
        let mut value = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "uri" => uri = Some(attr_value.clone()),
                "name" => name = Some(attr_value.clone()),
                "val" => value = Some(attr_value.clone()),
                _ => (),
            }
        }

        Ok(Self {
            uri: uri.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "uri"))?,
            name: name.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?,
            value: value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?,
        })
    }
}

#[cfg(test)]
impl Attr {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} uri="http://some/uri" name="Some name" val="Some value"></{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            uri: String::from("http://some/uri"),
            name: String::from("Some name"),
            value: String::from("Some value"),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_attr_from_xml() {
    let xml = Attr::test_xml("attr");
    let attr = Attr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(attr, Attr::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomXmlPr {
    pub placeholder: Option<String>,
    pub attributes: Vec<Attr>,
}

impl CustomXmlPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut placeholder = None;
        let mut attributes = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "placeholder" => placeholder = child_node.attributes.get("val").cloned(),
                "attr" => attributes.push(Attr::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(Self {
            placeholder,
            attributes,
        })
    }
}

#[cfg(test)]
impl CustomXmlPr {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name}>
            <placeholder val="Placeholder" />
            {}
        </{node_name}>"#,
            Attr::test_xml("attr"),
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            placeholder: Some(String::from("Placeholder")),
            attributes: vec![Attr::test_instance()],
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_custom_xml_pr_from_xml() {
    let xml = CustomXmlPr::test_xml("customXmlPr");
    let custom_xml_pr = CustomXmlPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(custom_xml_pr, CustomXmlPr::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct SimpleField {
    pub paragraph_contents: Vec<PContent>,
    pub field_codes: String,
    pub field_lock: Option<OnOff>,
    pub dirty: Option<OnOff>,
}

impl SimpleField {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut field_codes = None;
        let mut field_lock = None;
        let mut dirty = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "instr" => field_codes = Some(value.clone()),
                "fldLock" => field_lock = Some(parse_xml_bool(value)?),
                "dirty" => dirty = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let mut paragraph_contents = Vec::new();
        for child_node in &xml_node.child_nodes {
            if PContent::is_choice_member(child_node.local_name()) {
                paragraph_contents.push(PContent::from_xml_element(child_node)?);
            }
        }

        let field_codes = field_codes.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "instr"))?;

        Ok(Self {
            field_codes,
            field_lock,
            dirty,
            paragraph_contents,
        })
    }
}

#[cfg(test)]
impl SimpleField {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} instr="AUTHOR" fldLock="false" dirty="false"></{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_xml_recursive(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} instr="AUTHOR" fldLock="false" dirty="false">
            {}
        </{node_name}>"#,
            Self::test_xml("fldSimple"),
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            paragraph_contents: Vec::new(),
            field_codes: String::from("AUTHOR"),
            field_lock: Some(false),
            dirty: Some(false),
        }
    }

    pub fn test_instance_recursive() -> Self {
        Self {
            paragraph_contents: vec![PContent::SimpleField(Self::test_instance())],
            ..Self::test_instance()
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_simple_field_from_xml() {
    let xml = SimpleField::test_xml_recursive("simpleField");
    let simple_field = SimpleField::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(simple_field, SimpleField::test_instance_recursive());
}

#[derive(Debug, Clone, PartialEq)]
pub struct Hyperlink {
    pub paragraph_contents: Vec<PContent>,
    pub target_frame: Option<String>,
    pub tooltip: Option<String>,
    pub document_location: Option<String>,
    pub history: Option<OnOff>,
    pub anchor: Option<String>,
    pub rel_id: RelationshipId,
}

impl Hyperlink {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut target_frame = None;
        let mut tooltip = None;
        let mut document_location = None;
        let mut history = None;
        let mut anchor = None;
        let mut rel_id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "tgtFrame" => target_frame = Some(value.clone()),
                "tooltip" => tooltip = Some(value.clone()),
                "docLocation" => document_location = Some(value.clone()),
                "history" => history = Some(parse_xml_bool(value)?),
                "anchor" => anchor = Some(value.clone()),
                "r:id" => rel_id = Some(value.clone()),
                _ => (),
            }
        }

        let mut paragraph_contents = Vec::new();
        for child_node in &xml_node.child_nodes {
            if PContent::is_choice_member(child_node.local_name()) {
                paragraph_contents.push(PContent::from_xml_element(child_node)?);
            }
        }

        let rel_id = rel_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;
        Ok(Self {
            paragraph_contents,
            target_frame,
            tooltip,
            document_location,
            history,
            anchor,
            rel_id,
        })
    }
}

#[cfg(test)]
impl Hyperlink {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} tgtFrame="_blank" tooltip="Some tooltip" docLocation="table" history="true" anchor="chapter1" r:id="rId1"></{node_name}>"#, node_name=node_name)
    }

    pub fn test_xml_recursive(node_name: &'static str) -> String {
        format!(r#"<{node_name} tgtFrame="_blank" tooltip="Some tooltip" docLocation="table" history="true" anchor="chapter1" r:id="rId1">
            {}
        </{node_name}>"#, SimpleField::test_xml("fldSimple"), node_name=node_name)
    }

    pub fn test_instance() -> Self {
        Self {
            paragraph_contents: Vec::new(),
            target_frame: Some(String::from("_blank")),
            tooltip: Some(String::from("Some tooltip")),
            document_location: Some(String::from("table")),
            history: Some(true),
            anchor: Some(String::from("chapter1")),
            rel_id: RelationshipId::from("rId1"),
        }
    }

    pub fn test_instance_recursive() -> Self {
        Self {
            paragraph_contents: vec![PContent::test_simple_field_instance()],
            ..Self::test_instance()
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_hyperlink_from_xml() {
    let xml = Hyperlink::test_xml_recursive("hyperlink");
    let hyperlink = Hyperlink::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(hyperlink, Hyperlink::test_instance_recursive());
}

#[derive(Debug, Clone, PartialEq)]
pub struct Rel {
    pub rel_id: RelationshipId,
}

impl Rel {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let rel_id = xml_node
            .attributes
            .get("r:id")
            .cloned()
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self { rel_id })
    }
}

#[cfg(test)]
impl Rel {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} r:id="rId1"></{node_name}>"#, node_name = node_name)
    }

    pub fn test_instance() -> Self {
        Self {
            rel_id: RelationshipId::from("rId1"),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_rel_from_xml() {
    let xml = Rel::test_xml("rel");
    let rel = Rel::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(rel, Rel::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub enum PContent {
    ContentRunContent(ContentRunContent),
    SimpleField(SimpleField),
    Hyperlink(Hyperlink),
    SubDocument(Rel),
}

impl PContent {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "fldSimple" | "hyperlink" | "subDoc" => true,
            _ => ContentRunContent::is_choice_member(&node_name),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            node_name @ _ if ContentRunContent::is_choice_member(node_name) => Ok(PContent::ContentRunContent(
                ContentRunContent::from_xml_element(xml_node)?,
            )),
            "fldSimple" => Ok(PContent::SimpleField(SimpleField::from_xml_element(xml_node)?)),
            "hyperlink" => Ok(PContent::Hyperlink(Hyperlink::from_xml_element(xml_node)?)),
            "subDoc" => Ok(PContent::SubDocument(Rel::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "PContent"))),
        }
    }
}

#[cfg(test)]
impl PContent {
    pub fn test_simple_field_xml() -> String {
        SimpleField::test_xml("fldSimple")
    }

    pub fn test_simple_field_instance() -> Self {
        PContent::SimpleField(SimpleField::test_instance())
    }
}

#[cfg(test)]
#[test]
pub fn test_pcontent_content_run_content_from_xml() {
    // TODO
}

#[cfg(test)]
#[test]
pub fn test_pcontent_simple_field_from_xml() {
    let xml = SimpleField::test_xml("fldSimple");
    let pcontent = PContent::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(pcontent, PContent::SimpleField(SimpleField::test_instance()));
}

#[cfg(test)]
#[test]
pub fn test_pcontent_hyperlink_from_xml() {
    let xml = Hyperlink::test_xml("hyperlink");
    let pcontent = PContent::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(pcontent, PContent::Hyperlink(Hyperlink::test_instance()));
}

#[cfg(test)]
#[test]
pub fn test_pcontent_subdocument_from_xml() {
    let xml = Rel::test_xml("subDoc");
    let pcontent = PContent::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(pcontent, PContent::SubDocument(Rel::test_instance()));
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomXmlRun {
    pub custom_xml_properties: Option<CustomXmlPr>,
    pub paragraph_contents: Vec<PContent>,

    pub uri: String,
    pub element: String,
}

impl CustomXmlRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut uri = None;
        let mut element = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "uri" => uri = Some(value.clone()),
                "element" => element = Some(value.clone()),
                _ => (),
            }
        }

        let mut custom_xml_properties = None;
        let mut paragraph_contents = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "customXmlPr" => custom_xml_properties = Some(CustomXmlPr::from_xml_element(child_node)?),
                node_name @ _ if PContent::is_choice_member(node_name) => {
                    paragraph_contents.push(PContent::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        let uri = uri.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "uri"))?;
        let element = element.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "element"))?;
        Ok(Self {
            custom_xml_properties,
            paragraph_contents,
            uri,
            element,
        })
    }
}

#[cfg(test)]
impl CustomXmlRun {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} uri="http://some/uri" element="Some element">
            {}
            {}
        </{node_name}>"#,
            CustomXmlPr::test_xml("customXmlPr"),
            PContent::test_simple_field_xml(),
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            custom_xml_properties: Some(CustomXmlPr::test_instance()),
            paragraph_contents: vec![PContent::test_simple_field_instance()],
            uri: String::from("http://some/uri"),
            element: String::from("Some element"),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_custom_xml_run_from_xml() {
    let xml = CustomXmlRun::test_xml("customXmlRun");
    let custom_xml_run = CustomXmlRun::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(custom_xml_run, CustomXmlRun::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct SmartTagPr {
    pub attributes: Vec<Attr>,
}

impl SmartTagPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut attributes = Vec::new();
        for child_node in &xml_node.child_nodes {
            if child_node.local_name() == "attr" {
                attributes.push(Attr::from_xml_element(child_node)?);
            }
        }

        Ok(Self { attributes })
    }
}

#[cfg(test)]
impl SmartTagPr {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name}>
            {}
        </{node_name}>"#,
            Attr::test_xml("attr"),
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            attributes: vec![Attr::test_instance()],
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_smart_tag_pr_from_xml() {
    let xml = SmartTagPr::test_xml("smartTagPr");
    let smart_tag_pr = SmartTagPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(smart_tag_pr, SmartTagPr::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct SmartTagRun {
    pub smart_tag_properties: Option<SmartTagPr>,
    pub paragraph_contents: Vec<PContent>,
    pub uri: String,
    pub element: String,
}

impl SmartTagRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut uri = None;
        let mut element = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "uri" => uri = Some(value.clone()),
                "element" => element = Some(value.clone()),
                _ => (),
            }
        }

        let mut smart_tag_properties = None;
        let mut paragraph_contents = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "smartTagPr" => smart_tag_properties = Some(SmartTagPr::from_xml_element(child_node)?),
                node_name @ _ if PContent::is_choice_member(node_name) => {
                    paragraph_contents.push(PContent::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        let uri = uri.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "uri"))?;
        let element = element.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "element"))?;

        Ok(Self {
            uri,
            element,
            smart_tag_properties,
            paragraph_contents,
        })
    }
}

#[cfg(test)]
impl SmartTagRun {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} uri="http://some/uri" element="Some element">
            {}
            {}
        </{node_name}>"#,
            SmartTagPr::test_xml("smartTagPr"),
            PContent::test_simple_field_xml(),
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            smart_tag_properties: Some(SmartTagPr::test_instance()),
            paragraph_contents: vec![PContent::test_simple_field_instance()],
            uri: String::from("http://some/uri"),
            element: String::from("Some element"),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_smart_tag_run_from_xml() {
    let xml = SmartTagRun::test_xml("smartTagRun");
    let smart_tag_run = SmartTagRun::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(smart_tag_run, SmartTagRun::test_instance());
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum Hint {
    #[strum(serialize = "default")]
    Default,
    #[strum(serialize = "eastAsia")]
    EastAsia,
    #[strum(serialize = "cs")]
    ComplexScript,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum Theme {
    #[strum(serialize = "majorEastAsia")]
    MajorEastAsia,
    #[strum(serialize = "majorBidi")]
    MajorBidirectional,
    #[strum(serialize = "majorAscii")]
    MajorAscii,
    #[strum(serialize = "majorHAnsi")]
    MajorHighAnsi,
    #[strum(serialize = "minorEastAsia")]
    MinorEastAsia,
    #[strum(serialize = "minorBidi")]
    MinorBidirectional,
    #[strum(serialize = "minorAscii")]
    MinorAscii,
    #[strum(serialize = "minorHAnsi")]
    MinorHighAnsi,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Fonts {
    pub hint: Option<Hint>,
    pub ascii: Option<String>,
    pub high_ansi: Option<String>,
    pub east_asia: Option<String>,
    pub complex_script: Option<String>,
    pub ascii_theme: Option<Theme>,
    pub high_ansi_theme: Option<Theme>,
    pub east_asia_theme: Option<Theme>,
    pub complex_script_theme: Option<Theme>,
}

impl Fonts {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Fonts = Default::default();
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "hint" => instance.hint = Some(value.parse()?),
                "ascii" => instance.ascii = Some(value.clone()),
                "hAnsi" => instance.high_ansi = Some(value.clone()),
                "eastAsia" => instance.east_asia = Some(value.clone()),
                "cs" => instance.complex_script = Some(value.clone()),
                "asciiTheme" => instance.ascii_theme = Some(value.parse()?),
                "hAnsiTheme" => instance.high_ansi_theme = Some(value.parse()?),
                "eastAsiaTheme" => instance.east_asia_theme = Some(value.parse()?),
                "cstheme" => instance.complex_script_theme = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[cfg(test)]
impl Fonts {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} hint="default" ascii="Arial" hAnsi="Arial" eastAsia="Arial" cs="Arial"
            asciiTheme="majorAscii" hAnsiTheme="majorHAnsi" eastAsiaTheme="majorEastAsia" cstheme="majorBidi">
        </{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            hint: Some(Hint::Default),
            ascii: Some(String::from("Arial")),
            high_ansi: Some(String::from("Arial")),
            east_asia: Some(String::from("Arial")),
            complex_script: Some(String::from("Arial")),
            ascii_theme: Some(Theme::MajorAscii),
            high_ansi_theme: Some(Theme::MajorHighAnsi),
            east_asia_theme: Some(Theme::MajorEastAsia),
            complex_script_theme: Some(Theme::MajorBidirectional),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_fonts_from_xml() {
    let xml = Fonts::test_xml("fonts");
    let fonts = Fonts::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(fonts, Fonts::test_instance());
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum UnderlineType {
    #[strum(serialize = "single")]
    Single,
    #[strum(serialize = "words")]
    Words,
    #[strum(serialize = "double")]
    Double,
    #[strum(serialize = "thick")]
    Thick,
    #[strum(serialize = "dotted")]
    Dotted,
    #[strum(serialize = "dottedHeavy")]
    DottedHeavy,
    #[strum(serialize = "dash")]
    Dash,
    #[strum(serialize = "dashedHeavy")]
    DashedHeavy,
    #[strum(serialize = "dashLong")]
    DashLong,
    #[strum(serialize = "dashLongHeavy")]
    DashLongHeavy,
    #[strum(serialize = "dotDash")]
    DotDash,
    #[strum(serialize = "dashDotHeavy")]
    DashDotHeavy,
    #[strum(serialize = "dotDotDash")]
    DotDotDash,
    #[strum(serialize = "dashDotDotHeavy")]
    DashDotDotHeavy,
    #[strum(serialize = "wave")]
    Wave,
    #[strum(serialize = "wavyHeavy")]
    WavyHeavy,
    #[strum(serialize = "wavyDouble")]
    WavyDouble,
    #[strum(serialize = "none")]
    None,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Underline {
    pub value: Option<UnderlineType>,
    pub color: Option<HexColor>,
    pub theme_color: Option<ThemeColor>,
    pub theme_tint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,
}

impl Underline {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Underline = Default::default();
        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => instance.value = Some(attr_value.parse()?),
                "color" => instance.color = Some(attr_value.parse()?),
                "themeColor" => instance.theme_color = Some(attr_value.parse()?),
                "themeTint" => instance.theme_tint = Some(u8::from_str_radix(attr_value, 16)?),
                "themeShade" => instance.theme_shade = Some(u8::from_str_radix(attr_value, 16)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[cfg(test)]
impl Underline {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} val="single" color="ffffff" themeColor="accent1" themeTint="ff" themeShade="ff">
        </{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            value: Some(UnderlineType::Single),
            color: Some(HexColor::RGB([0xff, 0xff, 0xff])),
            theme_color: Some(ThemeColor::Accent1),
            theme_tint: Some(0xff),
            theme_shade: Some(0xff),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_underline_from_xml() {
    let xml = Underline::test_xml("underline");
    let underline = Underline::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(underline, Underline::test_instance());
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TextEffect {
    #[strum(serialize = "blinkBackground")]
    BlinkBackground,
    #[strum(serialize = "lights")]
    Lights,
    #[strum(serialize = "antsBlack")]
    AntsBlack,
    #[strum(serialize = "antsRed")]
    AntsRed,
    #[strum(serialize = "shimmer")]
    Shimmer,
    #[strum(serialize = "sparkle")]
    Sparkle,
    #[strum(serialize = "none")]
    None,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum BorderType {
    #[strum(serialize = "nil")]
    Nil,
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "single")]
    Single,
    #[strum(serialize = "thick")]
    Thick,
    #[strum(serialize = "double")]
    Double,
    #[strum(serialize = "dotted")]
    Dotted,
    #[strum(serialize = "dashed")]
    Dashed,
    #[strum(serialize = "dotDash")]
    DotDash,
    #[strum(serialize = "dotDotDash")]
    DotDotDash,
    #[strum(serialize = "triple")]
    Triple,
    #[strum(serialize = "thinThickSmallGap")]
    ThinThickSmallGap,
    #[strum(serialize = "thickThinSmallGap")]
    ThickThinSmallGap,
    #[strum(serialize = "thinThickThinSmallGap")]
    ThinThickThinSmallGap,
    #[strum(serialize = "thinThickMediumGap")]
    ThinThickMediumGap,
    #[strum(serialize = "thickThinMediumGap")]
    ThickThinMediumGap,
    #[strum(serialize = "thinThickThinMediumGap")]
    ThinThickThinMediumGap,
    #[strum(serialize = "thinThickLargeGap")]
    ThinThickLargeGap,
    #[strum(serialize = "thickThinLargeGap")]
    ThickThinLargeGap,
    #[strum(serialize = "thinThickThinLargeGap")]
    ThinThickThinLargeGap,
    #[strum(serialize = "wave")]
    Wave,
    #[strum(serialize = "doubleWave")]
    DoubleWave,
    #[strum(serialize = "dashSmallGap")]
    DashSmallGap,
    #[strum(serialize = "dashDotStroked")]
    DashDotStroked,
    #[strum(serialize = "threeDEmboss")]
    ThreeDEmboss,
    #[strum(serialize = "threeDEngrave")]
    ThreeDEngrave,
    #[strum(serialize = "outset")]
    Outset,
    #[strum(serialize = "inset")]
    Inset,
    #[strum(serialize = "apples")]
    Apples,
    #[strum(serialize = "archedScallops")]
    ArchedScallops,
    #[strum(serialize = "babyPacifier")]
    BabyPacifier,
    #[strum(serialize = "babyRattle")]
    BabyRattle,
    #[strum(serialize = "balloons3Colors")]
    Balloons3Colors,
    #[strum(serialize = "balloonsHotAir")]
    BalloonsHotAir,
    #[strum(serialize = "basicBlackDashes")]
    BasicBlackDashes,
    #[strum(serialize = "basicBlackDots")]
    BasicBlackDots,
    #[strum(serialize = "basicBlackSquares")]
    BasicBlackSquares,
    #[strum(serialize = "basicThinLines")]
    BasicThinLines,
    #[strum(serialize = "basicWhiteDashes")]
    BasicWhiteDashes,
    #[strum(serialize = "basicWhiteDots")]
    BasicWhiteDots,
    #[strum(serialize = "basicWhiteSquares")]
    BasicWhiteSquares,
    #[strum(serialize = "basicWideInline")]
    BasicWideInline,
    #[strum(serialize = "basicWideMidline")]
    BasicWideMidline,
    #[strum(serialize = "basicWideOutline")]
    BasicWideOutline,
    #[strum(serialize = "bats")]
    Bats,
    #[strum(serialize = "birds")]
    Birds,
    #[strum(serialize = "birdsFlight")]
    BirdsFlight,
    #[strum(serialize = "cabins")]
    Cabins,
    #[strum(serialize = "cakeSlice")]
    CakeSlice,
    #[strum(serialize = "candyCorn")]
    CandyCorn,
    #[strum(serialize = "celticKnotwork")]
    CelticKnotwork,
    #[strum(serialize = "certificateBanner")]
    CertificateBanner,
    #[strum(serialize = "chainLink")]
    ChainLink,
    #[strum(serialize = "champagneBottle")]
    ChampagneBottle,
    #[strum(serialize = "checkedBarBlack")]
    CheckedBarBlack,
    #[strum(serialize = "checkedBarColor")]
    CheckedBarColor,
    #[strum(serialize = "checkered")]
    Checkered,
    #[strum(serialize = "christmasTree")]
    ChristmasTree,
    #[strum(serialize = "circlesLines")]
    CirclesLines,
    #[strum(serialize = "circlesRectangles")]
    CirclesRectangles,
    #[strum(serialize = "classicalWave")]
    ClassicalWave,
    #[strum(serialize = "clocks")]
    Clocks,
    #[strum(serialize = "compass")]
    Compass,
    #[strum(serialize = "confetti")]
    Confetti,
    #[strum(serialize = "confettiGrays")]
    ConfettiGrays,
    #[strum(serialize = "confettiOutline")]
    ConfettiOutline,
    #[strum(serialize = "confettiStreamers")]
    ConfettiStreamers,
    #[strum(serialize = "confettiWhite")]
    ConfettiWhite,
    #[strum(serialize = "cornerTriangles")]
    CornerTriangles,
    #[strum(serialize = "couponCutoutDashes")]
    CouponCutoutDashes,
    #[strum(serialize = "couponCutoutDots")]
    CouponCutoutDots,
    #[strum(serialize = "crazyMaze")]
    CrazyMaze,
    #[strum(serialize = "creaturesButterfly")]
    CreaturesButterfly,
    #[strum(serialize = "creaturesFish")]
    CreaturesFish,
    #[strum(serialize = "creaturesInsects")]
    CreaturesInsects,
    #[strum(serialize = "creaturesLadyBug")]
    CreaturesLadyBug,
    #[strum(serialize = "crossStitch")]
    CrossStitch,
    #[strum(serialize = "cup")]
    Cup,
    #[strum(serialize = "decoArch")]
    DecoArch,
    #[strum(serialize = "decoArchColor")]
    DecoArchColor,
    #[strum(serialize = "decoBlocks")]
    DecoBlocks,
    #[strum(serialize = "diamondsGray")]
    DiamondsGray,
    #[strum(serialize = "doubleD")]
    DoubleD,
    #[strum(serialize = "doubleDiamonds")]
    DoubleDiamonds,
    #[strum(serialize = "earth1")]
    Earth1,
    #[strum(serialize = "earth2")]
    Earth2,
    #[strum(serialize = "earth3")]
    Earth3,
    #[strum(serialize = "eclipsingSquares1")]
    EclipsingSquares1,
    #[strum(serialize = "eclipsingSquares2")]
    EclipsingSquares2,
    #[strum(serialize = "eggsBlack")]
    EggsBlack,
    #[strum(serialize = "fans")]
    Fans,
    #[strum(serialize = "film")]
    Film,
    #[strum(serialize = "firecrackers")]
    Firecrackers,
    #[strum(serialize = "flowersBlockPrint")]
    FlowersBlockPrint,
    #[strum(serialize = "flowersDaisies")]
    FlowersDaisies,
    #[strum(serialize = "flowersModern1")]
    FlowersModern1,
    #[strum(serialize = "flowersModern2")]
    FlowersModern2,
    #[strum(serialize = "flowersPansy")]
    FlowersPansy,
    #[strum(serialize = "flowersRedRose")]
    FlowersRedRose,
    #[strum(serialize = "flowersRoses")]
    FlowersRoses,
    #[strum(serialize = "flowersTeacup")]
    FlowersTeacup,
    #[strum(serialize = "flowersTiny")]
    FlowersTiny,
    #[strum(serialize = "gems")]
    Gems,
    #[strum(serialize = "gingerbreadMan")]
    GingerbreadMan,
    #[strum(serialize = "gradient")]
    Gradient,
    #[strum(serialize = "handmade1")]
    Handmade1,
    #[strum(serialize = "handmade2")]
    Handmade2,
    #[strum(serialize = "heartBalloon")]
    HeartBalloon,
    #[strum(serialize = "heartGray")]
    HeartGray,
    #[strum(serialize = "hearts")]
    Hearts,
    #[strum(serialize = "heebieJeebies")]
    HeebieJeebies,
    #[strum(serialize = "holly")]
    Holly,
    #[strum(serialize = "houseFunky")]
    HouseFunky,
    #[strum(serialize = "hypnotic")]
    Hypnotic,
    #[strum(serialize = "iceCreamCones")]
    IceCreamCones,
    #[strum(serialize = "lightBulb")]
    LightBulb,
    #[strum(serialize = "lightning1")]
    Lightning1,
    #[strum(serialize = "lightning2")]
    Lightning2,
    #[strum(serialize = "mapPins")]
    MapPins,
    #[strum(serialize = "mapleLeaf")]
    MapleLeaf,
    #[strum(serialize = "mapleMuffins")]
    MapleMuffins,
    #[strum(serialize = "marquee")]
    Marquee,
    #[strum(serialize = "marqueeToothed")]
    MarqueeToothed,
    #[strum(serialize = "moons")]
    Moons,
    #[strum(serialize = "mosaic")]
    Mosaic,
    #[strum(serialize = "musicNotes")]
    MusicNotes,
    #[strum(serialize = "northwest")]
    Northwest,
    #[strum(serialize = "ovals")]
    Ovals,
    #[strum(serialize = "packages")]
    Packages,
    #[strum(serialize = "palmsBlack")]
    PalmsBlack,
    #[strum(serialize = "palmsColor")]
    PalmsColor,
    #[strum(serialize = "paperClips")]
    PaperClips,
    #[strum(serialize = "papyrus")]
    Papyrus,
    #[strum(serialize = "partyFavor")]
    PartyFavor,
    #[strum(serialize = "partyGlass")]
    PartyGlass,
    #[strum(serialize = "pencils")]
    Pencils,
    #[strum(serialize = "people")]
    People,
    #[strum(serialize = "peopleWaving")]
    PeopleWaving,
    #[strum(serialize = "peopleHats")]
    PeopleHats,
    #[strum(serialize = "poinsettias")]
    Poinsettias,
    #[strum(serialize = "postageStamp")]
    PostageStamp,
    #[strum(serialize = "pumpkin1")]
    Pumpkin1,
    #[strum(serialize = "pushPinNote2")]
    PushPinNote2,
    #[strum(serialize = "pushPinNote1")]
    PushPinNote1,
    #[strum(serialize = "pyramids")]
    Pyramids,
    #[strum(serialize = "pyramidsAbove")]
    PyramidsAbove,
    #[strum(serialize = "quadrants")]
    Quadrants,
    #[strum(serialize = "rings")]
    Rings,
    #[strum(serialize = "safari")]
    Safari,
    #[strum(serialize = "sawtooth")]
    Sawtooth,
    #[strum(serialize = "sawtoothGray")]
    SawtoothGray,
    #[strum(serialize = "scaredCat")]
    ScaredCat,
    #[strum(serialize = "seattle")]
    Seattle,
    #[strum(serialize = "shadowedSquares")]
    ShadowedSquares,
    #[strum(serialize = "sharksTeeth")]
    SharksTeeth,
    #[strum(serialize = "shorebirdTracks")]
    ShorebirdTracks,
    #[strum(serialize = "skyrocket")]
    Skyrocket,
    #[strum(serialize = "snowflakeFancy")]
    SnowflakeFancy,
    #[strum(serialize = "snowflakes")]
    Snowflakes,
    #[strum(serialize = "sombrero")]
    Sombrero,
    #[strum(serialize = "southwest")]
    Southwest,
    #[strum(serialize = "stars")]
    Stars,
    #[strum(serialize = "starsTop")]
    StarsTop,
    #[strum(serialize = "stars3d")]
    Stars3d,
    #[strum(serialize = "starsBlack")]
    StarsBlack,
    #[strum(serialize = "starsShadowed")]
    StarsShadowed,
    #[strum(serialize = "sun")]
    Sun,
    #[strum(serialize = "swirligig")]
    Swirligig,
    #[strum(serialize = "tornPaper")]
    TornPaper,
    #[strum(serialize = "tornPaperBlack")]
    TornPaperBlack,
    #[strum(serialize = "trees")]
    Trees,
    #[strum(serialize = "triangleParty")]
    TriangleParty,
    #[strum(serialize = "triangles")]
    Triangles,
    #[strum(serialize = "triangle1")]
    Triangle1,
    #[strum(serialize = "triangle2")]
    Triangle2,
    #[strum(serialize = "triangleCircle1")]
    TriangleCircle1,
    #[strum(serialize = "triangleCircle2")]
    TriangleCircle2,
    #[strum(serialize = "shapes1")]
    Shapes1,
    #[strum(serialize = "shapes2")]
    Shapes2,
    #[strum(serialize = "twistedLines1")]
    TwistedLines1,
    #[strum(serialize = "twistedLines2")]
    TwistedLines2,
    #[strum(serialize = "vine")]
    Vine,
    #[strum(serialize = "waveline")]
    Waveline,
    #[strum(serialize = "weavingAngles")]
    WeavingAngles,
    #[strum(serialize = "weavingBraid")]
    WeavingBraid,
    #[strum(serialize = "weavingRibbon")]
    WeavingRibbon,
    #[strum(serialize = "weavingStrips")]
    WeavingStrips,
    #[strum(serialize = "whiteFlowers")]
    WhiteFlowers,
    #[strum(serialize = "woodwork")]
    Woodwork,
    #[strum(serialize = "xIllusions")]
    XIllusions,
    #[strum(serialize = "zanyTriangles")]
    ZanyTriangles,
    #[strum(serialize = "zigZag")]
    ZigZag,
    #[strum(serialize = "zigZagStitch")]
    ZigZagStitch,
    #[strum(serialize = "custom")]
    Custom,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Border {
    pub value: BorderType,
    pub color: Option<HexColor>,
    pub theme_color: Option<ThemeColor>,
    pub theme_tint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,
    pub size: Option<EightPointMeasure>,
    pub spacing: Option<PointMeasure>,
    pub shadow: Option<OnOff>,
    pub frame: Option<OnOff>,
}

impl Border {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut value = None;
        let mut color = None;
        let mut theme_color = None;
        let mut theme_tint = None;
        let mut theme_shade = None;
        let mut size = None;
        let mut spacing = None;
        let mut shadow = None;
        let mut frame = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => value = Some(attr_value.parse()?),
                "color" => color = Some(attr_value.parse()?),
                "themeColor" => theme_color = Some(attr_value.parse()?),
                "themeTint" => theme_tint = Some(u8::from_str_radix(attr_value, 16)?),
                "themeShade" => theme_shade = Some(u8::from_str_radix(attr_value, 16)?),
                "sz" => size = Some(attr_value.parse()?),
                "space" => spacing = Some(attr_value.parse()?),
                "shadow" => shadow = Some(parse_xml_bool(attr_value)?),
                "frame" => frame = Some(parse_xml_bool(attr_value)?),
                _ => (),
            }
        }

        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;

        Ok(Self {
            value,
            color,
            theme_color,
            theme_tint,
            theme_shade,
            size,
            spacing,
            shadow,
            frame,
        })
    }
}

#[cfg(test)]
impl Border {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} val="single" color="ffffff" themeColor="accent1" themeTint="ff" themeShade="ff" sz="100" space="100" shadow="true" frame="true">
        </{node_name}>"#,
            node_name=node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            value: BorderType::Single,
            color: Some(HexColor::RGB([0xff, 0xff, 0xff])),
            theme_color: Some(ThemeColor::Accent1),
            theme_tint: Some(0xff),
            theme_shade: Some(0xff),
            size: Some(100),
            spacing: Some(100),
            shadow: Some(true),
            frame: Some(true),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_border_from_xml() {
    let xml = Border::test_xml("border");
    let border = Border::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(border, Border::test_instance());
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum ShdType {
    #[strum(serialize = "nil")]
    Nil,
    #[strum(serialize = "clear")]
    Clear,
    #[strum(serialize = "solid")]
    Solid,
    #[strum(serialize = "horzStripe")]
    HorizontalStripe,
    #[strum(serialize = "vertStripe")]
    VerticalStripe,
    #[strum(serialize = "reverseDiagStripe")]
    ReverseDiagonalStripe,
    #[strum(serialize = "diagStripe")]
    DiagonalStripe,
    #[strum(serialize = "horzCross")]
    HorizontalCross,
    #[strum(serialize = "diagCross")]
    DiagonalCross,
    #[strum(serialize = "thinHorzStripe")]
    ThinHorizontalStripe,
    #[strum(serialize = "thinVertStripe")]
    ThinVerticalStripe,
    #[strum(serialize = "thinReverseDiagStripe")]
    ThinReverseDiagonalStripe,
    #[strum(serialize = "thinDiagStripe")]
    ThinDiagonalStripe,
    #[strum(serialize = "thinHorzCross")]
    ThinHorizontalCross,
    #[strum(serialize = "thinDiagCross")]
    ThinDiagonalCross,
    #[strum(serialize = "pct5")]
    Percent5,
    #[strum(serialize = "pct10")]
    Percent10,
    #[strum(serialize = "pct12")]
    Percent12,
    #[strum(serialize = "pct15")]
    Percent15,
    #[strum(serialize = "pct20")]
    Percent20,
    #[strum(serialize = "pct25")]
    Percent25,
    #[strum(serialize = "pct30")]
    Percent30,
    #[strum(serialize = "pct35")]
    Percent35,
    #[strum(serialize = "pct37")]
    Percent37,
    #[strum(serialize = "pct40")]
    Percent40,
    #[strum(serialize = "pct45")]
    Percent45,
    #[strum(serialize = "pct50")]
    Percent50,
    #[strum(serialize = "pct55")]
    Percent55,
    #[strum(serialize = "pct60")]
    Percent60,
    #[strum(serialize = "pct62")]
    Percent62,
    #[strum(serialize = "pct65")]
    Percent65,
    #[strum(serialize = "pct70")]
    Percent70,
    #[strum(serialize = "pct75")]
    Percent75,
    #[strum(serialize = "pct80")]
    Percent80,
    #[strum(serialize = "pct85")]
    Percent85,
    #[strum(serialize = "pct87")]
    Percent87,
    #[strum(serialize = "pct90")]
    Percent90,
    #[strum(serialize = "pct95")]
    Percent95,
}

#[derive(Debug, Clone, PartialEq)]
pub struct Shd {
    pub value: ShdType,
    pub color: Option<HexColor>,
    pub theme_color: Option<ThemeColor>,
    pub theme_tint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,
    pub fill: Option<HexColor>,
    pub theme_fill: Option<ThemeColor>,
    pub theme_fill_tint: Option<UcharHexNumber>,
    pub theme_fill_shade: Option<UcharHexNumber>,
}

impl Shd {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut value = None;
        let mut color = None;
        let mut theme_color = None;
        let mut theme_tint = None;
        let mut theme_shade = None;
        let mut fill = None;
        let mut theme_fill = None;
        let mut theme_fill_tint = None;
        let mut theme_fill_shade = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => value = Some(attr_value.parse()?),
                "color" => color = Some(attr_value.parse()?),
                "themeColor" => theme_color = Some(attr_value.parse()?),
                "themeTint" => theme_tint = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                "themeShade" => theme_shade = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                "fill" => fill = Some(attr_value.parse()?),
                "themeFill" => theme_fill = Some(attr_value.parse()?),
                "themeFillTint" => theme_fill_tint = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                "themeFillShade" => theme_fill_shade = Some(UcharHexNumber::from_str_radix(attr_value, 16)?),
                _ => (),
            }
        }

        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "value"))?;
        Ok(Self {
            value,
            color,
            theme_color,
            theme_tint,
            theme_shade,
            fill,
            theme_fill,
            theme_fill_tint,
            theme_fill_shade,
        })
    }
}

#[cfg(test)]
impl Shd {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} val="solid" color="ffffff" themeColor="accent1" themeTint="ff" themeShade="ff" fill="ffffff" themeFill="accent1" themeFillTint="ff" themeFillShade="ff">
        </{node_name}>"#,
            node_name=node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            value: ShdType::Solid,
            color: Some(HexColor::RGB([0xff, 0xff, 0xff])),
            theme_color: Some(ThemeColor::Accent1),
            theme_tint: Some(0xff),
            theme_shade: Some(0xff),
            fill: Some(HexColor::RGB([0xff, 0xff, 0xff])),
            theme_fill: Some(ThemeColor::Accent1),
            theme_fill_tint: Some(0xff),
            theme_fill_shade: Some(0xff),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_shd_from_xml() {
    let xml = Shd::test_xml("shd");
    let shd = Shd::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(shd, Shd::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct FitText {
    pub value: TwipsMeasure,
    pub id: Option<DecimalNumber>,
}

impl FitText {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut value = None;
        let mut id = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => value = Some(attr_value.parse()?),
                "id" => id = Some(attr_value.parse()?),
                _ => (),
            }
        }

        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;

        Ok(Self { value, id })
    }
}

#[cfg(test)]
impl FitText {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} val="123.456mm" id="1"></{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            value: TwipsMeasure::UniversalMeasure(PositiveUniversalMeasure::new(
                123.456,
                UniversalMeasureUnit::Millimeter,
            )),
            id: Some(1),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_fit_text_from_xml() {
    let xml = FitText::test_xml("fitText");
    let fit_text = FitText::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(fit_text, FitText::test_instance());
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum Em {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "dot")]
    Dot,
    #[strum(serialize = "comma")]
    Comma,
    #[strum(serialize = "circle")]
    Circle,
    #[strum(serialize = "underDot")]
    UnderDot,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Language {
    pub value: Option<Lang>,
    pub east_asia: Option<Lang>,
    pub bidirectional: Option<Lang>,
}

impl Language {
    pub fn from_xml_element(xml_node: &XmlNode) -> Self {
        let mut instance: Self = Default::default();
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => instance.value = Some(value.clone()),
                "eastAsia" => instance.east_asia = Some(value.clone()),
                "bidi" => instance.bidirectional = Some(value.clone()),
                _ => (),
            }
        }

        instance
    }
}

#[cfg(test)]
impl Language {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} val="en" eastAsia="jp" bidi="fa"></{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            value: Some(Lang::from("en")),
            east_asia: Some(Lang::from("jp")),
            bidirectional: Some(Lang::from("fa")),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_language_from_xml() {
    let xml = Language::test_xml("language");
    let language = Language::from_xml_element(&XmlNode::from_str(xml).unwrap());
    assert_eq!(language, Language::test_instance());
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum CombineBrackets {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "round")]
    Round,
    #[strum(serialize = "square")]
    Square,
    #[strum(serialize = "angle")]
    Angle,
    #[strum(serialize = "curly")]
    Curly,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct EastAsianLayout {
    pub id: Option<DecimalNumber>,
    pub combine: Option<OnOff>,
    pub combine_brackets: Option<CombineBrackets>,
    pub vertical: Option<OnOff>,
    pub vertical_compress: Option<OnOff>,
}

impl EastAsianLayout {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "id" => instance.id = Some(value.parse()?),
                "combine" => instance.combine = Some(parse_xml_bool(value)?),
                "combineBrackets" => instance.combine_brackets = Some(value.parse()?),
                "vert" => instance.vertical = Some(parse_xml_bool(value)?),
                "vertCompress" => instance.vertical_compress = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[cfg(test)]
impl EastAsianLayout {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} id="1" combine="true" combineBrackets="square" vert="true" vertCompress="true">
        </{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            id: Some(1),
            combine: Some(true),
            combine_brackets: Some(CombineBrackets::Square),
            vertical: Some(true),
            vertical_compress: Some(true),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_east_asian_layout_from_xml() {
    let xml = EastAsianLayout::test_xml("eastAsianLayout");
    let east_asian_layout = EastAsianLayout::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(east_asian_layout, EastAsianLayout::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub enum RPrBase {
    RunStyle(String),
    RunFonts(Fonts),
    Bold(Option<OnOff>),
    ComplexScriptBold(Option<OnOff>),
    Italic(Option<OnOff>),
    ComplexScriptItalic(Option<OnOff>),
    Capitals(Option<OnOff>),
    SmallCapitals(Option<OnOff>),
    Strikethrough(Option<OnOff>),
    DoubleStrikethrough(Option<OnOff>),
    Outline(Option<OnOff>),
    Shadow(Option<OnOff>),
    Emboss(Option<OnOff>),
    Imprint(Option<OnOff>),
    NoProofing(Option<OnOff>),
    SnapToGrid(Option<OnOff>),
    Vanish(Option<OnOff>),
    WebHidden(Option<OnOff>),
    Color(Color),
    Spacing(SignedTwipsMeasure),
    Width(Option<TextScale>),
    Kerning(HpsMeasure),
    Position(SignedHpsMeasure),
    Size(HpsMeasure),
    ComplexScriptSize(HpsMeasure),
    Highlight(HighlightColor),
    Underline(Underline),
    Effect(TextEffect),
    Border(Border),
    Shading(Shd),
    FitText(FitText),
    VerticalAlignment(VerticalAlignRun),
    Rtl(Option<OnOff>),
    ComplexScript(Option<OnOff>),
    EmphasisMark(Em),
    Language(Language),
    EastAsianLayout(EastAsianLayout),
    SpecialVanish(Option<OnOff>),
    OMath(Option<OnOff>),
}

impl RPrBase {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "rStyle" | "rFonts" | "b" | "bCs" | "i" | "iCs" | "caps" | "smallCaps" | "strike" | "dstrike"
            | "outline" | "shadow" | "emboss" | "imprint" | "noProof" | "snapToGrid" | "vanish" | "webHidden"
            | "color" | "spacing" | "w" | "kern" | "position" | "sz" | "szCs" | "highlight" | "u" | "effect"
            | "bdr" | "shd" | "fitText" | "vertAlign" | "rtl" | "cs" | "em" | "lang" | "eastAsianLayout"
            | "specVanish" | "oMath" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "rStyle" => {
                let val = xml_node
                    .attributes
                    .get("val")
                    .cloned()
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(RPrBase::RunStyle(val))
            }
            "rFonts" => Ok(RPrBase::RunFonts(Fonts::from_xml_element(xml_node)?)),
            "b" => Ok(RPrBase::Bold(parse_on_off_xml_element(xml_node)?)),
            "bCs" => Ok(RPrBase::ComplexScriptBold(parse_on_off_xml_element(xml_node)?)),
            "i" => Ok(RPrBase::Italic(parse_on_off_xml_element(xml_node)?)),
            "iCs" => Ok(RPrBase::ComplexScriptItalic(parse_on_off_xml_element(xml_node)?)),
            "caps" => Ok(RPrBase::Capitals(parse_on_off_xml_element(xml_node)?)),
            "smallCaps" => Ok(RPrBase::SmallCapitals(parse_on_off_xml_element(xml_node)?)),
            "strike" => Ok(RPrBase::Strikethrough(parse_on_off_xml_element(xml_node)?)),
            "dstrike" => Ok(RPrBase::DoubleStrikethrough(parse_on_off_xml_element(xml_node)?)),
            "outline" => Ok(RPrBase::Outline(parse_on_off_xml_element(xml_node)?)),
            "shadow" => Ok(RPrBase::Shadow(parse_on_off_xml_element(xml_node)?)),
            "emboss" => Ok(RPrBase::Emboss(parse_on_off_xml_element(xml_node)?)),
            "imprint" => Ok(RPrBase::Imprint(parse_on_off_xml_element(xml_node)?)),
            "noProof" => Ok(RPrBase::NoProofing(parse_on_off_xml_element(xml_node)?)),
            "snapToGrid" => Ok(RPrBase::SnapToGrid(parse_on_off_xml_element(xml_node)?)),
            "vanish" => Ok(RPrBase::Vanish(parse_on_off_xml_element(xml_node)?)),
            "webHidden" => Ok(RPrBase::WebHidden(parse_on_off_xml_element(xml_node)?)),
            "color" => Ok(RPrBase::Color(Color::from_xml_element(xml_node)?)),
            "spacing" => Ok(RPrBase::Spacing(SignedTwipsMeasure::from_xml_element(xml_node)?)),
            "w" => {
                let val = xml_node.attributes.get("val").map(|val| parse_text_scale_percent(val)).transpose()?;
                Ok(RPrBase::Width(val))
            }
            "kern" => Ok(RPrBase::Kerning(HpsMeasure::from_xml_element(xml_node)?)),
            "position" => Ok(RPrBase::Position(SignedHpsMeasure::from_xml_element(xml_node)?)),
            "sz" => Ok(RPrBase::Size(HpsMeasure::from_xml_element(xml_node)?)),
            "szCs" => Ok(RPrBase::ComplexScriptSize(HpsMeasure::from_xml_element(xml_node)?)),
            "highlight" => {
                let val_attr = xml_node
                    .attributes
                    .get("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(RPrBase::Highlight(val_attr.parse()?))
            }
            "u" => Ok(RPrBase::Underline(Underline::from_xml_element(xml_node)?)),
            "effect" => {
                let val_attr = xml_node
                    .attributes
                    .get("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(RPrBase::Effect(val_attr.parse()?))
            }
            "bdr" => Ok(RPrBase::Border(Border::from_xml_element(xml_node)?)),
            "shd" => Ok(RPrBase::Shading(Shd::from_xml_element(xml_node)?)),
            "fitText" => Ok(RPrBase::FitText(FitText::from_xml_element(xml_node)?)),
            "vertAlign" => {
                let val_attr = xml_node
                    .attributes
                    .get("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(RPrBase::VerticalAlignment(val_attr.parse()?))
            }
            "rtl" => Ok(RPrBase::Rtl(parse_on_off_xml_element(xml_node)?)),
            "cs" => Ok(RPrBase::ComplexScript(parse_on_off_xml_element(xml_node)?)),
            "em" => {
                let val_attr = xml_node
                    .attributes
                    .get("val")
                    .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
                Ok(RPrBase::EmphasisMark(val_attr.parse()?))
            }
            "lang" => Ok(RPrBase::Language(Language::from_xml_element(xml_node))),
            "eastAsianLayout" => Ok(RPrBase::EastAsianLayout(EastAsianLayout::from_xml_element(xml_node)?)),
            "specVanish" => Ok(RPrBase::SpecialVanish(parse_on_off_xml_element(xml_node)?)),
            "oMath" => Ok(RPrBase::OMath(parse_on_off_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "RPrBase"))),
        }
    }
}

#[cfg(test)]
impl RPrBase {
    pub fn test_run_style_xml() -> &'static str {
        r#"<rStyle val="Arial"></rStyle>"#
    }

    pub fn test_run_style_instance() -> Self {
        RPrBase::RunStyle(String::from("Arial"))
    }
}

// TODO Write some more unit tests
#[cfg(test)]
#[test]
pub fn test_r_pr_base_run_style_from_xml() {
    let xml = RPrBase::test_run_style_xml();
    let r_pr_base = RPrBase::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(r_pr_base, RPrBase::test_run_style_instance());
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RPrOriginal {
    pub r_pr_bases: Vec<RPrBase>,
}

impl RPrOriginal {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();
        for child_node in &xml_node.child_nodes {
            if RPrBase::is_choice_member(child_node.local_name()) {
                instance.r_pr_bases.push(RPrBase::from_xml_element(child_node)?);
            }
        }

        Ok(instance)
    }
}

#[cfg(test)]
impl RPrOriginal {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name}>{}</{node_name}>"#,
            RPrBase::test_run_style_xml(),
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            r_pr_bases: vec![RPrBase::test_run_style_instance()],
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_r_pr_original_from_xml() {
    let xml = RPrOriginal::test_xml("rPrOriginal");
    let r_pr_original = RPrOriginal::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(r_pr_original, RPrOriginal::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct RPrChange {
    pub base: TrackChange,
    pub run_properties: RPrOriginal,
}

impl RPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = TrackChange::from_xml_element(xml_node)?;
        let run_properties_node = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "rPr")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rPr"))?;

        let run_properties = RPrOriginal::from_xml_element(run_properties_node)?;
        Ok(Self { base, run_properties })
    }
}

#[cfg(test)]
impl RPrChange {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} id="0" author="John Smith" date="2001-10-26T21:32:52">
            {}
        </{node_name}>"#,
            RPrOriginal::test_xml("rPr"),
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            base: TrackChange::test_instance(),
            run_properties: RPrOriginal::test_instance(),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_r_pr_change_from_xml() {
    let xml = RPrChange::test_xml("rRpChange");
    let r_pr_change = RPrChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(r_pr_change, RPrChange::test_instance());
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RPr {
    pub r_pr_bases: Vec<RPrBase>,
    pub run_properties_change: Option<RPrChange>,
}

impl RPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: RPr = Default::default();
        for child_node in &xml_node.child_nodes {
            let child_node_name = child_node.local_name();
            if RPrBase::is_choice_member(child_node_name) {
                instance.r_pr_bases.push(RPrBase::from_xml_element(child_node)?);
            } else if child_node_name == "rPrChange" {
                instance.run_properties_change = Some(RPrChange::from_xml_element(child_node)?);
            }
        }

        Ok(instance)
    }
}

#[cfg(test)]
impl RPr {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name}>
            {}
            {}
        </{node_name}>"#,
            RPrBase::test_run_style_xml(),
            RPrChange::test_xml("rPrChange"),
            node_name = node_name,
        )
    }

    pub fn test_instance() -> Self {
        Self {
            r_pr_bases: vec![RPrBase::test_run_style_instance()],
            run_properties_change: Some(RPrChange::test_instance()),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_r_pr_from_xml() {
    let xml = RPr::test_xml("rPr");
    let r_pr_content = RPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(r_pr_content, RPr::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct SdtListItem {
    pub display_text: String,
    pub value: String,
}

impl SdtListItem {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let display_text = xml_node
            .attributes
            .get("displayText")
            .cloned()
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "displayText"))?;

        let value = xml_node
            .attributes
            .get("value")
            .cloned()
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "value"))?;

        Ok(Self { display_text, value })
    }
}

#[cfg(test)]
impl SdtListItem {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} displayText="Displayed" value="Some value"></{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            display_text: String::from("Displayed"),
            value: String::from("Some value"),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_sdt_list_item_from_xml() {
    let xml = SdtListItem::test_xml("sdtListItem");
    let sdt_list_item = SdtListItem::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(sdt_list_item, SdtListItem::test_instance());
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtComboBox {
    pub list_items: Vec<SdtListItem>,
    pub last_value: Option<String>,
}

impl SdtComboBox {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let last_value = xml_node.attributes.get("lastValue").cloned();

        let list_items = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| {
                if child_node.local_name() == "listItem" {
                    Some(SdtListItem::from_xml_element(child_node))
                } else {
                    None
                }
            })
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { list_items, last_value })
    }
}

#[cfg(test)]
impl SdtComboBox {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} lastValue="Some value">
            {}
            {}
        </{node_name}>"#,
            SdtListItem::test_xml("listItem"),
            SdtListItem::test_xml("listItem"),
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            list_items: vec![SdtListItem::test_instance(), SdtListItem::test_instance()],
            last_value: Some(String::from("Some value")),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_sdt_combo_box_from_xml() {
    let xml = SdtComboBox::test_xml("sdtComboBox");
    let sdt_combo_box = SdtComboBox::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(sdt_combo_box, SdtComboBox::test_instance());
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum SdtDateMappingType {
    #[strum(serialize = "text")]
    Text,
    #[strum(serialize = "date")]
    Date,
    #[strum(serialize = "dateTime")]
    DateTime,
}

impl SdtDateMappingType {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Option<Self>> {
        Ok(xml_node.attributes.get("val").map(|val| val.parse()).transpose()?)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtDate {
    pub date_format: Option<String>,
    pub language_id: Option<Lang>,
    pub store_mapped_data_as: Option<SdtDateMappingType>,
    pub calendar: Option<CalendarType>,

    pub full_date: Option<DateTime>,
}

impl SdtDate {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();
        instance.full_date = xml_node.attributes.get("fullDate").cloned();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "dateFormat" => instance.date_format = child_node.attributes.get("val").cloned(),
                "lid" => instance.language_id = child_node.attributes.get("val").cloned(),
                "storeMappedDataAs" => {
                    instance.store_mapped_data_as = SdtDateMappingType::from_xml_element(child_node)?
                }
                "calendar" => {
                    instance.calendar = child_node.attributes.get("val").map(|val| val.parse()).transpose()?;
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[cfg(test)]
impl SdtDate {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} fullDate="2001-10-26T21:32:52">
            <dateFormat val="MM-YYYY" />
            <lid val="ja-JP" />
            <storeMappedDataAs val="dateTime" />
            <calendar val="gregorian" />
        </{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            date_format: Some(String::from("MM-YYYY")),
            language_id: Some(Lang::from("ja-JP")),
            store_mapped_data_as: Some(SdtDateMappingType::DateTime),
            calendar: Some(CalendarType::Gregorian),
            full_date: Some(DateTime::from("2001-10-26T21:32:52")),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_sdt_date_from_xml() {
    let xml = SdtDate::test_xml("sdtDate");
    let sdt_date = SdtDate::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(sdt_date, SdtDate::test_instance());
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtDocPart {
    pub doc_part_gallery: Option<String>,
    pub doc_part_category: Option<String>,
    pub doc_part_unique: Option<OnOff>,
}

impl SdtDocPart {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "docPartGallery" => instance.doc_part_gallery = child_node.attributes.get("val").cloned(),
                "docPartCategory" => instance.doc_part_category = child_node.attributes.get("val").cloned(),
                "docPartUnique" => instance.doc_part_unique = parse_on_off_xml_element(child_node)?,
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[cfg(test)]
impl SdtDocPart {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name}>
            <docPartGallery val="Some string" />
            <docPartCategory val="Some string" />
            <docPartUnique val="true" />
        </{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            doc_part_gallery: Some(String::from("Some string")),
            doc_part_category: Some(String::from("Some string")),
            doc_part_unique: Some(true),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_sdt_doc_part_from_xml() {
    let xml = SdtDocPart::test_xml("sdtDocPart");
    let sdt_doc_part = SdtDocPart::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(sdt_doc_part, SdtDocPart::test_instance());
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtDropDownList {
    pub list_items: Vec<SdtListItem>,
    pub last_value: Option<String>,
}

impl SdtDropDownList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let last_value = xml_node.attributes.get("lastValue").cloned();

        let list_items = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| {
                if child_node.local_name() == "listItem" {
                    Some(SdtListItem::from_xml_element(child_node))
                } else {
                    None
                }
            })
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { list_items, last_value })
    }
}

#[cfg(test)]
impl SdtDropDownList {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name} lastValue="Some value">
            {}
            {}
        </{node_name}>"#,
            SdtListItem::test_xml("listItem"),
            SdtListItem::test_xml("listItem"),
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            list_items: vec![SdtListItem::test_instance(), SdtListItem::test_instance()],
            last_value: Some(String::from("Some value")),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_sdt_drop_down_list_from_xml() {
    let xml = SdtDropDownList::test_xml("sdtDropDownList");
    let sdt_combo_box = SdtDropDownList::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(sdt_combo_box, SdtDropDownList::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub struct SdtText {
    pub is_multi_line: OnOff,
}

impl SdtText {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let is_multi_line_attr = xml_node
            .attributes
            .get("multiLine")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "multiLine"))?;

        Ok(Self {
            is_multi_line: parse_xml_bool(is_multi_line_attr)?,
        })
    }
}

#[cfg(test)]
impl SdtText {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} multiLine="true"></{node_name}>"#, node_name = node_name)
    }

    pub fn test_instance() -> Self {
        Self { is_multi_line: true }
    }
}

#[cfg(test)]
#[test]
pub fn test_sdt_text_from_xml() {
    let xml = SdtText::test_xml("sdtText");
    let sdt_text = SdtText::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(sdt_text, SdtText::test_instance());
}

#[derive(Debug, Clone, PartialEq)]
pub enum SdtPrChoice {
    Equation,
    ComboBox(SdtComboBox),
    Date(SdtDate),
    DocumentPartObject(SdtDocPart),
    DocumentPartList(SdtDocPart),
    DropDownList(SdtDropDownList),
    Picture,
    RichText,
    Text(SdtText),
    Citation,
    Group,
    Bibliography,
}

impl SdtPrChoice {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "equation" | "comboBox" | "date" | "docPartObj" | "docPartList" | "dropDownList" | "picture"
            | "richText" | "text" | "citation" | "group" | "bibliography" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "equation" => Ok(SdtPrChoice::Equation),
            "comboBox" => Ok(SdtPrChoice::ComboBox(SdtComboBox::from_xml_element(xml_node)?)),
            "date" => Ok(SdtPrChoice::Date(SdtDate::from_xml_element(xml_node)?)),
            "docPartObj" => Ok(SdtPrChoice::DocumentPartObject(SdtDocPart::from_xml_element(xml_node)?)),
            "docPartList" => Ok(SdtPrChoice::DocumentPartList(SdtDocPart::from_xml_element(xml_node)?)),
            "dropDownList" => Ok(SdtPrChoice::DropDownList(SdtDropDownList::from_xml_element(xml_node)?)),
            "picture" => Ok(SdtPrChoice::Picture),
            "richText" => Ok(SdtPrChoice::RichText),
            "text" => Ok(SdtPrChoice::Text(SdtText::from_xml_element(xml_node)?)),
            "citation" => Ok(SdtPrChoice::Citation),
            "group" => Ok(SdtPrChoice::Group),
            "bibliography" => Ok(SdtPrChoice::Bibliography),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "SdtPrChoice"))),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_sdt_pr_control_choice_from_xml() {
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str("<equation></equation>").unwrap()).unwrap(),
        SdtPrChoice::Equation,
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtComboBox::test_xml("comboBox")).unwrap()).unwrap(),
        SdtPrChoice::ComboBox(SdtComboBox::test_instance()),
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtDate::test_xml("date")).unwrap()).unwrap(),
        SdtPrChoice::Date(SdtDate::test_instance()),
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtDocPart::test_xml("docPartObj")).unwrap()).unwrap(),
        SdtPrChoice::DocumentPartObject(SdtDocPart::test_instance()),
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtDocPart::test_xml("docPartList")).unwrap()).unwrap(),
        SdtPrChoice::DocumentPartList(SdtDocPart::test_instance()),
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtDropDownList::test_xml("dropDownList")).unwrap()).unwrap(),
        SdtPrChoice::DropDownList(SdtDropDownList::test_instance()),
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str("<picture></picture>").unwrap()).unwrap(),
        SdtPrChoice::Picture,
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str("<richText></richText>").unwrap()).unwrap(),
        SdtPrChoice::RichText,
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtText::test_xml("text")).unwrap()).unwrap(),
        SdtPrChoice::Text(SdtText::test_instance()),
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str("<citation></citation>").unwrap()).unwrap(),
        SdtPrChoice::Citation,
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str("<group></group>").unwrap()).unwrap(),
        SdtPrChoice::Group,
    );
    assert_eq!(
        SdtPrChoice::from_xml_element(&XmlNode::from_str("<bibliography></bibliography>").unwrap()).unwrap(),
        SdtPrChoice::Bibliography,
    );
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum Lock {
    #[strum(serialize = "sdtLocked")]
    SdtLocked,
    #[strum(serialize = "contentLocked")]
    ContentLocked,
    #[strum(serialize = "unlocked")]
    Unlocked,
    #[strum(serialize = "sdtContentLocked")]
    SdtContentLocked,
}

impl Lock {
    pub fn from_xml_element(xml_node: &XmlNode) -> std::result::Result<Option<Self>, strum::ParseError> {
        xml_node.attributes.get("val").map(|val| val.parse()).transpose()
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Placeholder {
    pub document_part: String,
}

impl Placeholder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let document_part_node = xml_node
            .child_nodes
            .first()
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "docPart"))?;

        let document_part = document_part_node
            .attributes
            .get("val")
            .cloned()
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;

        Ok(Self { document_part })
    }
}

#[cfg(test)]
impl Placeholder {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(
            r#"<{node_name}>
            <docPart val="title" />
        </{node_name}>"#,
            node_name = node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            document_part: String::from("title"),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_placeholder_from_xml() {
    let xml = Placeholder::test_xml("placeholder");
    assert_eq!(
        Placeholder::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
        Placeholder::test_instance()
    );
}

#[derive(Debug, Clone, PartialEq)]
pub struct DataBinding {
    pub prefix_mappings: Option<String>,
    pub xpath: String,
    pub store_item_id: String,
}

impl DataBinding {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut prefix_mappings = None;
        let mut xpath = None;
        let mut store_item_id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "prefixMappings" => prefix_mappings = Some(value.clone()),
                "xpath" => xpath = Some(value.clone()),
                "storeItemID" => store_item_id = Some(value.clone()),
                _ => (),
            }
        }

        let xpath = xpath.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "xpath"))?;
        let store_item_id =
            store_item_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "storeItemId"))?;

        Ok(Self {
            prefix_mappings,
            xpath,
            store_item_id,
        })
    }
}

#[cfg(test)]
impl DataBinding {
    pub fn test_xml(node_name: &'static str) -> String {
        format!(r#"<{node_name} prefixMappings="xmlns:ns0='http://example.com/example'" xpath="//ns0:book" storeItemID="testXmlPart">
        </{node_name}>"#
            , node_name=node_name
        )
    }

    pub fn test_instance() -> Self {
        Self {
            prefix_mappings: Some(String::from("xmlns:ns0='http://example.com/example'")),
            xpath: String::from("//ns0:book"),
            store_item_id: String::from("testXmlPart"),
        }
    }
}

#[cfg(test)]
#[test]
pub fn test_data_binding_from_xml() {
    let xml = DataBinding::test_xml("dataBinding");
    assert_eq!(
        DataBinding::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
        DataBinding::test_instance()
    );
}

/*
<xsd:complexType name="CT_SdtPr">
    <xsd:sequence>
      <xsd:element name="rPr" type="CT_RPr" minOccurs="0"/>
      <xsd:element name="alias" type="CT_String" minOccurs="0"/>
      <xsd:element name="tag" type="CT_String" minOccurs="0"/>
      <xsd:element name="id" type="CT_DecimalNumber" minOccurs="0"/>
      <xsd:element name="lock" type="CT_Lock" minOccurs="0"/>
      <xsd:element name="placeholder" type="CT_Placeholder" minOccurs="0"/>
      <xsd:element name="temporary" type="CT_OnOff" minOccurs="0"/>
      <xsd:element name="showingPlcHdr" type="CT_OnOff" minOccurs="0"/>
      <xsd:element name="dataBinding" type="CT_DataBinding" minOccurs="0"/>
      <xsd:element name="label" type="CT_DecimalNumber" minOccurs="0"/>
      <xsd:element name="tabIndex" type="CT_UnsignedDecimalNumber" minOccurs="0"/>
      <xsd:choice minOccurs="0" maxOccurs="1">
        <xsd:element name="equation" type="CT_Empty"/>
        <xsd:element name="comboBox" type="CT_SdtComboBox"/>
        <xsd:element name="date" type="CT_SdtDate"/>
        <xsd:element name="docPartObj" type="CT_SdtDocPart"/>
        <xsd:element name="docPartList" type="CT_SdtDocPart"/>
        <xsd:element name="dropDownList" type="CT_SdtDropDownList"/>
        <xsd:element name="picture" type="CT_Empty"/>
        <xsd:element name="richText" type="CT_Empty"/>
        <xsd:element name="text" type="CT_SdtText"/>
        <xsd:element name="citation" type="CT_Empty"/>
        <xsd:element name="group" type="CT_Empty"/>
        <xsd:element name="bibliography" type="CT_Empty"/>
      </xsd:choice>
    </xsd:sequence>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtPr {
    pub run_properties: Option<RPr>,
    pub alias: Option<String>,
    pub tag: Option<String>,
    pub id: Option<DecimalNumber>,
    pub lock: Option<Lock>,
    pub placeholder: Option<Placeholder>,
    pub temporary: Option<OnOff>,
    pub showing_placeholder_header: Option<OnOff>,
    pub data_binding: Option<DataBinding>,
    pub label: Option<DecimalNumber>,
    pub tab_index: Option<UnsignedDecimalNumber>,
    pub control_choice: Option<SdtPrChoice>,
}

/*
<xsd:complexType name="CT_SdtRun">
    <xsd:sequence>
      <xsd:element name="sdtPr" type="CT_SdtPr" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sdtEndPr" type="CT_SdtEndPr" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="sdtContent" type="CT_SdtContentRun" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq)]
pub struct SdtRun {
    pub sdt_properties: Option<SdtPr>,
    //pub sdt_end_properties: Option<SdtEndPr>,
    //pub sdt_content: Option<SdtContentRun>,
}

/*
<xsd:group name="EG_ContentRunContent">
    <xsd:choice>
      <xsd:element name="customXml" type="CT_CustomXmlRun"/>
      <xsd:element name="smartTag" type="CT_SmartTagRun"/>
      <xsd:element name="sdt" type="CT_SdtRun"/>
      <xsd:element name="dir" type="CT_DirContentRun"/>
      <xsd:element name="bdo" type="CT_BdoContentRun"/>
      <xsd:element name="r" type="CT_R"/>
      <xsd:group ref="EG_RunLevelElts" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:choice>
  </xsd:group>
*/
#[derive(Debug, Clone, PartialEq)]
pub enum ContentRunContent {
    CustomXml(CustomXmlRun),
    SmartTag(SmartTagRun),
    Sdt(SdtRun),
    // Bidirectional(DirContentRun),
    // BidirectionalOverride(BdoContentRun),
    // Run(R),
    RunLevelElts(Vec<RunLevelElts>),
}

impl ContentRunContent {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "customXml" | "smartTag" | "sdt" | "dir" | "bdo" | "r" => true,
            _ => RunLevelElts::is_choice_member(&node_name),
        }
    }

    pub fn from_xml_element(_xml_node: &XmlNode) -> Result<Self> {
        unimplemented!();
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RunTrackChangeChoice {
    ContentRunContent(ContentRunContent),
    // TODO
    // OMathMathElements(OMathMathElements),
}

/*
<xsd:complexType name="CT_RunTrackChange">
    <xsd:complexContent>
      <xsd:extension base="CT_TrackChange">
        <xsd:choice minOccurs="0" maxOccurs="unbounded">
          <xsd:group ref="EG_ContentRunContent"/>
          <xsd:group ref="m:EG_OMathMathElements"/>
        </xsd:choice>
      </xsd:extension>
    </xsd:complexContent>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq)]
pub struct RunTrackChange {
    pub base: TrackChange,
    pub choices: Vec<RunTrackChangeChoice>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum RangeMarkupElements {
    BookmarkStart(Bookmark),
    BookmarkEnd(MarkupRange),
    MoveFromRangeStart(MoveBookmark),
    MoveFromRangeEnd(MarkupRange),
    MoveToRangeStart(MoveBookmark),
    MoveToRangeEnd(MarkupRange),
    CommentRangeStart(MarkupRange),
    CommentRangeEnd(MarkupRange),
    CustomXmlInsertRangeStart(TrackChange),
    CustomXmlInsertRangeEnd(Markup),
    CustomXmlDeleteRangeStart(TrackChange),
    CustomXmlDeleteRangeEnd(Markup),
    CustomXmlMoveFromRangeStart(TrackChange),
    CustomXmlMoveFromRangeEnd(Markup),
    CustomXmlMoveToRangeStart(TrackChange),
    CustomXmlMoveToRangeEnd(Markup),
}

/*
<xsd:group name="EG_RangeMarkupElements">
    <xsd:choice>
      <xsd:element name="bookmarkStart" type="CT_Bookmark"/>
      <xsd:element name="bookmarkEnd" type="CT_MarkupRange"/>
      <xsd:element name="moveFromRangeStart" type="CT_MoveBookmark"/>
      <xsd:element name="moveFromRangeEnd" type="CT_MarkupRange"/>
      <xsd:element name="moveToRangeStart" type="CT_MoveBookmark"/>
      <xsd:element name="moveToRangeEnd" type="CT_MarkupRange"/>
      <xsd:element name="commentRangeStart" type="CT_MarkupRange"/>
      <xsd:element name="commentRangeEnd" type="CT_MarkupRange"/>
      <xsd:element name="customXmlInsRangeStart" type="CT_TrackChange"/>
      <xsd:element name="customXmlInsRangeEnd" type="CT_Markup"/>
      <xsd:element name="customXmlDelRangeStart" type="CT_TrackChange"/>
      <xsd:element name="customXmlDelRangeEnd" type="CT_Markup"/>
      <xsd:element name="customXmlMoveFromRangeStart" type="CT_TrackChange"/>
      <xsd:element name="customXmlMoveFromRangeEnd" type="CT_Markup"/>
      <xsd:element name="customXmlMoveToRangeStart" type="CT_TrackChange"/>
      <xsd:element name="customXmlMoveToRangeEnd" type="CT_Markup"/>
    </xsd:choice>
  </xsd:group>
*/
impl RangeMarkupElements {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "bookmarkStart"
            | "bookmarkEnd"
            | "moveFromRangeStart"
            | "moveFromRangeEnd"
            | "moveToRangeStart"
            | "moveToRangeEnd"
            | "commentRangeStart"
            | "commentRangeEnd"
            | "customXmlInsRangeStart"
            | "customXmlInsRangeEnd"
            | "customXmlDelRangeStart"
            | "customXmlDelRangeEnd"
            | "customXmlMoveFromRangeStart"
            | "customXmlMoveFromRangeEnd"
            | "customXmlMoveToRangeStart"
            | "customXmlMoveToRangeEnd" => true,
            _ => false,
        }
    }
}

/*
<xsd:group name="EG_MathContent">
    <xsd:choice>
      <xsd:element ref="m:oMathPara"/>
      <xsd:element ref="m:oMath"/>
    </xsd:choice>
  </xsd:group>
*/
// TODO
#[derive(Debug, Clone, PartialEq)]
pub enum MathContent {
    // OMathParagraph(OMathParagraph),
// OMath(OMath),
}

impl MathContent {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "oMathPara" | "oMath" => true,
            _ => false,
        }
    }
}

/*
<xsd:group name="EG_RunLevelElts">
    <xsd:choice>
      <xsd:element name="proofErr" minOccurs="0" type="CT_ProofErr"/>
      <xsd:element name="permStart" minOccurs="0" type="CT_PermStart"/>
      <xsd:element name="permEnd" minOccurs="0" type="CT_Perm"/>
      <xsd:group ref="EG_RangeMarkupElements" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="ins" type="CT_RunTrackChange" minOccurs="0"/>
      <xsd:element name="del" type="CT_RunTrackChange" minOccurs="0"/>
      <xsd:element name="moveFrom" type="CT_RunTrackChange"/>
      <xsd:element name="moveTo" type="CT_RunTrackChange"/>
      <xsd:group ref="EG_MathContent" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:choice>
  </xsd:group>
*/
#[derive(Debug, Clone, PartialEq)]
pub enum RunLevelElts {
    ProofError(Option<ProofErr>),
    PermissionStart(Option<PermStart>),
    PermissionEnd(Option<Perm>),
    RangeMarkupElements(Vec<RangeMarkupElements>),
    Insert(Option<RunTrackChange>),
    Delete(Option<RunTrackChange>),
    MoveFrom(RunTrackChange),
    MoveTo(RunTrackChange),
    MathContent(Vec<MathContent>),
}

impl RunLevelElts {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "proofErr" | "permStart" | "permEnd" | "ins" | "del" | "moveFrom" | "moveTo" => true,
            _ => RangeMarkupElements::is_choice_member(&node_name) || MathContent::is_choice_member(&node_name),
        }
    }
}

/*
<xsd:group name="EG_ContentBlockContent">
    <xsd:choice>
      <xsd:element name="customXml" type="CT_CustomXmlBlock"/>
      <xsd:element name="sdt" type="CT_SdtBlock"/>
      <xsd:element name="p" type="CT_P" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="tbl" type="CT_Tbl" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:group ref="EG_RunLevelElts" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:choice>
  </xsd:group>
*/
#[derive(Debug, Clone, PartialEq)]
pub enum ContentBlockContent {
    // CustomXml(CustomXmlBlock),
    // Sdt(SdtBlock),
    // Paragraph(Vec<P>),
    // Table(Vec<Tbl>),
    RunLevelElement(Vec<RunLevelElts>),
}

/*
<xsd:group name="EG_BlockLevelChunkElts">
    <xsd:choice>
      <xsd:group ref="EG_ContentBlockContent" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:choice>
  </xsd:group>
*/
#[derive(Debug, Clone, PartialEq)]
pub enum BlockLevelChunkElts {
    Content(Vec<ContentBlockContent>),
}
// <xsd:group name="EG_BlockLevelElts">
//     <xsd:choice>
//       <xsd:group ref="EG_BlockLevelChunkElts" minOccurs="0" maxOccurs="unbounded"/>
//       <xsd:element name="altChunk" type="CT_AltChunk" minOccurs="0" maxOccurs="unbounded"/>
//     </xsd:choice>
//   </xsd:group>
#[derive(Debug, Clone, PartialEq)]
pub enum BlockLevelElts {
    Chunks(Vec<BlockLevelChunkElts>),
    //AltChunks(Vec<AltChunk>),
}
