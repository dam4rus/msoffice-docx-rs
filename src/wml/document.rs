use super::{
    drawing::{Anchor, Inline},
    error::ParseHexColorError,
};
use msoffice_shared::{
    drawingml::{parse_hex_color_rgb, HexColorRGB},
    error::{
        LimitViolationError, MaxOccurs, MissingAttributeError, MissingChildNodeError, NotGroupMemberError,
        ParseBoolError, PatternRestrictionError,
    },
    relationship::RelationshipId,
    sharedtypes::{
        CalendarType, Lang, OnOff, Percentage, PositiveUniversalMeasure, TwipsMeasure, UniversalMeasure,
        VerticalAlignRun, XAlign, XmlName, YAlign,
    },
    util::XmlNodeExt,
    xml::{parse_xml_bool, XmlNode},
    xsdtypes::XsdChoice,
};
use regex::Regex;
use std::str::FromStr;

pub type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

pub type UcharHexNumber = u8;
pub type ShortHexNumber = u16;
pub type LongHexNumber = u32;
pub type UnqualifiedPercentage = i32;
pub type DecimalNumber = i32;
pub type UnsignedDecimalNumber = u32;
pub type DateTime = String;
pub type MacroName = String; // maxLength=33
pub type FFName = String; // maxLength=65
pub type FFHelpTextVal = String; // maxLength=256
pub type FFStatusTextVal = String; // maxLength=140
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
    Decimal(UnqualifiedPercentage),
    Percentage(Percentage),
}

impl FromStr for DecimalNumberOrPercent {
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> std::result::Result<Self, Self::Err> {
        if let Ok(value) = s.parse::<UnqualifiedPercentage>() {
            Ok(DecimalNumberOrPercent::Decimal(value))
        } else {
            Ok(DecimalNumberOrPercent::Percentage(s.parse()?))
        }
    }
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
        Ok(xml_node.get_val_attribute()?.parse()?)
    }
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
        Ok(xml_node.get_val_attribute()?.parse()?)
    }
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
        Ok(xml_node.get_val_attribute()?.parse()?)
    }
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
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "name"))?
            .clone();

        Ok(Self { base, name })
    }
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
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "author"))?
            .clone();

        let date = xml_node
            .attributes
            .get("date")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "date"))?
            .clone();

        Ok(Self { base, author, date })
    }
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
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "author"))?
            .clone();

        let date = xml_node.attributes.get("date").cloned();

        Ok(Self { base, author, date })
    }
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
                "placeholder" => placeholder = Some(child_node.get_val_attribute()?.clone()),
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

        let paragraph_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| PContent::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        let field_codes = field_codes.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "instr"))?;

        Ok(Self {
            field_codes,
            field_lock,
            dirty,
            paragraph_contents,
        })
    }
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

        let paragraph_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| PContent::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

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

#[derive(Debug, Clone, PartialEq)]
pub struct Rel {
    pub rel_id: RelationshipId,
}

impl Rel {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let rel_id = xml_node
            .attributes
            .get("r:id")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?
            .clone();

        Ok(Self { rel_id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum PContent {
    ContentRunContent(ContentRunContent),
    SimpleField(SimpleField),
    Hyperlink(Hyperlink),
    SubDocument(Rel),
}

impl XsdChoice for PContent {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "fldSimple" | "hyperlink" | "subDoc" => true,
            _ => ContentRunContent::is_choice_member(&node_name),
        }
    }

    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
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

        Ok(Self {
            value: value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?,
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

impl XsdChoice for RPrBase {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "rStyle" | "rFonts" | "b" | "bCs" | "i" | "iCs" | "caps" | "smallCaps" | "strike" | "dstrike"
            | "outline" | "shadow" | "emboss" | "imprint" | "noProof" | "snapToGrid" | "vanish" | "webHidden"
            | "color" | "spacing" | "w" | "kern" | "position" | "sz" | "szCs" | "highlight" | "u" | "effect"
            | "bdr" | "shd" | "fitText" | "vertAlign" | "rtl" | "cs" | "em" | "lang" | "eastAsianLayout"
            | "specVanish" | "oMath" => true,
            _ => false,
        }
    }

    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "rStyle" => Ok(RPrBase::RunStyle(xml_node.get_val_attribute()?.clone())),
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
                let val = xml_node
                    .attributes
                    .get("val")
                    .map(|val| parse_text_scale_percent(val))
                    .transpose()?;
                Ok(RPrBase::Width(val))
            }
            "kern" => Ok(RPrBase::Kerning(HpsMeasure::from_xml_element(xml_node)?)),
            "position" => Ok(RPrBase::Position(SignedHpsMeasure::from_xml_element(xml_node)?)),
            "sz" => Ok(RPrBase::Size(HpsMeasure::from_xml_element(xml_node)?)),
            "szCs" => Ok(RPrBase::ComplexScriptSize(HpsMeasure::from_xml_element(xml_node)?)),
            "highlight" => Ok(RPrBase::Highlight(xml_node.get_val_attribute()?.parse()?)),
            "u" => Ok(RPrBase::Underline(Underline::from_xml_element(xml_node)?)),
            "effect" => Ok(RPrBase::Effect(xml_node.get_val_attribute()?.parse()?)),
            "bdr" => Ok(RPrBase::Border(Border::from_xml_element(xml_node)?)),
            "shd" => Ok(RPrBase::Shading(Shd::from_xml_element(xml_node)?)),
            "fitText" => Ok(RPrBase::FitText(FitText::from_xml_element(xml_node)?)),
            "vertAlign" => Ok(RPrBase::VerticalAlignment(xml_node.get_val_attribute()?.parse()?)),
            "rtl" => Ok(RPrBase::Rtl(parse_on_off_xml_element(xml_node)?)),
            "cs" => Ok(RPrBase::ComplexScript(parse_on_off_xml_element(xml_node)?)),
            "em" => Ok(RPrBase::EmphasisMark(xml_node.get_val_attribute()?.parse()?)),
            "lang" => Ok(RPrBase::Language(Language::from_xml_element(xml_node))),
            "eastAsianLayout" => Ok(RPrBase::EastAsianLayout(EastAsianLayout::from_xml_element(xml_node)?)),
            "specVanish" => Ok(RPrBase::SpecialVanish(parse_on_off_xml_element(xml_node)?)),
            "oMath" => Ok(RPrBase::OMath(parse_on_off_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "RPrBase"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RPrOriginal {
    pub r_pr_bases: Vec<RPrBase>,
}

impl RPrOriginal {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let r_pr_bases = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| RPrBase::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { r_pr_bases })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct RPrChange {
    pub base: TrackChange,
    pub run_properties: RPrOriginal,
}

impl RPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = TrackChange::from_xml_element(xml_node)?;
        let run_properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "rPr")
            .map(|child_node| RPrOriginal::from_xml_element(child_node))
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rPr"))?;

        Ok(Self { base, run_properties })
    }
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
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "displayText"))?
            .clone();

        let value = xml_node
            .attributes
            .get("value")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "value"))?
            .clone();

        Ok(Self { display_text, value })
    }
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
            .filter(|child_node| child_node.local_name() == "listItem")
            .map(|child_node| SdtListItem::from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { list_items, last_value })
    }
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
                "dateFormat" => instance.date_format = Some(child_node.get_val_attribute()?.clone()),
                "lid" => instance.language_id = Some(child_node.get_val_attribute()?.clone()),
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
                "docPartGallery" => instance.doc_part_gallery = Some(child_node.get_val_attribute()?.clone()),
                "docPartCategory" => instance.doc_part_category = Some(child_node.get_val_attribute()?.clone()),
                "docPartUnique" => instance.doc_part_unique = parse_on_off_xml_element(child_node)?,
                _ => (),
            }
        }

        Ok(instance)
    }
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
            .filter(|child_node| child_node.local_name() == "listItem")
            .map(|child_node| SdtListItem::from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { list_items, last_value })
    }
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
        let document_part = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "docPart")
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "docPart"))?
            .get_val_attribute()?
            .clone();

        Ok(Self { document_part })
    }
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

impl SdtPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => instance.run_properties = Some(RPr::from_xml_element(child_node)?),
                "alias" => instance.alias = Some(child_node.get_val_attribute()?.clone()),
                "tag" => instance.tag = Some(child_node.get_val_attribute()?.clone()),
                "id" => instance.id = Some(child_node.get_val_attribute()?.parse()?),
                "lock" => instance.lock = child_node.attributes.get("val").map(|val| val.parse()).transpose()?,
                "placeholder" => instance.placeholder = Some(Placeholder::from_xml_element(child_node)?),
                "temporary" => instance.temporary = parse_on_off_xml_element(child_node)?,
                "showingPlcHdr" => instance.showing_placeholder_header = parse_on_off_xml_element(child_node)?,
                "dataBinding" => instance.data_binding = Some(DataBinding::from_xml_element(child_node)?),
                "label" => instance.label = Some(child_node.get_val_attribute()?.parse()?),
                "tabIndex" => instance.tab_index = Some(child_node.get_val_attribute()?.parse()?),
                node_name @ _ if SdtPrChoice::is_choice_member(node_name) => {
                    instance.control_choice = Some(SdtPrChoice::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtEndPr {
    pub run_properties_vec: Vec<RPr>,
}

impl SdtEndPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let run_properties_vec = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "rPr")
            .map(|child_node| RPr::from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { run_properties_vec })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtContentRun {
    pub p_contents: Vec<PContent>,
}

impl SdtContentRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let p_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| PContent::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { p_contents })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtRun {
    pub sdt_properties: Option<SdtPr>,
    pub sdt_end_properties: Option<SdtEndPr>,
    pub sdt_content: Option<SdtContentRun>,
}

impl SdtRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "sdtPr" => instance.sdt_properties = Some(SdtPr::from_xml_element(child_node)?),
                "sdtEndPr" => instance.sdt_end_properties = Some(SdtEndPr::from_xml_element(child_node)?),
                "sdtContent" => instance.sdt_content = Some(SdtContentRun::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum Direction {
    #[strum(serialize = "ltr")]
    LeftToRight,
    #[strum(serialize = "rtl")]
    RightToLeft,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct DirContentRun {
    pub p_contents: Vec<PContent>,
    pub value: Option<Direction>,
}

impl DirContentRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let value = xml_node.attributes.get("val").map(|val| val.parse()).transpose()?;

        let p_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| PContent::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { p_contents, value })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct BdoContentRun {
    pub p_contents: Vec<PContent>,
    pub value: Option<Direction>,
}

impl BdoContentRun {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let value = xml_node.attributes.get("val").map(|val| val.parse()).transpose()?;

        let p_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| PContent::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { p_contents, value })
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum BrType {
    #[strum(serialize = "page")]
    Page,
    #[strum(serialize = "column")]
    Column,
    #[strum(serialzie = "textWrapping")]
    TextWrapping,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum BrClear {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "all")]
    All,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Br {
    pub break_type: Option<BrType>,
    pub clear: Option<BrClear>,
}

impl Br {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();
        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "type" => instance.break_type = Some(value.parse()?),
                "clear" => instance.clear = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Text {
    pub text: String,
    pub xml_space: Option<String>, // default or preserve
}

impl Text {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let xml_space = xml_node.attributes.get("xml:space").cloned();

        let text = xml_node
            .text
            .as_ref()
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "Text node"))?
            .clone();

        Ok(Self { text, xml_space })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Sym {
    pub font: Option<String>,
    pub character: Option<ShortHexNumber>,
}

impl Sym {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "font" => instance.font = Some(value.clone()),
                "char" => instance.character = Some(ShortHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Control {
    pub name: Option<String>,
    pub shapeid: Option<String>,
    pub rel_id: Option<RelationshipId>,
}

impl Control {
    pub fn from_xml_element(xml_node: &XmlNode) -> Self {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "name" => instance.name = Some(value.clone()),
                "shapeid" => instance.shapeid = Some(value.clone()),
                "r:id" => instance.rel_id = Some(value.clone()),
                _ => (),
            }
        }

        instance
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum ObjectDrawAspect {
    #[strum(serialize = "content")]
    Content,
    #[strum(serialize = "icon")]
    Icon,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ObjectEmbed {
    pub draw_aspect: Option<ObjectDrawAspect>,
    pub rel_id: RelationshipId,
    pub application_id: Option<String>,
    pub shape_id: Option<String>,
    pub field_codes: Option<String>,
}

impl ObjectEmbed {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut draw_aspect = None;
        let mut rel_id = None;
        let mut application_id = None;
        let mut shape_id = None;
        let mut field_codes = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "drawAspect" => draw_aspect = Some(value.parse()?),
                "r:id" => rel_id = Some(value.clone()),
                "progId" => application_id = Some(value.clone()),
                "shapeId" => shape_id = Some(value.clone()),
                "fieldCodes" => field_codes = Some(value.clone()),
                _ => (),
            }
        }

        let rel_id = rel_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:id"))?;

        Ok(Self {
            draw_aspect,
            rel_id,
            application_id,
            shape_id,
            field_codes,
        })
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum ObjectUpdateMode {
    #[strum(serialize = "always")]
    Always,
    #[strum(serialize = "onCall")]
    OnCall,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ObjectLink {
    pub base: ObjectEmbed,
    pub update_mode: ObjectUpdateMode,
    pub locked_field: Option<OnOff>,
}

impl ObjectLink {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = ObjectEmbed::from_xml_element(xml_node)?;
        let mut update_mode = None;
        let mut locked_field = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "updateMode" => update_mode = Some(value.parse()?),
                "lockedField" => locked_field = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let update_mode = update_mode.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "updateMode"))?;

        Ok(Self {
            base,
            update_mode,
            locked_field,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ObjectChoice {
    Control(Control),
    ObjectLink(ObjectLink),
    ObjectEmbed(ObjectEmbed),
    Movie(Rel),
}

impl ObjectChoice {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "control" | "objectLink" | "objectEmbed" | "movie" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "control" => Ok(ObjectChoice::Control(Control::from_xml_element(xml_node))),
            "objectLink" => Ok(ObjectChoice::ObjectLink(ObjectLink::from_xml_element(xml_node)?)),
            "objectEmbed" => Ok(ObjectChoice::ObjectEmbed(ObjectEmbed::from_xml_element(xml_node)?)),
            "movie" => Ok(ObjectChoice::Movie(Rel::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "ObjectChoice",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum DrawingChoice {
    Anchor(Anchor),
    Inline(Inline),
}

impl XsdChoice for DrawingChoice {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "anchor" | "inline" => true,
            _ => false,
        }
    }

    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "anchor" => Ok(DrawingChoice::Anchor(Anchor::from_xml_element(xml_node)?)),
            "inline" => Ok(DrawingChoice::Inline(Inline::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "DrawingChoice",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Drawing {
    pub anchor_or_inline_vec: Vec<DrawingChoice>,
}

impl Drawing {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let anchor_or_inline_vec = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| DrawingChoice::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { anchor_or_inline_vec })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Object {
    pub drawing: Option<Drawing>,
    pub choice: Option<ObjectChoice>,
    pub original_image_width: Option<TwipsMeasure>,
    pub original_image_height: Option<TwipsMeasure>,
}

impl Object {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "dxaOrig" => instance.original_image_width = Some(value.parse()?),
                "dyaOrig" => instance.original_image_height = Some(value.parse()?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "drawing" => instance.drawing = Some(Drawing::from_xml_element(child_node)?),
                node_name @ _ if ObjectChoice::is_choice_member(node_name) => {
                    instance.choice = Some(ObjectChoice::from_xml_element(child_node)?)
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum InfoTextType {
    #[strum(serialize = "text")]
    Text,
    #[strum(serialize = "autoText")]
    AutoText,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FFHelpText {
    pub info_text_type: Option<InfoTextType>,
    pub value: Option<FFHelpTextVal>,
}

impl FFHelpText {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "type" => instance.info_text_type = Some(value.parse()?),
                "val" => instance.value = Some(value.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FFStatusText {
    pub info_text_type: Option<InfoTextType>,
    pub value: Option<FFStatusTextVal>,
}

impl FFStatusText {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "type" => instance.info_text_type = Some(value.parse()?),
                "val" => instance.value = Some(value.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FFCheckBoxSizeChoice {
    Explicit(HpsMeasure),
    Auto(Option<OnOff>),
}

impl FFCheckBoxSizeChoice {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "size" | "sizeAuto" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "size" => Ok(FFCheckBoxSizeChoice::Explicit(HpsMeasure::from_xml_element(xml_node)?)),
            "sizeAuto" => Ok(FFCheckBoxSizeChoice::Auto(parse_on_off_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "FFCheckBoxSizeChoice",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct FFCheckBox {
    pub size: FFCheckBoxSizeChoice,
    pub is_default: Option<OnOff>,
    pub is_checked: Option<OnOff>,
}

impl FFCheckBox {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut size = None;
        let mut is_default = None;
        let mut is_checked = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                node_name @ _ if FFCheckBoxSizeChoice::is_choice_member(node_name) => {
                    size = Some(FFCheckBoxSizeChoice::from_xml_element(child_node)?)
                }
                "default" => is_default = parse_on_off_xml_element(child_node)?,
                "checked" => is_checked = parse_on_off_xml_element(child_node)?,
                _ => (),
            }
        }

        let size = size.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "size|sizeAuto"))?;

        Ok(Self {
            size,
            is_default,
            is_checked,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FFDDList {
    pub result: Option<DecimalNumber>,
    pub default: Option<DecimalNumber>,
    pub list_entries: Vec<String>,
}

impl FFDDList {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "result" => instance.result = Some(child_node.get_val_attribute()?.parse()?),
                "default" => instance.default = Some(child_node.get_val_attribute()?.parse()?),
                "listEntry" => instance.list_entries.push(child_node.get_val_attribute()?.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum FFTextType {
    #[strum(serialize = "regular")]
    Regular,
    #[strum(serialize = "number")]
    Number,
    #[strum(serialize = "date")]
    Date,
    #[strum(serialize = "currentTime")]
    CurrentTime,
    #[strum(serialize = "currentDate")]
    CurrentDate,
    #[strum(serialize = "calculated")]
    Calculated,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FFTextInput {
    pub text_type: Option<FFTextType>,
    pub default: Option<String>,
    pub max_length: Option<DecimalNumber>,
    pub format: Option<String>,
}

impl FFTextInput {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "type" => instance.text_type = Some(child_node.get_val_attribute()?.parse()?),
                "default" => instance.default = Some(child_node.get_val_attribute()?.clone()),
                "maxLength" => instance.max_length = Some(child_node.get_val_attribute()?.parse()?),
                "format" => instance.format = Some(child_node.get_val_attribute()?.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum FFData {
    Name(FFName),
    Label(DecimalNumber),
    TabIndex(UnsignedDecimalNumber),
    Enabled(Option<OnOff>),
    RecalculateOnExit(Option<OnOff>),
    EntryMacro(MacroName),
    ExitMacro(MacroName),
    HelpText(FFHelpText),
    StatusText(FFStatusText),
    CheckBox(FFCheckBox),
    DropDownList(FFDDList),
    TextInput(FFTextInput),
}

impl XsdChoice for FFData {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "name" | "label" | "tabIndex" | "enabled" | "calcOnExit" | "entryMacro" | "exitMacro" | "helpText"
            | "statusText" | "checkBox" | "ddList" | "textInput" => true,
            _ => false,
        }
    }

    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "name" => Ok(FFData::Name(xml_node.get_val_attribute()?.clone())),
            "label" => Ok(FFData::Label(xml_node.get_val_attribute()?.parse()?)),
            "tabIndex" => Ok(FFData::TabIndex(xml_node.get_val_attribute()?.parse()?)),
            "enabled" => Ok(FFData::Enabled(parse_on_off_xml_element(xml_node)?)),
            "calcOnExit" => Ok(FFData::RecalculateOnExit(parse_on_off_xml_element(xml_node)?)),
            "entryMacro" => Ok(FFData::EntryMacro(xml_node.get_val_attribute()?.clone())),
            "exitMacro" => Ok(FFData::ExitMacro(xml_node.get_val_attribute()?.clone())),
            "helpText" => Ok(FFData::HelpText(FFHelpText::from_xml_element(xml_node)?)),
            "statusText" => Ok(FFData::StatusText(FFStatusText::from_xml_element(xml_node)?)),
            "checkBox" => Ok(FFData::CheckBox(FFCheckBox::from_xml_element(xml_node)?)),
            "ddList" => Ok(FFData::DropDownList(FFDDList::from_xml_element(xml_node)?)),
            "textInput" => Ok(FFData::TextInput(FFTextInput::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "FFData"))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum FldCharType {
    #[strum(serialize = "begin")]
    Begin,
    #[strum(serialize = "separate")]
    Separate,
    #[strum(serialize = "end")]
    End,
}

#[derive(Debug, Clone, PartialEq)]
pub struct FldChar {
    pub form_field_properties: Option<FFData>,
    pub field_char_type: FldCharType,
    pub field_lock: Option<OnOff>,
    pub dirty: Option<OnOff>,
}

impl FldChar {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut field_char_type = None;
        let mut field_lock = None;
        let mut dirty = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "fldCharType" => field_char_type = Some(value.parse()?),
                "fldLock" => field_lock = Some(parse_xml_bool(value)?),
                "dirty" => dirty = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        let form_field_properties = xml_node
            .child_nodes
            .iter()
            .find_map(|child_node| FFData::try_from_xml_element(child_node))
            .transpose()?;

        let field_char_type =
            field_char_type.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "fldCharType"))?;

        Ok(Self {
            form_field_properties,
            field_char_type,
            field_lock,
            dirty,
        })
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum RubyAlign {
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "distributeLetter")]
    DistributeLetter,
    #[strum(serialize = "distributeSpace")]
    DistributeSpace,
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "rightVertical")]
    RightVertical,
}

#[derive(Debug, Clone, PartialEq)]
pub struct RubyPr {
    pub ruby_align: RubyAlign,
    pub hps: HpsMeasure,
    pub hps_raise: HpsMeasure,
    pub hps_base_text: HpsMeasure,
    pub language_id: Lang,
    pub dirty: Option<OnOff>,
}

impl RubyPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut ruby_align = None;
        let mut hps = None;
        let mut hps_raise = None;
        let mut hps_base_text = None;
        let mut language_id = None;
        let mut dirty = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rubyAlign" => ruby_align = Some(child_node.get_val_attribute()?.parse()?),
                "hps" => hps = Some(child_node.get_val_attribute()?.parse()?),
                "hpsRaise" => hps_raise = Some(child_node.get_val_attribute()?.parse()?),
                "hpsBaseText" => hps_base_text = Some(child_node.get_val_attribute()?.parse()?),
                "lid" => language_id = Some(child_node.get_val_attribute()?.clone()),
                "dirty" => dirty = parse_on_off_xml_element(child_node)?,
                _ => (),
            }
        }

        let ruby_align = ruby_align.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rubyAlign"))?;
        let hps = hps.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "hps"))?;
        let hps_raise = hps_raise.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "hpsRaise"))?;
        let hps_base_text =
            hps_base_text.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "hpsBaseText"))?;
        let language_id = language_id.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "lid"))?;

        Ok(Self {
            ruby_align,
            hps,
            hps_raise,
            hps_base_text,
            language_id,
            dirty,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RubyContentChoice {
    Run(R),
    RunLevelElement(RunLevelElts),
}

impl XsdChoice for RubyContentChoice {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "r" => true,
            _ => RunLevelElts::is_choice_member(&node_name),
        }
    }

    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "r" => Ok(RubyContentChoice::Run(R::from_xml_element(xml_node)?)),
            node_name @ _ if RunLevelElts::is_choice_member(node_name) => Ok(RubyContentChoice::RunLevelElement(
                RunLevelElts::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RubyContentChoice",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct RubyContent {
    pub ruby_contents: Vec<RubyContentChoice>,
}

impl RubyContent {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let ruby_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| RubyContentChoice::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { ruby_contents })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Ruby {
    pub ruby_properties: RubyPr,
    pub ruby_content: RubyContent,
    pub ruby_base: RubyContent,
}

impl Ruby {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut ruby_properties = None;
        let mut ruby_content = None;
        let mut ruby_base = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rubyPr" => ruby_properties = Some(RubyPr::from_xml_element(child_node)?),
                "rt" => ruby_content = Some(RubyContent::from_xml_element(child_node)?),
                "rubyBase" => ruby_base = Some(RubyContent::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let ruby_properties =
            ruby_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rubyPr"))?;
        let ruby_content = ruby_content.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rt"))?;
        let ruby_base = ruby_base.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rubyBase"))?;

        Ok(Self {
            ruby_properties,
            ruby_content,
            ruby_base,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct FtnEdnRef {
    pub custom_mark_follows: Option<OnOff>,
    pub id: DecimalNumber,
}

impl FtnEdnRef {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut custom_mark_follows = None;
        let mut id = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "customMarkFollows" => custom_mark_follows = Some(parse_xml_bool(value)?),
                "id" => id = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            custom_mark_follows,
            id: id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "id"))?,
        })
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum PTabAlignment {
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "right")]
    Right,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum PTabRelativeTo {
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "indent")]
    Indent,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum PTabLeader {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "dot")]
    Dot,
    #[strum(serialize = "hyphen")]
    Hyphen,
    #[strum(serialize = "underscore")]
    Underscore,
    #[strum(serialize = "middleDot")]
    MiddleDot,
}

#[derive(Debug, Clone, PartialEq)]
pub struct PTab {
    pub alignment: PTabAlignment,
    pub relative_to: PTabRelativeTo,
    pub leader: PTabLeader,
}

impl PTab {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut alignment = None;
        let mut relative_to = None;
        let mut leader = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "alignment" => alignment = Some(value.parse()?),
                "relativeTo" => relative_to = Some(value.parse()?),
                "leader" => leader = Some(value.parse()?),
                _ => (),
            }
        }

        let alignment = alignment.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "alignment"))?;
        let relative_to = relative_to.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "relativeTo"))?;
        let leader = leader.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "leader"))?;

        Ok(Self {
            alignment,
            relative_to,
            leader,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RunInnerContent {
    Break(Br),
    Text(Text),
    ContentPart(Rel),
    DeletedText(Text),
    InstructionText(Text),
    DeletedInstructionText(Text),
    NonBreakingHyphen,
    OptionalHypen,
    ShortDayFormat,
    ShortMonthFormat,
    ShortYearFormat,
    LongDayFormat,
    LongMonthFormat,
    LongYearFormat,
    AnnorationReferenceMark,
    FootnoteReferenceMark,
    EndnoteReferenceMark,
    Separator,
    ContinuationSeparator,
    Symbol(Sym),
    PageNum,
    CarriageReturn,
    Tab,
    Object(Object),
    FieldCharacter(FldChar),
    Ruby(Ruby),
    FootnoteReference(FtnEdnRef),
    EndnoteReference(FtnEdnRef),
    CommentReference(Markup),
    Drawing(Drawing),
    PositionTab(PTab),
    LastRenderedPageBreak,
}

impl RunInnerContent {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "br"
            | "t"
            | "contentPart"
            | "delText"
            | "instrText"
            | "delInstrText"
            | "noBreakHyphen"
            | "softHyphen"
            | "dayShort"
            | "monthShort"
            | "yearShort"
            | "dayLong"
            | "monthLong"
            | "yearLong"
            | "annotationRef"
            | "footnoteRef"
            | "endnoteRef"
            | "separator"
            | "continuationSeparator"
            | "sym"
            | "pgNum"
            | "cr"
            | "tab"
            | "object"
            | "fldChar"
            | "ruby"
            | "footnoteReference"
            | "endnoteReference"
            | "commentReference"
            | "drawing"
            | "ptab"
            | "lastRenderedPageBreak" => true,
            _ => false,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "br" => Ok(RunInnerContent::Break(Br::from_xml_element(xml_node)?)),
            "t" => Ok(RunInnerContent::Text(Text::from_xml_element(xml_node)?)),
            "contentPart" => Ok(RunInnerContent::ContentPart(Rel::from_xml_element(xml_node)?)),
            "delText" => Ok(RunInnerContent::DeletedText(Text::from_xml_element(xml_node)?)),
            "instrText" => Ok(RunInnerContent::InstructionText(Text::from_xml_element(xml_node)?)),
            "delInstrText" => Ok(RunInnerContent::DeletedInstructionText(Text::from_xml_element(
                xml_node,
            )?)),
            "noBreakHyphen" => Ok(RunInnerContent::NonBreakingHyphen),
            "softHyphen" => Ok(RunInnerContent::OptionalHypen),
            "dayShort" => Ok(RunInnerContent::ShortDayFormat),
            "monthShort" => Ok(RunInnerContent::ShortMonthFormat),
            "yearShort" => Ok(RunInnerContent::ShortYearFormat),
            "dayLong" => Ok(RunInnerContent::LongDayFormat),
            "monthLong" => Ok(RunInnerContent::LongMonthFormat),
            "yearLong" => Ok(RunInnerContent::LongYearFormat),
            "annotationRef" => Ok(RunInnerContent::AnnorationReferenceMark),
            "footnoteRef" => Ok(RunInnerContent::FootnoteReferenceMark),
            "endnoteRef" => Ok(RunInnerContent::EndnoteReferenceMark),
            "separator" => Ok(RunInnerContent::Separator),
            "continuationSeparator" => Ok(RunInnerContent::ContinuationSeparator),
            "sym" => Ok(RunInnerContent::Symbol(Sym::from_xml_element(xml_node)?)),
            "pgNum" => Ok(RunInnerContent::PageNum),
            "cr" => Ok(RunInnerContent::CarriageReturn),
            "tab" => Ok(RunInnerContent::Tab),
            "object" => Ok(RunInnerContent::Object(Object::from_xml_element(xml_node)?)),
            "fldChar" => Ok(RunInnerContent::FieldCharacter(FldChar::from_xml_element(xml_node)?)),
            "ruby" => Ok(RunInnerContent::Ruby(Ruby::from_xml_element(xml_node)?)),
            "footnoteReference" => Ok(RunInnerContent::FootnoteReference(FtnEdnRef::from_xml_element(
                xml_node,
            )?)),
            "endnoteReference" => Ok(RunInnerContent::EndnoteReference(FtnEdnRef::from_xml_element(
                xml_node,
            )?)),
            "commentReference" => Ok(RunInnerContent::CommentReference(Markup::from_xml_element(xml_node)?)),
            "drawing" => Ok(RunInnerContent::Drawing(Drawing::from_xml_element(xml_node)?)),
            "ptab" => Ok(RunInnerContent::PositionTab(PTab::from_xml_element(xml_node)?)),
            "lastRenderedPageBreak" => Ok(RunInnerContent::LastRenderedPageBreak),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RunInnerContent",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct R {
    pub run_properties: Option<RPr>,
    pub run_inner_contents: Vec<RunInnerContent>,
    pub run_properties_revision_id: Option<LongHexNumber>,
    pub deletion_revision_id: Option<LongHexNumber>,
    pub run_revision_id: Option<LongHexNumber>,
}

impl R {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "rsidRPr" => instance.run_properties_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "rsidDel" => instance.deletion_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "rsidR" => instance.run_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => instance.run_properties = Some(RPr::from_xml_element(child_node)?),
                node_name @ _ if RunInnerContent::is_choice_member(node_name) => instance
                    .run_inner_contents
                    .push(RunInnerContent::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum ContentRunContent {
    CustomXml(CustomXmlRun),
    SmartTag(SmartTagRun),
    Sdt(SdtRun),
    Bidirectional(DirContentRun),
    BidirectionalOverride(BdoContentRun),
    Run(R),
    RunLevelElements(RunLevelElts),
}

impl ContentRunContent {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "customXml" | "smartTag" | "sdt" | "dir" | "bdo" | "r" => true,
            _ => RunLevelElts::is_choice_member(&node_name),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "customXml" => Ok(ContentRunContent::CustomXml(CustomXmlRun::from_xml_element(xml_node)?)),
            "smartTag" => Ok(ContentRunContent::SmartTag(SmartTagRun::from_xml_element(xml_node)?)),
            "sdt" => Ok(ContentRunContent::Sdt(SdtRun::from_xml_element(xml_node)?)),
            "dir" => Ok(ContentRunContent::Bidirectional(DirContentRun::from_xml_element(
                xml_node,
            )?)),
            "bdo" => Ok(ContentRunContent::BidirectionalOverride(
                BdoContentRun::from_xml_element(xml_node)?,
            )),
            "r" => Ok(ContentRunContent::Run(R::from_xml_element(xml_node)?)),
            node_name @ _ if RunLevelElts::is_choice_member(node_name) => Ok(ContentRunContent::RunLevelElements(
                RunLevelElts::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "ContentRunContent",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum RunTrackChangeChoice {
    ContentRunContent(ContentRunContent),
    // TODO
    // OMathMathElements(OMathMathElements),
}

impl XsdChoice for RunTrackChangeChoice {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        ContentRunContent::is_choice_member(node_name)
    }

    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let local_name = xml_node.local_name();
        if ContentRunContent::is_choice_member(local_name) {
            Ok(RunTrackChangeChoice::ContentRunContent(
                ContentRunContent::from_xml_element(xml_node)?,
            ))
        } else {
            Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RunTrackChangeChoice",
            )))
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct RunTrackChange {
    pub base: TrackChange,
    pub choices: Vec<RunTrackChangeChoice>,
}

impl RunTrackChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = TrackChange::from_xml_element(xml_node)?;
        let choices = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| RunTrackChangeChoice::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { base, choices })
    }
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

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "bookmarkStart" => Ok(RangeMarkupElements::BookmarkStart(Bookmark::from_xml_element(
                xml_node,
            )?)),
            "bookmarkEnd" => Ok(RangeMarkupElements::BookmarkEnd(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "moveFromRangeStart" => Ok(RangeMarkupElements::MoveFromRangeStart(MoveBookmark::from_xml_element(
                xml_node,
            )?)),
            "moveFromRangeEnd" => Ok(RangeMarkupElements::MoveFromRangeEnd(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "moveToRangeStart" => Ok(RangeMarkupElements::MoveToRangeStart(MoveBookmark::from_xml_element(
                xml_node,
            )?)),
            "moveToRangeEnd" => Ok(RangeMarkupElements::MoveToRangeEnd(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "commentRangeStart" => Ok(RangeMarkupElements::CommentRangeStart(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "commentRangeEnd" => Ok(RangeMarkupElements::CommentRangeEnd(MarkupRange::from_xml_element(
                xml_node,
            )?)),
            "customXmlInsRangeStart" => Ok(RangeMarkupElements::CustomXmlInsertRangeStart(
                TrackChange::from_xml_element(xml_node)?,
            )),
            "customXmlInsRangeEnd" => Ok(RangeMarkupElements::CustomXmlInsertRangeEnd(Markup::from_xml_element(
                xml_node,
            )?)),
            "customXmlDelRangeStart" => Ok(RangeMarkupElements::CustomXmlDeleteRangeStart(
                TrackChange::from_xml_element(xml_node)?,
            )),
            "customXmlDelRangeEnd" => Ok(RangeMarkupElements::CustomXmlDeleteRangeEnd(Markup::from_xml_element(
                xml_node,
            )?)),
            "customXmlMoveFromRangeStart" => Ok(RangeMarkupElements::CustomXmlMoveFromRangeStart(
                TrackChange::from_xml_element(xml_node)?,
            )),
            "customXmlMoveFromRangeEnd" => Ok(RangeMarkupElements::CustomXmlMoveFromRangeEnd(
                Markup::from_xml_element(xml_node)?,
            )),
            "customXmlMoveToRangeStart" => Ok(RangeMarkupElements::CustomXmlMoveToRangeStart(
                TrackChange::from_xml_element(xml_node)?,
            )),
            "customXmlMoveToRangeEnd" => Ok(RangeMarkupElements::CustomXmlMoveToRangeEnd(Markup::from_xml_element(
                xml_node,
            )?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RangeMarkupElements",
            ))),
        }
    }
}

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

#[derive(Debug, Clone, PartialEq)]
pub enum RunLevelElts {
    ProofError(ProofErr),
    PermissionStart(PermStart),
    PermissionEnd(Perm),
    RangeMarkupElements(RangeMarkupElements),
    Insert(RunTrackChange),
    Delete(RunTrackChange),
    MoveFrom(RunTrackChange),
    MoveTo(RunTrackChange),
    MathContent(MathContent),
}

impl RunLevelElts {
    pub fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "proofErr" | "permStart" | "permEnd" | "ins" | "del" | "moveFrom" | "moveTo" => true,
            _ => RangeMarkupElements::is_choice_member(&node_name) || MathContent::is_choice_member(&node_name),
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let local_name = xml_node.local_name();
        match local_name {
            "proofErr" => Ok(RunLevelElts::ProofError(ProofErr::from_xml_element(xml_node)?)),
            "permStart" => Ok(RunLevelElts::PermissionStart(PermStart::from_xml_element(xml_node)?)),
            "permEnd" => Ok(RunLevelElts::PermissionEnd(Perm::from_xml_element(xml_node)?)),
            "ins" => Ok(RunLevelElts::Insert(RunTrackChange::from_xml_element(xml_node)?)),
            "del" => Ok(RunLevelElts::Delete(RunTrackChange::from_xml_element(xml_node)?)),
            "moveFrom" => Ok(RunLevelElts::MoveFrom(RunTrackChange::from_xml_element(xml_node)?)),
            "moveTo" => Ok(RunLevelElts::MoveTo(RunTrackChange::from_xml_element(xml_node)?)),
            _ if RangeMarkupElements::is_choice_member(local_name) => Ok(RunLevelElts::RangeMarkupElements(
                RangeMarkupElements::from_xml_element(xml_node)?,
            )),
            // TODO MathContent
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "RunLevelElts",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct CustomXmlBlock {
    pub custom_xml_properties: Option<CustomXmlPr>,
    pub block_contents: Vec<ContentBlockContent>,
    pub uri: Option<String>,
    pub element: XmlName,
}

impl CustomXmlBlock {
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
        let mut block_contents = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "customXmlPr" => custom_xml_properties = Some(CustomXmlPr::from_xml_element(child_node)?),
                node_name @ _ if ContentBlockContent::is_choice_member(node_name) => {
                    block_contents.push(ContentBlockContent::from_xml_element(child_node)?);
                }
                _ => (),
            }
        }

        let element = element.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "element"))?;

        Ok(Self {
            custom_xml_properties,
            block_contents,
            uri,
            element,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtContentBlock {
    pub block_contents: Vec<ContentBlockContent>,
}

impl SdtContentBlock {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let block_contents = xml_node
            .child_nodes
            .iter()
            .filter_map(|child_node| ContentBlockContent::try_from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { block_contents })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SdtBlock {
    pub sdt_properties: Option<SdtPr>,
    pub sdt_end_properties: Option<SdtEndPr>,
    pub sdt_content: Option<SdtContentBlock>,
}

impl SdtBlock {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "sdtPr" => instance.sdt_properties = Some(SdtPr::from_xml_element(child_node)?),
                "sdtEndPr" => instance.sdt_end_properties = Some(SdtEndPr::from_xml_element(child_node)?),
                "sdtContent" => instance.sdt_content = Some(SdtContentBlock::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum DropCap {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "drop")]
    Drop,
    #[strum(serialize = "margin")]
    Margin,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum HeightRule {
    #[strum(serialize = "auto")]
    Auto,
    #[strum(serialize = "exact")]
    Exact,
    #[strum(serialize = "atLeast")]
    AtLeast,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum Wrap {
    #[strum(serialize = "auto")]
    Auto,
    #[strum(serialize = "notBeside")]
    NotBeside,
    #[strum(serialize = "around")]
    Around,
    #[strum(serialize = "tight")]
    Tight,
    #[strum(serialize = "through")]
    Throught,
    #[strum(serialize = "none")]
    None,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum VAnchor {
    #[strum(serialize = "text")]
    Text,
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "page")]
    Page,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum HAnchor {
    #[strum(serialize = "text")]
    Text,
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "page")]
    Page,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FramePr {
    pub drop_cap: Option<DropCap>,
    pub lines: Option<DecimalNumber>,
    pub width: Option<TwipsMeasure>,
    pub height: Option<TwipsMeasure>,
    pub vertical_space: Option<TwipsMeasure>,
    pub horizontal_space: Option<TwipsMeasure>,
    pub wrap: Option<Wrap>,
    pub horizontal_anchor: Option<HAnchor>,
    pub vertical_anchor: Option<VAnchor>,
    pub x: Option<SignedTwipsMeasure>,
    pub x_align: Option<XAlign>,
    pub y: Option<SignedTwipsMeasure>,
    pub y_align: Option<YAlign>,
    pub height_rule: Option<HeightRule>,
    pub anchor_lock: Option<OnOff>,
}

impl FramePr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "dropCap" => instance.drop_cap = Some(value.parse()?),
                "lines" => instance.lines = Some(value.parse()?),
                "w" => instance.width = Some(value.parse()?),
                "h" => instance.height = Some(value.parse()?),
                "vSpace" => instance.vertical_space = Some(value.parse()?),
                "hSpace" => instance.horizontal_space = Some(value.parse()?),
                "wrap" => instance.wrap = Some(value.parse()?),
                "hAnchor" => instance.horizontal_anchor = Some(value.parse()?),
                "vAnchor" => instance.vertical_anchor = Some(value.parse()?),
                "x" => instance.x = Some(value.parse()?),
                "xAlign" => instance.x_align = Some(value.parse()?),
                "y" => instance.y = Some(value.parse()?),
                "yAlign" => instance.y_align = Some(value.parse()?),
                "hRule" => instance.height_rule = Some(value.parse()?),
                "anchorLock" => instance.anchor_lock = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct NumPr {
    pub indent_level: Option<DecimalNumber>,
    pub numbering_id: Option<DecimalNumber>,
    pub inserted: Option<TrackChange>,
}

impl NumPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();
        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "ilvl" => instance.indent_level = Some(child_node.get_val_attribute()?.parse()?),
                "numId" => instance.numbering_id = Some(child_node.get_val_attribute()?.parse()?),
                "ins" => instance.inserted = Some(TrackChange::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PBdr {
    pub top: Option<Border>,
    pub left: Option<Border>,
    pub bottom: Option<Border>,
    pub right: Option<Border>,
    pub between: Option<Border>,
    pub bar: Option<Border>,
}

impl PBdr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(Border::from_xml_element(child_node)?),
                "left" => instance.left = Some(Border::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(Border::from_xml_element(child_node)?),
                "right" => instance.right = Some(Border::from_xml_element(child_node)?),
                "between" => instance.between = Some(Border::from_xml_element(child_node)?),
                "bar" => instance.bar = Some(Border::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TabJc {
    #[strum(serialize = "clear")]
    Clear,
    #[strum(serialize = "start")]
    Start,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "end")]
    End,
    #[strum(serialize = "decimal")]
    Decimal,
    #[strum(serialize = "bar")]
    Bar,
    #[strum(serialize = "num")]
    Number,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TabTlc {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "dot")]
    Dot,
    #[strum(serialize = "hyphen")]
    Hyphen,
    #[strum(serialize = "underscore")]
    Underscore,
    #[strum(serialize = "heavy")]
    Heavy,
    #[strum(serialize = "middleDot")]
    MiddleDot,
}

#[derive(Debug, Clone, PartialEq)]
pub struct TabStop {
    pub value: TabJc,
    pub leader: Option<TabTlc>,
    pub position: SignedTwipsMeasure,
}

impl TabStop {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut value = None;
        let mut leader = None;
        let mut position = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => value = Some(attr_value.parse()?),
                "leader" => leader = Some(attr_value.parse()?),
                "pos" => position = Some(attr_value.parse()?),
                _ => (),
            }
        }

        let value = value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?;
        let position = position.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "pos"))?;

        Ok(Self {
            value,
            leader,
            position,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct Tabs {
    pub tabs: Vec<TabStop>,
}

impl Tabs {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let tabs = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "tab")
            .map(|child_node| TabStop::from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        if tabs.is_empty() {
            Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "tab",
                1,
                MaxOccurs::Unbounded,
                0,
            )))
        } else {
            Ok(Self { tabs })
        }
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum LineSpacingRule {
    #[strum(serialize = "auto")]
    Auto,
    #[strum(serialize = "exact")]
    Exact,
    #[strum(serialize = "atLeast")]
    AtLeast,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Spacing {
    pub before: Option<TwipsMeasure>,
    pub before_lines: Option<DecimalNumber>,
    pub before_autospacing: Option<OnOff>,
    pub after: Option<TwipsMeasure>,
    pub after_lines: Option<DecimalNumber>,
    pub after_autospacing: Option<OnOff>,
    pub line: Option<SignedTwipsMeasure>,
    pub line_rule: Option<LineSpacingRule>,
}

impl Spacing {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "before" => instance.before = Some(value.parse()?),
                "beforeLines" => instance.before_lines = Some(value.parse()?),
                "beforeAutospacing" => instance.before_autospacing = Some(parse_xml_bool(value)?),
                "after" => instance.after = Some(value.parse()?),
                "afterLines" => instance.after_lines = Some(value.parse()?),
                "afterAutospacing" => instance.after_autospacing = Some(parse_xml_bool(value)?),
                "line" => instance.line = Some(value.parse()?),
                "lineRule" => instance.line_rule = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Ind {
    pub start: Option<SignedTwipsMeasure>,
    pub start_chars: Option<DecimalNumber>,
    pub end: Option<SignedTwipsMeasure>,
    pub end_chars: Option<DecimalNumber>,
    pub hanging: Option<TwipsMeasure>,
    pub hanging_chars: Option<DecimalNumber>,
    pub first_line: Option<TwipsMeasure>,
    pub first_line_chars: Option<DecimalNumber>,
}

impl Ind {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "start" => instance.start = Some(value.parse()?),
                "startChars" => instance.start_chars = Some(value.parse()?),
                "end" => instance.end = Some(value.parse()?),
                "endChars" => instance.end_chars = Some(value.parse()?),
                "hanging" => instance.hanging = Some(value.parse()?),
                "hangingChars" => instance.hanging_chars = Some(value.parse()?),
                "firstLine" => instance.first_line = Some(value.parse()?),
                "firstLineChars" => instance.first_line_chars = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum Jc {
    #[strum(serialize = "start")]
    Start,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "end")]
    End,
    #[strum(serialize = "both")]
    Both,
    #[strum(serialize = "mediumKashida")]
    MediumKashida,
    #[strum(serialize = "distribute")]
    Distribute,
    #[strum(serialize = "numTab")]
    NumTab,
    #[strum(serialize = "highKashida")]
    HighKashida,
    #[strum(serialize = "lowKashida")]
    LowKashida,
    #[strum(serialize = "thaiDistribute")]
    ThaiDistribute,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TextDirection {
    #[strum(serialize = "tb")]
    TopToBottom,
    #[strum(serialize = "rl")]
    RightToLeft,
    #[strum(serialize = "lr")]
    LeftToRight,
    #[strum(serialize = "tbV")]
    TopToBottomRotated,
    #[strum(serialize = "rlV")]
    RightToLeftRotated,
    #[strum(serialize = "lrV")]
    LeftToRightRotated,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TextAlignment {
    #[strum(serialize = "top")]
    Top,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "baseline")]
    Baseline,
    #[strum(serialize = "bottom")]
    Bottom,
    #[strum(serialize = "auto")]
    Auto,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TextboxTightWrap {
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "allLines")]
    AllLines,
    #[strum(serialize = "firstAndLastLine")]
    FirstAndLastLine,
    #[strum(serialize = "firstLineOnly")]
    FirstLineOnly,
    #[strum(serialize = "lastLineOnly")]
    LastLineOnly,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Cnf {
    pub first_row: Option<OnOff>,
    pub last_row: Option<OnOff>,
    pub first_column: Option<OnOff>,
    pub last_column: Option<OnOff>,
    pub odd_vertical_band: Option<OnOff>,
    pub even_vertical_band: Option<OnOff>,
    pub odd_horizontal_band: Option<OnOff>,
    pub even_horizontal_band: Option<OnOff>,
    pub first_row_first_column: Option<OnOff>,
    pub first_row_last_column: Option<OnOff>,
    pub last_row_first_column: Option<OnOff>,
    pub last_row_last_column: Option<OnOff>,
}

impl Cnf {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "firstRow" => instance.first_row = Some(parse_xml_bool(value)?),
                "lastRow" => instance.last_row = Some(parse_xml_bool(value)?),
                "firstColumn" => instance.first_column = Some(parse_xml_bool(value)?),
                "lastColumn" => instance.last_column = Some(parse_xml_bool(value)?),
                "oddVBand" => instance.odd_vertical_band = Some(parse_xml_bool(value)?),
                "evenVBand" => instance.even_vertical_band = Some(parse_xml_bool(value)?),
                "oddHBand" => instance.odd_horizontal_band = Some(parse_xml_bool(value)?),
                "evenHBand" => instance.even_horizontal_band = Some(parse_xml_bool(value)?),
                "firstRowFirstColumn" => instance.first_row_first_column = Some(parse_xml_bool(value)?),
                "firstRowLastColumn" => instance.first_row_last_column = Some(parse_xml_bool(value)?),
                "lastRowFirstColumn" => instance.last_row_first_column = Some(parse_xml_bool(value)?),
                "lastRowLastColumn" => instance.last_row_last_column = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PPrBase {
    pub style: Option<String>,
    pub keep_with_next: Option<OnOff>,
    pub keep_lines_on_one_page: Option<OnOff>,
    pub start_on_next_page: Option<OnOff>,
    pub frame_properties: Option<FramePr>,
    pub widow_control: Option<OnOff>,
    pub numbering_properties: Option<NumPr>,
    pub suppress_line_numbers: Option<OnOff>,
    pub borders: Option<PBdr>,
    pub shading: Option<Shd>,
    pub tabs: Option<Tabs>,
    pub suppress_auto_hyphens: Option<OnOff>,
    pub kinsoku: Option<OnOff>,
    pub word_wrapping: Option<OnOff>,
    pub overflow_punctuations: Option<OnOff>,
    pub top_line_punctuations: Option<OnOff>,
    pub auto_space_latin_and_east_asian: Option<OnOff>,
    pub auto_space_east_asian_and_numbers: Option<OnOff>,
    pub bidirectional: Option<OnOff>,
    pub adjust_right_indent: Option<OnOff>,
    pub snap_to_grid: Option<OnOff>,
    pub spacing: Option<Spacing>,
    pub indent: Option<Ind>,
    pub contextual_spacing: Option<OnOff>,
    pub mirror_indents: Option<OnOff>,
    pub suppress_overlapping: Option<OnOff>,
    pub alignment: Option<Jc>,
    pub text_direction: Option<TextDirection>,
    pub text_alignment: Option<TextAlignment>,
    pub textbox_tight_wrap: Option<TextboxTightWrap>,
    pub outline_level: Option<DecimalNumber>,
    pub div_id: Option<DecimalNumber>,
    pub conditional_formatting: Option<Cnf>,
}

impl PPrBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            instance.try_update_from_xml_element(child_node)?;
        }

        Ok(instance)
    }

    pub fn try_update_from_xml_element(&mut self, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "pStyle" => {
                self.style = Some(xml_node.get_val_attribute()?.clone());
                Ok(true)
            }
            "keepNext" => {
                self.keep_with_next = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "keepLines" => {
                self.keep_lines_on_one_page = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "pageBreakBefore" => {
                self.start_on_next_page = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "framePr" => {
                self.frame_properties = Some(FramePr::from_xml_element(xml_node)?);
                Ok(true)
            }
            "widowControl" => {
                self.widow_control = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "numPr" => {
                self.numbering_properties = Some(NumPr::from_xml_element(xml_node)?);
                Ok(true)
            }
            "suppressLineNumbers" => {
                self.suppress_line_numbers = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "pBdr" => {
                self.borders = Some(PBdr::from_xml_element(xml_node)?);
                Ok(true)
            }
            "shd" => {
                self.shading = Some(Shd::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tabs" => {
                self.tabs = Some(Tabs::from_xml_element(xml_node)?);
                Ok(true)
            }
            "suppressAutoHyphens" => {
                self.suppress_auto_hyphens = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "kinsoku" => {
                self.kinsoku = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "wordWrap" => {
                self.word_wrapping = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "overflowPunct" => {
                self.overflow_punctuations = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "topLinePunct" => {
                self.top_line_punctuations = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "autoSpaceDE" => {
                self.auto_space_latin_and_east_asian = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "autoSpaceDN" => {
                self.auto_space_east_asian_and_numbers = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "bidi" => {
                self.bidirectional = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "adjustRightInd" => {
                self.adjust_right_indent = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "snapToGrid" => {
                self.snap_to_grid = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "spacing" => {
                self.spacing = Some(Spacing::from_xml_element(xml_node)?);
                Ok(true)
            }
            "ind" => {
                self.indent = Some(Ind::from_xml_element(xml_node)?);
                Ok(true)
            }
            "contextualSpacing" => {
                self.contextual_spacing = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "mirrorIndents" => {
                self.mirror_indents = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "suppressOverlap" => {
                self.suppress_overlapping = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "jc" => {
                self.alignment = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "textDirection" => {
                self.text_direction = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "textAlignment" => {
                self.text_alignment = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "textboxTightWrap" => {
                self.textbox_tight_wrap = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "outlineLvl" => {
                self.outline_level = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "divId" => {
                self.div_id = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "cnfStyle" => {
                self.conditional_formatting = Some(Cnf::from_xml_element(xml_node)?);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParaRPrTrackChanges {
    pub inserted: Option<TrackChange>,
    pub deleted: Option<TrackChange>,
    pub move_from: Option<TrackChange>,
    pub move_to: Option<TrackChange>,
}

impl ParaRPrTrackChanges {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Option<Self>> {
        let mut instance: Option<Self> = None;

        for child_node in &xml_node.child_nodes {
            Self::try_parse_group_node(&mut instance, child_node)?;
        }

        Ok(instance)
    }

    pub fn try_parse_group_node(instance: &mut Option<Self>, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "ins" => {
                instance.get_or_insert_with(Default::default).inserted = Some(TrackChange::from_xml_element(xml_node)?);
                Ok(true)
            }
            "del" => {
                instance.get_or_insert_with(Default::default).deleted = Some(TrackChange::from_xml_element(xml_node)?);
                Ok(true)
            }
            "moveFrom" => {
                instance.get_or_insert_with(Default::default).move_from =
                    Some(TrackChange::from_xml_element(xml_node)?);
                Ok(true)
            }
            "moveTo" => {
                instance.get_or_insert_with(Default::default).move_to = Some(TrackChange::from_xml_element(xml_node)?);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParaRPrOriginal {
    pub track_changes: Option<ParaRPrTrackChanges>,
    pub bases: Vec<RPrBase>,
}

impl ParaRPrOriginal {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            if ParaRPrTrackChanges::try_parse_group_node(&mut instance.track_changes, child_node)? {
                continue;
            }

            if RPrBase::is_choice_member(child_node.local_name()) {
                instance.bases.push(RPrBase::from_xml_element(child_node)?);
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct ParaRPrChange {
    base: TrackChange,
    run_properties: ParaRPrOriginal,
}

impl ParaRPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = TrackChange::from_xml_element(xml_node)?;
        let run_properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "rPr")
            .map(|child_node| ParaRPrOriginal::from_xml_element(child_node))
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "rPr"))?;

        Ok(Self { base, run_properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct ParaRPr {
    pub track_changes: Option<ParaRPrTrackChanges>,
    pub bases: Vec<RPrBase>,
    pub change: Option<ParaRPrChange>,
}

impl ParaRPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            if ParaRPrTrackChanges::try_parse_group_node(&mut instance.track_changes, child_node)? {
                continue;
            }

            let local_name = child_node.local_name();
            if RPrBase::is_choice_member(local_name) {
                instance.bases.push(RPrBase::from_xml_element(child_node)?);
            } else if local_name == "rPrChange" {
                instance.change = Some(ParaRPrChange::from_xml_element(child_node)?);
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum HdrFtr {
    #[strum(serialize = "even")]
    Even,
    #[strum(serialize = "default")]
    Default,
    #[strum(serialize = "first")]
    First,
}

#[derive(Debug, Clone, PartialEq)]
pub struct HdrFtrRef {
    pub base: Rel,
    pub header_footer_type: HdrFtr,
}

impl HdrFtrRef {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = Rel::from_xml_element(xml_node)?;
        let header_footer_type = xml_node
            .attributes
            .get("type")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "type"))?
            .parse()?;

        Ok(Self {
            base,
            header_footer_type,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum HdrFtrReferences {
    Header(HdrFtrRef),
    Footer(HdrFtrRef),
}

impl XsdChoice for HdrFtrReferences {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "headerReference" | "footerReference" => true,
            _ => false,
        }
    }

    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "headerReference" => Ok(HdrFtrReferences::Header(HdrFtrRef::from_xml_element(xml_node)?)),
            "footerReference" => Ok(HdrFtrReferences::Footer(HdrFtrRef::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "HdrFtrReferences",
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum FtnPos {
    #[strum(serialize = "pageBottom")]
    PageBottom,
    #[strum(serialize = "beneathText")]
    BeneathText,
    #[strum(serialize = "sectEnd")]
    SectionEnd,
    #[strum(serialize = "docEnd")]
    DocumentEnd,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum NumberFormat {
    #[strum(serialize = "decimal")]
    Decimal,
    #[strum(serialize = "upperRoman")]
    UpperRoman,
    #[strum(serialize = "lowerRoman")]
    LowerRoman,
    #[strum(serialize = "upperLetter")]
    UpperLetter,
    #[strum(serialize = "lowerLetter")]
    LowerLetter,
    #[strum(serialize = "ordinal")]
    Ordinal,
    #[strum(serialize = "cardinalText")]
    CardinalText,
    #[strum(serialize = "ordinalText")]
    OrdinalText,
    #[strum(serialize = "hex")]
    Hex,
    #[strum(serialize = "chicago")]
    Chicago,
    #[strum(serialize = "ideographDigital")]
    IdeographDigital,
    #[strum(serialize = "japaneseCounting")]
    JapaneseCounting,
    #[strum(serialize = "aiueo")]
    Aiueo,
    #[strum(serialize = "iroha")]
    Iroha,
    #[strum(serialize = "decimalFullWidth")]
    DecimalFullWidth,
    #[strum(serialize = "decimalHalfWidth")]
    DecimalHalfWidth,
    #[strum(serialize = "japaneseLegal")]
    JapaneseLegal,
    #[strum(serialize = "japaneseDigitalTenThousand")]
    JapaneseDigitalTenThousand,
    #[strum(serialize = "decimalEnclosedCircle")]
    DecimalEnclosedCircle,
    #[strum(serialize = "decimalFullWidth2")]
    DecimalFullWidth2,
    #[strum(serialize = "aiueoFullWidth")]
    AiueoFullWidth,
    #[strum(serialize = "irohaFullWidth")]
    IrohaFullWidth,
    #[strum(serialize = "decimalZero")]
    DecimalZero,
    #[strum(serialize = "bullet")]
    Bullet,
    #[strum(serialize = "ganada")]
    Ganada,
    #[strum(serialize = "chosung")]
    Chosung,
    #[strum(serialize = "decimalEnclosedFullstop")]
    DecimalEnclosedFullstop,
    #[strum(serialize = "decimalEnclosedParen")]
    DecimalEnclosedParen,
    #[strum(serialize = "decimalEnclosedCircleChinese")]
    DecimalEnclosedCircleChinese,
    #[strum(serialize = "ideographEnclosedCircle")]
    IdeographEnclosedCircle,
    #[strum(serialize = "ideographTraditional")]
    IdeographTraditional,
    #[strum(serialize = "ideographZodiac")]
    IdeographZodiac,
    #[strum(serialize = "ideographZodiacTraditional")]
    IdeographZodiacTraditional,
    #[strum(serialize = "taiwaneseCounting")]
    TaiwaneseCounting,
    #[strum(serialize = "ideographLegalTraditional")]
    IdeographLegalTraditional,
    #[strum(serialize = "taiwaneseCountingThousand")]
    TaiwaneseCountingThousand,
    #[strum(serialize = "taiwaneseDigital")]
    TaiwaneseDigital,
    #[strum(serialize = "chineseCounting")]
    ChineseCounting,
    #[strum(serialize = "chineseLegalSimplified")]
    ChineseLegalSimplified,
    #[strum(serialize = "chineseCountingThousand")]
    ChineseCountingThousand,
    #[strum(serialize = "koreanDigital")]
    KoreanDigital,
    #[strum(serialize = "koreanCounting")]
    KoreanCounting,
    #[strum(serialize = "koreanLegal")]
    KoreanLegal,
    #[strum(serialize = "koreanDigital2")]
    KoreanDigital2,
    #[strum(serialize = "vietnameseCounting")]
    VietnameseCounting,
    #[strum(serialize = "russianLower")]
    RussianLower,
    #[strum(serialize = "russianUpper")]
    RussianUpper,
    #[strum(serialize = "none")]
    None,
    #[strum(serialize = "numberInDash")]
    NumberInDash,
    #[strum(serialize = "hebrew1")]
    Hebrew1,
    #[strum(serialize = "hebrew2")]
    Hebrew2,
    #[strum(serialize = "arabicAlpha")]
    ArabicAlpha,
    #[strum(serialize = "arabicAbjad")]
    ArabicAbjad,
    #[strum(serialize = "hindiVowels")]
    HindiVowels,
    #[strum(serialize = "hindiConsonants")]
    HindiConsonants,
    #[strum(serialize = "hindiNumbers")]
    HindiNumbers,
    #[strum(serialize = "hindiCounting")]
    HindiCounting,
    #[strum(serialize = "thaiLetters")]
    ThaiLetters,
    #[strum(serialize = "thaiNumbers")]
    ThaiNumbers,
    #[strum(serialize = "thaiCounting")]
    ThaiCounting,
    #[strum(serialize = "bahtText")]
    BahtText,
    #[strum(serialize = "dollarText")]
    DollarText,
    #[strum(serialize = "custom")]
    Custom,
}

#[derive(Debug, Clone, PartialEq)]
pub struct NumFmt {
    pub value: NumberFormat,
    pub format: Option<String>,
}

impl NumFmt {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut value = None;
        let mut format = None;

        for (attr, attr_value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => value = Some(attr_value.parse()?),
                "format" => format = Some(attr_value.clone()),
                _ => (),
            }
        }

        Ok(Self {
            value: value.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "val"))?,
            format,
        })
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum RestartNumber {
    #[strum(serialize = "continuous")]
    Continuous,
    #[strum(serialize = "eachSect")]
    EachSection,
    #[strum(serialize = "eachPage")]
    EachPage,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FtnEdnNumProps {
    pub numbering_start: Option<DecimalNumber>,
    pub numbering_restart: Option<RestartNumber>,
}

impl FtnEdnNumProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Option<Self>> {
        let mut instance: Option<Self> = None;

        for child_node in &xml_node.child_nodes {
            Self::try_parse_group_node(&mut instance, child_node)?;
        }

        Ok(instance)
    }

    pub fn try_parse_group_node(instance: &mut Option<Self>, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "numStart" => {
                instance.get_or_insert_with(Default::default).numbering_start =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "numRestart" => {
                instance.get_or_insert_with(Default::default).numbering_restart =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct FtnProps {
    pub position: Option<FtnPos>,
    pub numbering_format: Option<NumFmt>,
    pub numbering_properties: Option<FtnEdnNumProps>,
}

impl FtnProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "pos" => instance.position = Some(child_node.get_val_attribute()?.parse()?),
                "numFmt" => instance.numbering_format = Some(NumFmt::from_xml_element(child_node)?),
                _ => {
                    FtnEdnNumProps::try_parse_group_node(&mut instance.numbering_properties, child_node)?;
                }
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum EdnPos {
    #[strum(serialize = "sectEnd")]
    SectionEnd,
    #[strum(serialize = "docEnd")]
    DocumentEnd,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct EdnProps {
    pub position: Option<EdnPos>,
    pub numbering_format: Option<NumFmt>,
    pub numbering_properties: Option<FtnEdnNumProps>,
}

impl EdnProps {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "pos" => instance.position = Some(child_node.get_val_attribute()?.parse()?),
                "numFmt" => instance.numbering_format = Some(NumFmt::from_xml_element(child_node)?),
                _ => {
                    FtnEdnNumProps::try_parse_group_node(&mut instance.numbering_properties, child_node)?;
                }
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum SectionMark {
    #[strum(serialize = "nextPage")]
    NextPage,
    #[strum(serialize = "nextColumn")]
    NextColumn,
    #[strum(serialize = "continuous")]
    Continuous,
    #[strum(serialize = "evenPage")]
    EvenPage,
    #[strum(serialize = "oddPage")]
    OddPage,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum PageOrientation {
    #[strum(serialize = "portrait")]
    Portrait,
    #[strum(serialize = "landscape")]
    Landscape,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PageSz {
    pub width: Option<TwipsMeasure>,
    pub height: Option<TwipsMeasure>,
    pub orientation: Option<PageOrientation>,
    pub code: Option<DecimalNumber>,
}

impl PageSz {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w" => instance.width = Some(value.parse()?),
                "h" => instance.height = Some(value.parse()?),
                "orient" => instance.orientation = Some(value.parse()?),
                "code" => instance.code = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PageMar {
    pub top: SignedTwipsMeasure,
    pub right: TwipsMeasure,
    pub bottom: SignedTwipsMeasure,
    pub left: TwipsMeasure,
    pub header: TwipsMeasure,
    pub footer: TwipsMeasure,
    pub gutter: TwipsMeasure,
}

impl PageMar {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut top = None;
        let mut right = None;
        let mut bottom = None;
        let mut left = None;
        let mut header = None;
        let mut footer = None;
        let mut gutter = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "top" => top = Some(value.parse()?),
                "right" => right = Some(value.parse()?),
                "bottom" => bottom = Some(value.parse()?),
                "left" => left = Some(value.parse()?),
                "header" => header = Some(value.parse()?),
                "footer" => footer = Some(value.parse()?),
                "gutter" => gutter = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            top: top.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "top"))?,
            right: right.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "right"))?,
            bottom: bottom.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "bottom"))?,
            left: left.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "left"))?,
            header: header.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "header"))?,
            footer: footer.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "footer"))?,
            gutter: gutter.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "gutter"))?,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PaperSource {
    pub first: Option<DecimalNumber>,
    pub other: Option<DecimalNumber>,
}

impl PaperSource {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "first" => instance.first = Some(value.parse()?),
                "other" => instance.other = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PageBorder {
    pub base: Border,
    pub rel_id: Option<RelationshipId>,
}

impl PageBorder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = Border::from_xml_element(xml_node)?;
        let rel_id = xml_node.attributes.get("r:id").map(|value| value.parse()).transpose()?;

        Ok(Self { base, rel_id })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TopPageBorder {
    pub base: PageBorder,
    pub top_left: Option<RelationshipId>,
    pub top_right: Option<RelationshipId>,
}

impl TopPageBorder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = PageBorder::from_xml_element(xml_node)?;
        let top_left = xml_node.attributes.get("r:topLeft").cloned();
        let top_right = xml_node.attributes.get("r:topRight").cloned();

        Ok(Self {
            base,
            top_left,
            top_right,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct BottomPageBorder {
    pub base: PageBorder,
    pub bottom_left: Option<RelationshipId>,
    pub bottom_right: Option<RelationshipId>,
}

impl BottomPageBorder {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = PageBorder::from_xml_element(xml_node)?;
        let bottom_left = xml_node.attributes.get("r:bottomLeft").cloned();
        let bottom_right = xml_node.attributes.get("r:bottomRight").cloned();

        Ok(Self {
            base,
            bottom_left,
            bottom_right,
        })
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum PageBorderZOrder {
    #[strum(serialize = "front")]
    Front,
    #[strum(serialize = "back")]
    Back,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum PageBorderDisplay {
    #[strum(serialize = "allPages")]
    AllPages,
    #[strum(serialize = "firstPage")]
    FirstPage,
    #[strum(serialize = "notFirstPage")]
    NotFirstPage,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum PageBorderOffset {
    #[strum(serialize = "page")]
    Page,
    #[strum(serialize = "text")]
    Text,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PageBorders {
    pub top: Option<TopPageBorder>,
    pub left: Option<PageBorder>,
    pub bottom: Option<BottomPageBorder>,
    pub right: Option<PageBorder>,
    pub z_order: Option<PageBorderZOrder>,
    pub display: Option<PageBorderDisplay>,
    pub offset_from: Option<PageBorderOffset>,
}

impl PageBorders {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "zOrder" => instance.z_order = Some(value.parse()?),
                "display" => instance.display = Some(value.parse()?),
                "offsetFrom" => instance.offset_from = Some(value.parse()?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(TopPageBorder::from_xml_element(child_node)?),
                "left" => instance.left = Some(PageBorder::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(BottomPageBorder::from_xml_element(child_node)?),
                "right" => instance.right = Some(PageBorder::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum LineNumberRestart {
    #[strum(serialize = "newPage")]
    NewPage,
    #[strum(serialize = "newSection")]
    NewSection,
    #[strum(serialize = "continuous")]
    Continuous,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct LineNumber {
    pub count_by: Option<DecimalNumber>,
    pub start: Option<DecimalNumber>,
    pub distance: Option<TwipsMeasure>,
    pub restart: Option<LineNumberRestart>,
}

impl LineNumber {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "countBy" => instance.count_by = Some(value.parse()?),
                "start" => instance.start = Some(value.parse()?),
                "distance" => instance.distance = Some(value.parse()?),
                "restart" => instance.restart = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum ChapterSep {
    #[strum(serialize = "hyphen")]
    Hyphen,
    #[strum(serialize = "period")]
    Period,
    #[strum(serialize = "colon")]
    Color,
    #[strum(serialize = "emDash")]
    EmDash,
    #[strum(serialize = "enDash")]
    EnDash,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PageNumber {
    pub format: Option<NumberFormat>,
    pub start: Option<DecimalNumber>,
    pub chapter_style: Option<DecimalNumber>,
    pub chapter_separator: Option<ChapterSep>,
}

impl PageNumber {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "fmt" => instance.format = Some(value.parse()?),
                "start" => instance.start = Some(value.parse()?),
                "chapStyle" => instance.chapter_style = Some(value.parse()?),
                "chapSep" => instance.chapter_separator = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Column {
    pub width: Option<TwipsMeasure>,
    pub spacing: Option<TwipsMeasure>,
}

impl Column {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w" => instance.width = Some(value.parse()?),
                "space" => instance.spacing = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Columns {
    pub columns: Vec<Column>,
    pub equal_width: Option<OnOff>,
    pub spacing: Option<TwipsMeasure>,
    pub number: Option<DecimalNumber>,
    pub separator: Option<OnOff>,
}

impl Columns {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "equalWidth" => instance.equal_width = Some(parse_xml_bool(value)?),
                "space" => instance.spacing = Some(value.parse()?),
                "num" => instance.number = Some(value.parse()?),
                "sep" => instance.separator = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        instance.columns = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "col")
            .map(|child_node| Column::from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        match instance.columns.len() {
            0..=45 => Ok(instance),
            occurs @ _ => Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "col",
                0,
                MaxOccurs::Value(45),
                occurs as u32,
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum VerticalJc {
    #[strum(serialize = "top")]
    Top,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "both")]
    Both,
    #[strum(serialize = "bottom")]
    Bottom,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum DocGridType {
    #[strum(serialize = "default")]
    Default,
    #[strum(serialize = "lines")]
    Lines,
    #[strum(serialize = "linesAndChars")]
    LinesAndChars,
    #[strum(serialize = "snapToChars")]
    SnapToChars,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct DocGrid {
    pub doc_grid_type: Option<DocGridType>,
    pub line_pitch: Option<DecimalNumber>,
    pub char_spacing: Option<DecimalNumber>, // defaults to 0
}

impl DocGrid {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "type" => instance.doc_grid_type = Some(value.parse()?),
                "linePitch" => instance.line_pitch = Some(value.parse()?),
                "charSpace" => instance.char_spacing = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SectPrContents {
    pub footnote_properties: Option<FtnProps>,
    pub endnote_properties: Option<EdnProps>,
    pub section_type: Option<SectionMark>,
    pub page_size: Option<PageSz>,
    pub page_margin: Option<PageMar>,
    pub paper_source: Option<PaperSource>,
    pub page_borders: Option<PageBorders>,
    pub line_number_type: Option<LineNumber>,
    pub page_number_type: Option<PageNumber>,
    pub columns: Option<Columns>,
    pub protect_form_fields: Option<OnOff>,
    pub vertical_align: Option<VerticalJc>,
    pub no_endnote: Option<OnOff>,
    pub title_page: Option<OnOff>,
    pub text_direction: Option<TextDirection>,
    pub bidirectional: Option<OnOff>,
    pub rtl_gutter: Option<OnOff>,
    pub document_grid: Option<DocGrid>,
    pub printer_settings: Option<Rel>,
}

impl SectPrContents {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Option<Self>> {
        let mut instance: Option<Self> = None;

        for child_node in &xml_node.child_nodes {
            Self::try_parse_group_node(&mut instance, child_node)?;
        }

        Ok(instance)
    }

    pub fn try_parse_group_node(instance: &mut Option<Self>, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "footnotePr" => {
                instance.get_or_insert_with(Default::default).footnote_properties =
                    Some(FtnProps::from_xml_element(xml_node)?);
                Ok(true)
            }
            "endnotePr" => {
                instance.get_or_insert_with(Default::default).endnote_properties =
                    Some(EdnProps::from_xml_element(xml_node)?);
                Ok(true)
            }
            "type" => {
                instance.get_or_insert_with(Default::default).section_type =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "pgSz" => {
                instance.get_or_insert_with(Default::default).page_size = Some(PageSz::from_xml_element(xml_node)?);
                Ok(true)
            }
            "pgMar" => {
                instance.get_or_insert_with(Default::default).page_margin = Some(PageMar::from_xml_element(xml_node)?);
                Ok(true)
            }
            "paperSrc" => {
                instance.get_or_insert_with(Default::default).paper_source =
                    Some(PaperSource::from_xml_element(xml_node)?);
                Ok(true)
            }
            "pgBorders" => {
                instance.get_or_insert_with(Default::default).page_borders =
                    Some(PageBorders::from_xml_element(xml_node)?);
                Ok(true)
            }
            "lnNumType" => {
                instance.get_or_insert_with(Default::default).line_number_type =
                    Some(LineNumber::from_xml_element(xml_node)?);
                Ok(true)
            }
            "pgNumType" => {
                instance.get_or_insert_with(Default::default).page_number_type =
                    Some(PageNumber::from_xml_element(xml_node)?);
                Ok(true)
            }
            "cols" => {
                instance.get_or_insert_with(Default::default).columns = Some(Columns::from_xml_element(xml_node)?);
                Ok(true)
            }
            "formProt" => {
                instance.get_or_insert_with(Default::default).protect_form_fields = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "vAlign" => {
                instance.get_or_insert_with(Default::default).vertical_align =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "noEndnote" => {
                instance.get_or_insert_with(Default::default).no_endnote = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "titlePg" => {
                instance.get_or_insert_with(Default::default).title_page = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "textDirection" => {
                instance.get_or_insert_with(Default::default).text_direction =
                    Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "bidi" => {
                instance.get_or_insert_with(Default::default).bidirectional = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "rtlGutter" => {
                instance.get_or_insert_with(Default::default).rtl_gutter = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "docGrid" => {
                instance.get_or_insert_with(Default::default).document_grid =
                    Some(DocGrid::from_xml_element(xml_node)?);
                Ok(true)
            }
            "printerSettings" => {
                instance.get_or_insert_with(Default::default).printer_settings = Some(Rel::from_xml_element(xml_node)?);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SectPrAttributes {
    pub run_properties_revision_id: Option<LongHexNumber>,
    pub deletion_revision_id: Option<LongHexNumber>,
    pub run_revision_id: Option<LongHexNumber>,
    pub section_revision_id: Option<LongHexNumber>,
}

impl SectPrAttributes {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "rsidRPr" => instance.run_properties_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "rsidDel" => instance.deletion_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "rsidR" => instance.run_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "rsidSect" => instance.section_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SectPrBase {
    pub contents: Option<SectPrContents>,
    pub attributes: SectPrAttributes,
}

impl SectPrBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        Ok(Self {
            contents: SectPrContents::from_xml_element(xml_node)?,
            attributes: SectPrAttributes::from_xml_element(xml_node)?,
        })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct SectPrChange {
    pub base: TrackChange,
    pub section_properties: Option<SectPrBase>,
}

impl SectPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = TrackChange::from_xml_element(xml_node)?;
        let section_properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "sectPr")
            .map(|child_node| SectPrBase::from_xml_element(child_node))
            .transpose()?;

        Ok(Self {
            base,
            section_properties,
        })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct SectPr {
    pub header_footer_references: Vec<HdrFtrReferences>,
    pub contents: Option<SectPrContents>,
    pub change: Option<SectPrChange>,
    pub attributes: SectPrAttributes,
}

impl SectPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        instance.attributes = SectPrAttributes::from_xml_element(xml_node)?;

        for child_node in &xml_node.child_nodes {
            if let Some(result) = HdrFtrReferences::try_from_xml_element(child_node) {
                instance.header_footer_references.push(result?);
                continue;
            }

            if SectPrContents::try_parse_group_node(&mut instance.contents, child_node)? {
                continue;
            }

            if child_node.local_name() == "sectPrChange" {
                instance.change = Some(SectPrChange::from_xml_element(child_node)?);
            }
        }

        match instance.header_footer_references.len() {
            0..=6 => Ok(instance),
            occurs @ _ => Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "headerReference|footerReference",
                0,
                MaxOccurs::Value(6),
                occurs as u32,
            ))),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct PPrChange {
    pub base: TrackChange,
    pub properties: PPrBase,
}

impl PPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = TrackChange::from_xml_element(xml_node)?;
        let properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "pPr")
            .map(|child_node| PPrBase::from_xml_element(child_node))
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "pPr"))?;

        Ok(Self { base, properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct PPr {
    pub base: PPrBase,
    pub run_properties: Option<ParaRPr>,
    pub section_properties: Option<SectPr>,
    pub properties_change: Option<PPrChange>,
}

impl PPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "rPr" => instance.run_properties = Some(ParaRPr::from_xml_element(child_node)?),
                "sectPr" => instance.section_properties = Some(SectPr::from_xml_element(child_node)?),
                "pPrChange" => instance.properties_change = Some(PPrChange::from_xml_element(child_node)?),
                _ => {
                    instance.base.try_update_from_xml_element(child_node)?;
                }
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct P {
    pub properties: Option<PPr>,
    pub contents: Vec<PContent>,
    pub run_properties_revision_id: Option<LongHexNumber>,
    pub run_revision_id: Option<LongHexNumber>,
    pub deletion_revision_id: Option<LongHexNumber>,
    pub paragraph_revision_id: Option<LongHexNumber>,
    pub run_default_revision_id: Option<LongHexNumber>,
}

impl P {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "rsidRPr" => instance.run_properties_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "rsidR" => instance.run_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "rsidDel" => instance.deletion_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "rsidP" => instance.paragraph_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                "rsidRDefault" => instance.run_default_revision_id = Some(LongHexNumber::from_str_radix(value, 16)?),
                _ => (),
            }
        }

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "pPr" => instance.properties = Some(PPr::from_xml_element(child_node)?),
                node_name @ _ if PContent::is_choice_member(node_name) => {
                    instance.contents.push(PContent::from_xml_element(child_node)?);
                }
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPPr {
    pub left_from_text: Option<TwipsMeasure>,
    pub right_from_text: Option<TwipsMeasure>,
    pub top_from_text: Option<TwipsMeasure>,
    pub bottom_from_text: Option<TwipsMeasure>,
    pub vertical_anchor: Option<VAnchor>,
    pub horizontal_anchor: Option<HAnchor>,
    pub horizontal_alignment: Option<XAlign>,
    pub horizontal_distance: Option<SignedTwipsMeasure>,
    pub vertical_alignment: Option<YAlign>,
    pub vertical_distance: Option<SignedTwipsMeasure>,
}

impl TblPPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "leftFromText" => instance.left_from_text = Some(value.parse()?),
                "rightFromText" => instance.right_from_text = Some(value.parse()?),
                "topFromText" => instance.top_from_text = Some(value.parse()?),
                "bottomFromText" => instance.bottom_from_text = Some(value.parse()?),
                "vertAnchor" => instance.vertical_anchor = Some(value.parse()?),
                "horzAnchor" => instance.horizontal_anchor = Some(value.parse()?),
                "tblpXSpec" => instance.horizontal_alignment = Some(value.parse()?),
                "tblpX" => instance.horizontal_distance = Some(value.parse()?),
                "tblpYSpec" => instance.vertical_alignment = Some(value.parse()?),
                "tblpY" => instance.vertical_distance = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TblOverlap {
    #[strum(serialize = "never")]
    Never,
    #[strum(serialize = "overlap")]
    Overlap,
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TblWidthType {
    #[strum(serialize = "nil")]
    NoWidth,
    #[strum(serialize = "percent")]
    Percent,
    #[strum(serialize = "dxa")]
    TwentiethsOfPoint,
    #[strum(serialize = "auto")]
    Auto,
}

#[derive(Debug, Clone, PartialEq)]
pub enum MeasurementOrPercent {
    DecimalOrPercent(DecimalNumberOrPercent),
    UniversalMeasure(UniversalMeasure),
}

impl FromStr for MeasurementOrPercent {
    type Err = Box<dyn std::error::Error>;

    fn from_str(s: &str) -> std::result::Result<Self, Self::Err> {
        if let Ok(value) = s.parse::<DecimalNumberOrPercent>() {
            Ok(MeasurementOrPercent::DecimalOrPercent(value))
        } else {
            Ok(MeasurementOrPercent::UniversalMeasure(s.parse()?))
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblWidth {
    pub width: Option<MeasurementOrPercent>,
    pub width_type: Option<TblWidthType>,
}

impl TblWidth {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "w" => instance.width = Some(value.parse()?),
                "type" => instance.width_type = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum JcTable {
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "end")]
    End,
    #[strum(serialize = "start")]
    Start,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblBorders {
    pub top: Option<Border>,
    pub start: Option<Border>,
    pub bottom: Option<Border>,
    pub end: Option<Border>,
    pub inside_horizontal: Option<Border>,
    pub inside_vertical: Option<Border>,
}

impl TblBorders {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(Border::from_xml_element(child_node)?),
                "start" => instance.start = Some(Border::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(Border::from_xml_element(child_node)?),
                "end" => instance.end = Some(Border::from_xml_element(child_node)?),
                "insideH" => instance.inside_horizontal = Some(Border::from_xml_element(child_node)?),
                "insideV" => instance.inside_vertical = Some(Border::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum TblLayoutType {
    #[strum(serialize = "fixed")]
    Fixed,
    #[strum(serialize = "autofit")]
    Autofit,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblCellMar {
    pub top: Option<TblWidth>,
    pub start: Option<TblWidth>,
    pub bottom: Option<TblWidth>,
    pub end: Option<TblWidth>,
}

impl TblCellMar {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(TblWidth::from_xml_element(child_node)?),
                "start" => instance.start = Some(TblWidth::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(TblWidth::from_xml_element(child_node)?),
                "end" => instance.end = Some(TblWidth::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblLook {
    pub first_row: Option<OnOff>,
    pub last_row: Option<OnOff>,
    pub first_column: Option<OnOff>,
    pub last_column: Option<OnOff>,
    pub no_horizontal_band: Option<OnOff>,
    pub no_vertical_band: Option<OnOff>,
}

impl TblLook {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "firstRow" => instance.first_row = Some(parse_xml_bool(value)?),
                "lastRow" => instance.last_row = Some(parse_xml_bool(value)?),
                "firstColumn" => instance.first_column = Some(parse_xml_bool(value)?),
                "lastColumn" => instance.last_column = Some(parse_xml_bool(value)?),
                "noHBand" => instance.no_horizontal_band = Some(parse_xml_bool(value)?),
                "noVBand" => instance.no_vertical_band = Some(parse_xml_bool(value)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPrBase {
    pub style: Option<String>,
    pub paragraph_properties: Option<TblPPr>,
    pub overlap: Option<TblOverlap>,
    pub bidirectional_visual: Option<OnOff>,
    pub style_row_band_size: Option<DecimalNumber>,
    pub style_column_band_size: Option<DecimalNumber>,
    pub width: Option<TblWidth>,
    pub alignment: Option<JcTable>,
    pub cell_spacing: Option<TblWidth>,
    pub indent: Option<TblWidth>,
    pub borders: Option<TblBorders>,
    pub shading: Option<Shd>,
    pub layout: Option<TblLayoutType>,
    pub cell_margin: Option<TblCellMar>,
    pub look: Option<TblLook>,
    pub caption: Option<String>,
    pub description: Option<String>,
}

impl TblPrBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            instance.try_update_from_xml_element(child_node)?;
        }

        Ok(instance)
    }

    pub fn try_update_from_xml_element(&mut self, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "tblStyle" => {
                self.style = Some(xml_node.get_val_attribute()?.clone());
                Ok(true)
            }
            "tblpPr" => {
                self.paragraph_properties = Some(TblPPr::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblOverlap" => {
                self.overlap = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "bidiVisual" => {
                self.bidirectional_visual = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "tblStyleRowBandSize" => {
                self.style_row_band_size = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "tblStyleColBandSize" => {
                self.style_column_band_size = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "tblW" => {
                self.width = Some(TblWidth::from_xml_element(xml_node)?);
                Ok(true)
            }
            "jc" => {
                self.alignment = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "tblCellSpacing" => {
                self.cell_spacing = Some(TblWidth::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblInd" => {
                self.indent = Some(TblWidth::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblBorders" => {
                self.borders = Some(TblBorders::from_xml_element(xml_node)?);
                Ok(true)
            }
            "shd" => {
                self.shading = Some(Shd::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblLayout" => {
                self.layout = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "tblCellMar" => {
                self.cell_margin = Some(TblCellMar::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblLook" => {
                self.look = Some(TblLook::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblCaption" => {
                self.caption = Some(xml_node.get_val_attribute()?.clone());
                Ok(true)
            }
            "tblDescription" => {
                self.description = Some(xml_node.get_val_attribute()?.clone());
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TblPrChange {
    pub base: TrackChange,
    pub properties: TblPrBase,
}

impl TblPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = TrackChange::from_xml_element(xml_node)?;
        let properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "tblPr")
            .map(|child_node| TblPrBase::from_xml_element(child_node))
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tblPr"))?;

        Ok(Self { base, properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPr {
    pub base: TblPrBase,
    pub change: Option<TblPrChange>,
}

impl TblPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tblPrChange" => instance.change = Some(TblPrChange::from_xml_element(child_node)?),
                _ => {
                    instance.base.try_update_from_xml_element(child_node)?;
                }
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblGridCol {
    pub width: Option<TwipsMeasure>,
}

impl TblGridCol {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let width = xml_node.attributes.get("w").map(|value| value.parse()).transpose()?;

        Ok(Self { width })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TblGridChange {
    pub base: Markup,
    pub grid: TblGridBase,
}

impl TblGridChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = Markup::from_xml_element(xml_node)?;
        let grid = TblGridBase::from_xml_element(xml_node)?;

        Ok(Self { base, grid })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblGridBase {
    pub columns: Vec<TblGridCol>,
}

impl TblGridBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let columns = xml_node
            .child_nodes
            .iter()
            .filter(|child_node| child_node.local_name() == "gridCol")
            .map(|child_node| TblGridCol::from_xml_element(child_node))
            .collect::<Result<Vec<_>>>()?;

        Ok(Self { columns })
    }

    pub fn try_update_from_xml_element(&mut self, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "gridCol" => {
                self.columns.push(TblGridCol::from_xml_element(xml_node)?);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblGrid {
    pub base: TblGridBase,
    pub change: Option<TblGridChange>,
}

impl TblGrid {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tblGridChange" => instance.change = Some(TblGridChange::from_xml_element(child_node)?),
                _ => {
                    instance.base.try_update_from_xml_element(child_node)?;
                }
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPrExBase {
    pub width: Option<TblWidth>,
    pub alignment: Option<JcTable>,
    pub cell_spacing: Option<TblWidth>,
    pub indent: Option<TblWidth>,
    pub borders: Option<TblBorders>,
    pub shading: Option<Shd>,
    pub layout: Option<TblLayoutType>,
    pub cell_margin: Option<TblCellMar>,
    pub look: Option<TblLook>,
}

impl TblPrExBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            instance.try_update_from_xml_element(child_node)?;
        }

        Ok(instance)
    }

    pub fn try_update_from_xml_element(&mut self, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "tblW" => {
                self.width = Some(TblWidth::from_xml_element(xml_node)?);
                Ok(true)
            }
            "jc" => {
                self.alignment = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "tblCellSpacing" => {
                self.cell_spacing = Some(TblWidth::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblInd" => {
                self.indent = Some(TblWidth::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblBorders" => {
                self.borders = Some(TblBorders::from_xml_element(xml_node)?);
                Ok(true)
            }
            "shd" => {
                self.shading = Some(Shd::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblLayout" => {
                self.layout = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "tblCellMar" => {
                self.cell_margin = Some(TblCellMar::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblLook" => {
                self.look = Some(TblLook::from_xml_element(xml_node)?);
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TblPrExChange {
    pub base: TrackChange,
    pub properties_ex: TblPrExBase,
}

impl TblPrExChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = TrackChange::from_xml_element(xml_node)?;
        let properties_ex = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "tblPrEx")
            .map(|child_node| TblPrExBase::from_xml_element(child_node))
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "tblPrEx"))?;

        Ok(Self { base, properties_ex })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TblPrEx {
    pub base: TblPrExBase,
    pub change: Option<TblPrExChange>,
}

impl TblPrEx {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "tblPrExChange" => instance.change = Some(TblPrExChange::from_xml_element(child_node)?),
                _ => {
                    instance.base.try_update_from_xml_element(child_node)?;
                }
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct Height {
    pub value: Option<TwipsMeasure>,
    pub height_rule: Option<HeightRule>,
}

impl Height {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => instance.value = Some(value.parse()?),
                "hRule" => instance.height_rule = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TrPrBase {
    pub conditional_formatting: Option<Cnf>,
    pub div_id: Option<DecimalNumber>,
    pub grid_column_before_first_cell: Option<DecimalNumber>,
    pub grid_column_after_last_cell: Option<DecimalNumber>,
    pub width_before_row: Option<TblWidth>,
    pub width_after_row: Option<TblWidth>,
    pub cant_split: Option<OnOff>,
    pub row_height: Option<Height>,
    pub header: Option<OnOff>,
    pub cell_spacing: Option<TblWidth>,
    pub alignment: Option<JcTable>,
    pub hidden: Option<OnOff>,
}

impl TrPrBase {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            instance.try_update_from_xml_element(child_node)?;
        }

        Ok(instance)
    }

    pub fn try_update_from_xml_element(&mut self, xml_node: &XmlNode) -> Result<bool> {
        match xml_node.local_name() {
            "cnfStyle" => {
                self.conditional_formatting = Some(Cnf::from_xml_element(xml_node)?);
                Ok(true)
            }
            "divId" => {
                self.div_id = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "gridBefore" => {
                self.grid_column_before_first_cell = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "gridAfter" => {
                self.grid_column_after_last_cell = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "wBefore" => {
                self.width_before_row = Some(TblWidth::from_xml_element(xml_node)?);
                Ok(true)
            }
            "wAfter" => {
                self.width_after_row = Some(TblWidth::from_xml_element(xml_node)?);
                Ok(true)
            }
            "cantSplit" => {
                self.cant_split = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "trHeight" => {
                self.row_height = Some(Height::from_xml_element(xml_node)?);
                Ok(true)
            }
            "tblHeader" => {
                self.header = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            "tblCellSpacing" => {
                self.cell_spacing = Some(TblWidth::from_xml_element(xml_node)?);
                Ok(true)
            }
            "jc" => {
                self.alignment = Some(xml_node.get_val_attribute()?.parse()?);
                Ok(true)
            }
            "hidden" => {
                self.hidden = parse_on_off_xml_element(xml_node)?;
                Ok(true)
            }
            _ => Ok(false),
        }
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct TrPrChange {
    pub base: TrackChange,
    pub properties: TrPrBase,
}

impl TrPrChange {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let base = TrackChange::from_xml_element(xml_node)?;
        let properties = xml_node
            .child_nodes
            .iter()
            .find(|child_node| child_node.local_name() == "trPr")
            .map(|child_node| TrPrBase::from_xml_element(child_node))
            .transpose()?
            .ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "trPr"))?;

        Ok(Self { base, properties })
    }
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TrPr {
    pub base: TrPrBase,
    pub inserted: Option<TrackChange>,
    pub deleted: Option<TrackChange>,
    pub change: Option<TrPrChange>,
}

impl TrPr {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "ins" => instance.inserted = Some(TrackChange::from_xml_element(child_node)?),
                "del" => instance.deleted = Some(TrackChange::from_xml_element(child_node)?),
                "trPrChange" => instance.change = Some(TrPrChange::from_xml_element(child_node)?),
                _ => {
                    instance.base.try_update_from_xml_element(child_node)?;
                }
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone, PartialEq, EnumString)]
pub enum Merge {
    #[strum(serialize = "continue")]
    Continue,
    #[strum(serialize = "restart")]
    Restart,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub struct TcBorders {
    pub top: Option<Border>,
    pub start: Option<Border>,
    pub bottom: Option<Border>,
    pub end: Option<Border>,
    pub inside_horizontal: Option<Border>,
    pub inside_vertical: Option<Border>,
    pub top_left_to_bottom_right: Option<Border>,
    pub top_right_to_bottom_left: Option<Border>,
}

impl TcBorders {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Self = Default::default();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "top" => instance.top = Some(Border::from_xml_element(child_node)?),
                "start" => instance.start = Some(Border::from_xml_element(child_node)?),
                "bottom" => instance.bottom = Some(Border::from_xml_element(child_node)?),
                "end" => instance.end = Some(Border::from_xml_element(child_node)?),
                "insideH" => instance.inside_horizontal = Some(Border::from_xml_element(child_node)?),
                "insideV" => instance.inside_vertical = Some(Border::from_xml_element(child_node)?),
                "tl2br" => instance.top_left_to_bottom_right = Some(Border::from_xml_element(child_node)?),
                "tr2bl" => instance.top_right_to_bottom_left = Some(Border::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(instance)
    }
}

/*
<xsd:complexType name="CT_TcPrBase">
    <xsd:sequence>
      <xsd:element name="cnfStyle" type="CT_Cnf" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="tcW" type="CT_TblWidth" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="gridSpan" type="CT_DecimalNumber" minOccurs="0"/>
      <xsd:element name="vMerge" type="CT_VMerge" minOccurs="0"/>
      <xsd:element name="tcBorders" type="CT_TcBorders" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="shd" type="CT_Shd" minOccurs="0"/>
      <xsd:element name="noWrap" type="CT_OnOff" minOccurs="0"/>
      <xsd:element name="tcMar" type="CT_TcMar" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="textDirection" type="CT_TextDirection" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="tcFitText" type="CT_OnOff" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="vAlign" type="CT_VerticalJc" minOccurs="0"/>
      <xsd:element name="hideMark" type="CT_OnOff" minOccurs="0"/>
      <xsd:element name="headers" type="CT_Headers" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>
  */
#[derive(Debug, Clone, PartialEq, Default)]
pub struct TcPrBase {
    pub conditional_formatting: Option<Cnf>,
    pub width: Option<TblWidth>,
    pub grid_span: Option<DecimalNumber>,
    pub vertical_merge: Option<Merge>,
    pub borders: Option<TcBorders>,
    pub shading: Option<Shd>,
    pub no_wrapping: Option<OnOff>,
    //pub margin: Option<TcMar>,
    pub text_direction: Option<TextDirection>,
    pub fit_text: Option<OnOff>,
    pub vertical_alignment: Option<VerticalJc>,
    pub hide_marker: Option<OnOff>,
    //pub headers: Option<Headers>,
}

/*
<xsd:complexType name="CT_TcPr">
  <xsd:complexContent>
    <xsd:extension base="CT_TcPrInner">
      <xsd:sequence>
        <xsd:element name="tcPrChange" type="CT_TcPrChange" minOccurs="0"/>
      </xsd:sequence>
    </xsd:extension>
  </xsd:complexContent>
</xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq, Default)]
pub struct TcPr {
    pub base: TcPrInner,
    //pub change: Option<TcPrChange>,
}
/*
  <xsd:complexType name="CT_TcPrInner">
    <xsd:complexContent>
      <xsd:extension base="CT_TcPrBase">
        <xsd:sequence>
          <xsd:group ref="EG_CellMarkupElements" minOccurs="0" maxOccurs="1"/>
        </xsd:sequence>
      </xsd:extension>
    </xsd:complexContent>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq, Default)]
pub struct TcPrInner {
    pub base: TcPrBase,
    //pub markup_element: Option<CellMarkupElements>,
}

/*
<xsd:complexType name="CT_Tc">
    <xsd:sequence>
      <xsd:element name="tcPr" type="CT_TcPr" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_BlockLevelElts" minOccurs="1" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="id" type="s:ST_String" use="optional"/>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Tc {
    pub properties: Option<TcPr>,
    pub block_level_elements: Vec<BlockLevelElts>, // minOccurs="1"
}

/*
<xsd:group name="EG_ContentCellContent">
    <xsd:choice>
      <xsd:element name="tc" type="CT_Tc" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="customXml" type="CT_CustomXmlCell"/>
      <xsd:element name="sdt" type="CT_SdtCell"/>
      <xsd:group ref="EG_RunLevelElts" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:choice>
  </xsd:group>
*/
#[derive(Debug, Clone, PartialEq)]
pub enum ContentCellContent {
    Cell(Tc),
    // CustomXml(CustomXmlCell),
    // Sdt(SdtCell),
    RunLevelElements(RunLevelElts),
}

/*
<xsd:complexType name="CT_Row">
    <xsd:sequence>
      <xsd:element name="tblPrEx" type="CT_TblPrEx" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="trPr" type="CT_TrPr" minOccurs="0" maxOccurs="1"/>
      <xsd:group ref="EG_ContentCellContent" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="rsidRPr" type="ST_LongHexNumber"/>
    <xsd:attribute name="rsidR" type="ST_LongHexNumber"/>
    <xsd:attribute name="rsidDel" type="ST_LongHexNumber"/>
    <xsd:attribute name="rsidTr" type="ST_LongHexNumber"/>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq, Default)]
pub struct Row {
    pub property_exceptions: Option<TblPrEx>,
    pub properties: Option<TrPr>,
    pub contents: Vec<ContentCellContent>,
    pub run_properties_revision_id: Option<LongHexNumber>,
    pub run_revision_id: Option<LongHexNumber>,
    pub deletion_revision_id: Option<LongHexNumber>,
    pub row_revision_id: Option<LongHexNumber>,
}

/*
<xsd:group name="EG_ContentRowContent">
    <xsd:choice>
      <xsd:element name="tr" type="CT_Row" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="customXml" type="CT_CustomXmlRow"/>
      <xsd:element name="sdt" type="CT_SdtRow"/>
      <xsd:group ref="EG_RunLevelElts" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:choice>
  </xsd:group>
*/
#[derive(Debug, Clone, PartialEq)]
pub enum ContentRowContent {
    Table(Row),
    // CustomXml(CustomXmlRow),
    // Sdt(SdtRow),
    RunLevelElements(RunLevelElts),
}

/*
<xsd:complexType name="CT_Tbl">
    <xsd:sequence>
      <xsd:group ref="EG_RangeMarkupElements" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="tblPr" type="CT_TblPr"/>
      <xsd:element name="tblGrid" type="CT_TblGrid"/>
      <xsd:group ref="EG_ContentRowContent" minOccurs="0" maxOccurs="unbounded"/>
    </xsd:sequence>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq)]
pub struct Tbl {
    pub range_markup_elements: Vec<RangeMarkupElements>,
    pub properties: TblPr,
    pub grid: TblGrid,
    pub row_contents: Vec<ContentRowContent>,
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
    CustomXml(CustomXmlBlock),
    Sdt(SdtBlock),
    Paragraph(P),
    Table(Tbl),
    RunLevelElement(RunLevelElts),
}

impl XsdChoice for ContentBlockContent {
    fn is_choice_member<T: AsRef<str>>(node_name: T) -> bool {
        match node_name.as_ref() {
            "customXml" | "sdt" | "p" | "tbl" => true,
            _ => RunLevelElts::is_choice_member(&node_name),
        }
    }

    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "customXml" => Ok(ContentBlockContent::CustomXml(CustomXmlBlock::from_xml_element(
                xml_node,
            )?)),
            node_name @ _ if RunLevelElts::is_choice_member(&node_name) => Ok(ContentBlockContent::RunLevelElement(
                RunLevelElts::from_xml_element(xml_node)?,
            )),
            _ => Err(Box::new(NotGroupMemberError::new(
                xml_node.name.clone(),
                "ContentBlockContent",
            ))),
        }
    }
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
    Content(ContentBlockContent),
}
// <xsd:group name="EG_BlockLevelElts">
//     <xsd:choice>
//       <xsd:group ref="EG_BlockLevelChunkElts" minOccurs="0" maxOccurs="unbounded"/>
//       <xsd:element name="altChunk" type="CT_AltChunk" minOccurs="0" maxOccurs="unbounded"/>
//     </xsd:choice>
//   </xsd:group>
#[derive(Debug, Clone, PartialEq)]
pub enum BlockLevelElts {
    Chunks(BlockLevelChunkElts),
    //AltChunks(AltChunk),
}

#[cfg(test)]
mod tests {
    use super::*;
    use msoffice_shared::sharedtypes::UniversalMeasureUnit;

    #[test]
    pub fn test_parse_text_scale_percent() {
        assert_eq!(parse_text_scale_percent("100%").unwrap(), 1.0);
        assert_eq!(parse_text_scale_percent("600%").unwrap(), 6.0);
        assert_eq!(parse_text_scale_percent("333%").unwrap(), 3.33);
        assert_eq!(parse_text_scale_percent("0%").unwrap(), 0.0);
    }

    impl SignedTwipsMeasure {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} val="123.456mm"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            SignedTwipsMeasure::UniversalMeasure(UniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter))
        }
    }

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

    #[test]
    pub fn test_signed_twips_measure_from_xml() {
        let xml = SignedTwipsMeasure::test_xml("signedTwipsMeasure");
        let signed_twips_measure = SignedTwipsMeasure::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(signed_twips_measure, SignedTwipsMeasure::test_instance());
    }

    impl HpsMeasure {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} val="123.456mm"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            HpsMeasure::UniversalMeasure(PositiveUniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter))
        }
    }

    #[test]
    pub fn test_hps_measure_from_str() {
        use msoffice_shared::sharedtypes::UniversalMeasureUnit;

        assert_eq!("123".parse::<HpsMeasure>().unwrap(), HpsMeasure::Decimal(123));
        assert_eq!(
            "123.456mm".parse::<HpsMeasure>().unwrap(),
            HpsMeasure::UniversalMeasure(PositiveUniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter)),
        );
    }

    #[test]
    pub fn test_hps_measure_from_xml() {
        let xml = HpsMeasure::test_xml("hpsMeasure");
        let hps_measure = HpsMeasure::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(hps_measure, HpsMeasure::test_instance());
    }

    impl SignedHpsMeasure {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} val="123.456mm"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            SignedHpsMeasure::UniversalMeasure(UniversalMeasure::new(123.456, UniversalMeasureUnit::Millimeter))
        }
    }

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

    #[test]
    pub fn test_signed_hps_measure_from_xml() {
        let xml = SignedHpsMeasure::test_xml("signedHpsMeasure");
        let hps_measure = SignedHpsMeasure::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(hps_measure, SignedHpsMeasure::test_instance());
    }

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

    #[test]
    pub fn test_color_from_xml() {
        let xml = Color::test_xml("color");
        let color = Color::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(color, Color::test_instance());
    }

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

    #[test]
    pub fn test_proof_err_from_xml() {
        let xml = ProofErr::test_xml("proofErr");
        let proof_err = ProofErr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(proof_err, ProofErr::test_instance());
    }

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

    #[test]
    pub fn test_perm_from_xml() {
        let xml = Perm::test_xml("perm");
        let perm = Perm::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(perm, Perm::test_instance());
    }

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

    #[test]
    pub fn test_perm_start_from_xml() {
        let xml = PermStart::test_xml("permStart");
        let perm_start = PermStart::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(perm_start, PermStart::test_instance());
    }

    impl Markup {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} id="0"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self { id: 0 }
        }
    }

    #[test]
    pub fn test_markup_from_xml() {
        let xml = Markup::test_xml("markup");
        let markup = Markup::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(markup, Markup::test_instance());
    }

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

    #[test]
    pub fn test_markup_range_from_xml() {
        let xml = MarkupRange::test_xml("markupRange");
        let markup_range = MarkupRange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(markup_range, MarkupRange::test_instance());
    }

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

    #[test]
    pub fn test_bookmark_range_from_xml() {
        let xml = BookmarkRange::test_xml("bookmarkRange");
        let bookmark_range = BookmarkRange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(bookmark_range, BookmarkRange::test_instance());
    }

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

    #[test]
    fn test_bookmark_from_xml() {
        let xml = Bookmark::test_xml("bookmark");
        let bookmark = Bookmark::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(bookmark, Bookmark::test_instance());
    }

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

    #[test]
    fn test_move_bookmark_from_xml() {
        let xml = MoveBookmark::test_xml("moveBookmark");
        let move_bookmark = MoveBookmark::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(move_bookmark, MoveBookmark::test_instance());
    }

    impl TrackChange {
        const TEST_ATTRIBUTES: &'static str = r#"id="0" author="John Smith" date="2001-10-26T21:32:52""#;

        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}></{node_name}>"#,
                Self::TEST_ATTRIBUTES,
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

    #[test]
    fn test_track_change_from_xml() {
        let xml = TrackChange::test_xml("trackChange");
        let track_change = TrackChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(track_change, TrackChange::test_instance());
    }

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

    #[test]
    pub fn test_attr_from_xml() {
        let xml = Attr::test_xml("attr");
        let attr = Attr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(attr, Attr::test_instance());
    }

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

    #[test]
    pub fn test_custom_xml_pr_from_xml() {
        let xml = CustomXmlPr::test_xml("customXmlPr");
        let custom_xml_pr = CustomXmlPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(custom_xml_pr, CustomXmlPr::test_instance());
    }

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

    #[test]
    pub fn test_simple_field_from_xml() {
        let xml = SimpleField::test_xml_recursive("simpleField");
        let simple_field = SimpleField::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(simple_field, SimpleField::test_instance_recursive());
    }

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

    #[test]
    pub fn test_hyperlink_from_xml() {
        let xml = Hyperlink::test_xml_recursive("hyperlink");
        let hyperlink = Hyperlink::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(hyperlink, Hyperlink::test_instance_recursive());
    }

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

    #[test]
    pub fn test_rel_from_xml() {
        let xml = Rel::test_xml("rel");
        let rel = Rel::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(rel, Rel::test_instance());
    }

    impl PContent {
        pub fn test_simple_field_xml() -> String {
            SimpleField::test_xml("fldSimple")
        }

        pub fn test_simple_field_instance() -> Self {
            PContent::SimpleField(SimpleField::test_instance())
        }
    }

    #[test]
    pub fn test_pcontent_content_run_content_from_xml() {
        // TODO
    }

    #[test]
    pub fn test_pcontent_simple_field_from_xml() {
        let xml = SimpleField::test_xml("fldSimple");
        let pcontent = PContent::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(pcontent, PContent::SimpleField(SimpleField::test_instance()));
    }

    #[test]
    pub fn test_pcontent_hyperlink_from_xml() {
        let xml = Hyperlink::test_xml("hyperlink");
        let pcontent = PContent::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(pcontent, PContent::Hyperlink(Hyperlink::test_instance()));
    }

    #[test]
    pub fn test_pcontent_subdocument_from_xml() {
        let xml = Rel::test_xml("subDoc");
        let pcontent = PContent::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(pcontent, PContent::SubDocument(Rel::test_instance()));
    }

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

    #[test]
    pub fn test_custom_xml_run_from_xml() {
        let xml = CustomXmlRun::test_xml("customXmlRun");
        let custom_xml_run = CustomXmlRun::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(custom_xml_run, CustomXmlRun::test_instance());
    }

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

    #[test]
    pub fn test_smart_tag_pr_from_xml() {
        let xml = SmartTagPr::test_xml("smartTagPr");
        let smart_tag_pr = SmartTagPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(smart_tag_pr, SmartTagPr::test_instance());
    }

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

    #[test]
    pub fn test_smart_tag_run_from_xml() {
        let xml = SmartTagRun::test_xml("smartTagRun");
        let smart_tag_run = SmartTagRun::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(smart_tag_run, SmartTagRun::test_instance());
    }

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

    #[test]
    pub fn test_fonts_from_xml() {
        let xml = Fonts::test_xml("fonts");
        let fonts = Fonts::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(fonts, Fonts::test_instance());
    }

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

    #[test]
    pub fn test_underline_from_xml() {
        let xml = Underline::test_xml("underline");
        let underline = Underline::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(underline, Underline::test_instance());
    }

    impl Border {
        const TEST_ATTRIBUTES: &'static str = r#"val="single" color="ffffff" themeColor="accent1" themeTint="ff"
            themeShade="ff" sz="100" space="100" shadow="true" frame="true""#;

        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                </{node_name}>"#,
                Self::TEST_ATTRIBUTES,
                node_name = node_name,
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

    #[test]
    pub fn test_border_from_xml() {
        let xml = Border::test_xml("border");
        let border = Border::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(border, Border::test_instance());
    }

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

    #[test]
    pub fn test_shd_from_xml() {
        let xml = Shd::test_xml("shd");
        let shd = Shd::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(shd, Shd::test_instance());
    }

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

    #[test]
    pub fn test_fit_text_from_xml() {
        let xml = FitText::test_xml("fitText");
        let fit_text = FitText::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(fit_text, FitText::test_instance());
    }

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

    #[test]
    pub fn test_language_from_xml() {
        let xml = Language::test_xml("language");
        let language = Language::from_xml_element(&XmlNode::from_str(xml).unwrap());
        assert_eq!(language, Language::test_instance());
    }

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

    #[test]
    pub fn test_east_asian_layout_from_xml() {
        let xml = EastAsianLayout::test_xml("eastAsianLayout");
        let east_asian_layout = EastAsianLayout::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(east_asian_layout, EastAsianLayout::test_instance());
    }

    impl RPrBase {
        pub fn test_run_style_xml() -> &'static str {
            r#"<rStyle val="Arial"></rStyle>"#
        }

        pub fn test_run_style_instance() -> Self {
            RPrBase::RunStyle(String::from("Arial"))
        }
    }

    // TODO Write some more unit tests

    #[test]
    pub fn test_r_pr_base_run_style_from_xml() {
        let xml = RPrBase::test_run_style_xml();
        let r_pr_base = RPrBase::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(r_pr_base, RPrBase::test_run_style_instance());
    }

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

    #[test]
    pub fn test_r_pr_original_from_xml() {
        let xml = RPrOriginal::test_xml("rPrOriginal");
        let r_pr_original = RPrOriginal::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(r_pr_original, RPrOriginal::test_instance());
    }

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

    #[test]
    pub fn test_r_pr_change_from_xml() {
        let xml = RPrChange::test_xml("rRpChange");
        let r_pr_change = RPrChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(r_pr_change, RPrChange::test_instance());
    }

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

    #[test]
    pub fn test_r_pr_from_xml() {
        let xml = RPr::test_xml("rPr");
        let r_pr_content = RPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(r_pr_content, RPr::test_instance());
    }

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

    #[test]
    pub fn test_sdt_list_item_from_xml() {
        let xml = SdtListItem::test_xml("sdtListItem");
        let sdt_list_item = SdtListItem::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(sdt_list_item, SdtListItem::test_instance());
    }

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

    #[test]
    pub fn test_sdt_combo_box_from_xml() {
        let xml = SdtComboBox::test_xml("sdtComboBox");
        let sdt_combo_box = SdtComboBox::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(sdt_combo_box, SdtComboBox::test_instance());
    }

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

    #[test]
    pub fn test_sdt_date_from_xml() {
        let xml = SdtDate::test_xml("sdtDate");
        let sdt_date = SdtDate::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(sdt_date, SdtDate::test_instance());
    }

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

    #[test]
    pub fn test_sdt_doc_part_from_xml() {
        let xml = SdtDocPart::test_xml("sdtDocPart");
        let sdt_doc_part = SdtDocPart::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(sdt_doc_part, SdtDocPart::test_instance());
    }

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

    #[test]
    pub fn test_sdt_drop_down_list_from_xml() {
        let xml = SdtDropDownList::test_xml("sdtDropDownList");
        let sdt_combo_box = SdtDropDownList::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(sdt_combo_box, SdtDropDownList::test_instance());
    }

    impl SdtText {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} multiLine="true"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self { is_multi_line: true }
        }
    }

    #[test]
    pub fn test_sdt_text_from_xml() {
        let xml = SdtText::test_xml("sdtText");
        let sdt_text = SdtText::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
        assert_eq!(sdt_text, SdtText::test_instance());
    }

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
            SdtPrChoice::from_xml_element(&XmlNode::from_str(SdtDropDownList::test_xml("dropDownList")).unwrap())
                .unwrap(),
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

    #[test]
    pub fn test_placeholder_from_xml() {
        let xml = Placeholder::test_xml("placeholder");
        assert_eq!(
            Placeholder::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Placeholder::test_instance()
        );
    }

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

    #[test]
    pub fn test_data_binding_from_xml() {
        let xml = DataBinding::test_xml("dataBinding");
        assert_eq!(
            DataBinding::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            DataBinding::test_instance()
        );
    }

    impl SdtPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            {}
            <alias val="Alias" />
            <tag val="Tag"/>
            <id val="1" />
            <lock val="unlocked" />
            {}
            <temporary val="false" />
            <showingPlcHdr val="false" />
            {}
            <label val="1" />
            <tabIndex val="1" />
            <equation />
        </{node_name}>"#,
                RPr::test_xml("rPr"),
                Placeholder::test_xml("placeholder"),
                DataBinding::test_xml("dataBinding"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                run_properties: Some(RPr::test_instance()),
                alias: Some(String::from("Alias")),
                tag: Some(String::from("Tag")),
                id: Some(1),
                lock: Some(Lock::Unlocked),
                placeholder: Some(Placeholder::test_instance()),
                temporary: Some(false),
                showing_placeholder_header: Some(false),
                data_binding: Some(DataBinding::test_instance()),
                label: Some(1),
                tab_index: Some(1),
                control_choice: Some(SdtPrChoice::Equation),
            }
        }
    }

    #[test]
    pub fn test_sdt_pr_from_xml() {
        let xml = SdtPr::test_xml("sdtPr");
        assert_eq!(
            SdtPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SdtPr::test_instance()
        );
    }

    impl SdtEndPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                "<{node_name}>
            {rpr}
            {rpr}
        </{node_name}>",
                rpr = RPr::test_xml("rPr"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                run_properties_vec: vec![RPr::test_instance(), RPr::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_std_end_pr_from_xml() {
        let xml = SdtEndPr::test_xml("sdtEndPr");
        assert_eq!(
            SdtEndPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SdtEndPr::test_instance(),
        );
    }

    impl SdtContentRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            {pcontent}
            {pcontent}
        </{node_name}>"#,
                pcontent = PContent::test_simple_field_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                p_contents: vec![
                    PContent::test_simple_field_instance(),
                    PContent::test_simple_field_instance(),
                ],
            }
        }
    }

    #[test]
    pub fn test_sdt_content_run_from_xml() {
        let xml = SdtContentRun::test_xml("sdtContentRun");
        assert_eq!(
            SdtContentRun::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SdtContentRun::test_instance()
        );
    }

    impl SdtRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
            {}
            {}
            {}
        </{node_name}>"#,
                SdtPr::test_xml("sdtPr"),
                SdtEndPr::test_xml("sdtEndPr"),
                SdtContentRun::test_xml("sdtContent"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                sdt_properties: Some(SdtPr::test_instance()),
                sdt_end_properties: Some(SdtEndPr::test_instance()),
                sdt_content: Some(SdtContentRun::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_sdt_run_from_xml() {
        let xml = SdtRun::test_xml("sdtRun");
        assert_eq!(
            SdtRun::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SdtRun::test_instance()
        );
    }

    impl DirContentRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} val="ltr">
                {}
            </{node_name}>"#,
                PContent::test_simple_field_xml(),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                p_contents: vec![PContent::test_simple_field_instance()],
                value: Some(Direction::LeftToRight),
            }
        }
    }

    #[test]
    pub fn test_dir_content_run_from_xml() {
        let xml = DirContentRun::test_xml("dirContentRun");
        assert_eq!(
            DirContentRun::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            DirContentRun::test_instance()
        );
    }

    impl BdoContentRun {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} val="ltr">
                {}
            </{node_name}>"#,
                PContent::test_simple_field_xml(),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                p_contents: vec![PContent::test_simple_field_instance()],
                value: Some(Direction::LeftToRight),
            }
        }
    }

    #[test]
    pub fn test_bdo_content_run_from_xml() {
        let xml = DirContentRun::test_xml("bdoContentRun");
        assert_eq!(
            DirContentRun::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            DirContentRun::test_instance()
        );
    }

    impl Br {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} type="page" clear="none"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                break_type: Some(BrType::Page),
                clear: Some(BrClear::None),
            }
        }
    }

    #[test]
    pub fn test_br_from_xml() {
        let xml = Br::test_xml("br");
        assert_eq!(
            Br::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Br::test_instance()
        );
    }

    impl Text {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} xml:space="default">Some text</{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                text: String::from("Some text"),
                xml_space: Some(String::from("default")),
            }
        }
    }

    #[test]
    pub fn test_text_from_xml() {
        let xml = Text::test_xml("text");
        assert_eq!(
            Text::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Text::test_instance()
        );
    }

    impl Sym {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} font="Arial" char="ffff"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                font: Some(String::from("Arial")),
                character: Some(0xffff),
            }
        }
    }

    #[test]
    pub fn test_sym_from_xml() {
        let xml = Sym::test_xml("sym");
        assert_eq!(
            Sym::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Sym::test_instance()
        );
    }

    impl Control {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} name="Name" shapeid="Id" r:id="rId1" >
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                name: Some(String::from("Name")),
                shapeid: Some(String::from("Id")),
                rel_id: Some(String::from("rId1")),
            }
        }
    }

    #[test]
    pub fn test_control_from_xml() {
        let xml = Control::test_xml("control");
        assert_eq!(
            Control::from_xml_element(&XmlNode::from_str(xml).unwrap()),
            Control::test_instance()
        );
    }

    impl ObjectEmbed {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} drawAspect="content" r:id="rId1" progId="AVIFile" shapeId="1" fieldCodes="\f 0">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                draw_aspect: Some(ObjectDrawAspect::Content),
                rel_id: String::from("rId1"),
                application_id: Some(String::from("AVIFile")),
                shape_id: Some(String::from("1")),
                field_codes: Some(String::from(r#"\f 0"#)),
            }
        }
    }

    #[test]
    pub fn test_object_embed_from_xml() {
        let xml = ObjectEmbed::test_xml("objectEmbed");
        assert_eq!(
            ObjectEmbed::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            ObjectEmbed::test_instance()
        );
    }

    impl ObjectLink {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} drawAspect="content" r:id="rId1" progId="AVIFile" shapeId="1" fieldCodes="\f 0" updateMode="always" lockedField="true">
            </{node_name}>"#,
                node_name=node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: ObjectEmbed::test_instance(),
                update_mode: ObjectUpdateMode::Always,
                locked_field: Some(true),
            }
        }
    }

    #[test]
    pub fn test_object_link_from_xml() {
        let xml = ObjectLink::test_xml("objectLink");
        assert_eq!(
            ObjectLink::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            ObjectLink::test_instance()
        );
    }

    impl Drawing {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                Anchor::test_xml("wp:anchor"),
                Inline::test_xml("wp:inline"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                anchor_or_inline_vec: vec![
                    DrawingChoice::Anchor(Anchor::test_instance()),
                    DrawingChoice::Inline(Inline::test_instance()),
                ],
            }
        }
    }

    #[test]
    pub fn test_drawing_from_xml() {
        let xml = Drawing::test_xml("drawing");
        assert_eq!(
            Drawing::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Drawing::test_instance()
        );
    }

    impl Object {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} dxaOrig="123.456mm" dyaOrig="123">
                {}
                {}
            </{node_name}>"#,
                Drawing::test_xml("drawing"),
                Control::test_xml("control"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                drawing: Some(Drawing::test_instance()),
                choice: Some(ObjectChoice::Control(Control::test_instance())),
                original_image_width: Some(TwipsMeasure::UniversalMeasure(UniversalMeasure::new(
                    123.456,
                    UniversalMeasureUnit::Millimeter,
                ))),
                original_image_height: Some(TwipsMeasure::Decimal(123)),
            }
        }
    }

    #[test]
    pub fn test_object_from_xml() {
        let xml = Object::test_xml("object");
        assert_eq!(
            Object::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Object::test_instance()
        );
    }

    impl FFHelpText {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} type="text" val="Help text"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                info_text_type: Some(InfoTextType::Text),
                value: Some(FFHelpTextVal::from("Help text")),
            }
        }
    }

    #[test]
    pub fn test_ff_help_text_from_xml() {
        let xml = FFHelpText::test_xml("ffHelpText");
        assert_eq!(
            FFHelpText::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FFHelpText::test_instance()
        );
    }

    impl FFStatusText {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} type="text" val="Status text"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                info_text_type: Some(InfoTextType::Text),
                value: Some(FFStatusTextVal::from("Status text")),
            }
        }
    }

    #[test]
    pub fn test_ff_status_text_from_xml() {
        let xml = FFStatusText::test_xml("ffStatusText");
        assert_eq!(
            FFStatusText::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FFStatusText::test_instance()
        );
    }

    #[test]
    pub fn test_ff_check_box_size_choice_from_xml() {
        let xml = r#"<size val="123"></size>"#;
        assert_eq!(
            FFCheckBoxSizeChoice::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FFCheckBoxSizeChoice::Explicit(HpsMeasure::Decimal(123)),
        );
        let xml = r#"<sizeAuto val="true"></sizeAuto>"#;
        assert_eq!(
            FFCheckBoxSizeChoice::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FFCheckBoxSizeChoice::Auto(Some(true)),
        );
    }

    impl FFCheckBox {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <sizeAuto val="true" />
                <default val="true" />
                <checked val="true" />
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                size: FFCheckBoxSizeChoice::Auto(Some(true)),
                is_default: Some(true),
                is_checked: Some(true),
            }
        }
    }

    #[test]
    pub fn test_ff_check_box_from_xml() {
        let xml = FFCheckBox::test_xml("ffCheckBox");
        assert_eq!(
            FFCheckBox::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FFCheckBox::test_instance()
        );
    }

    impl FFDDList {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <result val="1" />
                <default val="1" />
                <listEntry val="Entry1" />
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                result: Some(1),
                default: Some(1),
                list_entries: vec![String::from("Entry1")],
            }
        }
    }

    #[test]
    pub fn test_ff_ddlist_from_xml() {
        let xml = FFDDList::test_xml("ffDDList");
        assert_eq!(
            FFDDList::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FFDDList::test_instance()
        );
    }

    impl FFTextInput {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <type val="regular" />
                <default val="Default" />
                <maxLength val="100" />
                <format val=".*" />
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                text_type: Some(FFTextType::Regular),
                default: Some(String::from("Default")),
                max_length: Some(100),
                format: Some(String::from(".*")),
            }
        }
    }

    #[test]
    pub fn test_ff_text_input_from_xml() {
        let xml = FFTextInput::test_xml("ffTextInput");
        assert_eq!(
            FFTextInput::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FFTextInput::test_instance(),
        );
    }

    impl FldChar {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} fldCharType="begin", fldLock="false" dirty="false">
                <name val="Some name" />
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                form_field_properties: Some(FFData::Name(FFName::from("Some name"))),
                field_char_type: FldCharType::Begin,
                field_lock: Some(false),
                dirty: Some(false),
            }
        }
    }

    #[test]
    pub fn test_fld_char_from_xml() {
        let xml = FldChar::test_xml("fldChar");
        assert_eq!(
            FldChar::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FldChar::test_instance()
        );
    }

    impl RubyPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <rubyAlign val="left" />
                <hps val="123" />
                <hpsRaise val="123" />
                <hpsBaseText val="123" />
                <lid val="en-US" />
                <dirty val="true" />
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                ruby_align: RubyAlign::Left,
                hps: HpsMeasure::Decimal(123),
                hps_raise: HpsMeasure::Decimal(123),
                hps_base_text: HpsMeasure::Decimal(123),
                language_id: Lang::from("en-US"),
                dirty: Some(true),
            }
        }
    }

    #[test]
    pub fn test_ruby_pr_from_xml() {
        let xml = RubyPr::test_xml("rubyPr");
        assert_eq!(
            RubyPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            RubyPr::test_instance()
        );
    }

    #[test]
    pub fn test_ruby_content_choice_from_xml() {
        let xml = R::test_xml("r");
        assert_eq!(
            RubyContentChoice::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            RubyContentChoice::Run(R::test_instance())
        );
        let xml = ProofErr::test_xml("proofErr");
        assert_eq!(
            RubyContentChoice::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            RubyContentChoice::RunLevelElement(RunLevelElts::ProofError(ProofErr::test_instance()))
        );
    }

    impl RubyContent {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
            </{node_name}>"#,
                R::test_xml("r"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                ruby_contents: vec![RubyContentChoice::Run(R::test_instance())],
            }
        }
    }

    #[test]
    pub fn test_ruby_content_from_xml() {
        let xml = RubyContent::test_xml("rubyContent");
        assert_eq!(
            RubyContent::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            RubyContent::test_instance()
        );
    }

    impl R {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} rsidRPr="ffffffff" rsidDel="ffffffff" rsidR="ffffffff">
                {}
                {}
            </{node_name}>"#,
                RPr::test_xml("rPr"),
                Br::test_xml("br"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                run_properties: Some(RPr::test_instance()),
                run_inner_contents: vec![RunInnerContent::Break(Br::test_instance())],
                run_properties_revision_id: Some(0xffffffff),
                deletion_revision_id: Some(0xffffffff),
                run_revision_id: Some(0xffffffff),
            }
        }
    }

    #[test]
    pub fn test_r_from_xml() {
        let xml = R::test_xml("r");
        assert_eq!(
            R::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            R::test_instance()
        );
    }

    impl Ruby {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
            </{node_name}>"#,
                RubyPr::test_xml("rubyPr"),
                RubyContent::test_xml("rt"),
                RubyContent::test_xml("rubyBase"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                ruby_properties: RubyPr::test_instance(),
                ruby_content: RubyContent::test_instance(),
                ruby_base: RubyContent::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_ruby_from_xml() {
        let xml = Ruby::test_xml("ruby");
        assert_eq!(
            Ruby::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Ruby::test_instance()
        );
    }

    impl FtnEdnRef {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} customMarkFollows="true" id="1"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                custom_mark_follows: Some(true),
                id: 1,
            }
        }
    }

    #[test]
    pub fn test_ftn_edn_ref_from_xml() {
        let xml = FtnEdnRef::test_xml("ftnEdnRef");
        assert_eq!(
            FtnEdnRef::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FtnEdnRef::test_instance()
        );
    }

    impl PTab {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} alignment="left" relativeTo="margin" leader="none">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                alignment: PTabAlignment::Left,
                relative_to: PTabRelativeTo::Margin,
                leader: PTabLeader::None,
            }
        }
    }

    #[test]
    pub fn test_p_tab_from_xml() {
        let xml = PTab::test_xml("pTab");
        assert_eq!(
            PTab::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PTab::test_instance()
        );
    }

    impl RunTrackChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} id="0" author="John Smith" date="2001-10-26T21:32:52">
                {}
            </{node_name}>"#,
                R::test_xml("r"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                choices: vec![RunTrackChangeChoice::ContentRunContent(ContentRunContent::Run(
                    R::test_instance(),
                ))],
            }
        }
    }

    #[test]
    pub fn test_run_track_change_from_xml() {
        let xml = RunTrackChange::test_xml("runTrackChange");
        assert_eq!(
            RunTrackChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            RunTrackChange::test_instance()
        );
    }

    impl CustomXmlBlock {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} uri="https://some/uri" element="Some element">
                {}
            </{node_name}>"#,
                CustomXmlPr::test_xml("customXmlPr"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                custom_xml_properties: Some(CustomXmlPr::test_instance()),
                block_contents: Vec::new(),
                uri: Some(String::from("https://some/uri")),
                element: XmlName::from("Some element"),
            }
        }
    }

    #[test]
    pub fn test_custom_xml_block_from_xml() {
        let xml = CustomXmlBlock::test_xml("customXmlBlock");
        assert_eq!(
            CustomXmlBlock::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            CustomXmlBlock::test_instance()
        );
    }

    impl SdtContentBlock {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                "<{node_name}>
                {}
            </{node_name}>",
                CustomXmlBlock::test_xml("customXml"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                block_contents: vec![ContentBlockContent::CustomXml(CustomXmlBlock::test_instance())],
            }
        }
    }

    #[test]
    pub fn test_sdt_content_block_from_xml() {
        let xml = SdtContentBlock::test_xml("sdtContentBlock");
        assert_eq!(
            SdtContentBlock::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SdtContentBlock::test_instance()
        );
    }

    impl SdtBlock {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
            </{node_name}>"#,
                SdtPr::test_xml("sdtPr"),
                SdtEndPr::test_xml("sdtEndPr"),
                SdtContentBlock::test_xml("sdtContent"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                sdt_properties: Some(SdtPr::test_instance()),
                sdt_end_properties: Some(SdtEndPr::test_instance()),
                sdt_content: Some(SdtContentBlock::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_sdt_block_from_xml() {
        let xml = SdtBlock::test_xml("sdtBlock");
        assert_eq!(
            SdtBlock::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SdtBlock::test_instance()
        );
    }

    impl FramePr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} dropCap="drop" lines="1" w="100" h="100" vSpace="50" hSpace="50" wrap="auto"
                hAnchor="text" vAnchor="text" x="0" xAlign="left" y="0" yAlign="top" hRule="auto" anchorLock="true">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                drop_cap: Some(DropCap::Drop),
                lines: Some(1),
                width: Some(TwipsMeasure::Decimal(100)),
                height: Some(TwipsMeasure::Decimal(100)),
                vertical_space: Some(TwipsMeasure::Decimal(50)),
                horizontal_space: Some(TwipsMeasure::Decimal(50)),
                wrap: Some(Wrap::Auto),
                horizontal_anchor: Some(HAnchor::Text),
                vertical_anchor: Some(VAnchor::Text),
                x: Some(SignedTwipsMeasure::Decimal(0)),
                x_align: Some(XAlign::Left),
                y: Some(SignedTwipsMeasure::Decimal(0)),
                y_align: Some(YAlign::Top),
                height_rule: Some(HeightRule::Auto),
                anchor_lock: Some(true),
            }
        }
    }

    #[test]
    pub fn test_frame_pr_from_xml() {
        let xml = FramePr::test_xml("framePr");
        assert_eq!(
            FramePr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FramePr::test_instance()
        );
    }

    impl NumPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <ilvl val="1" />
                <numId val="1" />
                {}
            </{node_name}>"#,
                TrackChange::test_xml("ins"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                indent_level: Some(1),
                numbering_id: Some(1),
                inserted: Some(TrackChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_num_pr_from_xml() {
        let xml = NumPr::test_xml("numPr");
        assert_eq!(
            NumPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            NumPr::test_instance(),
        );
    }

    impl PBdr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                Border::test_xml("top"),
                Border::test_xml("left"),
                Border::test_xml("bottom"),
                Border::test_xml("right"),
                Border::test_xml("between"),
                Border::test_xml("bar"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(Border::test_instance()),
                left: Some(Border::test_instance()),
                bottom: Some(Border::test_instance()),
                right: Some(Border::test_instance()),
                between: Some(Border::test_instance()),
                bar: Some(Border::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_p_bdr_from_xml() {
        let xml = PBdr::test_xml("pBdr");
        assert_eq!(
            PBdr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PBdr::test_instance(),
        );
    }

    impl TabStop {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} val="start" leader="dot" pos="0"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: TabJc::Start,
                leader: Some(TabTlc::Dot),
                position: SignedTwipsMeasure::Decimal(0),
            }
        }
    }

    #[test]
    pub fn test_tab_stop_from_xml() {
        let xml = TabStop::test_xml("tabStop");
        assert_eq!(
            TabStop::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TabStop::test_instance(),
        );
    }

    impl Tabs {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                "<{node_name}>
                {tab_stop}
                {tab_stop}
            </{node_name}>",
                tab_stop = TabStop::test_xml("tab"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                tabs: vec![TabStop::test_instance(), TabStop::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_tabs_from_xml() {
        let xml = Tabs::test_xml("tabs");
        assert_eq!(
            Tabs::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Tabs::test_instance(),
        );
    }

    impl Spacing {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} before="10" beforeLines="1" beforeAutospacing="true"
                after="10" afterLines="1" afterAutospacing="true" line="50" lineRule="auto">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                before: Some(TwipsMeasure::Decimal(10)),
                before_lines: Some(1),
                before_autospacing: Some(true),
                after: Some(TwipsMeasure::Decimal(10)),
                after_lines: Some(1),
                after_autospacing: Some(true),
                line: Some(SignedTwipsMeasure::Decimal(50)),
                line_rule: Some(LineSpacingRule::Auto),
            }
        }
    }

    #[test]
    pub fn test_spacing_from_xml() {
        let xml = Spacing::test_xml("spacing");
        assert_eq!(
            Spacing::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Spacing::test_instance(),
        );
    }

    impl Ind {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} start="50" startChars="0" end="50" endChars="10" hanging="50" hangingChars="5"
                firstLine="50" firstLineChars="5">
            </{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                start: Some(SignedTwipsMeasure::Decimal(50)),
                start_chars: Some(0),
                end: Some(SignedTwipsMeasure::Decimal(50)),
                end_chars: Some(10),
                hanging: Some(TwipsMeasure::Decimal(50)),
                hanging_chars: Some(5),
                first_line: Some(TwipsMeasure::Decimal(50)),
                first_line_chars: Some(5),
            }
        }
    }

    #[test]
    pub fn test_ind_from_xml() {
        let xml = Ind::test_xml("ind");
        assert_eq!(
            Ind::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Ind::test_instance(),
        );
    }

    impl Cnf {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} firstRow="true" lastRow="true" firstColumn="true" lastColumn="true" oddVBand="true"
                evenVBand="true" oddHBand="true" evenHBand="true" firstRowFirstColumn="true" firstRowLastColumn="true"
                lastRowFirstColumn="true" lastRowLastColumn="true">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                first_row: Some(true),
                last_row: Some(true),
                first_column: Some(true),
                last_column: Some(true),
                odd_vertical_band: Some(true),
                even_vertical_band: Some(true),
                odd_horizontal_band: Some(true),
                even_horizontal_band: Some(true),
                first_row_first_column: Some(true),
                first_row_last_column: Some(true),
                last_row_first_column: Some(true),
                last_row_last_column: Some(true),
            }
        }
    }

    #[test]
    pub fn test_cnf_from_xml() {
        let xml = Cnf::test_xml("cnf");
        assert_eq!(
            Cnf::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Cnf::test_instance(),
        );
    }

    impl PPrBase {
        pub fn test_xml_nodes() -> String {
            format!(
                r#"<pStyle val="Normal" />
                <keepNext val="true" />
                <keepLines val="true" />
                <pageBreakBefore val="true" />
                {}
                <widowControl val="true" />
                {}
                <suppressLineNumbers val="true" />
                {}
                {}
                {}
                <suppressAutoHyphens val="true" />
                <kinsoku val="true" />
                <wordWrap val="true" />
                <overflowPunct val="true" />
                <topLinePunct val="true" />
                <autoSpaceDE val="true" />
                <autoSpaceDN val="true" />
                <bidi val="true" />
                <adjustRightInd val="true" />
                <snapToGrid val="true" />
                {}
                {}
                <contextualSpacing val="true" />
                <mirrorIndents val="true" />
                <suppressOverlap val="true" />
                <jc val="start" />
                <textDirection val="lr" />
                <textAlignment val="auto" />
                <textboxTightWrap val="none" />
                <outlineLvl val="1" />
                <divId val="1" />
                {}"#,
                FramePr::test_xml("framePr"),
                NumPr::test_xml("numPr"),
                PBdr::test_xml("pBdr"),
                Shd::test_xml("shd"),
                Tabs::test_xml("tabs"),
                Spacing::test_xml("spacing"),
                Ind::test_xml("ind"),
                Cnf::test_xml("cnfStyle"),
            )
        }

        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
            </{node_name}>"#,
                Self::test_xml_nodes(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                style: Some(String::from("Normal")),
                keep_with_next: Some(true),
                keep_lines_on_one_page: Some(true),
                start_on_next_page: Some(true),
                frame_properties: Some(FramePr::test_instance()),
                widow_control: Some(true),
                numbering_properties: Some(NumPr::test_instance()),
                suppress_line_numbers: Some(true),
                borders: Some(PBdr::test_instance()),
                shading: Some(Shd::test_instance()),
                tabs: Some(Tabs::test_instance()),
                suppress_auto_hyphens: Some(true),
                kinsoku: Some(true),
                word_wrapping: Some(true),
                overflow_punctuations: Some(true),
                top_line_punctuations: Some(true),
                auto_space_latin_and_east_asian: Some(true),
                auto_space_east_asian_and_numbers: Some(true),
                bidirectional: Some(true),
                adjust_right_indent: Some(true),
                snap_to_grid: Some(true),
                spacing: Some(Spacing::test_instance()),
                indent: Some(Ind::test_instance()),
                contextual_spacing: Some(true),
                mirror_indents: Some(true),
                suppress_overlapping: Some(true),
                alignment: Some(Jc::Start),
                text_direction: Some(TextDirection::LeftToRight),
                text_alignment: Some(TextAlignment::Auto),
                textbox_tight_wrap: Some(TextboxTightWrap::None),
                outline_level: Some(1),
                div_id: Some(1),
                conditional_formatting: Some(Cnf::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_p_pr_base_from_xml() {
        let xml = PPrBase::test_xml("pPrBase");
        assert_eq!(
            PPrBase::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PPrBase::test_instance(),
        );
    }

    impl ParaRPrTrackChanges {
        pub fn test_xml() -> String {
            format!(
                r#"{}
                {}
                {}
                {}
            "#,
                TrackChange::test_xml("ins"),
                TrackChange::test_xml("del"),
                TrackChange::test_xml("moveFrom"),
                TrackChange::test_xml("moveTo"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                inserted: Some(TrackChange::test_instance()),
                deleted: Some(TrackChange::test_instance()),
                move_from: Some(TrackChange::test_instance()),
                move_to: Some(TrackChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_para_r_pr_track_changes_from_xml() {
        let xml = format!("<node>{}</node>", ParaRPrTrackChanges::test_xml());
        assert_eq!(
            ParaRPrTrackChanges::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Some(ParaRPrTrackChanges::test_instance()),
        );
    }

    impl ParaRPrOriginal {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                ParaRPrTrackChanges::test_xml(),
                RPrBase::test_run_style_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                track_changes: Some(ParaRPrTrackChanges::test_instance()),
                bases: vec![RPrBase::test_run_style_instance()],
            }
        }
    }

    #[test]
    pub fn test_para_r_pr_original_from_xml() {
        let xml = ParaRPrOriginal::test_xml("paraRPrOriginal");
        assert_eq!(
            ParaRPrOriginal::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            ParaRPrOriginal::test_instance(),
        );
    }

    impl ParaRPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} id="0" author="John Smith" date="2001-10-26T21:32:52">
                {}
            </{node_name}>"#,
                ParaRPrOriginal::test_xml("rPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                run_properties: ParaRPrOriginal::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_para_r_change_from_xml() {
        let xml = ParaRPrChange::test_xml("paraRPrChange");
        assert_eq!(
            ParaRPrChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            ParaRPrChange::test_instance(),
        );
    }

    impl ParaRPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
            </{node_name}>"#,
                ParaRPrTrackChanges::test_xml(),
                RPrBase::test_run_style_xml(),
                ParaRPrChange::test_xml("rPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                track_changes: Some(ParaRPrTrackChanges::test_instance()),
                bases: vec![RPrBase::test_run_style_instance()],
                change: Some(ParaRPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_para_r_pr_from_xml() {
        let xml = ParaRPr::test_xml("paraRPr");
        assert_eq!(
            ParaRPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            ParaRPr::test_instance(),
        );
    }

    impl HdrFtrRef {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} r:id="rId1" type="default"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: Rel::test_instance(),
                header_footer_type: HdrFtr::Default,
            }
        }
    }

    #[test]
    pub fn test_hdr_ftr_ref_from_xml() {
        let xml = HdrFtrRef::test_xml("hdrFtrRef");
        assert_eq!(
            HdrFtrRef::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            HdrFtrRef::test_instance(),
        );
    }

    impl NumFmt {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} val="decimal" format="&#x30A2;"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: NumberFormat::Decimal,
                format: Some(String::from("&#x30A2;")),
            }
        }
    }

    #[test]
    pub fn test_num_fmt_from_xml() {
        let xml = NumFmt::test_xml("numFmt");
        assert_eq!(
            NumFmt::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            NumFmt::test_instance(),
        );
    }

    impl FtnEdnNumProps {
        pub fn test_xml() -> String {
            format!(
                r#"<numStart val="1" />
            <numRestart val="continuous" />
            "#
            )
        }

        pub fn test_instance() -> Self {
            Self {
                numbering_start: Some(1),
                numbering_restart: Some(RestartNumber::Continuous),
            }
        }
    }

    #[test]
    pub fn test_ftn_edn_num_props_from_xml() {
        let xml = format!("<node>{}</node>", FtnEdnNumProps::test_xml());
        assert_eq!(
            FtnEdnNumProps::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Some(FtnEdnNumProps::test_instance()),
        );
    }

    impl FtnProps {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <pos val="pageBottom" />
                {}
                {}
            </{node_name}>"#,
                NumFmt::test_xml("numFmt"),
                FtnEdnNumProps::test_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                position: Some(FtnPos::PageBottom),
                numbering_format: Some(NumFmt::test_instance()),
                numbering_properties: Some(FtnEdnNumProps::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_ftn_props_from_xml() {
        let xml = FtnProps::test_xml("ftnProps");
        assert_eq!(
            FtnProps::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            FtnProps::test_instance(),
        );
    }

    impl EdnProps {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                <pos val="docEnd" />
                {}
                {}
            </{node_name}>"#,
                NumFmt::test_xml("numFmt"),
                FtnEdnNumProps::test_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                position: Some(EdnPos::DocumentEnd),
                numbering_format: Some(NumFmt::test_instance()),
                numbering_properties: Some(FtnEdnNumProps::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_edn_props_from_xml() {
        let xml = EdnProps::test_xml("endProps");
        assert_eq!(
            EdnProps::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            EdnProps::test_instance(),
        );
    }

    impl PageSz {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w="100" h="100" orient="portrait" code="1"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(TwipsMeasure::Decimal(100)),
                height: Some(TwipsMeasure::Decimal(100)),
                orientation: Some(PageOrientation::Portrait),
                code: Some(1),
            }
        }
    }

    #[test]
    pub fn test_page_sz_from_xml() {
        let xml = PageSz::test_xml("pageSz");
        assert_eq!(
            PageSz::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PageSz::test_instance(),
        );
    }

    impl PageMar {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} top="10" right="10" bottom="10" left="10" header="10" footer="10" gutter="10">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: SignedTwipsMeasure::Decimal(10),
                right: TwipsMeasure::Decimal(10),
                bottom: SignedTwipsMeasure::Decimal(10),
                left: TwipsMeasure::Decimal(10),
                header: TwipsMeasure::Decimal(10),
                footer: TwipsMeasure::Decimal(10),
                gutter: TwipsMeasure::Decimal(10),
            }
        }
    }

    #[test]
    pub fn test_page_mar_from_xml() {
        let xml = PageMar::test_xml("pageMar");
        assert_eq!(
            PageMar::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PageMar::test_instance(),
        );
    }

    impl PaperSource {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} first="1" other="1"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                first: Some(1),
                other: Some(1),
            }
        }
    }

    #[test]
    pub fn test_paper_source_from_xml() {
        let xml = PaperSource::test_xml("paperSource");
        assert_eq!(
            PaperSource::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PaperSource::test_instance(),
        );
    }

    impl PageBorder {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {} r:id="rId1"></{node_name}>"#,
                Border::TEST_ATTRIBUTES,
                node_name = node_name
            )
        }

        pub fn test_attributes() -> String {
            format!(r#"{} r:id="rId1""#, Border::TEST_ATTRIBUTES)
        }

        pub fn test_instance() -> Self {
            Self {
                base: Border::test_instance(),
                rel_id: Some(RelationshipId::from("rId1")),
            }
        }
    }

    #[test]
    pub fn test_page_border_from_xml() {
        let xml = PageBorder::test_xml("pageBorder");
        assert_eq!(
            PageBorder::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PageBorder::test_instance(),
        );
    }

    impl TopPageBorder {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {} r:topLeft="rId2" r:topRight="rId3"></{node_name}>"#,
                PageBorder::test_attributes(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: PageBorder::test_instance(),
                top_left: Some(RelationshipId::from("rId2")),
                top_right: Some(RelationshipId::from("rId3")),
            }
        }
    }

    #[test]
    pub fn test_top_page_border_from_xml() {
        let xml = TopPageBorder::test_xml("topPageBorder");
        assert_eq!(
            TopPageBorder::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TopPageBorder::test_instance(),
        );
    }

    impl BottomPageBorder {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {} r:bottomLeft="rId2" r:bottomRight="rId3"></{node_name}>"#,
                PageBorder::test_attributes(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: PageBorder::test_instance(),
                bottom_left: Some(RelationshipId::from("rId2")),
                bottom_right: Some(RelationshipId::from("rId3")),
            }
        }
    }

    #[test]
    pub fn test_bottom_page_border_from_xml() {
        let xml = BottomPageBorder::test_xml("bottomPageBorder");
        assert_eq!(
            BottomPageBorder::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            BottomPageBorder::test_instance(),
        );
    }

    impl PageBorders {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} zOrder="front" display="allPages" offsetFrom="page">
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                TopPageBorder::test_xml("top"),
                PageBorder::test_xml("left"),
                BottomPageBorder::test_xml("bottom"),
                PageBorder::test_xml("right"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(TopPageBorder::test_instance()),
                left: Some(PageBorder::test_instance()),
                bottom: Some(BottomPageBorder::test_instance()),
                right: Some(PageBorder::test_instance()),
                z_order: Some(PageBorderZOrder::Front),
                display: Some(PageBorderDisplay::AllPages),
                offset_from: Some(PageBorderOffset::Page),
            }
        }
    }

    #[test]
    pub fn test_page_borders_from_xml() {
        let xml = PageBorders::test_xml("pageBorders");
        assert_eq!(
            PageBorders::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PageBorders::test_instance(),
        );
    }

    impl LineNumber {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} countBy="1" start="1" distance="100" restart="newPage"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                count_by: Some(1),
                start: Some(1),
                distance: Some(TwipsMeasure::Decimal(100)),
                restart: Some(LineNumberRestart::NewPage),
            }
        }
    }

    #[test]
    pub fn test_line_number_from_xml() {
        let xml = LineNumber::test_xml("lineNumber");
        assert_eq!(
            LineNumber::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            LineNumber::test_instance(),
        );
    }

    impl PageNumber {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} fmt="decimal" start="1" chapStyle="1" chapSep="hyphen"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                format: Some(NumberFormat::Decimal),
                start: Some(1),
                chapter_style: Some(1),
                chapter_separator: Some(ChapterSep::Hyphen),
            }
        }
    }

    #[test]
    pub fn test_page_number_from_xml() {
        let xml = PageNumber::test_xml("pageNumber");
        assert_eq!(
            PageNumber::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PageNumber::test_instance(),
        );
    }

    impl Column {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w="100" space="10"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(TwipsMeasure::Decimal(100)),
                spacing: Some(TwipsMeasure::Decimal(10)),
            }
        }
    }

    #[test]
    pub fn test_column_from_xml() {
        let xml = Column::test_xml("column");
        assert_eq!(
            Column::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Column::test_instance(),
        );
    }

    impl Columns {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} equalWidth="true" space="10" num="2" sep="true">
                {col}
                {col}
            </{node_name}>"#,
                col = Column::test_xml("col"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                columns: vec![Column::test_instance(), Column::test_instance()],
                equal_width: Some(true),
                spacing: Some(TwipsMeasure::Decimal(10)),
                number: Some(2),
                separator: Some(true),
            }
        }
    }

    #[test]
    pub fn test_columns_from_xml() {
        let xml = Columns::test_xml("columns");
        assert_eq!(
            Columns::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Columns::test_instance(),
        );
    }

    impl DocGrid {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} type="default" linePitch="1" charSpace="10"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                doc_grid_type: Some(DocGridType::Default),
                line_pitch: Some(1),
                char_spacing: Some(10),
            }
        }
    }

    #[test]
    pub fn test_doc_grid_from_xml() {
        let xml = DocGrid::test_xml("docGrid");
        assert_eq!(
            DocGrid::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            DocGrid::test_instance(),
        );
    }

    impl SectPrContents {
        pub fn test_xml() -> String {
            format!(
                r#"{}
                {}
                <type val="nextPage" />
                {}
                {}
                {}
                {}
                {}
                {}
                {}
                <formProt val="false" />
                <vAlign val="top" />
                <noEndnote val="false" />
                <titlePg val="true" />
                <textDirection val="lr" />
                <bidi val="false" />
                <rtlGutter val="false" />
                {}
                <printerSettings r:id="rId1" />"#,
                FtnProps::test_xml("footnotePr"),
                EdnProps::test_xml("endnotePr"),
                PageSz::test_xml("pgSz"),
                PageMar::test_xml("pgMar"),
                PaperSource::test_xml("paperSrc"),
                PageBorders::test_xml("pgBorders"),
                LineNumber::test_xml("lnNumType"),
                PageNumber::test_xml("pgNumType"),
                Columns::test_xml("cols"),
                DocGrid::test_xml("docGrid"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                footnote_properties: Some(FtnProps::test_instance()),
                endnote_properties: Some(EdnProps::test_instance()),
                section_type: Some(SectionMark::NextPage),
                page_size: Some(PageSz::test_instance()),
                page_margin: Some(PageMar::test_instance()),
                paper_source: Some(PaperSource::test_instance()),
                page_borders: Some(PageBorders::test_instance()),
                line_number_type: Some(LineNumber::test_instance()),
                page_number_type: Some(PageNumber::test_instance()),
                columns: Some(Columns::test_instance()),
                protect_form_fields: Some(false),
                vertical_align: Some(VerticalJc::Top),
                no_endnote: Some(false),
                title_page: Some(true),
                text_direction: Some(TextDirection::LeftToRight),
                bidirectional: Some(false),
                rtl_gutter: Some(false),
                document_grid: Some(DocGrid::test_instance()),
                printer_settings: Some(Rel::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_contents_from_xml() {
        let xml = format!(r#"<node>{}</node>"#, SectPrContents::test_xml());
        assert_eq!(
            SectPrContents::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Some(SectPrContents::test_instance()),
        );
    }

    impl SectPrAttributes {
        const TEST_ATTRIBUTES: &'static str =
            r#"rsidRPr="ffffffff" rsidDel="fefefefe" rsidR="fdfdfdfd" rsidSect="fcfcfcfc""#;

        pub fn test_instance() -> Self {
            Self {
                run_properties_revision_id: Some(0xffffffff),
                deletion_revision_id: Some(0xfefefefe),
                run_revision_id: Some(0xfdfdfdfd),
                section_revision_id: Some(0xfcfcfcfc),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_attributes_from_xml() {
        let xml = format!(r#"<node {}></node>"#, SectPrAttributes::TEST_ATTRIBUTES);
        assert_eq!(
            SectPrAttributes::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SectPrAttributes::test_instance(),
        );
    }

    impl SectPrBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                SectPrAttributes::TEST_ATTRIBUTES,
                SectPrContents::test_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                contents: Some(SectPrContents::test_instance()),
                attributes: SectPrAttributes::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_base_from_xml() {
        let xml = SectPrBase::test_xml("sectPrBase");
        assert_eq!(
            SectPrBase::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SectPrBase::test_instance(),
        );
    }

    impl SectPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                SectPrBase::test_xml("sectPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                section_properties: Some(SectPrBase::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_change_from_xml() {
        let xml = SectPrChange::test_xml("sectPrChange");
        assert_eq!(
            SectPrChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SectPrChange::test_instance(),
        );
    }

    impl SectPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {header_ref}
                {footer_ref}
                {}
                {}
            </{node_name}>"#,
                SectPrAttributes::TEST_ATTRIBUTES,
                SectPrContents::test_xml(),
                SectPrChange::test_xml("sectPrChange"),
                node_name = node_name,
                header_ref = HdrFtrRef::test_xml("headerReference"),
                footer_ref = HdrFtrRef::test_xml("footerReference"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                header_footer_references: vec![
                    HdrFtrReferences::Header(HdrFtrRef::test_instance()),
                    HdrFtrReferences::Footer(HdrFtrRef::test_instance()),
                ],
                contents: Some(SectPrContents::test_instance()),
                change: Some(SectPrChange::test_instance()),
                attributes: SectPrAttributes::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_sect_pr_from_xml() {
        let xml = SectPr::test_xml("sectPr");
        assert_eq!(
            SectPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            SectPr::test_instance(),
        );
    }

    impl PPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                PPrBase::test_xml("pPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                properties: PPrBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_p_pr_change_from_xml() {
        let xml = PPrChange::test_xml("pPrChange");
        assert_eq!(
            PPrChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PPrChange::test_instance(),
        );
    }

    impl PPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                PPrBase::test_xml_nodes(),
                ParaRPr::test_xml("rPr"),
                SectPr::test_xml("sectPr"),
                PPrChange::test_xml("pPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: PPrBase::test_instance(),
                run_properties: Some(ParaRPr::test_instance()),
                section_properties: Some(SectPr::test_instance()),
                properties_change: Some(PPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_p_pr_from_xml() {
        let xml = PPr::test_xml("pPr");
        assert_eq!(
            PPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            PPr::test_instance(),
        );
    }

    impl P {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} rsidRPr="ffffffff" rsidR="fefefefe" rsidDel="fdfdfdfd" rsidP="fcfcfcfc" rsidRDefault="fbfbfbfb">
                {}
                {}
            </{node_name}>"#,
                PPr::test_xml("pPr"),
                PContent::test_simple_field_xml(),
                node_name=node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                properties: Some(PPr::test_instance()),
                contents: vec![PContent::test_simple_field_instance()],
                run_properties_revision_id: Some(0xffffffff),
                run_revision_id: Some(0xfefefefe),
                deletion_revision_id: Some(0xfdfdfdfd),
                paragraph_revision_id: Some(0xfcfcfcfc),
                run_default_revision_id: Some(0xfbfbfbfb),
            }
        }
    }

    #[test]
    pub fn test_p_from_xml() {
        let xml = P::test_xml("p");
        assert_eq!(
            P::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            P::test_instance(),
        );
    }

    impl TblPPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} leftFromText="10" rightFromText="10" topFromText="10" bottomFromText="10"
                vertAnchor="text" horzAnchor="text" tblpXSpec="left" tblpX="10" tblpYSpec="top" tblpY="10">
            </{node_name}>"#,
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                left_from_text: Some(TwipsMeasure::Decimal(10)),
                right_from_text: Some(TwipsMeasure::Decimal(10)),
                top_from_text: Some(TwipsMeasure::Decimal(10)),
                bottom_from_text: Some(TwipsMeasure::Decimal(10)),
                vertical_anchor: Some(VAnchor::Text),
                horizontal_anchor: Some(HAnchor::Text),
                horizontal_alignment: Some(XAlign::Left),
                horizontal_distance: Some(SignedTwipsMeasure::Decimal(10)),
                vertical_alignment: Some(YAlign::Top),
                vertical_distance: Some(SignedTwipsMeasure::Decimal(10)),
            }
        }
    }

    #[test]
    pub fn test_tbl_p_pr_from_xml() {
        let xml = TblPPr::test_xml("tblPPr");
        assert_eq!(
            TblPPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblPPr::test_instance(),
        );
    }

    #[test]
    pub fn test_decimal_number_or_percent_from_str() {
        assert_eq!(
            "123".parse::<DecimalNumberOrPercent>().unwrap(),
            DecimalNumberOrPercent::Decimal(123),
        );
        assert_eq!(
            "-123".parse::<DecimalNumberOrPercent>().unwrap(),
            DecimalNumberOrPercent::Decimal(-123),
        );
        assert_eq!(
            "123%".parse::<DecimalNumberOrPercent>().unwrap(),
            DecimalNumberOrPercent::Percentage(Percentage(123.0)),
        );
        assert_eq!(
            "-123%".parse::<DecimalNumberOrPercent>().unwrap(),
            DecimalNumberOrPercent::Percentage(Percentage(-123.0)),
        );

        match DecimalNumberOrPercent::Percentage(Percentage(100.0)) {
            DecimalNumberOrPercent::Percentage(Percentage(value)) => println!("{}", value),
            _ => (),
        }
    }

    #[test]
    pub fn test_measurement_or_percent_from_str() {
        assert_eq!(
            "123".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Decimal(123)),
        );
        assert_eq!(
            "-123".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Decimal(-123)),
        );
        assert_eq!(
            "123%".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Percentage(Percentage(123.0))),
        );
        assert_eq!(
            "-123%".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Percentage(Percentage(-123.0))),
        );
        assert_eq!(
            "123mm".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::UniversalMeasure(UniversalMeasure::new(123.0, UniversalMeasureUnit::Millimeter)),
        );
        assert_eq!(
            "-123mm".parse::<MeasurementOrPercent>().unwrap(),
            MeasurementOrPercent::UniversalMeasure(UniversalMeasure::new(-123.0, UniversalMeasureUnit::Millimeter)),
        );
    }

    impl TblWidth {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} w="100" type="auto"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(MeasurementOrPercent::DecimalOrPercent(DecimalNumberOrPercent::Decimal(
                    100,
                ))),
                width_type: Some(TblWidthType::Auto),
            }
        }
    }

    #[test]
    pub fn test_tbl_width_from_xml() {
        let xml = TblWidth::test_xml("tblWidth");
        assert_eq!(
            TblWidth::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblWidth::test_instance(),
        );
    }

    impl TblBorders {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                Border::test_xml("top"),
                Border::test_xml("start"),
                Border::test_xml("bottom"),
                Border::test_xml("end"),
                Border::test_xml("insideH"),
                Border::test_xml("insideV"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(Border::test_instance()),
                start: Some(Border::test_instance()),
                bottom: Some(Border::test_instance()),
                end: Some(Border::test_instance()),
                inside_horizontal: Some(Border::test_instance()),
                inside_vertical: Some(Border::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_borders_from_xml() {
        let xml = TblBorders::test_xml("tblBorders");
        assert_eq!(
            TblBorders::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblBorders::test_instance(),
        );
    }

    impl TblCellMar {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                TblWidth::test_xml("top"),
                TblWidth::test_xml("start"),
                TblWidth::test_xml("bottom"),
                TblWidth::test_xml("end"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(TblWidth::test_instance()),
                start: Some(TblWidth::test_instance()),
                bottom: Some(TblWidth::test_instance()),
                end: Some(TblWidth::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_cell_mar_from_xml() {
        let xml = TblCellMar::test_xml("tblCellMar");
        assert_eq!(
            TblCellMar::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblCellMar::test_instance(),
        );
    }

    impl TblLook {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} firstRow="true" lastRow="true" firstColumn="true" lastColumn="true" noHBand="true" noVBand="true">
            </{node_name}>"#,
                node_name=node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                first_row: Some(true),
                last_row: Some(true),
                first_column: Some(true),
                last_column: Some(true),
                no_horizontal_band: Some(true),
                no_vertical_band: Some(true),
            }
        }
    }

    #[test]
    pub fn test_tbl_look_from_xml() {
        let xml = TblLook::test_xml("tblLook");
        assert_eq!(
            TblLook::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblLook::test_instance(),
        );
    }

    impl TblPrBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            format!(
                r#"<tblStyle val="Normal" />
                {}
                <tblOverlap val="never" />
                <bidiVisual val="false" />
                <tblStyleRowBandSize val="100" />
                <tblStyleColBandSize val="100" />
                {}
                <jc val="center" />
                {}
                {}
                {}
                {}
                <tblLayout val="autofit" />
                {}
                {}
                <tblCaption val="Some caption" />
                <tblDescription val="Some description" />"#,
                TblPPr::test_xml("tblpPr"),
                TblWidth::test_xml("tblW"),
                TblWidth::test_xml("tblCellSpacing"),
                TblWidth::test_xml("tblInd"),
                TblBorders::test_xml("tblBorders"),
                Shd::test_xml("shd"),
                TblCellMar::test_xml("tblCellMar"),
                TblLook::test_xml("tblLook"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                style: Some(String::from("Normal")),
                paragraph_properties: Some(TblPPr::test_instance()),
                overlap: Some(TblOverlap::Never),
                bidirectional_visual: Some(false),
                style_row_band_size: Some(100),
                style_column_band_size: Some(100),
                width: Some(TblWidth::test_instance()),
                alignment: Some(JcTable::Center),
                cell_spacing: Some(TblWidth::test_instance()),
                indent: Some(TblWidth::test_instance()),
                borders: Some(TblBorders::test_instance()),
                shading: Some(Shd::test_instance()),
                layout: Some(TblLayoutType::Autofit),
                cell_margin: Some(TblCellMar::test_instance()),
                look: Some(TblLook::test_instance()),
                caption: Some(String::from("Some caption")),
                description: Some(String::from("Some description")),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_base_from_xml() {
        let xml = TblPrBase::test_xml("tblPrBase");
        assert_eq!(
            TblPrBase::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblPrBase::test_instance(),
        );
    }

    impl TblPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                TblPrBase::test_xml("tblPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                properties: TblPrBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_change_from_xml() {
        let xml = TblPrChange::test_xml("tblPrChange");
        assert_eq!(
            TblPrChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblPrChange::test_instance(),
        );
    }

    impl TblPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                TblPrBase::test_extension_xml(),
                TblPrChange::test_xml("tblPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TblPrBase::test_instance(),
                change: Some(TblPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_from_xml() {
        let xml = TblPr::test_xml("tblPr");
        assert_eq!(
            TblPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblPr::test_instance(),
        );
    }

    impl TblGridCol {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(r#"<{node_name} w="100"></{node_name}>"#, node_name = node_name)
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(TwipsMeasure::Decimal(100)),
            }
        }
    }

    #[test]
    pub fn test_tbl_grid_col_from_xml() {
        let xml = TblGridCol::test_xml("tblGridCol");
        assert_eq!(
            TblGridCol::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblGridCol::test_instance(),
        );
    }

    impl TblGridBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
            </{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name,
            )
        }

        pub fn test_extension_xml() -> String {
            format!("{grid_col}{grid_col}", grid_col = TblGridCol::test_xml("gridCol"))
        }

        pub fn test_instance() -> Self {
            Self {
                columns: vec![TblGridCol::test_instance(), TblGridCol::test_instance()],
            }
        }
    }

    #[test]
    pub fn test_tbl_grid_base_from_xml() {
        let xml = TblGridBase::test_xml("tblGridBase");
        assert_eq!(
            TblGridBase::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblGridBase::test_instance(),
        );
    }

    impl TblGridChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} id="0">{}</{node_name}>"#,
                TblGridBase::test_extension_xml(),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: Markup::test_instance(),
                grid: TblGridBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tbl_grid_change_from_xml() {
        let xml = TblGridChange::test_xml("tblGridChange");
        assert_eq!(
            TblGridChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblGridChange::test_instance(),
        );
    }

    impl TblGrid {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                TblGridBase::test_extension_xml(),
                TblGridChange::test_xml("tblGridChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TblGridBase::test_instance(),
                change: Some(TblGridChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_grid_from_xml() {
        let xml = TblGrid::test_xml("tblGrid");
        assert_eq!(
            TblGrid::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblGrid::test_instance(),
        );
    }

    impl TblPrExBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            format!(
                r#"
                {}
                <jc val="center" />
                {}
                {}
                {}
                {}
                <tblLayout val="autofit" />
                {}
                {}"#,
                TblWidth::test_xml("tblW"),
                TblWidth::test_xml("tblCellSpacing"),
                TblWidth::test_xml("tblInd"),
                TblBorders::test_xml("tblBorders"),
                Shd::test_xml("shd"),
                TblCellMar::test_xml("tblCellMar"),
                TblLook::test_xml("tblLook"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                width: Some(TblWidth::test_instance()),
                alignment: Some(JcTable::Center),
                cell_spacing: Some(TblWidth::test_instance()),
                indent: Some(TblWidth::test_instance()),
                borders: Some(TblBorders::test_instance()),
                shading: Some(Shd::test_instance()),
                layout: Some(TblLayoutType::Autofit),
                cell_margin: Some(TblCellMar::test_instance()),
                look: Some(TblLook::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_ex_base_from_xml() {
        let xml = TblPrExBase::test_xml("tblPrExBase");
        assert_eq!(
            TblPrExBase::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblPrExBase::test_instance(),
        );
    }

    impl TblPrExChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>
                {}
            </{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                TblPrExBase::test_xml("tblPrEx"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                properties_ex: TblPrExBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_ex_change_from_xml() {
        let xml = TblPrExChange::test_xml("tblPrExChange");
        assert_eq!(
            TblPrExChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblPrExChange::test_instance(),
        );
    }

    impl TblPrEx {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
            </{node_name}>"#,
                TblPrExBase::test_extension_xml(),
                TblPrExChange::test_xml("tblPrExChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TblPrExBase::test_instance(),
                change: Some(TblPrExChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tbl_pr_ex_from_xml() {
        let xml = TblPrEx::test_xml("tblPrEx");
        assert_eq!(
            TblPrEx::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TblPrEx::test_instance(),
        );
    }

    impl Height {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} val="100" hRule="auto"></{node_name}>"#,
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                value: Some(TwipsMeasure::Decimal(100)),
                height_rule: Some(HeightRule::Auto),
            }
        }
    }

    #[test]
    pub fn test_height_from_xml() {
        let xml = Height::test_xml("height");
        assert_eq!(
            Height::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            Height::test_instance(),
        );
    }

    impl TrPrBase {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>{}</{node_name}>"#,
                Self::test_extension_xml(),
                node_name = node_name
            )
        }

        pub fn test_extension_xml() -> String {
            format!(
                r#"{}
                <divId val="1" />
                <gridBefore val="1" />
                <gridAfter val="1" />
                {}
                {}
                <cantSplit val="false" />
                {}
                <tblHeader val="true" />
                {}
                <jc val="center" />
                <hidden val="false" />"#,
                Cnf::test_xml("cnfStyle"),
                TblWidth::test_xml("wBefore"),
                TblWidth::test_xml("wAfter"),
                Height::test_xml("trHeight"),
                TblWidth::test_xml("tblCellSpacing"),
            )
        }

        pub fn test_instance() -> Self {
            Self {
                conditional_formatting: Some(Cnf::test_instance()),
                div_id: Some(1),
                grid_column_before_first_cell: Some(1),
                grid_column_after_last_cell: Some(1),
                width_before_row: Some(TblWidth::test_instance()),
                width_after_row: Some(TblWidth::test_instance()),
                cant_split: Some(false),
                row_height: Some(Height::test_instance()),
                header: Some(true),
                cell_spacing: Some(TblWidth::test_instance()),
                alignment: Some(JcTable::Center),
                hidden: Some(false),
            }
        }
    }

    #[test]
    pub fn test_tr_pr_base_from_xml() {
        let xml = TrPrBase::test_xml("trPrBase");
        assert_eq!(
            TrPrBase::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TrPrBase::test_instance(),
        );
    }

    impl TrPrChange {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name} {}>{}</{node_name}>"#,
                TrackChange::TEST_ATTRIBUTES,
                TrPrBase::test_xml("trPr"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrackChange::test_instance(),
                properties: TrPrBase::test_instance(),
            }
        }
    }

    #[test]
    pub fn test_tr_pr_change_from_xml() {
        let xml = TrPrChange::test_xml("trPrChange");
        assert_eq!(
            TrPrChange::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TrPrChange::test_instance(),
        );
    }

    impl TrPr {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                TrPrBase::test_extension_xml(),
                TrackChange::test_xml("ins"),
                TrackChange::test_xml("del"),
                TrPrChange::test_xml("trPrChange"),
                node_name = node_name,
            )
        }

        pub fn test_instance() -> Self {
            Self {
                base: TrPrBase::test_instance(),
                inserted: Some(TrackChange::test_instance()),
                deleted: Some(TrackChange::test_instance()),
                change: Some(TrPrChange::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tr_pr_from_xml() {
        let xml = TrPr::test_xml("trPr");
        assert_eq!(
            TrPr::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TrPr::test_instance(),
        );
    }

    impl TcBorders {
        pub fn test_xml(node_name: &'static str) -> String {
            format!(
                r#"<{node_name}>
                {}
                {}
                {}
                {}
                {}
                {}
                {}
                {}
            </{node_name}>"#,
                Border::test_xml("top"),
                Border::test_xml("start"),
                Border::test_xml("bottom"),
                Border::test_xml("end"),
                Border::test_xml("insideH"),
                Border::test_xml("insideV"),
                Border::test_xml("tl2br"),
                Border::test_xml("tr2bl"),
                node_name = node_name
            )
        }

        pub fn test_instance() -> Self {
            Self {
                top: Some(Border::test_instance()),
                start: Some(Border::test_instance()),
                bottom: Some(Border::test_instance()),
                end: Some(Border::test_instance()),
                inside_horizontal: Some(Border::test_instance()),
                inside_vertical: Some(Border::test_instance()),
                top_left_to_bottom_right: Some(Border::test_instance()),
                top_right_to_bottom_left: Some(Border::test_instance()),
            }
        }
    }

    #[test]
    pub fn test_tc_borders_from_xml() {
        let xml = TcBorders::test_xml("tcBorders");
        assert_eq!(
            TcBorders::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap(),
            TcBorders::test_instance(),
        );
    }
}
