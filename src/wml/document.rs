use msoffice_shared::{
    drawingml::HexColorRGB,
    error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    relationship::RelationshipId,
    sharedtypes::OnOff,
    xml::{parse_xml_bool, XmlNode},
};

pub type Result<T> = std::result::Result<T, Box<dyn std::error::Error>>;

pub type UcharHexNumber = String;
pub type LongHexNumber = String; // length=4
pub type ShortHexNumber = String; // length=2
pub type UnqualifiedPercentage = i32;
pub type DecimalNumber = i32;
pub type UnsignedDecimalNumber = u32;
pub type DateTime = String;
pub type MacroName = String; // maxLength=33
pub type EightPointMeasure = u32;
pub type PointMeasure = u32;
pub type TextScalePercent = String; // pattern=0*(600|([0-5]?[0-9]?[0-9]))%
pub type TextScaleDecimal = i32; // 0 <= n <= 600

#[derive(Default, Debug, Clone)]
pub struct Charset {
    pub value: Option<UcharHexNumber>,
    pub character_set: Option<String>,
}

impl Charset {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut instance: Charset = Default::default();

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "val" => instance.value = Some(value.clone()),
                "characterSet" => instance.character_set = Some(value.clone()),
                _ => (),
            }
        }

        Ok(instance)
    }
}

#[derive(Debug, Clone)]
pub enum DecimalNumberOrPercent {
    Int(UnqualifiedPercentage),
    Percentage(String),
}

pub enum TextScale {
    Percent(TextScalePercent),
    Decimal(TextScaleDecimal),
}

#[derive(Debug, Clone, EnumString)]
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
                "placeholder" => placeholder = child_node.text.clone(),
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
            <placeholder>Placeholder</placeholder>
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
    pub field_lock: OnOff,
    pub dirty: OnOff,
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
        let field_lock = field_lock.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "fldLock"))?;
        let dirty = dirty.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "dirty"))?;

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
            field_lock: false,
            dirty: false,
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

// TODO
#[derive(Debug, Clone, PartialEq)]
pub enum RPrBase {
    RunStyle(String),
    // RunFonts(Fonts),
    // Bold(OnOff),
    // ComplexScriptBold(OnOff),
    // Italic(OnOff),
    // ComplexScriptItalic(OnOff),
    // Capitals(OnOff),
    // SmallCapitals(OnOff),
    // Strikethrough(OnOff),
    // DoubleStrikethrough(OnOff),
    // Outline(OnOff),
    // Shadow(OnOff),
    // Emboss(OnOff),
    // Imprint(OnOff),
    // NoProofing(OnOff),
    // SnapToGrid(OnOff),
    // Vanish(OnOff),
    // WebHidden(OnOff),
    // Color(Color),
    // Spacing(SignedTwipsMeasure),
    // Width(TextScale),
    // Kerning(HpsMeasure),
    // Position(SignedHpsMeasure),
    // Size(HpsMeasure),
    // ComplexScriptSize(HpsMeasure),
    // Highlight(Highlight),
    // Underline(Underline),
    // Effect(TextEffect),
    // Border(Border),
    // Shading(Shd),
    // FitText(FitText),
    // VertialAlignment(VerticalAlignRun),
    // Rtl(OnOff),
    // ComplexScript(OnOff),
    // EmphasisMark(Em),
    // Language(Language),
    // EastAsianLayout(EastAsianLayout),
    // SpecialVanish(OnOff),
    // OMath(OnOff),
}

/*
<xsd:group name="EG_RPrBase">
    <xsd:choice>
        <xsd:element name="rStyle" type="CT_String"/>
        <xsd:element name="rFonts" type="CT_Fonts"/>
        <xsd:element name="b" type="CT_OnOff"/>
        <xsd:element name="bCs" type="CT_OnOff"/>
        <xsd:element name="i" type="CT_OnOff"/>
        <xsd:element name="iCs" type="CT_OnOff"/>
        <xsd:element name="caps" type="CT_OnOff"/>
        <xsd:element name="smallCaps" type="CT_OnOff"/>
        <xsd:element name="strike" type="CT_OnOff"/>
        <xsd:element name="dstrike" type="CT_OnOff"/>
        <xsd:element name="outline" type="CT_OnOff"/>
        <xsd:element name="shadow" type="CT_OnOff"/>
        <xsd:element name="emboss" type="CT_OnOff"/>
        <xsd:element name="imprint" type="CT_OnOff"/>
        <xsd:element name="noProof" type="CT_OnOff"/>
        <xsd:element name="snapToGrid" type="CT_OnOff"/>
        <xsd:element name="vanish" type="CT_OnOff"/>
        <xsd:element name="webHidden" type="CT_OnOff"/>
        <xsd:element name="color" type="CT_Color"/>
        <xsd:element name="spacing" type="CT_SignedTwipsMeasure"/>
        <xsd:element name="w" type="CT_TextScale"/>
        <xsd:element name="kern" type="CT_HpsMeasure"/>
        <xsd:element name="position" type="CT_SignedHpsMeasure"/>
        <xsd:element name="sz" type="CT_HpsMeasure"/>
        <xsd:element name="szCs" type="CT_HpsMeasure"/>
        <xsd:element name="highlight" type="CT_Highlight"/>
        <xsd:element name="u" type="CT_Underline"/>
        <xsd:element name="effect" type="CT_TextEffect"/>
        <xsd:element name="bdr" type="CT_Border"/>
        <xsd:element name="shd" type="CT_Shd"/>
        <xsd:element name="fitText" type="CT_FitText"/>
        <xsd:element name="vertAlign" type="CT_VerticalAlignRun"/>
        <xsd:element name="rtl" type="CT_OnOff"/>
        <xsd:element name="cs" type="CT_OnOff"/>
        <xsd:element name="em" type="CT_Em"/>
        <xsd:element name="lang" type="CT_Language"/>
        <xsd:element name="eastAsianLayout" type="CT_EastAsianLayout"/>
        <xsd:element name="specVanish" type="CT_OnOff"/>
        <xsd:element name="oMath" type="CT_OnOff"/>
    </xsd:choice>
</xsd:group>
*/
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
                Ok(RPrBase::RunStyle(xml_node.text.as_ref().cloned().ok_or_else(|| {
                    MissingChildNodeError::new(xml_node.name.clone(), "Text node")
                })?))
            }
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "RPrBase"))),
        }
    }
}

#[cfg(test)]
impl RPrBase {
    pub fn test_run_style_xml() -> &'static str {
        r#"<rStyle>Arial</rStyle>"#
    }

    pub fn test_run_style_instance() -> Self {
        RPrBase::RunStyle(String::from("Arial"))
    }
}

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

/*
<xsd:group name="EG_RPrContent">
    <xsd:sequence>
      <xsd:group ref="EG_RPrBase" minOccurs="0" maxOccurs="unbounded"/>
      <xsd:element name="rPrChange" type="CT_RPrChange" minOccurs="0"/>
    </xsd:sequence>
  </xsd:group>
*/
#[derive(Debug, Clone, PartialEq)]
pub struct RPrContent {
    pub r_pr_bases: Vec<RPrBase>,
    pub run_properties_change: Option<RPrChange>,
}

/*
<xsd:complexType name="CT_RPr">
    <xsd:sequence>
      <xsd:group ref="EG_RPrContent" minOccurs="0"/>
    </xsd:sequence>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq)]
pub struct RPr {
    pub content: Option<RPrContent>,
}

#[derive(Debug, Clone, PartialEq)]
pub enum SdtPrControlChoice {
    // Equation,
// ComboBox(SdtComboBox),
// Date(SdtDate),
// DocumentPartObject(SdtDocPart),
// DocumentPartList(SdtDocPart),
// DropDownList(SdtDropDownList),
// Picture,
// RichText,
// Text(SdtText),
// Citation,
// Group,
// Bibliography,
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
#[derive(Debug, Clone, PartialEq)]
pub struct SdtPr {
    pub run_properties: Option<RPr>,
    pub alias: Option<String>,
    pub tag: Option<String>,
    pub id: Option<DecimalNumber>,
    //pub lock: Option<Lock>,
    //pub placeholder: Option<Placeholder>,
    pub temporary: Option<OnOff>,
    pub showing_placeholder_header: Option<OnOff>,
    //pub data_binding: Option<DataBinding>,
    pub label: Option<DecimalNumber>,
    pub tab_index: Option<UnsignedDecimalNumber>,
    pub control_choice: Option<SdtPrControlChoice>,
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
