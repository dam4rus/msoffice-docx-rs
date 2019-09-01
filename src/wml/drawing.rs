use msoffice_shared::{
    drawingml::{
        BlackWhiteMode, Coordinate, GraphicalObject, GraphicalObjectData, GroupShapeProperties,
        NonVisualConnectorProperties, NonVisualContentPartProperties, NonVisualDrawingProps,
        NonVisualDrawingShapeProps, NonVisualGraphicFrameProperties, Picture, Point2D, PositiveSize2D, ShapeProperties,
        ShapeStyle, TextBodyProperties, Transform2D,
    },
    error::{Limit, LimitViolationError, MissingAttributeError, MissingChildNodeError},
    relationship::RelationshipId,
    xml::XmlNode,
};

type Result<T> = ::std::result::Result<T, Box<dyn std::error::Error>>;

type PositionOffset = i32;
type WrapDistance = u32;

#[derive(Debug, Clone, PartialEq)]
pub struct EffectExtent {
    pub left: Coordinate,
    pub top: Coordinate,
    pub right: Coordinate,
    pub bottom: Coordinate,
}

impl EffectExtent {
    pub fn new(left: Coordinate, top: Coordinate, right: Coordinate, bottom: Coordinate) -> Self {
        Self {
            left,
            top,
            right,
            bottom,
        }
    }

    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut left = None;
        let mut top = None;
        let mut right = None;
        let mut bottom = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "l" => left = Some(value.parse()?),
                "t" => top = Some(value.parse()?),
                "r" => right = Some(value.parse()?),
                "b" => bottom = Some(value.parse()?),
                _ => (),
            }
        }

        Ok(Self {
            left: left.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "l"))?,
            top: top.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "t"))?,
            right: right.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r"))?,
            bottom: bottom.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "b"))?,
        })
    }
}

#[cfg(test)]
#[test]
pub fn effect_extent_from_xml_node_test() {
    let xml = r#"<effect l="0" t="0" r="100" b="100"></effect>"#;
    let effect_extent = EffectExtent::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    assert_eq!(
        effect_extent,
        EffectExtent {
            left: 0,
            top: 0,
            right: 100,
            bottom: 100
        }
    );
}

#[derive(Debug, Clone, PartialEq)]
pub struct Inline {
    pub extent: PositiveSize2D,
    pub effect_extent: Option<EffectExtent>,
    pub doc_properties: NonVisualDrawingProps,
    pub graphic_frame_properties: Option<NonVisualGraphicFrameProperties>,
    pub graphic: GraphicalObject,

    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

impl Inline {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut distance_top = None;
        let mut distance_bottom = None;
        let mut distance_left = None;
        let mut distance_right = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "distT" => distance_top = Some(value.parse()?),
                "distB" => distance_bottom = Some(value.parse()?),
                "distL" => distance_left = Some(value.parse()?),
                "distR" => distance_right = Some(value.parse()?),
                _ => (),
            }
        }

        let mut extent = None;
        let mut effect_extent = None;
        let mut doc_properties = None;
        let mut graphic_frame_properties = None;
        let mut graphic = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "extent" => extent = Some(PositiveSize2D::from_xml_element(child_node)?),
                "effectExtent" => effect_extent = Some(EffectExtent::from_xml_element(child_node)?),
                "docPr" => doc_properties = Some(NonVisualDrawingProps::from_xml_element(child_node)?),
                "cNvGraphicFramePr" => {
                    graphic_frame_properties = Some(NonVisualGraphicFrameProperties::from_xml_element(child_node)?)
                }
                "graphic" => graphic = Some(GraphicalObject::from_xml_element(child_node)?),
                _ => (),
            }
        }

        Ok(Self {
            extent: extent.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "extent"))?,
            effect_extent,
            doc_properties: doc_properties.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "docPr"))?,
            graphic_frame_properties,
            graphic: graphic.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "graphic"))?,
            distance_top,
            distance_bottom,
            distance_left,
            distance_right,
        })
    }
}

#[cfg(test)]
#[test]
pub fn test_inline_from_xml_element() {
    use msoffice_shared::drawingml::Hyperlink;

    let xml = r#"<inline distT="0" distB="100" distL="0" distR="100">
        <extent cx="10000" cy="10000" />
        <effectExtent l="0" t="0" r="1000" b="1000" />
        <docPr id="1" name="Object name" descr="Some description" title="Title of the object">
            <a:hlinkClick r:id="rId2" tooltip="Some Sample Text"/>
            <a:hlinkHover r:id="rId2" tooltip="Some Sample Text"/>
        </docPr>
        <graphic>
            <graphicData uri="http://some/url" />
        </graphic>
    </inline>"#;

    let inline = Inline::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    let inline_rhs = Inline {
        extent: PositiveSize2D::new(10000, 10000),
        effect_extent: Some(EffectExtent::new(0, 0, 1000, 1000)),
        doc_properties: NonVisualDrawingProps {
            id: 1,
            name: String::from("Object name"),
            description: Some(String::from("Some description")),
            hidden: None,
            title: Some(String::from("Title of the object")),
            hyperlink_click: Some(Box::new(Hyperlink {
                relationship_id: Some(String::from("rId2")),
                tooltip: Some(String::from("Some Sample Text")),
                ..Default::default()
            })),
            hyperlink_hover: Some(Box::new(Hyperlink {
                relationship_id: Some(String::from("rId2")),
                tooltip: Some(String::from("Some Sample Text")),
                ..Default::default()
            })),
        },
        graphic_frame_properties: None,
        graphic: GraphicalObject {
            graphic_data: GraphicalObjectData {
                uri: String::from("http://some/url"),
            },
        },
        distance_top: Some(0),
        distance_bottom: Some(100),
        distance_left: Some(0),
        distance_right: Some(100),
    };
    assert_eq!(inline, inline_rhs);
}

#[derive(Debug, Clone, EnumString, PartialEq)]
pub enum WrapText {
    #[strum(serialize = "bothSides")]
    BothSides,
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "largest")]
    Largest,
}

/*
<xsd:complexType name="CT_WrapPath">
    <xsd:sequence>
      <xsd:element name="start" type="a:CT_Point2D" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="lineTo" type="a:CT_Point2D" minOccurs="2" maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="edited" type="xsd:boolean" use="optional"/>
  </xsd:complexType>
*/
#[derive(Debug, Clone, PartialEq)]
pub struct WrapPath {
    pub start: Point2D,
    pub line_to: Vec<Point2D>,

    pub edited: Option<bool>,
}

impl WrapPath {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let edited = xml_node.attribute("edited").map(|value| value.parse()).transpose()?;

        let mut start = None;
        let mut line_to = Vec::new();

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "start" => start = Some(Point2D::from_xml_element(child_node)?),
                "lineTo" => line_to.push(Point2D::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let start = start.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "start"))?;
        if line_to.len() < 2 {
            return Err(Box::new(LimitViolationError::new(
                xml_node.name.clone(),
                "lineTo",
                Limit::Value(2),
                Limit::Unbounded,
                line_to.len() as u32,
            )));
        }

        Ok(Self { start, line_to, edited })
    }
}

#[cfg(test)]
#[test]
pub fn test_wrap_path_from_xml() {
    let xml = r#"<wrap_path edited="true">
        <start x="0" y="0" />
        <lineTo x="50" y="50" />
        <lineTo x="100" y="100" />
    </wrap_path>"#;

    let wrap_path = WrapPath::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    let wrap_path_rhs = WrapPath {
        start: Point2D::new(0, 0),
        line_to: vec![Point2D::new(50, 50), Point2D::new(100, 100)],
        edited: Some(true),
    };
    assert_eq!(wrap_path, wrap_path_rhs);
}

#[derive(Debug, Clone, PartialEq)]
pub struct WrapSquare {
    pub effect_extent: Option<EffectExtent>,

    pub wrap_text: WrapText,
    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

impl WrapSquare {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut wrap_text = None;
        let mut distance_top = None;
        let mut distance_bottom = None;
        let mut distance_left = None;
        let mut distance_right = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_ref() {
                "wrapText" => wrap_text = Some(value.parse()?),
                "distT" => distance_top = Some(value.parse()?),
                "distB" => distance_bottom = Some(value.parse()?),
                "distL" => distance_left = Some(value.parse()?),
                "distR" => distance_right = Some(value.parse()?),
                _ => (),
            }
        }

        let effect_extent = xml_node
            .child_nodes
            .first()
            .map(|child_node| EffectExtent::from_xml_element(child_node))
            .transpose()?;
        Ok(Self {
            effect_extent,
            wrap_text: wrap_text.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "wrapText"))?,
            distance_top,
            distance_bottom,
            distance_left,
            distance_right,
        })
    }
}

#[cfg(test)]
#[test]
pub fn test_wrap_square_from_xml() {
    let xml = r#"<wrap_square wrapText="bothSides" distT="0" distB="100" distL="0" distR="100">
        <effectExtent l="0" t="0" r="100" b="100" />
    </wrap_square>"#;

    let wrap_square = WrapSquare::from_xml_element(&XmlNode::from_str(xml).unwrap()).unwrap();
    let wrap_square_rhs = WrapSquare {
        effect_extent: Some(EffectExtent::new(0, 0, 100, 100)),
        wrap_text: WrapText::BothSides,
        distance_top: Some(0),
        distance_bottom: Some(100),
        distance_left: Some(0),
        distance_right: Some(100),
    };

    assert_eq!(wrap_square, wrap_square_rhs);
}

#[derive(Debug, Clone)]
pub struct WrapTight {
    pub wrap_polygon: WrapPath,

    pub wrap_text: WrapText,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

#[derive(Debug, Clone)]
pub struct WrapThrough {
    pub wrap_polygon: WrapPath,

    pub wrap_text: WrapText,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

#[derive(Debug, Clone, Default)]
pub struct WrapTopBottom {
    pub effect_extent: Option<EffectExtent>,

    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
}

#[derive(Debug, Clone)]
pub enum WrapType {
    None,
    Square(WrapSquare),
    Tight(WrapTight),
    Through(WrapThrough),
    TopAndBottom(WrapTopBottom),
}

#[derive(Debug, Clone, EnumString)]
pub enum AlignH {
    #[strum(serialize = "left")]
    Left,
    #[strum(serialize = "right")]
    Right,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "inside")]
    Inside,
    #[strum(serialize = "outside")]
    Outside,
}

#[derive(Debug, Clone, EnumString)]
pub enum RelFromH {
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "page")]
    Page,
    #[strum(serialize = "column")]
    Column,
    #[strum(serialize = "character")]
    Character,
    #[strum(serialize = "leftMargin")]
    LeftMargin,
    #[strum(serialize = "rightMargin")]
    RightMargin,
    #[strum(serialize = "insideMargin")]
    InsideMargin,
    #[strum(serialize = "outsideMargin")]
    OutsideMargin,
}

#[derive(Debug, Clone, EnumString)]
pub enum AlignV {
    #[strum(serialize = "top")]
    Top,
    #[strum(serialize = "bottom")]
    Bottom,
    #[strum(serialize = "center")]
    Center,
    #[strum(serialize = "inside")]
    Inside,
    #[strum(serialize = "outside")]
    Outside,
}

#[derive(Debug, Clone, EnumString)]
pub enum RelFromV {
    #[strum(serialize = "margin")]
    Margin,
    #[strum(serialize = "page")]
    Page,
    #[strum(serialize = "paragraph")]
    Paragraph,
    #[strum(serialize = "line")]
    Line,
    #[strum(serialize = "topMargin")]
    TopMargin,
    #[strum(serialize = "bottomMargin")]
    BottomMargin,
    #[strum(serialize = "insideMargin")]
    InsideMargin,
    #[strum(serialize = "outsideMargin")]
    OutsideMargin,
}

#[derive(Debug, Clone)]
pub enum PosHChoice {
    Align(AlignH),
    PositionOffset(PositionOffset),
}

#[derive(Debug, Clone)]
pub struct PosH {
    pub choice: PosHChoice,
    pub relative_from: RelFromH,
}

#[derive(Debug, Clone)]
pub enum PosVChoice {
    Align(AlignV),
    PositionOffset(PositionOffset),
}

#[derive(Debug, Clone)]
pub struct PosV {
    pub choice: PosVChoice,
    pub relative_from: RelFromV,
}

#[derive(Debug, Clone)]
pub struct Anchor {
    pub simple_position: Point2D,
    pub horizontal_position: PosH,
    pub vertical_position: PosV,
    pub extent: PositiveSize2D,
    pub effect_extent: Option<EffectExtent>,
    pub wrap_type: WrapType,
    pub document_properties: NonVisualDrawingProps,
    pub graphic_frame_properties: Option<NonVisualGraphicFrameProperties>,
    pub graphic: GraphicalObject,

    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
    pub use_simple_position: bool,
    pub relative_height: u32,
    pub behind_document_text: bool,
    pub locked: bool,
    pub layout_in_cell: bool,
    pub hidden: Option<bool>,
    pub allow_overlap: bool,
}

#[derive(Debug, Clone)]
pub struct TxbxContent {
    //pub block_level_elements: Vec<super::BlockLevelElts>, // minOccurs=1
}

#[derive(Debug, Clone)]
pub struct TextboxInfo {
    pub textbox_content: TxbxContent,
    pub id: Option<u16>, // default=0,
}

#[derive(Debug, Clone)]
pub struct LinkedTextboxInformation {
    pub id: u16,
    pub sequence: u16,
}

#[derive(Debug, Clone)]
pub enum WordprocessingShapePropertiesChoice {
    ShapeProperties(NonVisualDrawingShapeProps),
    Connector(NonVisualConnectorProperties),
}

#[derive(Debug, Clone)]
pub enum WordprocessingShapeTextboxInfoChoice {
    Textbox(TextboxInfo),
    LinkedTextbox(LinkedTextboxInformation),
}

#[derive(Debug, Clone)]
pub struct WordprocessingShape {
    pub non_visual_drawing_props: Option<NonVisualDrawingProps>,
    pub properties: WordprocessingShapePropertiesChoice,
    pub shape_properties: ShapeProperties,
    pub style: Option<ShapeStyle>,
    pub text_box_info: Option<WordprocessingShapeTextboxInfoChoice>,
    pub text_body_properties: TextBodyProperties,

    pub normal_east_asian_flow: Option<bool>, // default=false
}

#[derive(Debug, Clone)]
pub struct GraphicFrame {
    pub non_visual_drawing_props: NonVisualDrawingProps,
    pub non_visual_props: NonVisualGraphicFrameProperties,
    pub transform: Transform2D,
    pub graphic: GraphicalObject,
}

#[derive(Debug, Clone)]
pub struct WordprocessingContentPartNonVisual {
    pub non_visual_drawing_props: Option<NonVisualDrawingProps>,
    pub non_visual_props: Option<NonVisualContentPartProperties>,
}

#[derive(Debug, Clone)]
pub struct WordprocessingContentPart {
    pub properties: Option<WordprocessingContentPartNonVisual>,
    pub transform: Option<Transform2D>,

    pub black_and_white_mode: Option<BlackWhiteMode>,
    pub relationship_id: Option<RelationshipId>,
}

#[derive(Debug, Clone)]
pub enum WordprocessingGroupChoice {
    Shape(WordprocessingShape),
    Group(WordprocessingGroup),
    GraphicFrame(GraphicFrame),
    Picture(Picture),
    ContentPart(WordprocessingContentPart),
}

#[derive(Debug, Clone)]
pub struct WordprocessingGroup {
    pub non_visual_drawing_props: Option<NonVisualDrawingProps>,
    pub non_visual_drawing_shape_props: NonVisualDrawingShapeProps,
    pub group_shape_props: GroupShapeProperties,
    pub shapes: Vec<WordprocessingGroupChoice>,
}

pub enum WordprocessingCanvasChoice {
    Shape(WordprocessingShape),
    Picture(Picture),
    ContentPart(WordprocessingContentPart),
    Group(WordprocessingGroup),
    GraphicFrame(GraphicFrame),
}

pub struct WordprocessingCanvas {
    //pub background_formatting: Option<BackgroundFormatting>,
    //pub whole_formatting: Option<WholeE2oFormatting>,
    pub shapes: Vec<WordprocessingCanvasChoice>,
}
