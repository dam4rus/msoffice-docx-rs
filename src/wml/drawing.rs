use msoffice_shared::drawingml;

type PositionOffset = i32;
type WrapDistance = u32;

pub struct EffectExtent {
    pub left: drawingml::Coordinate,
    pub top: drawingml::Coordinate,
    pub right: drawingml::Coordinate,
    pub bottom: drawingml::Coordinate,
}

#[derive(Debug, Clone, Eq)]
pub struct Inline {
    pub extent: drawingml::PositiveSize2D,
    pub effect_extent: Option<EffectExtent>,
    pub doc_properties: drawingml::NonVisualDrawingProps,
    pub graphic_frame_properties: Option<drawingml::NonVisualGraphicFrameProperties>,
    pub graphic: drawingml::GraphicalObject,

    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

#[derive(Debug, Clone, Eq, EnumString)]
pub enum WrapText {
    #[strum(serialize="bothSides")]
    BothSides,
    #[strum(serialize="left")]
    Left,
    #[strum(serialize="right")]
    Right,
    #[strum(serialize="largest")]
    Largest,
}

#[derive(Debug, Clone, Eq)]
pub struct WrapPath {
    pub start: drawingml::Point2D,
    pub line_to: Vec<drawingml::Point2D>,

    pub edited: Option<bool>,
}

#[derive(Debug, Clone, Eq)]
pub struct WrapSquare {
    pub effect_extent: Option<EffectExtent>,

    pub wrap_text: WrapText,
    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

#[derive(Debug, Clone, Eq)]
pub struct WrapTight {
    pub wrap_polygon: WrapPath,

    pub wrap_text: WrapText,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

#[derive(Debug, Clone, Eq)]
pub struct WrapThrough {
    pub wrap_polygon: WrapPath,

    pub wrap_text: WrapText,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
}

#[derive(Debug, Clone, Eq, Default)]
pub struct WrapTopBottom {
    pub effect_extent: Option<EffectExtent>,

    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
}

#[derive(Debug, Eq, Clone)]
pub enum WrapType {
    None,
    Square(WrapSquare),
    Tight(WrapTight),
    Through(WrapThrough),
    TopAndBottom(WrapTopBottom),
}

#[derive(Debug, Clone, Eq, EnumString)]
pub enum AlignH {
    #[strum(serialize="left")]
    Left,
    #[strum(serialize="right")]
    Right,
    #[strum(serialize="center")]
    Center,
    #[strum(serialize="inside")]
    Inside,
    #[strum(serialize="outside")]
    Outside,
}

#[derive(Debug, Clone, Eq, EnumString)]
pub enum RelFromH {
    #[strum(serialize="margin")]
    Margin,
    #[strum(serialize="page")]
    Page,
    #[strum(serialize="column")]
    Column,
    #[strum(serialize="character")]
    Character,
    #[strum(serialize="leftMargin")]
    LeftMargin,
    #[strum(serialize="rightMargin")]
    RightMargin,
    #[strum(serialize="insideMargin")]
    InsideMargin,
    #[strum(serialize="outsideMargin")]
    OutsideMargin,
}

#[derive(Debug, Clone, Eq, EnumString)]
pub enum AlignV {
    #[strum(serialize="top")]
    Top,
    #[strum(serialize="bottom")]
    Bottom,
    #[strum(serialize="center")]
    Center,
    #[strum(serialize="inside")]
    Inside,
    #[strum(serialize="outside")]
    Outside,
}

#[derive(Debug, Clone, Eq, EnumString)]
pub enum RelFromV {
    #[strum(serialize="margin")]
    Margin,
    #[strum(serialize="page")]
    Page,
    #[strum(serialize="paragraph")]
    Paragraph,
    #[strum(serialize="line")]
    Line,
    #[strum(serialize="topMargin")]
    TopMargin,
    #[strum(serialize="bottomMargin")]
    BottomMargin,
    #[strum(serialize="insideMargin")]
    InsideMargin,
    #[strum(serialize="outsideMargin")]
    OutsideMargin,
}

#[derive(Debug, Clone, Eq)]
pub enum PosVChoice {
    Align(AlignV),
    PositionOffset(PositionOffset),
}

#[derive(Debug, Clone, Eq)]
pub struct PosV {
    pub choice: PosVChoice,
    pub relative_from: RelFromV,
}

#[derive(Debug, Clone, Eq)]
pub struct TxbxContent {
    pub block_level_elements: Vec<super::BlockLevelElts>, // minOccurs=1
}

#[derive(Debug, Clone, Eq)]
pub struct TextboxInfo {
    pub textbox_content: TxbxContent,
    pub id: Option<u16>, // default=0,
}

#[derive(Debug, Clone, Eq)]
pub struct LinkedTextboxInformation {
    pub id: u16,
    pub sequence: u16,
}

#[derive(Debug, Clone, Eq)]
enum WordprocessingShapePropertiesChoice {
    ShapeProperties(drawingml::NonVisualDrawingShapeProps),
    Connector(drawingml::NonVisualConnectorProperties),
}

#[derive(Debug, Clone, Eq)]
enum WordprocessingShapeTextboxInfoChoice {
    Textbox(TextBoxInfo),
    LinkedTextbox(LinkedTextboxInformation),
}

#[derive(Debug, Clone, Eq)]
pub struct WordprocessingShape {
    pub non_visual_drawing_props: Option<drawingml::NonVisualDrawingProps>,
    pub properties: WordprocessingShapePropertiesChoice,
    pub shape_properties: drawingml::ShapeProperties,
    pub style: Option<drawingml::ShapeStyle>,
    pub text_box_info: Option<WordprocessingShapeTextboxInfoChoice>,
    pub text_body_properties: drawingml::TextBodyProperties,

    pub normal_east_asian_flow: Option<bool>, // default=false
}
/*
  <xsd:complexType name="CT_GraphicFrame">
    <xsd:sequence>
      <xsd:element name="cNvPr" type="a:CT_NonVisualDrawingProps" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="cNvFrPr" type="a:CT_NonVisualGraphicFrameProperties" minOccurs="1"
        maxOccurs="1"/>
      <xsd:element name="xfrm" type="a:CT_Transform2D" minOccurs="1" maxOccurs="1"/>
      <xsd:element ref="a:graphic" minOccurs="1" maxOccurs="1"/>
      <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>
  <xsd:complexType name="CT_WordprocessingContentPartNonVisual">
    <xsd:sequence>
      <xsd:element name="cNvPr" type="a:CT_NonVisualDrawingProps" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="cNvContentPartPr" type="a:CT_NonVisualContentPartProperties" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>
  <xsd:complexType name="CT_WordprocessingContentPart">
    <xsd:sequence>
      <xsd:element name="nvContentPartPr" type="CT_WordprocessingContentPartNonVisual" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="xfrm" type="a:CT_Transform2D" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="bwMode" type="a:ST_BlackWhiteMode" use="optional"/>
    <xsd:attribute ref="r:id" use="required"/>
  </xsd:complexType>
  <xsd:complexType name="CT_WordprocessingGroup">
    <xsd:sequence minOccurs="1" maxOccurs="1">
      <xsd:element name="cNvPr" type="a:CT_NonVisualDrawingProps" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="cNvGrpSpPr" type="a:CT_NonVisualGroupDrawingShapeProps" minOccurs="1"
        maxOccurs="1"/>
      <xsd:element name="grpSpPr" type="a:CT_GroupShapeProperties" minOccurs="1" maxOccurs="1"/>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element ref="wsp"/>
        <xsd:element name="grpSp" type="CT_WordprocessingGroup"/>
        <xsd:element name="graphicFrame" type="CT_GraphicFrame"/>
        <xsd:element ref="dpct:pic"/>
        <xsd:element name="contentPart" type="CT_WordprocessingContentPart"/>
      </xsd:choice>
      <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>
  <xsd:complexType name="CT_WordprocessingCanvas">
    <xsd:sequence minOccurs="1" maxOccurs="1">
      <xsd:element name="bg" type="a:CT_BackgroundFormatting" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="whole" type="a:CT_WholeE2oFormatting" minOccurs="0" maxOccurs="1"/>
      <xsd:choice minOccurs="0" maxOccurs="unbounded">
        <xsd:element ref="wsp"/>
        <xsd:element ref="dpct:pic"/>
        <xsd:element name="contentPart" type="CT_WordprocessingContentPart"/>
        <xsd:element ref="wgp"/>
        <xsd:element name="graphicFrame" type="CT_GraphicFrame"/>
      </xsd:choice>
      <xsd:element name="extLst" type="a:CT_OfficeArtExtensionList" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
  </xsd:complexType>
*/

#[derive(Debug, Copy, Clone, Eq)]
pub enum PosHChoice {
    Align(AlignH),
    PositionOffset(PositionOffset),
}

#[derive(Debug, Copy, Clone, Eq)]
pub struct PosH {
    pub choice: PosHChoice,
    pub relative_from: RelFromH,
}

#[derive(Debug, Clone, Eq)]
pub struct Anchor {
    pub simple_position: drawingml::Point2D,
    pub horizontal_position: PosH,
    pub vertical_position: PosV,
    pub extent: drawingml::PositiveSize2D,
    pub effect_extent: Option<EffectExtent>,
    pub wrap_type: WrapType,
    pub document_properties: drawingml::NonVisualDrawingProps,
    pub graphic_frame_properties: Option<drawingml::NonVisualGraphicFrameProperties>,
    pub graphic: drawingml::GraphicalObject,

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