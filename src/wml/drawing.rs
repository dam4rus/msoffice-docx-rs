use msoffice_shared::{
    drawingml,
    relationship::RelationshipId
};

type PositionOffset = i32;
type WrapDistance = u32;

#[derive(Debug, Clone)]
pub struct EffectExtent {
    pub left: drawingml::Coordinate,
    pub top: drawingml::Coordinate,
    pub right: drawingml::Coordinate,
    pub bottom: drawingml::Coordinate,
}

#[derive(Debug, Clone)]
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

#[derive(Debug, Clone, EnumString)]
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

#[derive(Debug, Clone)]
pub struct WrapPath {
    pub start: drawingml::Point2D,
    pub line_to: Vec<drawingml::Point2D>,

    pub edited: Option<bool>,
}

#[derive(Debug, Clone)]
pub struct WrapSquare {
    pub effect_extent: Option<EffectExtent>,

    pub wrap_text: WrapText,
    pub distance_top: Option<WrapDistance>,
    pub distance_bottom: Option<WrapDistance>,
    pub distance_left: Option<WrapDistance>,
    pub distance_right: Option<WrapDistance>,
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

#[derive(Debug, Clone, EnumString)]
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

#[derive(Debug, Clone, EnumString)]
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

#[derive(Debug, Clone, EnumString)]
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

#[derive(Debug, Clone)]
pub struct TxbxContent {
    pub block_level_elements: Vec<super::BlockLevelElts>, // minOccurs=1
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

#[derive(Debug, Clone,)]
enum WordprocessingShapePropertiesChoice {
    ShapeProperties(drawingml::NonVisualDrawingShapeProps),
    Connector(drawingml::NonVisualConnectorProperties),
}

#[derive(Debug, Clone)]
enum WordprocessingShapeTextboxInfoChoice {
    Textbox(TextboxInfo),
    LinkedTextbox(LinkedTextboxInformation),
}

#[derive(Debug, Clone)]
pub struct WordprocessingShape {
    pub non_visual_drawing_props: Option<drawingml::NonVisualDrawingProps>,
    pub properties: WordprocessingShapePropertiesChoice,
    pub shape_properties: drawingml::ShapeProperties,
    pub style: Option<drawingml::ShapeStyle>,
    pub text_box_info: Option<WordprocessingShapeTextboxInfoChoice>,
    pub text_body_properties: drawingml::TextBodyProperties,

    pub normal_east_asian_flow: Option<bool>, // default=false
}

#[derive(Debug, Clone)]
pub struct GraphicFrame {
    pub non_visual_drawing_props: drawingml::NonVisualDrawingProps,
    pub non_visual_props: drawingml::NonVisualGraphicFrameProperties,
    pub transform: drawingml::Transform2D,
    pub graphic: drawingml::GraphicalObject,
}

#[derive(Debug, Clone)]
pub struct WordprocessingContentPartNonVisual {
    pub non_visual_drawing_props: Option<drawingml::NonVisualDrawingProps>,
    pub non_visual_props: Option<drawingml::NonVisualContentPartProperties>,
}

#[derive(Debug, Clone)]
pub struct WordprocessingContentPart {
    pub properties: Option<WordprocessingContentPartNonVisual>,
    pub transform: Option<drawingml::Transform2D>,

    pub black_and_white_mode: Option<drawingml::BlackWhiteMode>,
    pub relationship_id: Option<RelationshipId>,
}

#[derive(Debug, Clone)]
pub enum WordprocessingGroupChoice {
    Shape(WordprocessingShape),
    Group(WordprocessingGroup),
    GraphicFrame(GraphicFrame),
    Picture(drawingml::Picture),
    ContentPart(WordprocessingContentPart),
}

#[derive(Debug, Clone)]
pub struct WordprocessingGroup {
    pub non_visual_drawing_props: Option<drawingml::NonVisualDrawingProps>,
    pub non_visual_drawing_shape_props: drawingml::NonVisualDrawingShapeProps,
    pub group_shape_props: drawingml::GroupShapeProperties,
    pub shapes: Vec<WordprocessingGroupChoice>,
}

pub enum WordprocessingCanvasChoice {
    Shape(WordprocessingShape),
    Picture(drawingml::Picture),
    ContentPart(WordprocessingContentPart),
    Group(WordprocessingGroup),
    GraphicFrame(GraphicFrame),
}

pub struct WordprocessingCanvas {
    pub background_formatting: Option<drawingml::BackgroundFormatting>,
    pub whole_formatting: Option<drawingml::WholeE2oFormatting>,
    pub shapes: Vec<WordprocessingCanvasChoice>,
}
