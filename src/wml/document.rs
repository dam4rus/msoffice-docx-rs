use msoffice_shared::drawingml;

pub type UcharHexNumber = String;

#[derive(Debug, Copy, Clone, Eq, EnumString)]
enum ThemeColor {
    #[strum(serialize="dark1")]
    Dark1,
    #[strum(serialize="light1")]
    Light1,
    #[strum(serialize="dark2")]
    Dark2,
    #[strum(serialize="light2")]
    Light2,
    #[strum(serialize="accent1")]
    Accent1,
    #[strum(serialize="accent2")]
    Accent2,
    #[strum(serialize="accent3")]
    Accent3,
    #[strum(serialize="accent4")]
    Accent4,
    #[strum(serialize="accent5")]
    Accent5,
    #[strum(serialize="accent6")]
    Accent6,
    #[strum(serialize="hyperlink")]
    Hyperlink,
    #[strum(serialize="followedHyperlink")]
    FollowedHyperlink,
    #[strum(serialize="none")]
    None,
    #[strum(serialize="background1")]
    Background1,
    #[strum(serialize="text1")]
    Text1,
    #[strum(serialize="background2")]
    Background2,
    #[strum(serialize="text2")]
    Text2,
}

#[derive(Debug, Clone, Eq)]
enum HexColor {
    Auto,
    RGB(drawingml::HexColorRGB),
}

#[derive(Debug, Clone, Default, Eq)]
struct Background {
    pub color: Option<HexColor>,
    pub theme_color: Option<ThemeColor>,
    pub theme_hint: Option<UcharHexNumber>,
    pub theme_shade: Option<UcharHexNumber>,

    pub drawing: Option<Drawing>,
}

/*
<xsd:complexType name="CT_Drawing">
    <xsd:choice minOccurs="1" maxOccurs="unbounded">
      <xsd:element ref="wp:anchor" minOccurs="0"/>
      <xsd:element ref="wp:inline" minOccurs="0"/>
    </xsd:choice>
  </xsd:complexType>
*/
struct Drawing {

}

struct Document {
    pub background: Option<Background>,
    pub body: Option<Body>,
    // pub conformance: Option<ConformanceClass>,
}

#[derive(Debug, Clone, Eq)]
pub enum RunLevelElts {
    ProofError(Option<ProofErr>),
    PermStart(Option<PermStart>),
    PermEnd(Option<PermEnd>),
    RangeMarkupElements(Vec<RangeMarkupElements>),
    Insert(Option<RunTrackChange>),
    Delete(Option<RunTrackChange>),
    MoveFrom(RunTrackChange),
    MoveTo(RunTrackChange),
    MathContent(Vec<MathContent>),
}

pub enum ContentBlockContent {
    CustomXml(CustomXmlBlock),
    StructureDocumentTag(SdtBlock),
    Paragraph(Vec<Paragraph>),
    Table(Vec<Table>),
    RunLevelElements(Vec<RunLevelElts>),
}

pub struct BlockLevelChunksElts {
    pub content_block_contents: Vec<ContentBlockContent>,
}

pub enum BlockLevelElts {
    BlockLevelChunkElements(Vec<BlockLevelChunksElts>),
    AltChunk(Vec<AltChunk>),
}