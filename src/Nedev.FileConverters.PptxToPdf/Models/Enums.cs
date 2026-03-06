namespace Nedev.FileConverters.PptxToPdf;

public enum ShapeType
{
    // ??????
    Rectangle,
    Ellipse,
    RoundRectangle,
    Triangle,
    Line,
    AutoShape,
    Custom,
    
    // MS-PPT ?????? 186 ????????
    Parallelogram,
    Trapezoid,
    Diamond,
    Pentagon,
    Hexagon,
    Heptagon,
    Octagon,
    Decagon,
    Dodecagon,
    Pie,
    Chord,
    Teardrop,
    Frame,
    HalfFrame,
    LShape,
    DiagonalStripe,
    Cross,
    Plus,
    Plaque,
    Can,
    Cube,
    Bevel,
    Donut,
    NoSmoking,
    BlockArc,
    FoldedCorner,
    SmileyFace,
    Heart,
    LightningBolt,
    Sun,
    Moon,
    Cloud,
    Arc,
    BracketPair,
    BracePair,
    LeftBracket,
    RightBracket,
    LeftBrace,
    RightBrace,
    
    // ???
    RightArrow,
    LeftArrow,
    UpArrow,
    DownArrow,
    LeftRightArrow,
    UpDownArrow,
    QuadArrow,
    LeftRightUpArrow,
    BentArrow,
    UTurnArrow,
    LeftUpArrow,
    BentUpArrow,
    CurvedRightArrow,
    CurvedLeftArrow,
    CurvedUpArrow,
    CurvedDownArrow,
    StripedRightArrow,
    NotchedRightArrow,
    PentagonArrow,
    Chevron,
    RightArrowCallout,
    LeftArrowCallout,
    UpArrowCallout,
    DownArrowCallout,
    LeftRightArrowCallout,
    UpDownArrowCallout,
    QuadArrowCallout,
    CircularArrow,
    
    // ?????
    FlowChartProcess,
    FlowChartDecision,
    FlowChartInputOutput,
    FlowChartPredefinedProcess,
    FlowChartInternalStorage,
    FlowChartDocument,
    FlowChartMultidocument,
    FlowChartTerminator,
    FlowChartPreparation,
    FlowChartManualInput,
    FlowChartManualOperation,
    FlowChartConnector,
    FlowChartPunchedCard,
    FlowChartPunchedTape,
    FlowChartSummingJunction,
    FlowChartOr,
    FlowChartCollate,
    FlowChartSort,
    FlowChartExtract,
    FlowChartMerge,
    FlowChartOfflineStorage,
    FlowChartOnlineStorage,
    FlowChartMagneticTape,
    FlowChartMagneticDisk,
    FlowChartMagneticDrum,
    FlowChartDisplay,
    FlowChartDelay,
    FlowChartAlternateProcess,
    FlowChartOffpageConnector,
    
    // ???
    RectangularCallout,
    RoundedRectangularCallout,
    OvalCallout,
    CloudCallout,
    LineCallout1,
    LineCallout2,
    LineCallout3,
    LineCallout1AccentBar,
    LineCallout2AccentBar,
    LineCallout3AccentBar,
    LineCallout1NoBorder,
    LineCallout2NoBorder,
    LineCallout3NoBorder,
    LineCallout1BorderAndAccentBar,
    LineCallout2BorderAndAccentBar,
    LineCallout3BorderAndAccentBar,
    
    // ????????
    Star4,
    Star5,
    Star6,
    Star7,
    Star8,
    Star10,
    Star12,
    Star16,
    Star24,
    Star32,
    Ribbon,
    Ribbon2,
    EllipseRibbon,
    EllipseRibbon2,
    VerticalScroll,
    HorizontalScroll,
    Wave,
    DoubleWave,
    
    // ??????
    ActionButtonBlank,
    ActionButtonHome,
    ActionButtonHelp,
    ActionButtonInformation,
    ActionButtonForwardNext,
    ActionButtonBackPrevious,
    ActionButtonEnd,
    ActionButtonBeginning,
    ActionButtonReturn,
    ActionButtonDocument,
    ActionButtonSound,
    ActionButtonMovie,
    
    // ?????
    StraightConnector,
    ElbowConnector,
    CurvedConnector,
    
    // ??????
    Group,
    
    // ???
    Table,
    
    // ???
    Picture,
    
    // ???
    Chart,
    
    // OLE ???
    OleObject,
    
    // ???
    Media,
    
    // ?????
    TextBox
}

public enum FillType
{
    None,
    Solid,
    Gradient,
    Pattern,
    Picture,
    Group
}

public enum GradientType
{
    Linear,
    Radial,
    Rectangular,
    Path
}

public enum PatternType
{
    Cross,
    CrossDiag,
    DiagBrick,
    Divot,
    DkDnDiag,
    DkHorz,
    DkUpDiag,
    DkVert,
    DnDiag,
    DotDmnd,
    DotGrid,
    Horz,
    HorzBrick,
    LgCheck,
    LgConfetti,
    LgGrid,
    LtDnDiag,
    LtHorz,
    LtUpDiag,
    LtVert,
    NarHorz,
    NarVert,
    OpenDmnd,
    pct10,
    pct20,
    pct25,
    pct30,
    pct40,
    pct5,
    pct50,
    pct60,
    pct70,
    pct75,
    pct80,
    pct90,
    Plaid,
    Shingle,
    SmCheck,
    SmConfetti,
    SmGrid,
    SolidDmnd,
    Sphere,
    Trellis,
    UpDiag,
    Vert,
    Wave,
    Weave,
    ZigZag
}

public enum TextAlignment
{
    Left,
    Center,
    Right,
    Justify,
    Distributed
}

public enum LineCap
{
    Flat,
    Round,
    Square
}

public enum LineJoin
{
    Miter,
    Round,
    Bevel
}

public enum LineDashType
{
    Solid,
    Dot,
    Dash,
    DashDot,
    DashDotDot,
    SystemDot,
    SystemDash,
    SystemDashDot
}

public enum BulletType
{
    None,
    AutoNumber,
    Char,
    Blip
}

public enum TextDirection
{
    Horizontal,
    Vertical,
    Vertical270,
    WordArtVertical,
    EastAsianVertical,
    MongolianVertical,
    WordArtRightToLeft
}

public enum TextAnchor
{
    Top,
    Middle,
    Bottom,
    TopCentered,
    MiddleCentered,
    BottomCentered
}

public enum UnderlineType
{
    None,
    Single,
    Double,
    SingleAccounting,
    DoubleAccounting,
    Words
}

public enum StrikeType
{
    None,
    Single,
    Double
}

public enum CapsType
{
    None,
    Small,
    All
}

public enum SchemeColor
{
    Background1,
    Text1,
    Background2,
    Text2,
    Accent1,
    Accent2,
    Accent3,
    Accent4,
    Accent5,
    Accent6,
    Hyperlink,
    FollowedHyperlink,
    Dark1,
    Light1,
    Dark2,
    Light2
}

public enum PresetColor
{
    AliceBlue, AntiqueWhite, Aqua, Aquamarine, Azure, Beige, Bisque, Black,
    BlanchedAlmond, Blue, BlueViolet, Brown, BurlyWood, CadetBlue, Chartreuse,
    Chocolate, Coral, CornflowerBlue, Cornsilk, Crimson, Cyan, DarkBlue,
    DarkCyan, DarkGoldenrod, DarkGray, DarkGreen, DarkKhaki, DarkMagenta,
    DarkOliveGreen, DarkOrange, DarkOrchid, DarkRed, DarkSalmon, DarkSeaGreen,
    DarkSlateBlue, DarkSlateGray, DarkTurquoise, DarkViolet, DeepPink,
    DeepSkyBlue, DimGray, DodgerBlue, Firebrick, FloralWhite, ForestGreen,
    Fuchsia, Gainsboro, GhostWhite, Gold, Goldenrod, Gray, Green, GreenYellow,
    Honeydew, HotPink, IndianRed, Indigo, Ivory, Khaki, Lavender, LavenderBlush,
    LawnGreen, LemonChiffon, LightBlue, LightCoral, LightCyan, LightGoldenrodYellow,
    LightGray, LightGreen, LightPink, LightSalmon, LightSeaGreen, LightSkyBlue,
    LightSlateGray, LightSteelBlue, LightYellow, Lime, LimeGreen, Linen,
    Magenta, Maroon, MediumAquamarine, MediumBlue, MediumOrchid, MediumPurple,
    MediumSeaGreen, MediumSlateBlue, MediumSpringGreen, MediumTurquoise,
    MediumVioletRed, MidnightBlue, MintCream, MistyRose, Moccasin, NavajoWhite,
    Navy, OldLace, Olive, OliveDrab, Orange, OrangeRed, Orchid, PaleGoldenrod,
    PaleGreen, PaleTurquoise, PaleVioletRed, PapayaWhip, PeachPuff, Peru,
    Pink, Plum, PowderBlue, Purple, Red, RosyBrown, RoyalBlue, SaddleBrown,
    Salmon, SandyBrown, SeaGreen, SeaShell, Sienna, Silver, SkyBlue, SlateBlue,
    SlateGray, Snow, SpringGreen, SteelBlue, Tan, Teal, Thistle, Tomato,
    Turquoise, Violet, Wheat, White, WhiteSmoke, Yellow, YellowGreen
}

public enum SlideLayoutType
{
    Title,
    Text,
    TwoColumnText,
    Table,
    TextAndChart,
    ChartAndText,
    Diagram,
    Chart,
    TextAndClipArt,
    ClipArtAndText,
    TitleOnly,
    Blank,
    TextAndObject,
    ObjectAndText,
    LargeObject,
    Object,
    TitleSlide,
    TitleAndObject,
    TitleAndMedia,
    MediaAndTitle,
    ObjectOverText,
    TextOverObject,
    TextAndTwoObjects,
    TwoObjectsAndText,
    TwoObjectsOverText,
    FourObjects,
    VerticalTitleAndText,
    VerticalTwoColumnText,
    Custom
}

public enum PlaceholderType
{
    None,
    Title,
    Body,
    CenterTitle,
    SubTitle,
    Date,
    SlideNumber,
    Footer,
    Header,
    Object,
    Chart,
    Table,
    ClipArt,
    SmartArt,
    Media,
    Picture,
    VerticalObject,
    VerticalTitle,
    VerticalBody,
    SlideImage
}

public enum TransitionType
{
    None,
    Cut,
    Fade,
    Push,
    Wipe,
    Split,
    Reveal,
    RandomBars,
    Cover,
    Uncover,
    Clock,
    Zoom,
    Morph
}
