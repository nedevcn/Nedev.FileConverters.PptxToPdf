using System.Xml.Linq;

namespace NPptxToPdf.Pptx;

public class Shape
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";
    private static readonly XNamespace R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    public string? Id { get; }
    public string? Name { get; }
    public ShapeType ShapeType { get; }
    public bool IsPlaceholder { get; }
    public PlaceholderType PlaceholderType { get; }
    public bool HasText => !string.IsNullOrEmpty(Text);

    public Rect Bounds { get; }
    public string? Text { get; }
    public TextProperties? TextProperties { get; }
    public Fill? Fill { get; }
    public Outline? Outline { get; }
    public List<Paragraph> Paragraphs { get; } = new();
    public ShapeEffects? Effects { get; }
    public Transform2D? Transform { get; }
    public Geometry? Geometry { get; }
    public int? ZOrder { get; }
    public Hyperlink? Hyperlink { get; }

    public Shape(XElement element)
    {
        _element = element;

        var nvSpPr = element.Element(P + "nvSpPr");
        if (nvSpPr != null)
        {
            var cNvPr = nvSpPr.Element(P + "cNvPr");
            Id = cNvPr?.Attribute("id")?.Value;
            Name = cNvPr?.Attribute("name")?.Value;
            ZOrder = int.TryParse(cNvPr?.Attribute("id")?.Value, out var z) ? z : null;

            var nvPr = nvSpPr.Element(P + "nvPr");
            if (nvPr != null)
            {
                var ph = nvPr.Element(P + "ph");
                IsPlaceholder = ph != null;
                if (ph != null)
                {
                    PlaceholderType = ParsePlaceholderType(ph.Attribute("type")?.Value);
                }
            }
        }

        var spPr = element.Element(P + "spPr");
        if (spPr != null)
        {
            ShapeType = ParseShapeType(spPr);
            Bounds = ParseBounds(spPr);
            Fill = ParseFill(spPr);
            Outline = ParseOutline(spPr);
            Effects = ParseEffects(spPr);
            Transform = ParseTransform(spPr);
            Geometry = ParseGeometry(spPr);
        }
        else
        {
            Bounds = new Rect();
        }

        var txBody = element.Element(P + "txBody");
        if (txBody != null)
        {
            Text = ParseText(txBody);
            TextProperties = ParseTextProperties(txBody);
            Paragraphs = ParseParagraphs(txBody);
        }

        // Parse hyperlink
        var hlinkClick = element.Element(P + "nvSpPr")?.Element(P + "cNvPr")?.Element(A + "hlinkClick");
        if (hlinkClick != null)
        {
            Hyperlink = new Hyperlink(hlinkClick);
        }
    }

    private static ShapeType ParseShapeType(XElement spPr)
    {
        var prstGeom = spPr.Element(A + "prstGeom");
        if (prstGeom != null)
        {
            var prst = prstGeom.Attribute("prst")?.Value;
            return ShapeTypeMapping.FromPresetGeometry(prst);
        }

        if (spPr.Element(A + "custGeom") != null)
            return ShapeType.Custom;

        return ShapeType.Rectangle;
    }

    private static PlaceholderType ParsePlaceholderType(string? type)
    {
        return type?.ToLower() switch
        {
            "title" => PlaceholderType.Title,
            "body" => PlaceholderType.Body,
            "ctrTitle" => PlaceholderType.CenterTitle,
            "subTitle" => PlaceholderType.SubTitle,
            "dt" => PlaceholderType.Date,
            "sldNum" => PlaceholderType.SlideNumber,
            "ftr" => PlaceholderType.Footer,
            "hdr" => PlaceholderType.Header,
            "obj" => PlaceholderType.Object,
            "chart" => PlaceholderType.Chart,
            "tbl" => PlaceholderType.Table,
            "clipArt" => PlaceholderType.ClipArt,
            "dgm" => PlaceholderType.SmartArt,
            "media" => PlaceholderType.Media,
            "pic" => PlaceholderType.Picture,
            "sldImg" => PlaceholderType.SlideImage,
            _ => PlaceholderType.None
        };
    }

    private static Rect ParseBounds(XElement spPr)
    {
        var xfrm = spPr.Element(A + "xfrm");
        if (xfrm == null) return new Rect();

        var off = xfrm.Element(A + "off");
        var ext = xfrm.Element(A + "ext");

        if (off == null || ext == null) return new Rect();

        return new Rect
        {
            X = long.TryParse(off.Attribute("x")?.Value, out var x) ? x : 0,
            Y = long.TryParse(off.Attribute("y")?.Value, out var y) ? y : 0,
            Width = long.TryParse(ext.Attribute("cx")?.Value, out var w) ? w : 0,
            Height = long.TryParse(ext.Attribute("cy")?.Value, out var h) ? h : 0
        };
    }

    public static Fill? ParseFill(XElement spPr)
    {
        // Solid fill
        var solidFill = spPr.Element(A + "solidFill");
        if (solidFill != null)
        {
            var color = ParseColor(solidFill);
            if (color.HasValue)
                return new Fill { Type = FillType.Solid, Color = color.Value };
        }

        // Gradient fill
        var gradFill = spPr.Element(A + "gradFill");
        if (gradFill != null)
        {
            return ParseGradientFill(gradFill);
        }

        // Pattern fill
        var pattFill = spPr.Element(A + "pattFill");
        if (pattFill != null)
        {
            return ParsePatternFill(pattFill);
        }

        // Picture fill
        var blipFill = spPr.Element(A + "blipFill");
        if (blipFill != null)
        {
            return ParsePictureFill(blipFill);
        }

        // Group fill
        var grpFill = spPr.Element(A + "grpFill");
        if (grpFill != null)
        {
            return new Fill { Type = FillType.Group };
        }

        // No fill
        var noFill = spPr.Element(A + "noFill");
        if (noFill != null)
            return new Fill { Type = FillType.None };

        return null;
    }

    private static Fill ParseGradientFill(XElement gradFill)
    {
        var fill = new Fill { Type = FillType.Gradient };

        var gsLst = gradFill.Element(A + "gsLst");
        if (gsLst != null)
        {
            fill.GradientStops = new List<GradientStop>();
            foreach (var gs in gsLst.Elements(A + "gs"))
            {
                var pos = int.TryParse(gs.Attribute("pos")?.Value, out var p) ? p / 1000.0 : 0;
                var color = ParseColor(gs);
                if (color.HasValue)
                {
                    fill.GradientStops.Add(new GradientStop { Position = pos, Color = color.Value });
                }
            }
        }

        var lin = gradFill.Element(A + "lin");
        if (lin != null)
        {
            fill.GradientType = GradientType.Linear;
            fill.GradientAngle = int.TryParse(lin.Attribute("ang")?.Value, out var ang) ? ang / 60000 : 0;
        }

        var path = gradFill.Element(A + "path");
        if (path != null)
        {
            var pathType = path.Attribute("path")?.Value;
            fill.GradientType = pathType switch
            {
                "circle" => GradientType.Radial,
                "rect" => GradientType.Rectangular,
                "shape" => GradientType.Path,
                _ => GradientType.Linear
            };
        }

        return fill;
    }

    private static Fill ParsePatternFill(XElement pattFill)
    {
        var fill = new Fill { Type = FillType.Pattern };

        var prst = pattFill.Attribute("prst")?.Value;
        if (Enum.TryParse<PatternType>(prst, true, out var patternType))
        {
            fill.PatternType = patternType;
        }

        var fgClr = pattFill.Element(A + "fgClr");
        if (fgClr != null)
        {
            var color = ParseColor(fgClr);
            if (color.HasValue)
                fill.PatternForegroundColor = color.Value;
        }

        var bgClr = pattFill.Element(A + "bgClr");
        if (bgClr != null)
        {
            var color = ParseColor(bgClr);
            if (color.HasValue)
                fill.PatternBackgroundColor = color.Value;
        }

        return fill;
    }

    private static Fill ParsePictureFill(XElement blipFill)
    {
        var fill = new Fill { Type = FillType.Picture };

        var blip = blipFill.Element(A + "blip");
        if (blip != null)
        {
            fill.PictureRelationshipId = blip.Attribute(R + "embed")?.Value;
        }

        var stretch = blipFill.Element(A + "stretch");
        if (stretch != null)
        {
            fill.PictureFillMode = PictureFillMode.Stretch;
        }

        var tile = blipFill.Element(A + "tile");
        if (tile != null)
        {
            fill.PictureFillMode = PictureFillMode.Tile;
        }

        return fill;
    }

    private static Outline? ParseOutline(XElement spPr)
    {
        var ln = spPr.Element(A + "ln");
        if (ln == null) return null;

        var noFill = ln.Element(A + "noFill");
        if (noFill != null)
            return new Outline { Width = 0 };

        var outline = new Outline();

        var width = int.TryParse(ln.Attribute("w")?.Value, out var w) ? w : 12700;
        outline.Width = width;

        var cap = ln.Attribute("cap")?.Value;
        outline.LineCap = cap switch
        {
            "rnd" => LineCap.Round,
            "sq" => LineCap.Square,
            _ => LineCap.Flat
        };

        var cmpd = ln.Attribute("cmpd")?.Value;
        outline.CompoundType = cmpd switch
        {
            "dbl" => CompoundType.Double,
            "thickThin" => CompoundType.ThickThin,
            "thinThick" => CompoundType.ThinThick,
            "tri" => CompoundType.Triple,
            _ => CompoundType.Single
        };

        var algn = ln.Attribute("algn")?.Value;
        outline.Alignment = algn switch
        {
            "ctr" => LineAlignment.Center,
            "in" => LineAlignment.Inside,
            _ => LineAlignment.Outside
        };

        // Dash type
        var prstDash = ln.Element(A + "prstDash");
        if (prstDash != null)
        {
            var val = prstDash.Attribute("val")?.Value;
            outline.DashType = val switch
            {
                "dot" => LineDashType.Dot,
                "dash" => LineDashType.Dash,
                "dashDot" => LineDashType.DashDot,
                "dashDotDot" => LineDashType.DashDotDot,
                "sysDot" => LineDashType.SystemDot,
                "sysDash" => LineDashType.SystemDash,
                "sysDashDot" => LineDashType.SystemDashDot,
                _ => LineDashType.Solid
            };
        }

        // Line color
        var solidFill = ln.Element(A + "solidFill");
        if (solidFill != null)
        {
            outline.Color = ParseColor(solidFill);
        }

        var gradFill = ln.Element(A + "gradFill");
        if (gradFill != null)
        {
            outline.GradientFill = ParseGradientFill(gradFill);
        }

        // Line join
        var round = ln.Element(A + "round");
        if (round != null) outline.LineJoin = LineJoin.Round;

        var bevel = ln.Element(A + "bevel");
        if (bevel != null) outline.LineJoin = LineJoin.Bevel;

        var miter = ln.Element(A + "miter");
        if (miter != null)
        {
            outline.LineJoin = LineJoin.Miter;
            outline.MiterLimit = int.TryParse(miter.Attribute("lim")?.Value, out var lim) ? lim / 1000.0 : 8;
        }

        return outline;
    }

    private static ShapeEffects? ParseEffects(XElement spPr)
    {
        var effectLst = spPr.Element(A + "effectLst");
        if (effectLst == null) return null;

        var effects = new ShapeEffects();

        // Shadow
        var shadow = effectLst.Element(A + "outerShdw") ?? effectLst.Element(A + "innerShdw") ?? effectLst.Element(A + "prstShdw");
        if (shadow != null)
        {
            effects.Shadow = new ShadowEffect
            {
                Type = shadow.Name.LocalName switch
                {
                    "outerShdw" => ShadowType.Outer,
                    "innerShdw" => ShadowType.Inner,
                    "prstShdw" => ShadowType.Preset,
                    _ => ShadowType.Outer
                },
                BlurRadius = int.TryParse(shadow.Attribute("blurRad")?.Value, out var br) ? br / 12700.0 : 0,
                Distance = int.TryParse(shadow.Attribute("dist")?.Value, out var d) ? d / 12700.0 : 0,
                Direction = int.TryParse(shadow.Attribute("dir")?.Value, out var dir) ? dir / 60000 : 0,
                Color = ParseColor(shadow)
            };
        }

        // Reflection
        var reflection = effectLst.Element(A + "reflection");
        if (reflection != null)
        {
            effects.Reflection = new ReflectionEffect
            {
                BlurRadius = int.TryParse(reflection.Attribute("blurRad")?.Value, out var br) ? br / 12700.0 : 0,
                Distance = int.TryParse(reflection.Attribute("dist")?.Value, out var d) ? d / 12700.0 : 0,
                Direction = int.TryParse(reflection.Attribute("dir")?.Value, out var dir) ? dir / 60000 : 0,
                FadeDirection = int.TryParse(reflection.Attribute("fadeDir")?.Value, out var fd) ? fd / 60000 : 0,
                StartOpacity = int.TryParse(reflection.Attribute("stA")?.Value, out var so) ? so / 1000.0 : 1,
                EndOpacity = int.TryParse(reflection.Attribute("endA")?.Value, out var eo) ? eo / 1000.0 : 0
            };
        }

        // Glow
        var glow = effectLst.Element(A + "glow");
        if (glow != null)
        {
            effects.Glow = new GlowEffect
            {
                Radius = int.TryParse(glow.Attribute("rad")?.Value, out var r) ? r / 12700.0 : 0,
                Color = ParseColor(glow)
            };
        }

        // Soft edge
        var softEdge = effectLst.Element(A + "softEdge");
        if (softEdge != null)
        {
            effects.SoftEdge = new SoftEdgeEffect
            {
                Radius = int.TryParse(softEdge.Attribute("rad")?.Value, out var r) ? r / 12700.0 : 0
            };
        }

        return effects;
    }

    private static Transform2D? ParseTransform(XElement spPr)
    {
        var xfrm = spPr.Element(A + "xfrm");
        if (xfrm == null) return null;

        var transform = new Transform2D();

        var rot = xfrm.Attribute("rot")?.Value;
        if (rot != null && int.TryParse(rot, out var rotation))
        {
            transform.Rotation = rotation / 60000.0;
        }

        var flipH = xfrm.Attribute("flipH")?.Value;
        transform.FlipHorizontal = flipH == "1";

        var flipV = xfrm.Attribute("flipV")?.Value;
        transform.FlipVertical = flipV == "1";

        return transform;
    }

    private static Geometry? ParseGeometry(XElement spPr)
    {
        var prstGeom = spPr.Element(A + "prstGeom");
        if (prstGeom == null) return null;

        var geometry = new Geometry
        {
            Preset = prstGeom.Attribute("prst")?.Value
        };

        var avLst = prstGeom.Element(A + "avLst");
        if (avLst != null)
        {
            geometry.Adjustments = new Dictionary<string, double>();
            foreach (var gd in avLst.Elements(A + "gd"))
            {
                var name = gd.Attribute("name")?.Value;
                var fmla = gd.Attribute("fmla")?.Value;
                if (name != null && fmla != null)
                {
                    // Parse formula like "val 16667" or "adj1 50000"
                    var parts = fmla.Split(' ');
                    if (parts.Length >= 2 && double.TryParse(parts[1], out var value))
                    {
                        geometry.Adjustments[name] = value / 100000.0;
                    }
                }
            }
        }

        return geometry;
    }

    private static string? ParseText(XElement txBody)
    {
        var paragraphs = txBody.Elements(A + "p").ToList();
        if (!paragraphs.Any()) return null;

        var texts = new List<string>();
        foreach (var p in paragraphs)
        {
            var paraText = new List<string>();
            foreach (var r in p.Descendants(A + "r"))
            {
                var t = r.Element(A + "t");
                if (t != null)
                    paraText.Add(t.Value);
            }

            if (paraText.Any())
                texts.Add(string.Join("", paraText));
        }

        return texts.Any() ? string.Join("\n", texts) : null;
    }

    private static List<Paragraph> ParseParagraphs(XElement txBody)
    {
        var paragraphs = new List<Paragraph>();

        foreach (var p in txBody.Elements(A + "p"))
        {
            var paragraph = new Paragraph();

            // Paragraph properties
            var pPr = p.Element(A + "pPr");
            if (pPr != null)
            {
                paragraph.Alignment = ParseTextAlignment(pPr.Attribute("algn")?.Value);
                paragraph.Level = int.TryParse(pPr.Attribute("lvl")?.Value, out var lvl) ? lvl : 0;
                paragraph.DefaultTabSize = int.TryParse(pPr.Attribute("defTabSz")?.Value, out var dts) ? dts : 914400;
                paragraph.RightToLeft = pPr.Attribute("rtl")?.Value == "1";
                paragraph.EastAsianLineBreak = pPr.Attribute("eaLnBrk")?.Value != "0";
                paragraph.LatinLineBreak = pPr.Attribute("latinLnBrk")?.Value == "1";
                paragraph.HangingPunctuation = pPr.Attribute("hangingPunct")?.Value == "1";

                // Bullet
                var buNone = pPr.Element(A + "buNone");
                if (buNone != null)
                {
                    paragraph.BulletType = BulletType.None;
                }
                else
                {
                    var buAutoNum = pPr.Element(A + "buAutoNum");
                    if (buAutoNum != null)
                    {
                        paragraph.BulletType = BulletType.AutoNumber;
                        paragraph.BulletAutoNumberType = buAutoNum.Attribute("type")?.Value;
                        paragraph.BulletStartAt = int.TryParse(buAutoNum.Attribute("startAt")?.Value, out var sa) ? sa : 1;
                    }
                    else
                    {
                        var buChar = pPr.Element(A + "buChar");
                        if (buChar != null)
                        {
                            paragraph.BulletType = BulletType.Char;
                            paragraph.BulletChar = buChar.Attribute("char")?.Value;
                        }
                        else
                        {
                            var buBlip = pPr.Element(A + "buBlip");
                            if (buBlip != null)
                            {
                                paragraph.BulletType = BulletType.Blip;
                            }
                        }
                    }

                    // Bullet size and color
                    var buSzPct = pPr.Element(A + "buSzPct");
                    if (buSzPct != null)
                    {
                        paragraph.BulletSize = int.TryParse(buSzPct.Attribute("val")?.Value, out var bs) ? bs / 1000.0 : 1;
                    }

                    var buFont = pPr.Element(A + "buFont");
                    if (buFont != null)
                    {
                        paragraph.BulletFont = buFont.Attribute("typeface")?.Value;
                    }

                    var buClr = pPr.Element(A + "buClr");
                    if (buClr != null)
                    {
                        paragraph.BulletColor = ParseColor(buClr);
                    }
                }

                // Indentation
                var marL = pPr.Attribute("marL");
                if (marL != null)
                    paragraph.MarginLeft = long.TryParse(marL.Value, out var ml) ? ml : 0;

                var marR = pPr.Attribute("marR");
                if (marR != null)
                    paragraph.MarginRight = long.TryParse(marR.Value, out var mr) ? mr : 0;

                var indent = pPr.Attribute("indent");
                if (indent != null)
                    paragraph.Indent = long.TryParse(indent.Value, out var ind) ? ind : 0;

                // Line spacing
                var lnSpc = pPr.Element(A + "lnSpc");
                if (lnSpc != null)
                {
                    paragraph.LineSpacing = ParseSpacing(lnSpc);
                }

                // Space before
                var spcBef = pPr.Element(A + "spcBef");
                if (spcBef != null)
                {
                    paragraph.SpaceBefore = ParseSpacing(spcBef);
                }

                // Space after
                var spcAft = pPr.Element(A + "spcAft");
                if (spcAft != null)
                {
                    paragraph.SpaceAfter = ParseSpacing(spcAft);
                }
            }

            // Runs
            foreach (var r in p.Elements(A + "r"))
            {
                var run = new Run();

                var t = r.Element(A + "t");
                if (t != null)
                {
                    run.Text = t.Value;
                }

                var rPr = r.Element(A + "rPr");
                if (rPr != null)
                {
                    run.Properties = ParseRunProperties(rPr);
                }

                paragraph.Runs.Add(run);
            }

            // Line breaks
            foreach (var br in p.Elements(A + "br"))
            {
                var run = new Run { Text = "\n" };
                var rPr = br.Element(A + "rPr");
                if (rPr != null)
                {
                    run.Properties = ParseRunProperties(rPr);
                }
                paragraph.Runs.Add(run);
            }

            paragraphs.Add(paragraph);
        }

        return paragraphs;
    }

    private static RunProperties ParseRunProperties(XElement rPr)
    {
        var props = new RunProperties();

        // Font size
        var sz = rPr.Attribute("sz");
        if (sz != null && int.TryParse(sz.Value, out var fontSize))
        {
            props.FontSize = fontSize / 100;
        }

        // Bold
        var b = rPr.Attribute("b");
        props.Bold = b?.Value == "1";

        // Italic
        var i = rPr.Attribute("i");
        props.Italic = i?.Value == "1";

        // Underline
        var u = rPr.Attribute("u");
        if (u != null)
        {
            props.Underline = u.Value switch
            {
                "sng" => UnderlineType.Single,
                "dbl" => UnderlineType.Double,
                "sngAccounting" => UnderlineType.SingleAccounting,
                "dblAccounting" => UnderlineType.DoubleAccounting,
                "words" => UnderlineType.Words,
                _ => UnderlineType.None
            };
        }

        // Strikethrough
        var strike = rPr.Attribute("strike");
        if (strike != null)
        {
            props.Strike = strike.Value switch
            {
                "sngStrike" => StrikeType.Single,
                "dblStrike" => StrikeType.Double,
                _ => StrikeType.None
            };
        }

        // Caps
        var cap = rPr.Attribute("cap");
        if (cap != null)
        {
            props.Caps = cap.Value switch
            {
                "small" => CapsType.Small,
                "all" => CapsType.All,
                _ => CapsType.None
            };
        }

        // Font family
        var latin = rPr.Element(A + "latin");
        if (latin != null)
        {
            props.FontFamily = latin.Attribute("typeface")?.Value;
        }

        var ea = rPr.Element(A + "ea");
        if (ea != null)
        {
            props.EastAsianFont = ea.Attribute("typeface")?.Value;
        }

        var cs = rPr.Element(A + "cs");
        if (cs != null)
        {
            props.ComplexScriptFont = cs.Attribute("typeface")?.Value;
        }

        // Color
        var solidFill = rPr.Element(A + "solidFill");
        if (solidFill != null)
        {
            props.Color = ParseColor(solidFill);
        }

        // Highlight
        var highlight = rPr.Element(A + "highlight");
        if (highlight != null)
        {
            props.HighlightColor = ParseColor(highlight);
        }

        // Language
        var lang = rPr.Attribute("lang");
        if (lang != null)
        {
            props.Language = lang.Value;
        }

        // Baseline
        var baseline = rPr.Attribute("baseline");
        if (baseline != null && int.TryParse(baseline.Value, out var bl))
        {
            props.BaselineOffset = bl / 1000.0;
        }

        return props;
    }

    private static Spacing? ParseSpacing(XElement spacingElement)
    {
        var spcPct = spacingElement.Element(A + "spcPct");
        if (spcPct != null && int.TryParse(spcPct.Attribute("val")?.Value, out var pct))
        {
            return new Spacing { Percent = pct / 1000.0 };
        }

        var spcPts = spacingElement.Element(A + "spcPts");
        if (spcPts != null && int.TryParse(spcPts.Attribute("val")?.Value, out var pts))
        {
            return new Spacing { Points = pts / 100 };
        }

        return null;
    }

    private static TextAlignment ParseTextAlignment(string? algn)
    {
        return algn switch
        {
            "ctr" => TextAlignment.Center,
            "r" => TextAlignment.Right,
            "just" => TextAlignment.Justify,
            "dist" => TextAlignment.Distributed,
            _ => TextAlignment.Left
        };
    }

    private static TextProperties? ParseTextProperties(XElement txBody)
    {
        var bodyPr = txBody.Element(A + "bodyPr");
        if (bodyPr == null) return null;

        var props = new TextProperties();

        // Text direction
        var vert = bodyPr.Attribute("vert");
        if (vert != null)
        {
            props.TextDirection = vert.Value switch
            {
                "vert" => TextDirection.Vertical,
                "vert270" => TextDirection.Vertical270,
                "wordArtVert" => TextDirection.WordArtVertical,
                "eaVert" => TextDirection.EastAsianVertical,
                "mongolianVert" => TextDirection.MongolianVertical,
                "wordArtVertRtl" => TextDirection.WordArtRightToLeft,
                _ => TextDirection.Horizontal
            };
        }

        // Anchor
        var anchor = bodyPr.Attribute("anchor");
        if (anchor != null)
        {
            props.Anchor = anchor.Value switch
            {
                "t" => TextAnchor.Top,
                "b" => TextAnchor.Bottom,
                "ctr" => TextAnchor.Middle,
                "just" => TextAnchor.TopCentered,
                "dist" => TextAnchor.BottomCentered,
                _ => TextAnchor.Middle
            };
        }

        // Wrap
        var wrap = bodyPr.Attribute("wrap");
        props.WrapText = wrap?.Value != "none";

        // Margins
        props.LeftInset = int.TryParse(bodyPr.Attribute("lIns")?.Value, out var lIns) ? lIns / 12700.0 : 0.1;
        props.TopInset = int.TryParse(bodyPr.Attribute("tIns")?.Value, out var tIns) ? tIns / 12700.0 : 0.05;
        props.RightInset = int.TryParse(bodyPr.Attribute("rIns")?.Value, out var rIns) ? rIns / 12700.0 : 0.1;
        props.BottomInset = int.TryParse(bodyPr.Attribute("bIns")?.Value, out var bIns) ? bIns / 12700.0 : 0.05;

        // Auto-fit
        var noAutofit = bodyPr.Element(A + "noAutofit");
        if (noAutofit != null)
        {
            props.AutoFit = TextAutoFit.None;
        }
        else
        {
            var normAutofit = bodyPr.Element(A + "normAutofit");
            if (normAutofit != null)
            {
                props.AutoFit = TextAutoFit.Normal;
                props.FontScale = int.TryParse(normAutofit.Attribute("fontScale")?.Value, out var fs) ? fs / 1000.0 : 1;
                props.LineSpaceReduction = int.TryParse(normAutofit.Attribute("lnSpcReduction")?.Value, out var lsr) ? lsr / 1000.0 : 0;
            }
            else
            {
                var spAutoFit = bodyPr.Element(A + "spAutoFit");
                if (spAutoFit != null)
                {
                    props.AutoFit = TextAutoFit.Shape;
                }
            }
        }

        // Columns
        var extLst = bodyPr.Element(A + "extLst");
        if (extLst != null)
        {
            var ext = extLst.Element(A + "ext");
            if (ext != null)
            {
                var spPr = ext.Descendants(A + "spPr").FirstOrDefault();
                if (spPr != null)
                {
                    // Handle column properties
                }
            }
        }

        return props;
    }

    public static Color? ParseColor(XElement parent)
    {
        // sRGB color
        var srgbClr = parent.Element(A + "srgbClr");
        if (srgbClr != null)
        {
            var val = srgbClr.Attribute("val")?.Value;
            if (val != null && val.Length == 6)
            {
                if (byte.TryParse(val.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, null, out var r) &&
                    byte.TryParse(val.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out var g) &&
                    byte.TryParse(val.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
                {
                    var color = new Color(r, g, b);

                    // Apply color transformations
                    var alpha = srgbClr.Element(A + "alpha");
                    if (alpha != null && int.TryParse(alpha.Attribute("val")?.Value, out var a))
                    {
                        color = color.WithAlpha((byte)(a / 1000));
                    }

                    return color;
                }
            }
        }

        // Scheme color
        var schemeClr = parent.Element(A + "schemeClr");
        if (schemeClr != null)
        {
            var val = schemeClr.Attribute("val")?.Value;
            if (Enum.TryParse<SchemeColor>(val, true, out var schemeColor))
            {
                return Color.FromSchemeColor(schemeColor);
            }
        }

        // Preset color
        var prstClr = parent.Element(A + "prstClr");
        if (prstClr != null)
        {
            var val = prstClr.Attribute("val")?.Value;
            if (Enum.TryParse<PresetColor>(val, true, out var presetColor))
            {
                return Color.FromPresetColor(presetColor);
            }
        }

        // System color
        var sysClr = parent.Element(A + "sysClr");
        if (sysClr != null)
        {
            var lastClr = sysClr.Attribute("lastClr")?.Value;
            if (lastClr != null && lastClr.Length == 6)
            {
                if (byte.TryParse(lastClr.Substring(0, 2), System.Globalization.NumberStyles.HexNumber, null, out var r) &&
                    byte.TryParse(lastClr.Substring(2, 2), System.Globalization.NumberStyles.HexNumber, null, out var g) &&
                    byte.TryParse(lastClr.Substring(4, 2), System.Globalization.NumberStyles.HexNumber, null, out var b))
                {
                    return new Color(r, g, b);
                }
            }
        }

        // HSL color
        var hslClr = parent.Element(A + "hslClr");
        if (hslClr != null)
        {
            var hue = int.TryParse(hslClr.Attribute("hue")?.Value, out var h) ? h / 60000 : 0;
            var sat = int.TryParse(hslClr.Attribute("sat")?.Value, out var s) ? s / 1000.0 : 0;
            var lum = int.TryParse(hslClr.Attribute("lum")?.Value, out var l) ? l / 1000.0 : 0;
            return Color.FromHsl(hue, sat, lum);
    }

        return null;
    }
}
