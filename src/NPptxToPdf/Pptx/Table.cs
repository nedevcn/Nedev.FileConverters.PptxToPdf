using System.Xml.Linq;

namespace NPptxToPdf.Pptx;

public class Table
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public string? Id { get; }
    public string? Name { get; }
    public Rect Bounds { get; }
    public List<TableRow> Rows { get; } = new();
    public List<TableColumn> Columns { get; } = new();
    public TableProperties? Properties { get; }

    public Table(XElement element)
    {
        _element = element;

        var nvGraphicFramePr = element.Element(P + "nvGraphicFramePr");
        if (nvGraphicFramePr != null)
        {
            var cNvPr = nvGraphicFramePr.Element(P + "cNvPr");
            Id = cNvPr?.Attribute("id")?.Value;
            Name = cNvPr?.Attribute("name")?.Value;
        }

        var xfrm = element.Element(P + "xfrm");
        if (xfrm != null)
        {
            Bounds = ParseBounds(xfrm);
        }

        var tbl = element.Element(A + "tbl");
        if (tbl != null)
        {
            Properties = ParseTableProperties(tbl.Element(A + "tblPr"));

            // Parse columns
            var tblGrid = tbl.Element(A + "tblGrid");
            if (tblGrid != null)
            {
                foreach (var gridCol in tblGrid.Elements(A + "gridCol"))
                {
                    var width = long.TryParse(gridCol.Attribute("w")?.Value, out var w) ? w : 0;
                    Columns.Add(new TableColumn { Width = width });
                }
            }

            // Parse rows
            foreach (var tr in tbl.Elements(A + "tr"))
            {
                var row = new TableRow();

                var height = tr.Attribute("h")?.Value;
                if (height != null)
                {
                    row.Height = long.TryParse(height, out var h) ? h : 0;
                }

                // Parse cells
                foreach (var tc in tr.Elements(A + "tc"))
                {
                    row.Cells.Add(ParseCell(tc));
                }

                Rows.Add(row);
            }
        }
    }

    private static Rect ParseBounds(XElement xfrm)
    {
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

    private static TableProperties? ParseTableProperties(XElement? tblPr)
    {
        if (tblPr == null) return null;

        var props = new TableProperties();

        // Table direction
        var rtl = tblPr.Attribute("rtl");
        props.RightToLeft = rtl?.Value == "1";

        // First row
        var firstRow = tblPr.Attribute("firstRow");
        props.HasHeaderRow = firstRow?.Value == "1";

        // First column
        var firstCol = tblPr.Attribute("firstCol");
        props.HasHeaderColumn = firstCol?.Value == "1";

        // Last row
        var lastRow = tblPr.Attribute("lastRow");
        props.HasTotalRow = lastRow?.Value == "1";

        // Last column
        var lastCol = tblPr.Attribute("lastCol");
        props.HasLastColumn = lastCol?.Value == "1";

        // Band rows
        var bandRow = tblPr.Attribute("bandRow");
        props.BandRows = bandRow?.Value != "0";

        // Band columns
        var bandCol = tblPr.Attribute("bandCol");
        props.BandColumns = bandCol?.Value == "1";

        // Table style ID
        var tableStyleId = tblPr.Element(A + "tableStyleId");
        if (tableStyleId != null)
        {
            props.StyleId = tableStyleId.Value;
        }

        return props;
    }

    private static TableCell ParseCell(XElement tc)
    {
        var cell = new TableCell();

        // Row span
        var rowSpan = tc.Attribute("rowSpan");
        if (rowSpan != null)
        {
            cell.RowSpan = int.TryParse(rowSpan.Value, out var rs) ? rs : 1;
        }

        // Column span
        var gridSpan = tc.Attribute("gridSpan");
        if (gridSpan != null)
        {
            cell.ColumnSpan = int.TryParse(gridSpan.Value, out var gs) ? gs : 1;
        }

        // Horizontal merge
        var hMerge = tc.Attribute("hMerge");
        cell.HorizontalMerge = hMerge != null;

        // Vertical merge
        var vMerge = tc.Attribute("vMerge");
        cell.VerticalMerge = vMerge != null;

        // Cell properties
        var tcPr = tc.Element(A + "tcPr");
        if (tcPr != null)
        {
            cell.Properties = ParseCellProperties(tcPr);
        }

        // Text body
        var txBody = tc.Element(P + "txBody");
        if (txBody != null)
        {
            cell.TextBody = txBody;
            cell.Paragraphs = ParseCellParagraphs(txBody);
        }

        return cell;
    }

    private static CellProperties? ParseCellProperties(XElement tcPr)
    {
        var props = new CellProperties();

        // Anchor
        var anchor = tcPr.Attribute("anchor");
        if (anchor != null)
        {
            props.Anchor = anchor.Value switch
            {
                "t" => TextAnchor.Top,
                "ctr" => TextAnchor.Middle,
                "b" => TextAnchor.Bottom,
                _ => TextAnchor.Middle
            };
        }

        // Anchor center
        var anchorCtr = tcPr.Attribute("anchorCtr");
        props.AnchorCenter = anchorCtr?.Value == "1";

        // Margins
        var marL = tcPr.Attribute("marL");
        if (marL != null)
            props.LeftMargin = long.TryParse(marL.Value, out var l) ? l : 91440;

        var marR = tcPr.Attribute("marR");
        if (marR != null)
            props.RightMargin = long.TryParse(marR.Value, out var r) ? r : 91440;

        var marT = tcPr.Attribute("marT");
        if (marT != null)
            props.TopMargin = long.TryParse(marT.Value, out var t) ? t : 45720;

        var marB = tcPr.Attribute("marB");
        if (marB != null)
            props.BottomMargin = long.TryParse(marB.Value, out var b) ? b : 45720;

        // No wrap
        var noWrap = tcPr.Attribute("noWrap");
        props.NoWrap = noWrap?.Value == "1";

        // Fill
        var solidFill = tcPr.Element(A + "solidFill");
        if (solidFill != null)
        {
            var color = Shape.ParseColor(solidFill);
            if (color.HasValue)
                props.Fill = new Fill { Type = FillType.Solid, Color = color.Value };
        }

        // Cell borders
        var tcBorders = tcPr.Element(A + "tcBorders");
        if (tcBorders != null)
        {
            props.Borders = new CellBorders
            {
                Left = ParseBorder(tcBorders.Element(A + "left")),
                Right = ParseBorder(tcBorders.Element(A + "right")),
                Top = ParseBorder(tcBorders.Element(A + "top")),
                Bottom = ParseBorder(tcBorders.Element(A + "bottom")),
                InsideHorizontal = ParseBorder(tcBorders.Element(A + "insideH")),
                InsideVertical = ParseBorder(tcBorders.Element(A + "insideV"))
            };
        }

        return props;
    }

    private static CellBorder? ParseBorder(XElement? borderElement)
    {
        if (borderElement == null) return null;

        var border = new CellBorder();

        var width = borderElement.Attribute("w");
        if (width != null)
            border.Width = int.TryParse(width.Value, out var w) ? w : 12700;

        var cap = borderElement.Attribute("cap");
        border.LineCap = cap?.Value switch
        {
            "rnd" => LineCap.Round,
            "sq" => LineCap.Square,
            _ => LineCap.Flat
        };

        var cmpd = borderElement.Attribute("cmpd");
        border.CompoundType = cmpd?.Value switch
        {
            "dbl" => CompoundType.Double,
            "thickThin" => CompoundType.ThickThin,
            "thinThick" => CompoundType.ThinThick,
            "tri" => CompoundType.Triple,
            _ => CompoundType.Single
        };

        var algn = borderElement.Attribute("algn");
        border.Alignment = algn?.Value switch
        {
            "ctr" => LineAlignment.Center,
            "in" => LineAlignment.Inside,
            _ => LineAlignment.Outside
        };

        // Dash type
        var prstDash = borderElement.Element(A + "prstDash");
        if (prstDash != null)
        {
            var val = prstDash.Attribute("val")?.Value;
            border.DashType = val switch
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

        // Border color
        var solidFill = borderElement.Element(A + "solidFill");
        if (solidFill != null)
        {
            border.Color = Shape.ParseColor(solidFill);
        }

        return border;
    }

    private static List<Paragraph> ParseCellParagraphs(XElement txBody)
    {
        var paragraphs = new List<Paragraph>();

        foreach (var p in txBody.Elements(A + "p"))
        {
            var paragraph = new Paragraph();

            var pPr = p.Element(A + "pPr");
            if (pPr != null)
            {
                paragraph.Alignment = ParseTextAlignment(pPr.Attribute("algn")?.Value);
                paragraph.Level = int.TryParse(pPr.Attribute("lvl")?.Value, out var lvl) ? lvl : 0;
            }

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

            paragraphs.Add(paragraph);
        }

        return paragraphs;
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

    private static RunProperties ParseRunProperties(XElement rPr)
    {
        var props = new RunProperties();

        var sz = rPr.Attribute("sz");
        if (sz != null && int.TryParse(sz.Value, out var fontSize))
        {
            props.FontSize = fontSize / 100;
        }

        var b = rPr.Attribute("b");
        props.Bold = b?.Value == "1";

        var i = rPr.Attribute("i");
        props.Italic = i?.Value == "1";

        var latin = rPr.Element(A + "latin");
        if (latin != null)
        {
            props.FontFamily = latin.Attribute("typeface")?.Value;
        }

        var solidFill = rPr.Element(A + "solidFill");
        if (solidFill != null)
        {
            props.Color = Shape.ParseColor(solidFill);
        }

        return props;
    }
}

public class TableRow
{
    public long Height { get; set; }
    public List<TableCell> Cells { get; set; } = new();
}

public class TableColumn
{
    public long Width { get; set; }
}

public class TableCell
{
    public int RowSpan { get; set; } = 1;
    public int ColumnSpan { get; set; } = 1;
    public bool HorizontalMerge { get; set; }
    public bool VerticalMerge { get; set; }
    public CellProperties? Properties { get; set; }
    public XElement? TextBody { get; set; }
    public List<Paragraph> Paragraphs { get; set; } = new();
}

public class TableProperties
{
    public bool RightToLeft { get; set; }
    public bool HasHeaderRow { get; set; }
    public bool HasHeaderColumn { get; set; }
    public bool HasTotalRow { get; set; }
    public bool HasLastColumn { get; set; }
    public bool BandRows { get; set; } = true;
    public bool BandColumns { get; set; }
    public string? StyleId { get; set; }
}

public class CellProperties
{
    public TextAnchor Anchor { get; set; } = TextAnchor.Middle;
    public bool AnchorCenter { get; set; }
    public long LeftMargin { get; set; } = 91440;
    public long RightMargin { get; set; } = 91440;
    public long TopMargin { get; set; } = 45720;
    public long BottomMargin { get; set; } = 45720;
    public bool NoWrap { get; set; }
    public Fill? Fill { get; set; }
    public CellBorders? Borders { get; set; }
}

public class CellBorders
{
    public CellBorder? Left { get; set; }
    public CellBorder? Right { get; set; }
    public CellBorder? Top { get; set; }
    public CellBorder? Bottom { get; set; }
    public CellBorder? InsideHorizontal { get; set; }
    public CellBorder? InsideVertical { get; set; }
}

public class CellBorder
{
    public int Width { get; set; } = 12700;
    public LineCap LineCap { get; set; } = LineCap.Flat;
    public CompoundType CompoundType { get; set; } = CompoundType.Single;
    public LineAlignment Alignment { get; set; } = LineAlignment.Outside;
    public LineDashType DashType { get; set; } = LineDashType.Solid;
    public Color? Color { get; set; }
}
