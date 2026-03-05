using System.Xml.Linq;

namespace Nedev.PptxToPdf.Pptx;

public class Chart
{
    private readonly XElement _element;
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public string? Title { get; }
    public ChartType Type { get; }
    public List<ChartSeries> Series { get; } = new();
    public ChartAxis? CategoryAxis { get; }
    public ChartAxis? ValueAxis { get; }
    public ChartLegend? Legend { get; }
    public ChartPlotArea PlotArea { get; }
    public Rect Bounds { get; }

    public Chart(XElement element, Rect bounds)
    {
        _element = element;
        Bounds = bounds;

        // Parse chart space
        var chartSpace = element.Element(C + "chartSpace");
        if (chartSpace == null) return;

        // Parse title
        var title = chartSpace.Element(C + "title");
        if (title != null)
        {
            Title = ParseTitle(title);
        }

        // Parse chart
        var chart = chartSpace.Element(C + "chart");
        if (chart == null) return;

        // Parse plot area
        var plotArea = chart.Element(C + "plotArea");
        if (plotArea != null)
        {
            PlotArea = new ChartPlotArea(plotArea);
            Type = PlotArea.Type;
            Series.AddRange(PlotArea.Series);
        }

        // Parse axes
        foreach (var axis in chart.Elements(C + "catAx").Concat(chart.Elements(C + "valAx")))
        {
            var chartAxis = new ChartAxis(axis);
            if (chartAxis.AxisType == AxisType.Category)
                CategoryAxis = chartAxis;
            else
                ValueAxis = chartAxis;
        }

        // Parse legend
        var legend = chart.Element(C + "legend");
        if (legend != null)
        {
            Legend = new ChartLegend(legend);
        }
    }

    private static string? ParseTitle(XElement title)
    {
        var tx = title.Element(title.Name.Namespace + "tx");
        var rich = tx?.Element(title.Name.Namespace + "rich");
        var p = rich?.Element(title.Name.Namespace + "p");
        var r = p?.Element(title.Name.Namespace + "r");
        return r?.Element(title.Name.Namespace + "t")?.Value;
    }
}

public enum ChartType
{
    Bar,
    Column,
    Line,
    Pie,
    Area,
    Scatter,
    Radar,
    Surface,
    Doughnut,
    Bubble,
    Stock,
    Combo,
    Unknown
}

public enum AxisType
{
    Category,
    Value,
    Date,
    Series
}

public enum LegendPosition
{
    Left,
    Top,
    Right,
    Bottom,
    TopRight
}

public class ChartPlotArea
{
    private readonly XElement _element;
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public ChartType Type { get; }
    public List<ChartSeries> Series { get; } = new();
    public ChartAxis? CategoryAxis { get; }
    public ChartAxis? ValueAxis { get; }

    public ChartPlotArea(XElement element)
    {
        _element = element;

        // Determine chart type and parse series
        foreach (var child in element.Elements())
        {
            switch (child.Name.LocalName)
            {
                case "barChart":
                    Type = ChartType.Bar;
                    ParseBarChart(child);
                    break;
                case "lineChart":
                    Type = ChartType.Line;
                    ParseLineChart(child);
                    break;
                case "pieChart":
                    Type = ChartType.Pie;
                    ParsePieChart(child);
                    break;
                case "areaChart":
                    Type = ChartType.Area;
                    ParseAreaChart(child);
                    break;
                case "scatterChart":
                    Type = ChartType.Scatter;
                    ParseScatterChart(child);
                    break;
                case "radarChart":
                    Type = ChartType.Radar;
                    ParseRadarChart(child);
                    break;
                case "doughnutChart":
                    Type = ChartType.Doughnut;
                    ParseDoughnutChart(child);
                    break;
                case "bubbleChart":
                    Type = ChartType.Bubble;
                    ParseBubbleChart(child);
                    break;
                case "stockChart":
                    Type = ChartType.Stock;
                    ParseStockChart(child);
                    break;
            }
        }
    }

    private void ParseBarChart(XElement barChart)
    {
        foreach (var ser in barChart.Elements(C + "ser"))
        {
            Series.Add(new ChartSeries(ser, ChartType.Bar));
        }
    }

    private void ParseLineChart(XElement lineChart)
    {
        foreach (var ser in lineChart.Elements(C + "ser"))
        {
            Series.Add(new ChartSeries(ser, ChartType.Line));
        }
    }

    private void ParsePieChart(XElement pieChart)
    {
        foreach (var ser in pieChart.Elements(C + "ser"))
        {
            Series.Add(new ChartSeries(ser, ChartType.Pie));
        }
    }

    private void ParseAreaChart(XElement areaChart)
    {
        foreach (var ser in areaChart.Elements(C + "ser"))
        {
            Series.Add(new ChartSeries(ser, ChartType.Area));
        }
    }

    private void ParseScatterChart(XElement scatterChart)
    {
        foreach (var ser in scatterChart.Elements(C + "ser"))
        {
            Series.Add(new ChartSeries(ser, ChartType.Scatter));
        }
    }

    private void ParseRadarChart(XElement radarChart)
    {
        foreach (var ser in radarChart.Elements(C + "ser"))
        {
            Series.Add(new ChartSeries(ser, ChartType.Radar));
        }
    }

    private void ParseDoughnutChart(XElement doughnutChart)
    {
        foreach (var ser in doughnutChart.Elements(C + "ser"))
        {
            Series.Add(new ChartSeries(ser, ChartType.Doughnut));
        }
    }

    private void ParseBubbleChart(XElement bubbleChart)
    {
        foreach (var ser in bubbleChart.Elements(C + "ser"))
        {
            Series.Add(new ChartSeries(ser, ChartType.Bubble));
        }
    }

    private void ParseStockChart(XElement stockChart)
    {
        foreach (var ser in stockChart.Elements(C + "ser"))
        {
            Series.Add(new ChartSeries(ser, ChartType.Stock));
        }
    }
}

public class ChartSeries
{
    private readonly XElement _element;
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";

    public int? Index { get; }
    public int? Order { get; }
    public string? Name { get; }
    public ChartType ChartType { get; }
    public List<ChartDataPoint> DataPoints { get; } = new();
    public ChartDataLabels? DataLabels { get; }
    public Fill? Fill { get; }
    public Outline? Outline { get; }

    public ChartSeries(XElement element, ChartType chartType)
    {
        _element = element;
        ChartType = chartType;

        // Parse index
        var idx = element.Element(C + "idx");
        if (idx != null && int.TryParse(idx.Attribute("val")?.Value, out var index))
            Index = index;

        // Parse order
        var order = element.Element(C + "order");
        if (order != null && int.TryParse(order.Attribute("val")?.Value, out var orderVal))
            Order = orderVal;

        // Parse series name
        var tx = element.Element(C + "tx");
        if (tx != null)
        {
            var strRef = tx.Element(C + "strRef");
            if (strRef != null)
            {
                var f = strRef.Element(C + "f")?.Value;
                if (!string.IsNullOrEmpty(f))
                {
                    // Try to get cached value
                    var strCache = strRef.Element(C + "strCache");
                    var pt = strCache?.Element(C + "pt");
                    Name = pt?.Element(C + "v")?.Value;
                }
            }
            else
            {
                var v = tx.Element(C + "v")?.Value;
                if (!string.IsNullOrEmpty(v))
                    Name = v;
            }
        }

        // Parse data points
        var cat = element.Element(C + "cat");
        var val = element.Element(C + "val");
        var xVal = element.Element(C + "xVal");
        var yVal = element.Element(C + "yVal");

        List<string>? categories = null;
        if (cat != null)
        {
            categories = ParseStringData(cat);
        }

        List<double>? values = null;
        if (val != null)
        {
            values = ParseNumericData(val);
        }
        else if (xVal != null && yVal != null)
        {
            // For scatter charts
            var xValues = ParseNumericData(xVal);
            var yValues = ParseNumericData(yVal);
            values = yValues;
        }

        // Create data points
        if (values != null)
        {
            for (int i = 0; i < values.Count; i++)
            {
                DataPoints.Add(new ChartDataPoint
                {
                    Index = i,
                    Category = categories != null && i < categories.Count ? categories[i] : null,
                    Value = values[i]
                });
            }
        }

        // Parse data labels
        var dLbls = element.Element(C + "dLbls");
        if (dLbls != null)
        {
            DataLabels = new ChartDataLabels(dLbls);
        }

        // Parse shape properties for fill/outline
        var spPr = element.Element(C + "spPr");
        if (spPr != null)
        {
            Fill = Shape.ParseFill(spPr);
            Outline = ParseOutline(spPr);
        }
    }

    private static List<string> ParseStringData(XElement element)
    {
        var result = new List<string>();

        var strRef = element.Element(element.Name.Namespace + "strRef");
        if (strRef != null)
        {
            var strCache = strRef.Element(element.Name.Namespace + "strCache");
            if (strCache != null)
            {
                foreach (var pt in strCache.Elements(element.Name.Namespace + "pt"))
                {
                    var v = pt.Element(element.Name.Namespace + "v")?.Value;
                    if (v != null)
                        result.Add(v);
                }
            }
        }
        else
        {
            var strLit = element.Element(element.Name.Namespace + "strLit");
            if (strLit != null)
            {
                foreach (var pt in strLit.Elements(element.Name.Namespace + "pt"))
                {
                    var v = pt.Element(element.Name.Namespace + "v")?.Value;
                    if (v != null)
                        result.Add(v);
                }
            }
        }

        return result;
    }

    private static List<double> ParseNumericData(XElement element)
    {
        var result = new List<double>();

        var numRef = element.Element(element.Name.Namespace + "numRef");
        if (numRef != null)
        {
            var numCache = numRef.Element(element.Name.Namespace + "numCache");
            if (numCache != null)
            {
                foreach (var pt in numCache.Elements(element.Name.Namespace + "pt"))
                {
                    var v = pt.Element(element.Name.Namespace + "v")?.Value;
                    if (v != null && double.TryParse(v, out var val))
                        result.Add(val);
                }
            }
        }
        else
        {
            var numLit = element.Element(element.Name.Namespace + "numLit");
            if (numLit != null)
            {
                foreach (var pt in numLit.Elements(element.Name.Namespace + "pt"))
                {
                    var v = pt.Element(element.Name.Namespace + "v")?.Value;
                    if (v != null && double.TryParse(v, out var val))
                        result.Add(val);
                }
            }
        }

        return result;
    }

    private static Outline? ParseOutline(XElement spPr)
    {
        var ln = spPr.Element(spPr.Name.Namespace + "ln");
        if (ln == null) return null;

        var noFill = ln.Element(spPr.Name.Namespace + "noFill");
        if (noFill != null)
            return new Outline { Width = 0 };

        var width = int.TryParse(ln.Attribute("w")?.Value, out var w) ? w : 12700;

        var solidFill = ln.Element(spPr.Name.Namespace + "solidFill");
        Color? color = null;
        if (solidFill != null)
        {
            color = Shape.ParseColor(solidFill);
        }

        return new Outline { Width = width, Color = color };
    }
}

public class ChartDataPoint
{
    public int Index { get; set; }
    public string? Category { get; set; }
    public double Value { get; set; }
    public Fill? Fill { get; set; }
    public Outline? Outline { get; set; }
}

public class ChartDataLabels
{
    private readonly XElement _element;
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public bool ShowCategoryName { get; }
    public bool ShowSeriesName { get; }
    public bool ShowValue { get; }
    public bool ShowPercentage { get; }
    public bool ShowBubbleSize { get; }
    public string? Separator { get; }

    public ChartDataLabels(XElement element)
    {
        _element = element;

        ShowCategoryName = element.Element(C + "showCatName")?.Attribute("val")?.Value == "1";
        ShowSeriesName = element.Element(C + "showSerName")?.Attribute("val")?.Value == "1";
        ShowValue = element.Element(C + "showVal")?.Attribute("val")?.Value == "1";
        ShowPercentage = element.Element(C + "showPercent")?.Attribute("val")?.Value == "1";
        ShowBubbleSize = element.Element(C + "showBubbleSize")?.Attribute("val")?.Value == "1";
        Separator = element.Element(C + "separator")?.Attribute("val")?.Value;
    }
}

public class ChartAxis
{
    private readonly XElement _element;
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public int AxisId { get; }
    public AxisType AxisType { get; }
    public string? Title { get; }
    public double? MinValue { get; }
    public double? MaxValue { get; }
    public double? MajorUnit { get; }
    public double? MinorUnit { get; }
    public bool Delete { get; }
    public bool MajorGridlines { get; }
    public bool MinorGridlines { get; }

    public ChartAxis(XElement element)
    {
        _element = element;

        // Determine axis type
        AxisType = element.Name.LocalName switch
        {
            "catAx" => AxisType.Category,
            "valAx" => AxisType.Value,
            "dateAx" => AxisType.Date,
            "serAx" => AxisType.Series,
            _ => AxisType.Value
        };

        // Parse axis ID
        var axId = element.Element(C + "axId");
        if (axId != null && int.TryParse(axId.Attribute("val")?.Value, out var id))
            AxisId = id;

        // Parse title
        var title = element.Element(C + "title");
        if (title != null)
        {
            var tx = title.Element(C + "tx");
            var rich = tx?.Element(C + "rich");
            var p = rich?.Element(C + "p");
            var r = p?.Element(C + "r");
            Title = r?.Element(C + "t")?.Value;
        }

        // Parse scaling
        var scaling = element.Element(C + "scaling");
        if (scaling != null)
        {
            var min = scaling.Element(C + "min");
            if (min != null && double.TryParse(min.Attribute("val")?.Value, out var minVal))
                MinValue = minVal;

            var max = scaling.Element(C + "max");
            if (max != null && double.TryParse(max.Attribute("val")?.Value, out var maxVal))
                MaxValue = maxVal;
        }

        // Parse major unit
        var majorUnit = element.Element(C + "majorUnit");
        if (majorUnit != null && double.TryParse(majorUnit.Attribute("val")?.Value, out var major))
            MajorUnit = major;

        // Parse minor unit
        var minorUnit = element.Element(C + "minorUnit");
        if (minorUnit != null && double.TryParse(minorUnit.Attribute("val")?.Value, out var minor))
            MinorUnit = minor;

        // Parse delete
        var delete = element.Element(C + "delete");
        Delete = delete?.Attribute("val")?.Value == "1";

        // Parse major gridlines
        MajorGridlines = element.Element(C + "majorGridlines") != null;

        // Parse minor gridlines
        MinorGridlines = element.Element(C + "minorGridlines") != null;
    }
}

public class ChartLegend
{
    private readonly XElement _element;
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";

    public LegendPosition Position { get; }
    public bool Overlay { get; }

    public ChartLegend(XElement element)
    {
        _element = element;

        // Parse position
        var legendPos = element.Element(C + "legendPos");
        Position = legendPos?.Attribute("val")?.Value switch
        {
            "l" => LegendPosition.Left,
            "t" => LegendPosition.Top,
            "r" => LegendPosition.Right,
            "b" => LegendPosition.Bottom,
            "tr" => LegendPosition.TopRight,
            _ => LegendPosition.Right
        };

        // Parse overlay
        var overlay = element.Element(C + "overlay");
        Overlay = overlay?.Attribute("val")?.Value == "1";
    }
}
