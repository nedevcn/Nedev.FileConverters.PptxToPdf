namespace Nedev.FileConverters.PptxToPdf;

public struct Rect
{
    public long X { get; }
    public long Y { get; }
    public long Width { get; }
    public long Height { get; }

    public Rect(long x, long y, long width, long height)
    {
        X = x;
        Y = y;
        Width = width;
        Height = height;
    }

    public Rect() : this(0, 0, 0, 0) { }

    public double XInches => X / 914400.0;
    public double YInches => Y / 914400.0;
    public double WidthInches => Width / 914400.0;
    public double HeightInches => Height / 914400.0;

    public double XPoints => XInches * 72;
    public double YPoints => YInches * 72;
    public double WidthPoints => WidthInches * 72;
    public double HeightPoints => HeightInches * 72;
}
