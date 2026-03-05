namespace NPptxToPdf;

public readonly record struct Rect(long X, long Y, long Width, long Height)
{
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
