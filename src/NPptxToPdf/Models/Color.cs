namespace NPptxToPdf;

public readonly record struct Color(byte R, byte G, byte B, byte A = 255)
{
    public static Color Black => new(0, 0, 0);
    public static Color White => new(255, 255, 255);
    public static Color Red => new(255, 0, 0);
    public static Color Green => new(0, 255, 0);
    public static Color Blue => new(0, 0, 255);
    public static Color Transparent => new(0, 0, 0, 0);

    public byte Alpha => A;

    public Color WithAlpha(byte alpha) => new(R, G, B, alpha);

    public override string ToString() => $"#{R:X2}{G:X2}{B:X2}";

    public static Color FromSchemeColor(SchemeColor schemeColor)
    {
        // Default colors for scheme colors
        return schemeColor switch
        {
            SchemeColor.Background1 => White,
            SchemeColor.Text1 => Black,
            SchemeColor.Background2 => new Color(240, 240, 240),
            SchemeColor.Text2 => new Color(64, 64, 64),
            SchemeColor.Accent1 => new Color(68, 114, 196),
            SchemeColor.Accent2 => new Color(237, 125, 49),
            SchemeColor.Accent3 => new Color(165, 165, 165),
            SchemeColor.Accent4 => new Color(255, 192, 0),
            SchemeColor.Accent5 => new Color(91, 155, 213),
            SchemeColor.Accent6 => new Color(112, 173, 71),
            SchemeColor.Hyperlink => new Color(5, 99, 193),
            SchemeColor.FollowedHyperlink => new Color(149, 79, 114),
            SchemeColor.Dark1 => Black,
            SchemeColor.Light1 => White,
            SchemeColor.Dark2 => new Color(64, 64, 64),
            SchemeColor.Light2 => new Color(240, 240, 240),
            _ => Black
        };
    }

    public static Color FromPresetColor(PresetColor presetColor)
    {
        return presetColor switch
        {
            PresetColor.AliceBlue => new Color(240, 248, 255),
            PresetColor.AntiqueWhite => new Color(250, 235, 215),
            PresetColor.Aqua => new Color(0, 255, 255),
            PresetColor.Aquamarine => new Color(127, 255, 212),
            PresetColor.Azure => new Color(240, 255, 255),
            PresetColor.Beige => new Color(245, 245, 220),
            PresetColor.Bisque => new Color(255, 228, 196),
            PresetColor.Black => new Color(0, 0, 0),
            PresetColor.BlanchedAlmond => new Color(255, 235, 205),
            PresetColor.Blue => new Color(0, 0, 255),
            PresetColor.BlueViolet => new Color(138, 43, 226),
            PresetColor.Brown => new Color(165, 42, 42),
            PresetColor.BurlyWood => new Color(222, 184, 135),
            PresetColor.CadetBlue => new Color(95, 158, 160),
            PresetColor.Chartreuse => new Color(127, 255, 0),
            PresetColor.Chocolate => new Color(210, 105, 30),
            PresetColor.Coral => new Color(255, 127, 80),
            PresetColor.CornflowerBlue => new Color(100, 149, 237),
            PresetColor.Cornsilk => new Color(255, 248, 220),
            PresetColor.Crimson => new Color(220, 20, 60),
            PresetColor.Cyan => new Color(0, 255, 255),
            PresetColor.DarkBlue => new Color(0, 0, 139),
            PresetColor.DarkCyan => new Color(0, 139, 139),
            PresetColor.DarkGoldenrod => new Color(184, 134, 11),
            PresetColor.DarkGray => new Color(169, 169, 169),
            PresetColor.DarkGreen => new Color(0, 100, 0),
            PresetColor.DarkKhaki => new Color(189, 183, 107),
            PresetColor.DarkMagenta => new Color(139, 0, 139),
            PresetColor.DarkOliveGreen => new Color(85, 107, 47),
            PresetColor.DarkOrange => new Color(255, 140, 0),
            PresetColor.DarkOrchid => new Color(153, 50, 204),
            PresetColor.DarkRed => new Color(139, 0, 0),
            PresetColor.DarkSalmon => new Color(233, 150, 122),
            PresetColor.DarkSeaGreen => new Color(143, 188, 143),
            PresetColor.DarkSlateBlue => new Color(72, 61, 139),
            PresetColor.DarkSlateGray => new Color(47, 79, 79),
            PresetColor.DarkTurquoise => new Color(0, 206, 209),
            PresetColor.DarkViolet => new Color(148, 0, 211),
            PresetColor.DeepPink => new Color(255, 20, 147),
            PresetColor.DeepSkyBlue => new Color(0, 191, 255),
            PresetColor.DimGray => new Color(105, 105, 105),
            PresetColor.DodgerBlue => new Color(30, 144, 255),
            PresetColor.Firebrick => new Color(178, 34, 34),
            PresetColor.FloralWhite => new Color(255, 250, 240),
            PresetColor.ForestGreen => new Color(34, 139, 34),
            PresetColor.Fuchsia => new Color(255, 0, 255),
            PresetColor.Gainsboro => new Color(220, 220, 220),
            PresetColor.GhostWhite => new Color(248, 248, 255),
            PresetColor.Gold => new Color(255, 215, 0),
            PresetColor.Goldenrod => new Color(218, 165, 32),
            PresetColor.Gray => new Color(128, 128, 128),
            PresetColor.Green => new Color(0, 128, 0),
            PresetColor.GreenYellow => new Color(173, 255, 47),
            PresetColor.Honeydew => new Color(240, 255, 240),
            PresetColor.HotPink => new Color(255, 105, 180),
            PresetColor.IndianRed => new Color(205, 92, 92),
            PresetColor.Indigo => new Color(75, 0, 130),
            PresetColor.Ivory => new Color(255, 255, 240),
            PresetColor.Khaki => new Color(240, 230, 140),
            PresetColor.Lavender => new Color(230, 230, 250),
            PresetColor.LavenderBlush => new Color(255, 240, 245),
            PresetColor.LawnGreen => new Color(124, 252, 0),
            PresetColor.LemonChiffon => new Color(255, 250, 205),
            PresetColor.LightBlue => new Color(173, 216, 230),
            PresetColor.LightCoral => new Color(240, 128, 128),
            PresetColor.LightCyan => new Color(224, 255, 255),
            PresetColor.LightGoldenrodYellow => new Color(250, 250, 210),
            PresetColor.LightGray => new Color(211, 211, 211),
            PresetColor.LightGreen => new Color(144, 238, 144),
            PresetColor.LightPink => new Color(255, 182, 193),
            PresetColor.LightSalmon => new Color(255, 160, 122),
            PresetColor.LightSeaGreen => new Color(32, 178, 170),
            PresetColor.LightSkyBlue => new Color(135, 206, 250),
            PresetColor.LightSlateGray => new Color(119, 136, 153),
            PresetColor.LightSteelBlue => new Color(176, 196, 222),
            PresetColor.LightYellow => new Color(255, 255, 224),
            PresetColor.Lime => new Color(0, 255, 0),
            PresetColor.LimeGreen => new Color(50, 205, 50),
            PresetColor.Linen => new Color(250, 240, 230),
            PresetColor.Magenta => new Color(255, 0, 255),
            PresetColor.Maroon => new Color(128, 0, 0),
            PresetColor.MediumAquamarine => new Color(102, 205, 170),
            PresetColor.MediumBlue => new Color(0, 0, 205),
            PresetColor.MediumOrchid => new Color(186, 85, 211),
            PresetColor.MediumPurple => new Color(147, 112, 219),
            PresetColor.MediumSeaGreen => new Color(60, 179, 113),
            PresetColor.MediumSlateBlue => new Color(123, 104, 238),
            PresetColor.MediumSpringGreen => new Color(0, 250, 154),
            PresetColor.MediumTurquoise => new Color(72, 209, 204),
            PresetColor.MediumVioletRed => new Color(199, 21, 133),
            PresetColor.MidnightBlue => new Color(25, 25, 112),
            PresetColor.MintCream => new Color(245, 255, 250),
            PresetColor.MistyRose => new Color(255, 228, 225),
            PresetColor.Moccasin => new Color(255, 228, 181),
            PresetColor.NavajoWhite => new Color(255, 222, 173),
            PresetColor.Navy => new Color(0, 0, 128),
            PresetColor.OldLace => new Color(253, 245, 230),
            PresetColor.Olive => new Color(128, 128, 0),
            PresetColor.OliveDrab => new Color(107, 142, 35),
            PresetColor.Orange => new Color(255, 165, 0),
            PresetColor.OrangeRed => new Color(255, 69, 0),
            PresetColor.Orchid => new Color(218, 112, 214),
            PresetColor.PaleGoldenrod => new Color(238, 232, 170),
            PresetColor.PaleGreen => new Color(152, 251, 152),
            PresetColor.PaleTurquoise => new Color(175, 238, 238),
            PresetColor.PaleVioletRed => new Color(219, 112, 147),
            PresetColor.PapayaWhip => new Color(255, 239, 213),
            PresetColor.PeachPuff => new Color(255, 218, 185),
            PresetColor.Peru => new Color(205, 133, 63),
            PresetColor.Pink => new Color(255, 192, 203),
            PresetColor.Plum => new Color(221, 160, 221),
            PresetColor.PowderBlue => new Color(176, 224, 230),
            PresetColor.Purple => new Color(128, 0, 128),
            PresetColor.Red => new Color(255, 0, 0),
            PresetColor.RosyBrown => new Color(188, 143, 143),
            PresetColor.RoyalBlue => new Color(65, 105, 225),
            PresetColor.SaddleBrown => new Color(139, 69, 19),
            PresetColor.Salmon => new Color(250, 128, 114),
            PresetColor.SandyBrown => new Color(244, 164, 96),
            PresetColor.SeaGreen => new Color(46, 139, 87),
            PresetColor.SeaShell => new Color(255, 245, 238),
            PresetColor.Sienna => new Color(160, 82, 45),
            PresetColor.Silver => new Color(192, 192, 192),
            PresetColor.SkyBlue => new Color(135, 206, 235),
            PresetColor.SlateBlue => new Color(106, 90, 205),
            PresetColor.SlateGray => new Color(112, 128, 144),
            PresetColor.Snow => new Color(255, 250, 250),
            PresetColor.SpringGreen => new Color(0, 255, 127),
            PresetColor.SteelBlue => new Color(70, 130, 180),
            PresetColor.Tan => new Color(210, 180, 140),
            PresetColor.Teal => new Color(0, 128, 128),
            PresetColor.Thistle => new Color(216, 191, 216),
            PresetColor.Tomato => new Color(255, 99, 71),
            PresetColor.Turquoise => new Color(64, 224, 208),
            PresetColor.Violet => new Color(238, 130, 238),
            PresetColor.Wheat => new Color(245, 222, 179),
            PresetColor.White => new Color(255, 255, 255),
            PresetColor.WhiteSmoke => new Color(245, 245, 245),
            PresetColor.Yellow => new Color(255, 255, 0),
            PresetColor.YellowGreen => new Color(154, 205, 50),
            _ => Black
        };
    }

    public static Color FromHsl(double hue, double saturation, double lightness)
    {
        // Convert HSL to RGB
        double c = (1 - Math.Abs(2 * lightness - 1)) * saturation;
        double x = c * (1 - Math.Abs((hue / 60) % 2 - 1));
        double m = lightness - c / 2;

        double r, g, b;

        if (hue < 60)
        {
            r = c; g = x; b = 0;
        }
        else if (hue < 120)
        {
            r = x; g = c; b = 0;
        }
        else if (hue < 180)
        {
            r = 0; g = c; b = x;
        }
        else if (hue < 240)
        {
            r = 0; g = x; b = c;
        }
        else if (hue < 300)
        {
            r = x; g = 0; b = c;
        }
        else
        {
            r = c; g = 0; b = x;
        }

        return new Color(
            (byte)((r + m) * 255),
            (byte)((g + m) * 255),
            (byte)((b + m) * 255)
        );
    }
}
