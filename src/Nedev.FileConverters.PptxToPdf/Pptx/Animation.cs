using System.Xml.Linq;

namespace Nedev.FileConverters.PptxToPdf.Pptx;

public class Animation
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public List<AnimationEffect> Effects { get; } = new();
    public List<AnimationTiming> Timing { get; } = new();

    public Animation(XElement element)
    {
        _element = element;
        Parse();
    }

    private void Parse()
    {
        // Parse animation effects
        var animEffects = _element.Descendants(P + "anim");
        foreach (var effect in animEffects)
        {
            Effects.Add(new AnimationEffect(effect));
        }

        // Parse set effects (property changes)
        var setEffects = _element.Descendants(P + "set");
        foreach (var effect in setEffects)
        {
            Effects.Add(new AnimationEffect(effect, AnimationType.Set));
        }

        // Parse motion path effects
        var motionEffects = _element.Descendants(P + "animMotion");
        foreach (var effect in motionEffects)
        {
            Effects.Add(new AnimationEffect(effect, AnimationType.Motion));
        }

        // Parse scale effects
        var scaleEffects = _element.Descendants(P + "animScale");
        foreach (var effect in scaleEffects)
        {
            Effects.Add(new AnimationEffect(effect, AnimationType.Scale));
        }

        // Parse rotation effects
        var rotEffects = _element.Descendants(P + "animRot");
        foreach (var effect in rotEffects)
        {
            Effects.Add(new AnimationEffect(effect, AnimationType.Rotate));
        }

        // Parse color effects
        var colorEffects = _element.Descendants(P + "animClr");
        foreach (var effect in colorEffects)
        {
            Effects.Add(new AnimationEffect(effect, AnimationType.Color));
        }
    }
}

public class AnimationEffect
{
    private readonly XElement _element;
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private static readonly XNamespace P = "http://schemas.openxmlformats.org/presentationml/2006/main";

    public AnimationType Type { get; }
    public string? TargetId { get; }
    public AnimationAttribute? Attribute { get; }
    public AnimationValue? From { get; }
    public AnimationValue? To { get; }
    public AnimationValue? By { get; }
    public AnimationTiming? Timing { get; }
    public AnimationMotion? Motion { get; }

    public AnimationEffect(XElement element, AnimationType type = AnimationType.Generic)
    {
        _element = element;
        Type = type;

        // Parse target
        var tgtEl = element.Element(P + "tgtEl");
        if (tgtEl != null)
        {
            var spTgt = tgtEl.Element(P + "spTgt");
            TargetId = spTgt?.Attribute("spid")?.Value;
        }

        // Parse attribute
        var attrName = element.Attribute("attrName")?.Value;
        if (!string.IsNullOrEmpty(attrName))
        {
            Attribute = new AnimationAttribute(attrName);
        }

        // Parse values
        var from = element.Attribute("from")?.Value;
        if (!string.IsNullOrEmpty(from))
        {
            From = new AnimationValue(from);
        }

        var to = element.Attribute("to")?.Value;
        if (!string.IsNullOrEmpty(to))
        {
            To = new AnimationValue(to);
        }

        var by = element.Attribute("by")?.Value;
        if (!string.IsNullOrEmpty(by))
        {
            By = new AnimationValue(by);
        }

        // Parse timing
        var cTn = element.Element(P + "cTn");
        if (cTn != null)
        {
            Timing = new AnimationTiming(cTn);
        }

        // Parse motion path
        if (type == AnimationType.Motion)
        {
            Motion = new AnimationMotion(element);
        }
    }
}

public enum AnimationType
{
    Generic,
    Set,
    Motion,
    Scale,
    Rotate,
    Color,
    Alpha,
    Filter,
    Command
}

public class AnimationAttribute
{
    public string Name { get; }
    public AnimationAttributeType Type { get; }

    public AnimationAttribute(string name)
    {
        Name = name;
        Type = ParseAttributeType(name);
    }

    private static AnimationAttributeType ParseAttributeType(string name)
    {
        return name.ToLower() switch
        {
            "style.opacity" => AnimationAttributeType.Opacity,
            "style.visibility" => AnimationAttributeType.Visibility,
            "style.color" => AnimationAttributeType.Color,
            "ppt_x" => AnimationAttributeType.PositionX,
            "ppt_y" => AnimationAttributeType.PositionY,
            "ppt_w" => AnimationAttributeType.Width,
            "ppt_h" => AnimationAttributeType.Height,
            "r" => AnimationAttributeType.Rotation,
            "xsx" => AnimationAttributeType.ScaleX,
            "xsy" => AnimationAttributeType.ScaleY,
            _ => AnimationAttributeType.Other
        };
    }
}

public enum AnimationAttributeType
{
    Other,
    Opacity,
    Visibility,
    Color,
    PositionX,
    PositionY,
    Width,
    Height,
    Rotation,
    ScaleX,
    ScaleY
}

public class AnimationValue
{
    public string RawValue { get; }
    public double? NumericValue { get; }
    public string? Formula { get; }

    public AnimationValue(string value)
    {
        RawValue = value;

        // Check if it's a formula
        if (value.StartsWith("="))
        {
            Formula = value;
        }
        // Try to parse as number
        else if (double.TryParse(value, out var num))
        {
            NumericValue = num;
        }
    }
}

public class AnimationTiming
{
    private readonly XElement _element;

    public int Duration { get; } // in milliseconds
    public int Delay { get; } // in milliseconds
    public AnimationTrigger? Trigger { get; }
    public AnimationRepeat? Repeat { get; }
    public AnimationFillMode Fill { get; }
    public AnimationRestartMode Restart { get; }

    public AnimationTiming(XElement element)
    {
        _element = element;

        // Parse duration
        var dur = element.Attribute("dur")?.Value;
        Duration = ParseTimeValue(dur);

        // Parse delay
        var delay = element.Attribute("delay")?.Value;
        Delay = ParseTimeValue(delay);

        // Parse fill mode
        var fill = element.Attribute("fill")?.Value;
        Fill = fill switch
        {
            "freeze" => AnimationFillMode.Freeze,
            "hold" => AnimationFillMode.Hold,
            "transition" => AnimationFillMode.Transition,
            _ => AnimationFillMode.Remove
        };

        // Parse restart mode
        var restart = element.Attribute("restart")?.Value;
        Restart = restart switch
        {
            "never" => AnimationRestartMode.Never,
            "whenNotActive" => AnimationRestartMode.WhenNotActive,
            _ => AnimationRestartMode.Always
        };

        // Parse trigger
        var stCondLst = element.Element(element.Name.Namespace + "stCondLst");
        if (stCondLst != null)
        {
            Trigger = new AnimationTrigger(stCondLst);
        }

        // Parse repeat
        var repeatCount = element.Attribute("repeatCount")?.Value;
        var repeatDur = element.Attribute("repeatDur")?.Value;
        if (!string.IsNullOrEmpty(repeatCount) || !string.IsNullOrEmpty(repeatDur))
        {
            Repeat = new AnimationRepeat(repeatCount, repeatDur);
        }
    }

    public static int ParseTimeValue(string? value)
    {
        if (string.IsNullOrEmpty(value)) return 0;

        // Check for "indefinite"
        if (value == "indefinite") return -1;

        // Parse as milliseconds
        if (value.EndsWith("ms"))
        {
            if (int.TryParse(value.Substring(0, value.Length - 2), out var ms))
                return ms;
        }
        else if (value.EndsWith("s"))
        {
            if (double.TryParse(value.Substring(0, value.Length - 1), out var s))
                return (int)(s * 1000);
        }
        else if (int.TryParse(value, out var num))
        {
            // Assume milliseconds
            return num;
        }

        return 0;
    }
}

public class AnimationTrigger
{
    private readonly XElement _element;

    public AnimationTriggerType Type { get; }
    public string? TargetId { get; }
    public int Delay { get; }

    public AnimationTrigger(XElement element)
    {
        _element = element;

        var cond = element.Element(element.Name.Namespace + "cond");
        if (cond != null)
        {
            var evt = cond.Attribute("evt")?.Value;
            Type = evt switch
            {
                "onClick" => AnimationTriggerType.OnClick,
                "onBegin" => AnimationTriggerType.OnBegin,
                "onEnd" => AnimationTriggerType.OnEnd,
                _ => AnimationTriggerType.OnClick
            };

            var delay = cond.Attribute("delay")?.Value;
            if (!string.IsNullOrEmpty(delay))
            {
                Delay = AnimationTiming.ParseTimeValue(delay);
            }

            var tgtEl = cond.Element(element.Name.Namespace + "tgtEl");
            if (tgtEl != null)
            {
                var spTgt = tgtEl.Element(element.Name.Namespace + "spTgt");
                TargetId = spTgt?.Attribute("spid")?.Value;
            }
        }
    }
}

public enum AnimationTriggerType
{
    OnClick,
    OnBegin,
    OnEnd,
    WithPrevious,
    AfterPrevious
}

public enum AnimationFillMode
{
    Remove,
    Freeze,
    Hold,
    Transition
}

public enum AnimationRestartMode
{
    Always,
    Never,
    WhenNotActive
}

public class AnimationRepeat
{
    public int? Count { get; }
    public int? Duration { get; } // in milliseconds

    public AnimationRepeat(string? count, string? duration)
    {
        if (!string.IsNullOrEmpty(count))
        {
            if (count == "indefinite")
            {
                Count = -1;
            }
            else if (int.TryParse(count, out var c))
            {
                Count = c;
            }
        }

        if (!string.IsNullOrEmpty(duration))
        {
            Duration = AnimationTiming.ParseTimeValue(duration);
        }
    }
}

public class AnimationMotion
{
    private readonly XElement _element;

    public AnimationPathType PathType { get; }
    public string? PathData { get; }
    public AnimationOrigin Origin { get; }
    public bool AutoReverse { get; }

    public AnimationMotion(XElement element)
    {
        _element = element;

        // Parse path type
        var pathType = element.Attribute("pathEditMode")?.Value;
        PathType = pathType switch
        {
            "relative" => AnimationPathType.Relative,
            "fixed" => AnimationPathType.Fixed,
            _ => AnimationPathType.Relative
        };

        // Parse origin
        var origin = element.Attribute("origin")?.Value;
        Origin = origin switch
        {
            "parent" => AnimationOrigin.Parent,
            "layout" => AnimationOrigin.Layout,
            _ => AnimationOrigin.Default
        };

        // Parse auto reverse
        var autoReverse = element.Attribute("autoRev")?.Value;
        AutoReverse = autoReverse == "1";

        // Parse path data
        var path = element.Element(element.Name.Namespace + "path");
        PathData = path?.Attribute("val")?.Value;
    }
}

public enum AnimationPathType
{
    Relative,
    Fixed
}

public enum AnimationOrigin
{
    Default,
    Parent,
    Layout
}

public class SlideAnimation
{
    public List<AnimationEffect> EntranceEffects { get; } = new();
    public List<AnimationEffect> EmphasisEffects { get; } = new();
    public List<AnimationEffect> ExitEffects { get; } = new();
    public List<AnimationEffect> MotionPathEffects { get; } = new();

    public void AddEffect(AnimationEffect effect)
    {
        // Categorize effect by type
        switch (effect.Type)
        {
            case AnimationType.Motion:
                MotionPathEffects.Add(effect);
                break;
            case AnimationType.Set:
            case AnimationType.Color:
            case AnimationType.Alpha:
                EmphasisEffects.Add(effect);
                break;
            default:
                EntranceEffects.Add(effect);
                break;
        }
    }
}
