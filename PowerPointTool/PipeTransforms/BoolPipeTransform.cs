using System;
using System.Text.RegularExpressions;

namespace PowerPointTool.PipeTransforms;

public class BoolPipeTransform() : BaseTransform<(string trueVal, string falseVal, string nullVal)>("bool")
{
    public static string DefaultTrueValue = "TRUE";
    public static string DefaultFalseValue = "FALSE";
    public static string DefaultNullValue = null;

    protected override (string trueVal, string falseVal, string nullVal) ParseArgs(string[] args)
    {
        return (
            args.Length > 0 ? Regex.Unescape(args[0] ?? string.Empty) : DefaultTrueValue,
            args.Length > 1 ? Regex.Unescape(args[1] ?? string.Empty) : DefaultFalseValue,
            args.Length > 2 ? Regex.Unescape(args[2] ?? string.Empty) : DefaultNullValue
        );
    }

    protected override object TransformItem(object obj, (string trueVal, string falseVal, string nullVal) args)
    {
        if (obj == null)
            return args.nullVal;

        if (obj is bool @bool)
            return @bool ? args.trueVal : args.falseVal;

        if (obj is string str)
            return !string.IsNullOrWhiteSpace(str) ? args.trueVal : args.falseVal;

        if (obj.ToString() == "0")
            return args.falseVal;

        return args.trueVal;
    }
}
