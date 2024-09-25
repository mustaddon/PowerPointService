using System;

namespace PowerPointTool.PipeTransforms;

public class BoolPipeTransform() : BaseTransform<(string trueVal, string falseVal, string nullVal)>("bool")
{
    public static string DefaultTrueValue = "TRUE";
    public static string DefaultFalseValue = "FALSE";
    public static string DefaultNullValue = null;

    protected override (string trueVal, string falseVal, string nullVal) ParseArgs(string[] args)
    {
        return (
            args.Length > 0 ? args[0] : DefaultTrueValue,
            args.Length > 1 ? args[1] : DefaultFalseValue,
            args.Length > 2 ? args[2] : DefaultNullValue
        );
    }

    protected override object TransformItem(object obj, (string trueVal, string falseVal, string nullVal) args)
    {
        if (obj == null)
            return args.nullVal;

        if (Convert.ToBoolean(obj))
            return args.trueVal;

        return args.falseVal;
    }
}
