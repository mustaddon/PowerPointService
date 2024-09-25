using System;
using System.Globalization;

namespace PowerPointTool.PipeTransforms;

public class DatePipeTransform() : BaseTransform<(string format, CultureInfo locale)>("date")
{
    public static string DefaultFormat = "dd.MM.yyyy";
    public static CultureInfo DefaultLocale = CultureInfo.CurrentCulture;

    protected override (string format, CultureInfo locale) ParseArgs(string[] args)
    {
        return (
            args.Length > 0 ? args[0] : DefaultFormat,
            args.Length > 1 ? CultureInfo.CreateSpecificCulture(args[1]) : DefaultLocale
        );
    }

    protected override object TransformItem(object obj, (string format, CultureInfo locale) args)
    {
        if (obj is DateTime dt)
            return dt.ToString(args.format, args.locale);

        if (obj is DateTimeOffset dto)
            return dto.ToString(args.format, args.locale);

#if NET6_0_OR_GREATER
        if (obj is DateOnly date)
            return date.ToString(args.format, args.locale);
#endif

        return null;
    }
}
