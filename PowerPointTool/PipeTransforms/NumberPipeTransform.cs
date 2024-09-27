using System;
using System.Globalization;
using System.Text.RegularExpressions;

namespace PowerPointTool.PipeTransforms;

public class NumberPipeTransform() : BaseTransform<(string format, string thousandSeparator, string decimalSeparator)>("number")
{

    public static string DefaultFormat = "1.0-3";

    internal bool IntegerMode = false;

    protected override (string format, string thousandSeparator, string decimalSeparator) ParseArgs(string[] args)
    {
        var match = args.Length > 0 ? _reFormatPattern.Match(args[0]) : null;

        if (match?.Success != true)
            match = _reFormatPattern.Match(DefaultFormat);

        var thousandSeparator = match.Groups[1].Value;
        var decimalSeparator = match.Groups[3].Value;
        var minIntegerDigits = int.Parse(match.Groups[2].Value);
        string format;
        
        if (IntegerMode)
        {
            format = string.Format("#,#{0}", new string('0', minIntegerDigits));
        }
        else
        {
            var minFractionDigits = int.Parse(match.Groups[4].Value);
            var maxFractionDigits = int.Parse(match.Groups[5].Value);

            format = string.Format("#,#{0}.{1}{2}",
                new string('0', minIntegerDigits),
                new string('0', minFractionDigits),
                new string('#', Math.Max(0, maxFractionDigits - minFractionDigits)));
        }

        return (format, thousandSeparator, decimalSeparator);
    }

    protected override object TransformItem(object obj, (string format, string thousandSeparator, string decimalSeparator) args)
    {
        if (obj == null) return null;
        var value = Convert.ToDouble(obj, _cultureInfo);
        var str = value.ToString(args.format, _cultureInfo);
        return _reResult.Replace(str, x => x.Value == "." ? args.decimalSeparator : args.thousandSeparator);
    }


    static readonly Regex _reResult = new(@"\.|,");
    static readonly Regex _reFormatPattern = new(@"^([\D]*?)([\d]+?)(\.|,)([\d]+?)-([\d]+?)$");
    static readonly CultureInfo _cultureInfo = CultureInfo.CreateSpecificCulture("en-US");
}
