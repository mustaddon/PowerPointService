using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace PowerPointTool.PipeTransforms;

public class NumberPipeTransform() : BaseTransform<(string format, string thousandSeparator, string decimalSeparator)>("number")
{

    public static string DefaultFormat = "1.0-3";

    protected override (string format, string thousandSeparator, string decimalSeparator) ParseArgs(string[] args)
    {
        var match = args.Length > 0 ? _reFormatPattern.Match(args[0]) : null;

        if (match?.Success != true)
            match = _reFormatPattern.Match(DefaultFormat);

        var thousandSeparator = match.Groups[1].Value;
        var decimalSeparator = match.Groups[3].Value;
        var minIntegerDigits = int.Parse(match.Groups[2].Value);
        var minFractionDigits = int.Parse(match.Groups[4].Value);
        var maxFractionDigits = int.Parse(match.Groups[5].Value);

        var format = string.Format("#,#{0}.{1}",
            string.Join("", Enumerable.Range(1, minIntegerDigits).Select(x => "0")),
            string.Join("", Enumerable.Range(1, maxFractionDigits).Select(x => x > minFractionDigits ? "#" : "0")));

        return (format, thousandSeparator, decimalSeparator);
    }

    protected override object TransformItem(object obj, (string format, string thousandSeparator, string decimalSeparator) args)
    {
        if (obj == null) return null;
        var value = Convert.ToDouble(obj, _cultureInfo);
        var str = value.ToString(args.format, _cultureInfo);
        return Regex.Replace(str, @"\.|,", x => x.Value == "." ? args.decimalSeparator : args.thousandSeparator);
    }


    static readonly Regex _reFormatPattern = new(@"^([\D]*?)([\d]+?)(\.|,)([\d]+?)-([\d]+?)$");
    static readonly CultureInfo _cultureInfo = CultureInfo.CreateSpecificCulture("en-US");
}
