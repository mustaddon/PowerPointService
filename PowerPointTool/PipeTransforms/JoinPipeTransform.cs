﻿using PowerPointTool._internal;
using System.Linq;
using System.Text.RegularExpressions;

namespace PowerPointTool.PipeTransforms;

public class JoinPipeTransform : IPipeTransform
{
    public string Name => "join";

    public static string DefaultSeparator = string.Empty;

    public object Transform(object obj, params string[] args)
    {
        var separator = args.Length > 0 ? Regex.Unescape(args[0] ?? string.Empty) : DefaultSeparator;
        var prop = args.Length > 1 ? args[1] : null;
        
        var array = obj.AsEnumerable()?
            .Select(x => x.GetValue(prop))
            .Where(x => x != null);

        return array != null ? string.Join(separator, array) : null;
    }
}
