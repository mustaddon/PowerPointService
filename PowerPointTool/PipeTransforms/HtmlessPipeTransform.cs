using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace PowerPointTool.PipeTransforms;

public class HtmlessPipeTransform() : BaseTransform<string>("htmless")
{
    public static string DefaultLinkFormat = "$2 ($1)";

    protected override string ParseArgs(string[] args)
    {
        return args.Length > 0 ? args[0] : DefaultLinkFormat;
    }

    protected override object TransformItem(object obj, string linkFormat)
    {
        var html = obj as string;

        if (string.IsNullOrEmpty(html))
            return null;

        // replace tabs
        var result = _reTab.Replace(html, "\n");

        // replace comments
        result = _reComment.Replace(result, string.Empty);

        // replace hyperlinks
        result = _reHyperlink.Replace(result, m =>
        {
            var href = m.Groups[1].Value.Trim();
            var text = m.Groups[2].Value.Trim();
            return href.Length > 0 && !string.Equals(href, text, StringComparison.InvariantCultureIgnoreCase)
                ? m.Result(linkFormat)
                : text;
        });

        // replace other tags
        result = _reTag.Replace(result, m =>
        {
            var tag = m.Groups[1].Value != string.Empty ? m.Groups[1].Value.ToLower() : m.Groups[2].Value != string.Empty ? m.Groups[2].Value.ToLower() : null;
            var open = m.Groups[1].Value != string.Empty;
            var res = tag == "li" && open ? "\r\n- "
                : _inlines.Contains(tag) ? string.Empty
                : _tabs.Contains(tag) ? (open ? string.Empty : " \t")
                : "\r\n";
            return res;
        });

        // replace newlines
        result = _reNewlines.Replace(result, "\r\n");

        return System.Net.WebUtility.HtmlDecode(result.Trim());
    }

    static readonly Regex _reTab = new(@"\n[\t\s]+");
    static readonly Regex _reComment = new(@"<!--(.*?)-->");
    static readonly Regex _reHyperlink = new(@"<a [^>]*?href=""([^""]+?)""[^>]*?>(.*?)</a>", RegexOptions.IgnoreCase);
    static readonly Regex _reTag = new(@"<([^<>\s/]*)[^<>]*?([^<>\s/]*)>");
    static readonly Regex _reNewlines = new(@"[\r\n]+");

    static readonly HashSet<string> _inlines = ["a", "span", "b", "big", "i", "small", "em", "strong", "button", "label", "tr"];
    static readonly HashSet<string> _tabs = ["td", "th"];

}
