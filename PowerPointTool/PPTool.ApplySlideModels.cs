﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPointTool._internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace PowerPointTool;

public partial class PPTool
{
    public virtual void ApplySlideModels(Stream targetTemplate, Func<ISlideContext, object> slideModelFactory)
    {
        using var doc = PresentationDocument.Open(targetTemplate, true);

        var slist = doc.PresentationPart.Presentation.SlideIdList;
        var slideIds = slist.ChildElements.Cast<SlideId>().ToArray();

        for (var i = 0; i < slideIds.Length; i++)
        {
            var slideId = slideIds[i];
            var ctx = new SlideUpdate(this, doc.PresentationPart, slideId, i, slideIds.Length);
            var model = slideModelFactory(ctx);
            if (model != null)
                ctx.ApplyModels(model as IEnumerable<object> ?? [model]);
        }
    }

    internal void ApplyModels<T>(PresentationPart presentationPart, SlideId slideId, SlidePart slidePart, IEnumerable<T> slideModels, Action<T, int, SlideId> action = null)
    {
        var prevSlideId = slideId;
        var slist = presentationPart.Presentation.SlideIdList;
        var nextId = GetMaxSlideId(slist) + 1;
        var i = 0;

        foreach (var m in slideModels)
        {
            var newSlidePart = presentationPart.AddNewPart<SlidePart>();
            newSlidePart.Slide = new Slide(InsertValues(m, slidePart.Slide.OuterXml));
            CopyPartsAndRelationships(slidePart, newSlidePart, m);
            prevSlideId = slist.InsertAfter(new SlideId { Id = nextId++, RelationshipId = presentationPart.GetIdOfPart(newSlidePart) }, prevSlideId);
            action?.Invoke(m, i++, prevSlideId);
        }

        slist.RemoveChild(slideId);
        CleanCustomShow(presentationPart.Presentation.CustomShowList, slideId.RelationshipId);
        presentationPart.DeletePart(slidePart);
    }

    string InsertValues(object model, string xml)
    {
        if (model == null)
            return null;

        var result = _insertRows(model, xml);

        return _reCmd.Replace(result,
            x => _escape(_getValue(model, x.Groups[1].Value)?.ToString()));
    }

    Uri _insertValues(object model, Uri uri)
    {
        if (model == null || uri == null)
            return uri;

        try
        {
            var source = Uri.UnescapeDataString(uri.OriginalString);
            var result = _reCmd.Replace(source,
                x => _getValue(model, x.Groups[1].Value)?.ToString());

            return new Uri(Uri.EscapeDataString(result), UriKind.RelativeOrAbsolute);
        }
        catch
        {
            return uri;
        }
    }


    string _insertRows(object model, string xml)
    {
        return _reInsert.Replace(xml, x =>
        {
            var itemsPath = _reCmd.Matches(x.Value).Cast<Match>()
                .Select(xx => _getRowSourcePath(model, xx.Groups[1].Value))
                .FirstOrDefault(xx => xx != null);

            if (itemsPath == null)
                return x.Value;

            var items = model.GetValue(itemsPath) as System.Collections.IEnumerable;
            var result = new StringBuilder();

            if (items != null)
                foreach (var item in items)
                    result.Append(_reCmd.Replace(x.Value,
                        xx =>
                        {
                            var cmd = _removeTags(xx.Groups[1].Value).Trim();
                            return cmd.StartsWith(itemsPath)
                                ? _escape(_getValue(item, cmd.Substring(Math.Min(itemsPath.Length + 1, cmd.Length)))?.ToString())
                                : xx.Value;
                        }));

            return result.ToString();
        });
    }

    static string _getRowSourcePath(object obj, string cmd)
    {
        var propNames = _removeTags(cmd).Split('|').First().Trim().Split('.');
        var type = obj.GetType();

        for (var i = 0; i < propNames.Length - 1; i++)
        {
            var pi = type.GetProperty(propNames[i]);

            if (pi == null)
                break;

            if (_isArray(pi.PropertyType))
                return string.Join(".", propNames.Take(i + 1));

            if (obj == null)
                break;

            type = pi.PropertyType;
            obj = pi.GetValue(obj);
        }

        return null;
    }


    void CopyPartsAndRelationships(SlidePart source, SlidePart target, object model)
    {
        source.Parts?.Where(x => x.OpenXmlPart.GetType() != typeof(NotesSlidePart)).ToList()
            .ForEach(x => target.AddPart(x.OpenXmlPart, x.RelationshipId));

        source.HyperlinkRelationships?.ToList()
            .ForEach(x => target.AddHyperlinkRelationship(_insertValues(model, x.Uri), x.IsExternal, x.Id));

        source.ExternalRelationships?.ToList()
            .ForEach(x => target.AddExternalRelationship(x.RelationshipType, x.Uri, x.Id));

        source.DataPartReferenceRelationships?.ToList()
            .ForEach(x =>
            {
                if (x is AudioReferenceRelationship)
                    target.AddAudioReferenceRelationship((MediaDataPart)x.DataPart, x.Id);
                else if (x is VideoReferenceRelationship)
                    target.AddVideoReferenceRelationship((MediaDataPart)x.DataPart, x.Id);
                else if (x is MediaReferenceRelationship)
                    target.AddMediaReferenceRelationship((MediaDataPart)x.DataPart, x.Id);
            });
    }

    object _getValue(object obj, string cmd)
    {
        var parts = _removeTags(cmd).Split('|');
        var value = obj.GetValue(parts.First());

        for (var i = 1; i < parts.Length; i++)
        {
            var pipe = parts[i].Trim().Split(':').First().Trim().ToLower();

            if (Pipes.Value.ContainsKey(pipe))
            {
                var args = _rePipeArgs.Matches(parts[i]).Cast<Match>().Select(x => x.Groups[1].Value).ToArray();
                value = Pipes.Value[pipe].Transform(value, args);
            }
        }

        return value;
    }

    static string _removeTags(string str)
    {
        return _reTags.Replace(str, string.Empty);
    }


    static bool _isArray(Type type)
    {
        return type != null && type != typeof(string) && typeof(System.Collections.IEnumerable).IsAssignableFrom(type);
    }

    static string _escape(string unescaped)
    {
        var doc = new XmlDocument();
        var node = doc.CreateElement("root");
        node.InnerText = _reEscape.Replace(unescaped ?? string.Empty, string.Empty);
        return node.InnerXml;
    }

    static readonly Regex _reEscape = new(@"[\v]");
    static readonly Regex _reTags = new(@"<[^<>]*?>");
    static readonly Regex _reInsert = new(@"<a:tr.+?</a:tr>");
    static readonly Regex _reCmd = new(@"{{(.*?)}}", RegexOptions.IgnorePatternWhitespace);
    static readonly Regex _rePipeArgs = new(@"'([^']*?)'");
}
