﻿using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace PowerPointTool;

public partial class PPTool(IEnumerable<IPipeTransform> pipes = null)
{
    protected readonly Lazy<Dictionary<string, IPipeTransform>> Pipes = new(() => (pipes ?? [])
        .Concat(_defaultPipes.Value)
        .GroupBy(x => x?.Name?.Trim().ToLower())
        .Where(g => !string.IsNullOrWhiteSpace(g.Key))
        .ToDictionary(g => g.Key, g => g.First()));


    static readonly Lazy<IPipeTransform[]> _defaultPipes = new(() =>
    {
        try
        {
            return typeof(IPipeTransform).Assembly.GetTypes()
                .Where(x => !x.IsAbstract && typeof(IPipeTransform).IsAssignableFrom(x))
                .Select(x => Activator.CreateInstance(x) as IPipeTransform)
                .ToArray();
        }
        catch
        {
            return [];
        }
    });

    internal static uint GetMaxSlideId(SlideIdList slideIdList)
    {
        // Slide identifiers have a minimum value of greater than or
        // equal to 256 and a maximum value of less than 2147483648. 
        return Math.Max(256, slideIdList?.Elements<SlideId>().Max(x => x.Id) ?? 0);
    }

    static uint _getMaxSlideMasterId(SlideMasterIdList slideMasterIdList)
    {
        // Slide master identifiers have a minimum value of greater than
        // or equal to 2147483648. 
        return Math.Max(2147483648, slideMasterIdList?.Elements<SlideMasterId>().Max(x => x.Id) ?? 0);
    }

}
