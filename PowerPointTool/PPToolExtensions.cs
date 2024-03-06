using System;
using System.IO;
using System.Text.RegularExpressions;

namespace PowerPointTool;

public static partial class PPToolExtensions
{
    public static void CreateFromTemplate(this PPTool pps, Stream targetOutput, Stream sourceTemplate, Func<ISlideContext, object> slideModelFactory)
        => pps.CreateFromTemplateAsync(targetOutput, sourceTemplate, slideModelFactory).Wait();

    public static byte[] CreateFromTemplate(this PPTool pps, byte[] templatePresentation, Func<ISlideContext, object> slideModelFactory)
    {
        using var ms = new MemoryStream();
        ms.Write(templatePresentation, 0, templatePresentation.Length);
        pps.ApplySlideModels(ms, slideModelFactory);
        return ms.ToArray();
    }

    public static byte[] UpdateSlides(this PPTool pps, byte[] presentation, Action<ISlideUpdateContext> slideUpdate)
    {
        using var target = new MemoryStream(); target.Write(presentation, 0, presentation.Length);
        pps.UpdateSlides(target, slideUpdate);
        return target.ToArray();
    }

    public static byte[] InsertSlides(this PPTool pps, byte[] sourcePresentation, byte[] targetPresentation, int targetInsertIndex = -1, Func<ISlideContext, bool> sourceSlideSelector = null)
    {
        using var source = new MemoryStream(sourcePresentation);
        using var target = new MemoryStream(); target.Write(targetPresentation, 0, targetPresentation.Length);
        pps.InsertSlides(target, targetInsertIndex, source, sourceSlideSelector);
        return target.ToArray();
    }

    public static byte[] MergeSlides(this PPTool pps, byte[] sourcePresentation, byte[] targetPresentation, Func<ISlideContext, int?> map)
    {
        using var source = new MemoryStream(sourcePresentation);
        using var target = new MemoryStream(); target.Write(targetPresentation, 0, targetPresentation.Length);
        pps.MergeSlides(target, source, map);
        return target.ToArray();
    }

    public static byte[] DeleteSlides(this PPTool pps, byte[] presentation, Func<ISlideContext, bool> slideSelector)
    {
        using var ms = new MemoryStream();
        ms.Write(presentation, 0, presentation.Length);
        pps.DeleteSlides(ms, slideSelector);
        return ms.ToArray();
    }

    public static int SlideIndex(this PPTool pps, byte[] presentation, Regex regex)
        => pps.SlideIndex(new MemoryStream(presentation), regex);
}
