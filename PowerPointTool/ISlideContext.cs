using System;
using System.Drawing;
using System.IO;

namespace PowerPointTool;

public interface ISlideContext
{
    int SlidesCount { get; }
    int SlideWidth { get; }
    int SlideHeight { get; }
    int SlideIndex { get; }
    string SlideXml { get; }
}

public interface ISlideUpdateContext : ISlideContext
{
    void ApplyModel(object model);
    void AddImage(Stream stream, string type = "image/png", Rectangle? shape = null);
}

public static class ISlideContextExtensions
{
    public static void AddImage(this ISlideUpdateContext ctx, byte[] image, string type = "image/png", Rectangle? shape = null)
        => ctx.AddImage(new MemoryStream(image), type, shape);

    public static void RemoveSlide(this ISlideUpdateContext ctx)
        => ctx.ApplyModel(Array.Empty<object>());
}