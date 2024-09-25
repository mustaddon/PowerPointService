using System;
using System.Collections.Generic;
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

    void AddImage(Stream stream, string type = "image/*", Rectangle? shape = null);
}

public interface ISlideUpdateContext : ISlideContext
{
    void ApplyModels<T>(IEnumerable<T> models, Action<ISlideContext, T, int> action = null);
}



public static class ISlideContextExtensions
{
    public static void AddImage(this ISlideContext ctx, byte[] image, string type = "image/*", Rectangle? shape = null)
        => ctx.AddImage(new MemoryStream(image), type, shape);

    public static void AddImage(this ISlideContext ctx, Stream stream, Rectangle shape)
        => ctx.AddImage(stream, shape: shape);

    public static void AddImage(this ISlideContext ctx, byte[] image, Rectangle shape)
        => ctx.AddImage(new MemoryStream(image), shape: shape);


    public static void RemoveSlide(this ISlideUpdateContext ctx)
        => ctx.ApplyModels(Array.Empty<object>());


    public static void ApplyModel(this ISlideUpdateContext ctx, object model)
        => ctx.ApplyModels(model as IEnumerable<object> ?? [model]);

    public static void ApplyModels<T>(this ISlideUpdateContext ctx, IEnumerable<T> models, Action<ISlideContext, T> action)
        => ctx.ApplyModels(models, action == null ? null : (a, x, i) => action(a, x));

    public static void ApplyModels<T>(this ISlideUpdateContext ctx, IEnumerable<T> models, Action<ISlideContext> action)
        => ctx.ApplyModels(models, action == null ? null : (a, x, i) => action(a));

}