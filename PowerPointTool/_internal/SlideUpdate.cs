using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;

namespace PowerPointTool._internal;

internal class SlideUpdate(
    PPTool pps,
    PresentationPart presentationPart,
    SlideId slideId,
    int index,
    int totalCount)
    : SlideContext(presentationPart, (SlidePart)presentationPart.GetPartById(slideId.RelationshipId), index, totalCount), ISlideUpdateContext
{
    readonly PPTool _service = pps;
    readonly SlideId _slideId = slideId;


    public void AddImage(Stream image, string type, Rectangle? shape)
    {
        _service.AddImage(_slidePart, image, type,
            shape.HasValue && shape != Rectangle.Empty ? shape.Value : new Rectangle(0, 0, SlideWidth, SlideHeight));

        _presentationPart.Presentation.Save();
    }

    public void ApplyModels<T>(IEnumerable<T> models, Action<ISlideUpdateModelContext, T> action)
    {
        _service.ApplyModels(_presentationPart, _slideId, _slidePart, models, action == null ? null
            : (x, i, slideId) => action(new SlideUpdate(_service, _presentationPart, slideId, SlideIndex + i, SlidesCount + i), x));
    }
}
