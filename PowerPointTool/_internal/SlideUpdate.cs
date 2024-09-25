using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Collections.Generic;

namespace PowerPointTool._internal;

internal class SlideUpdate(
    PPTool pps,
    PresentationPart presentationPart,
    SlideId slideId,
    int index,
    int totalCount)
    : SlideContext(pps, presentationPart, (SlidePart)presentationPart.GetPartById(slideId.RelationshipId), index, totalCount), ISlideUpdateContext
{
    readonly SlideId _slideId = slideId;


    public void ApplyModels<T>(IEnumerable<T> models, Action<ISlideContext, T, int> action = null)
    {
        _service.ApplyModels(_presentationPart, _slideId, _slidePart, models, action == null ? null
            : (x, i, slideId) => action(
                new SlideContext(
                    _service, 
                    _presentationPart, 
                    (SlidePart)_presentationPart.GetPartById(slideId.RelationshipId), 
                    SlideIndex + i, 
                    SlidesCount + i), 
                x, i));
    }
}
