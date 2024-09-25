using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPointTool._internal;
using System;
using System.IO;
using System.Linq;

namespace PowerPointTool;

public partial class PPTool
{

    public virtual void MergeSlides(Stream targetPresentation, Stream sourcePresentation, Func<ISlideContext, int?> sourceToTargetMap)
    {
        using var target = PresentationDocument.Open(targetPresentation, true);
        using var source = PresentationDocument.Open(sourcePresentation, false);
        _mergeSlides(source, target, sourceToTargetMap);
    }

    void _mergeSlides(PresentationDocument source, PresentationDocument target, Func<ISlideContext, int?> map)
    {
        if (target.PresentationPart.Presentation.SlideIdList == null)
            target.PresentationPart.Presentation.SlideIdList = new SlideIdList();

        var slideMasterPartsMap = _cloneSlideMasterParts(source, target);
        var nextId = GetMaxSlideId(target.PresentationPart.Presentation.SlideIdList) + 1;
        var sourceSlideIds = source.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToArray();

        for (var i = 0; i < sourceSlideIds.Length; i++)
        {
            var sourceSlideId = sourceSlideIds[i];
            var sourceSlide = (SlidePart)source.PresentationPart.GetPartById(sourceSlideId.RelationshipId);
            var ctx = new SlideContext(this, source.PresentationPart, sourceSlide, i, sourceSlideIds.Length);
            var insertAt = map(ctx);

            if (insertAt.HasValue)
            {
                var targetSlidesCount = target.PresentationPart.Presentation.SlideIdList.Count();
                _insertSlidePart(source, sourceSlide, target, slideMasterPartsMap, _getIndex(insertAt.Value, targetSlidesCount), nextId++);
            }
        }
    }

}
