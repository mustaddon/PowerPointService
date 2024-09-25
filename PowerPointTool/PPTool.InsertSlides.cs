using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPointTool._internal;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace PowerPointTool;

public partial class PPTool
{

    public virtual void InsertSlides(Stream targetPresentation, int targetInsertIndex, Stream sourcePresentation, Func<ISlideContext, bool> sourceSlideSelector = null)
    {
        using var target = PresentationDocument.Open(targetPresentation, true);
        using var source = PresentationDocument.Open(sourcePresentation, false);
        _insertSlides(source, target, targetInsertIndex, sourceSlideSelector);
    }

    void _insertSlides(PresentationDocument source, PresentationDocument target, int targetInsertIndex, Func<ISlideContext, bool> sourceSlideSelector)
    {
        if (target.PresentationPart.Presentation.SlideIdList == null)
            target.PresentationPart.Presentation.SlideIdList = new SlideIdList();

        var slideMasterPartsMap = _cloneSlideMasterParts(source, target);
        var targetSlidesCount = target.PresentationPart.Presentation.SlideIdList.Count();
        var index = _getIndex(targetInsertIndex, targetSlidesCount);
        var nextId = GetMaxSlideId(target.PresentationPart.Presentation.SlideIdList) + 1;
        var sourceSlideIds = source.PresentationPart.Presentation.SlideIdList.Elements<SlideId>().ToArray();

        for (var i = 0; i < sourceSlideIds.Length; i++)
        {
            var sourceSlideId = sourceSlideIds[i];
            var sourceSlide = (SlidePart)source.PresentationPart.GetPartById(sourceSlideId.RelationshipId);
            var ctx = new SlideContext(this, source.PresentationPart, sourceSlide, i, sourceSlideIds.Length);

            if (sourceSlideSelector?.Invoke(ctx) != false)
                _insertSlidePart(source, sourceSlide, target, slideMasterPartsMap, index++, nextId++);
        }
    }

    static void _insertSlidePart(PresentationDocument source, SlidePart sourceSlidePart, PresentationDocument target, Dictionary<SlideMasterPart, SlideMasterPart> slideMasterPartsMap, int targetIndex, uint targetSlideId)
    {
        var sourceSlideLayoutPartId = sourceSlidePart.SlideLayoutPart.SlideMasterPart.GetIdOfPart(sourceSlidePart.SlideLayoutPart);
        var targetSlideMasterPart = slideMasterPartsMap[sourceSlidePart.SlideLayoutPart.SlideMasterPart];
        var targetSlideLayoutPart = (SlideLayoutPart)targetSlideMasterPart.GetPartById(sourceSlideLayoutPartId);//.SlideLayoutParts.First(x => x.SlideLayoutParts).FirstOrDefault(x => x.SlideLayout.CommonSlideData.Name.Value.IndexOf(sourceSlidePart.SlideLayoutPart.SlideLayout.CommonSlideData.Name, StringComparison.InvariantCultureIgnoreCase) >= 0)
        var targetSlidePart = target.PresentationPart.AddPart(sourceSlidePart);

        targetSlidePart.DeleteParts(targetSlidePart.Parts.Select(x => x.OpenXmlPart)
            .Where(x => x == targetSlidePart.SlideLayoutPart
                || x.GetType() == typeof(NotesSlidePart)));

        targetSlidePart.AddPart(targetSlideLayoutPart);
        target.PresentationPart.Presentation.SlideIdList.InsertAt(
            new SlideId() { Id = targetSlideId, RelationshipId = target.PresentationPart.GetIdOfPart(targetSlidePart) },
            targetIndex);
    }

    static Dictionary<SlideMasterPart, SlideMasterPart> _cloneSlideMasterParts(PresentationDocument source, PresentationDocument target)
    {
        var nextId = _getMaxSlideMasterId(target.PresentationPart.Presentation.SlideMasterIdList) + 1;
        var mapping = new Dictionary<SlideMasterPart, SlideMasterPart>();

        foreach (var sourceSlideMasterId in source.PresentationPart.Presentation.SlideMasterIdList.Elements<SlideMasterId>())
        {
            var sourceSlideMasterPart = (SlideMasterPart)source.PresentationPart.GetPartById(sourceSlideMasterId.RelationshipId);
            var targetSlideMasterPart = target.PresentationPart.AddPart(sourceSlideMasterPart);
            var targetSlideMasterId = new SlideMasterId() { Id = nextId++, RelationshipId = target.PresentationPart.GetIdOfPart(targetSlideMasterPart) };
            target.PresentationPart.Presentation.SlideMasterIdList.Append(targetSlideMasterId);
            mapping.Add(sourceSlideMasterPart, targetSlideMasterPart);
        }

        foreach (var slideMasterPart in target.PresentationPart.SlideMasterParts)
        {
            if (slideMasterPart.SlideMaster.SlideLayoutIdList != null)
                foreach (SlideLayoutId slideLayoutId in slideMasterPart.SlideMaster.SlideLayoutIdList)
                    slideLayoutId.Id = nextId++;

            slideMasterPart.SlideMaster.Save();
        }

        return mapping;
    }

    static int _getIndex(int insertAt, int count)
    {
        return insertAt < 0 ? Math.Max(0, count + insertAt + 1) : Math.Min(count, insertAt);
    }
}
