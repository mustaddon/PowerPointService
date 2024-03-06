using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPointTool._internal;
using System;
using System.IO;
using System.Linq;

namespace PowerPointTool;

public partial class PPTool
{
    public virtual void UpdateSlides(Stream presentation, Action<ISlideUpdateContext> slideUpdate)
    {
        using var doc = PresentationDocument.Open(presentation, true);

        var slist = doc.PresentationPart.Presentation.SlideIdList;
        var slideIds = slist.ChildElements.Cast<SlideId>().ToArray();

        for (var i = 0; i < slideIds.Length; i++)
        {
            var slideId = slideIds[i];
            var ctx = new SlideUpdate(this, doc.PresentationPart, slideId, i, slideIds.Length);
            slideUpdate(ctx);
        }
    }

}
