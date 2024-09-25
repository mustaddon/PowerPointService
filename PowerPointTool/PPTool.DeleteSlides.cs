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

    public virtual void DeleteSlides(Stream presentation, Func<ISlideContext, bool> slideSelector)
    {
        using var doc = PresentationDocument.Open(presentation, true);
        var slist = doc.PresentationPart.Presentation.SlideIdList;
        var slideIds = slist.ChildElements.Cast<SlideId>().ToArray();

        for (var i = 0; i < slideIds.Length; i++)
        {
            var slideId = slideIds[i];
            var slide = (SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId);
            var ctx = new SlideContext(this, doc.PresentationPart, slide, i, slideIds.Length);

            if (!slideSelector(ctx))
                continue;

            slist.RemoveChild(slideId);
            CleanCustomShow(doc.PresentationPart.Presentation.CustomShowList, slideId.RelationshipId);
            doc.PresentationPart.Presentation.Save();
            doc.PresentationPart.DeletePart(slide);
        }
    }

    static void CleanCustomShow(CustomShowList customShowList, string slideRelId)
    {
        if (customShowList != null)
            foreach (var customShow in customShowList.Elements<CustomShow>())
                if (customShow.SlideList != null)
                {
                    var slideListEntries = new LinkedList<SlideListEntry>();

                    foreach (SlideListEntry slideListEntry in customShow.SlideList.Elements())
                        if (slideListEntry.Id != null && slideListEntry.Id == slideRelId)
                            slideListEntries.AddLast(slideListEntry);

                    foreach (var slideListEntry in slideListEntries)
                        customShow.SlideList.RemoveChild(slideListEntry);
                }
    }

}
