using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace PowerPointTool;

public partial class PPTool
{

    public virtual int SlideIndex(Stream presentation, Regex regex)
    {
        using var doc = PresentationDocument.Open(presentation, false);
        var slist = doc.PresentationPart.Presentation.SlideIdList;
        var slideIds = slist.ChildElements.Cast<SlideId>().ToArray();

        for (var i = 0; i < slideIds.Length; i++)
        {
            var slideId = slideIds[i];
            var slide = (SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId);
            if (regex.IsMatch(slide.Slide.OuterXml))
                return i;
        }
        return -1;
    }

}
