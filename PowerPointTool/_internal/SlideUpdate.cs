using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System.Drawing;
using System.IO;

namespace PowerPointTool._internal;

internal class SlideUpdate(PPTool pps, PresentationPart presentationPart, SlideId slideId, int index, int totalCount)
    : SlideContext(presentationPart, (SlidePart)presentationPart.GetPartById(slideId.RelationshipId), index, totalCount), ISlideUpdateContext
{
    readonly PPTool _service = pps;
    readonly SlideId _slideId = slideId;

    public void ApplyModel(object model)
    {
        if (model == null)
            return;

        var nextId = PPTool.GetMaxSlideId(_presentationPart.Presentation.SlideIdList) + 1;

        _service.ApplyModel(model, _presentationPart, _slideId, _slidePart, ref nextId);
    }

    public void AddImage(Stream image, string type, Rectangle? shape)
    {
        _service.AddImage(_slidePart, image, type, 
            shape.HasValue && shape != Rectangle.Empty ? shape.Value : new Rectangle(0, 0, SlideWidth, SlideHeight));

        _presentationPart.Presentation.Save();
    }
}
