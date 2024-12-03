using DocumentFormat.OpenXml.Packaging;
using System.Drawing;
using System.IO;

namespace PowerPointTool._internal;

internal class SlideContext(PPTool pps, PresentationPart presentationPart, SlidePart slidePart, int index, int totalCount) : ISlideContext
{
    protected readonly PPTool _service = pps;
    protected readonly PresentationPart _presentationPart = presentationPart;
    protected readonly SlidePart _slidePart = slidePart;

    public int SlidesCount => totalCount;
    public int SlideIndex => index;
    public string SlideXml => _slidePart.Slide.OuterXml;
    public int SlideWidth => _presentationPart.Presentation.SlideSize.Cx ?? 0;
    public int SlideHeight => _presentationPart.Presentation.SlideSize.Cy ?? 0;


    public void AddImage(Stream image, string type, Rectangle? shape, bool fit)
    {
        _service.AddImage(_slidePart, image, type,
            shape.HasValue && shape != Rectangle.Empty ? shape.Value : new Rectangle(0, 0, SlideWidth, SlideHeight), 
            fit);
    }
}
