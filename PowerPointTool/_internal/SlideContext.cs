using DocumentFormat.OpenXml.Packaging;

namespace PowerPointTool._internal;

internal class SlideContext(PresentationPart presentationPart, SlidePart slidePart, int index, int totalCount) : ISlideContext
{
    protected readonly PresentationPart _presentationPart = presentationPart;
    protected readonly SlidePart _slidePart = slidePart;

    public int SlidesCount => totalCount;
    public int SlideIndex => index;
    public string SlideXml => _slidePart.Slide.OuterXml;
    public int SlideWidth => _presentationPart.Presentation.SlideSize.Cx ?? 0;
    public int SlideHeight => _presentationPart.Presentation.SlideSize.Cy ?? 0;
}
