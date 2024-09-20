using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Drawing;
using System.IO;
using System.Linq;

namespace PowerPointTool;

public partial class PPTool
{
    //public virtual void AddImages(Stream presentation, Func<ISlideContext, (byte[], Rectangle?)[]> slideImages)
    //{
    //    using var doc = PresentationDocument.Open(presentation, true);

    //    var slist = doc.PresentationPart.Presentation.SlideIdList;
    //    var slideIds = slist.ChildElements.Cast<SlideId>().ToArray();
    //    var size = doc.PresentationPart.Presentation.SlideSize;

    //    for (var i = 0; i < slideIds.Length; i++)
    //    {
    //        var slideId = slideIds[i];
    //        var slide = (SlidePart)doc.PresentationPart.GetPartById(slideId.RelationshipId);
    //        var ctx = new SlideContext(doc.PresentationPart, slide, i, slideIds.Length);
    //        var images = slideImages(ctx);

    //        if (images != null)
    //            foreach (var (image, shape) in images)
    //                AddImage(slide, image, ImagePartType.Png,
    //                    shape.HasValue && shape != Rectangle.Empty ? shape.Value : new Rectangle(0, 0, size.Cx, size.Cy));
    //    }

    //    doc.PresentationPart.Presentation.Save();
    //}

    internal void AddImage(SlidePart slidePart, Stream image, string type, Rectangle shape)
    {
        if (image == null || image.Length == 0)
            return;

        var imagePart = slidePart.AddImagePart(type);
        imagePart.FeedData(image);

        var tree = slidePart
            .Slide
            .Descendants<ShapeTree>()
            .First();

        var guid = Guid.NewGuid();
        var picture = new Picture();

        picture.NonVisualPictureProperties = new NonVisualPictureProperties();
        picture.NonVisualPictureProperties.Append(new NonVisualDrawingProperties
        {
            Name = $"Image Shape {guid}",
            Id = (uint)tree.ChildElements.Count - 1
        });

        var nonVisualPictureDrawingProperties = new NonVisualPictureDrawingProperties();
        nonVisualPictureDrawingProperties.Append(new DocumentFormat.OpenXml.Drawing.PictureLocks()
        {
            NoChangeAspect = true,
        });
        picture.NonVisualPictureProperties.Append(nonVisualPictureDrawingProperties);
        picture.NonVisualPictureProperties.Append(new ApplicationNonVisualDrawingProperties());

        var blipFill = new BlipFill();
        var blip1 = new DocumentFormat.OpenXml.Drawing.Blip()
        {
            Embed = slidePart.GetIdOfPart(imagePart)
        };
        var blipExtensionList1 = new DocumentFormat.OpenXml.Drawing.BlipExtensionList();
        var blipExtension1 = new DocumentFormat.OpenXml.Drawing.BlipExtension()
        {
            Uri = guid.ToString("B"),// "{28A0092B-C50C-407E-A947-70E740481C1C}"
        };
        var useLocalDpi1 = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi()
        {
            Val = false
        };
        useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");
        blipExtension1.Append(useLocalDpi1);
        blipExtensionList1.Append(blipExtension1);
        blip1.Append(blipExtensionList1);
        var stretch = new DocumentFormat.OpenXml.Drawing.Stretch();
        stretch.Append(new DocumentFormat.OpenXml.Drawing.FillRectangle());
        blipFill.Append(blip1);
        blipFill.Append(stretch);
        picture.Append(blipFill);

        picture.ShapeProperties = new ShapeProperties();
        picture.ShapeProperties.Transform2D = new DocumentFormat.OpenXml.Drawing.Transform2D();
        picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Offset
        {
            X = shape.Left,
            Y = shape.Top,
        });
        picture.ShapeProperties.Transform2D.Append(new DocumentFormat.OpenXml.Drawing.Extents
        {
            Cx = shape.Width,
            Cy = shape.Height,
        });
        picture.ShapeProperties.Append(new DocumentFormat.OpenXml.Drawing.PresetGeometry
        {
            Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
        });

        tree.Append(picture);
    }



}
