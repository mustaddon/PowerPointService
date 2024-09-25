using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using System;
using System.Drawing;
using System.IO;
using System.Linq;

namespace PowerPointTool;

public partial class PPTool
{
    internal void AddImage(SlidePart slidePart, Stream image, string type, Rectangle shape)
    {
        if (image == null || image.Length == 0)
            return;

        var imagePart = slidePart.AddImagePart(type); 
        imagePart.FeedData(image);

        var tree = slidePart.Slide.Descendants<ShapeTree>().First();
        var guid = Guid.NewGuid();

        var useLocalDpi = new DocumentFormat.OpenXml.Office2010.Drawing.UseLocalDpi() { Val = false };
        useLocalDpi.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

        tree.Append(new Picture(
            new BlipFill(
                new DocumentFormat.OpenXml.Drawing.Blip(
                    new DocumentFormat.OpenXml.Drawing.BlipExtensionList(
                        new DocumentFormat.OpenXml.Drawing.BlipExtension(useLocalDpi)
                        {
                            Uri = guid.ToString("B"),
                        }))
                {
                    Embed = slidePart.GetIdOfPart(imagePart)
                },

                new DocumentFormat.OpenXml.Drawing.Stretch(
                    new DocumentFormat.OpenXml.Drawing.FillRectangle())))
        {
            NonVisualPictureProperties = new(
                new NonVisualDrawingProperties
                {
                    Name = $"Image Shape {guid}",
                    Id = (uint)tree.ChildElements.Count - 1
                },

                new NonVisualPictureDrawingProperties(
                    new DocumentFormat.OpenXml.Drawing.PictureLocks()
                    {
                        NoChangeAspect = true,
                    }),

                new ApplicationNonVisualDrawingProperties()
            ),

            ShapeProperties = new(
                new DocumentFormat.OpenXml.Drawing.PresetGeometry
                {
                    Preset = DocumentFormat.OpenXml.Drawing.ShapeTypeValues.Rectangle
                })
            {
                Transform2D = new(
                    new DocumentFormat.OpenXml.Drawing.Offset
                    {
                        X = shape.Left,
                        Y = shape.Top
                    },
                    new DocumentFormat.OpenXml.Drawing.Extents
                    {
                        Cx = Math.Abs(shape.Width),
                        Cy = Math.Abs(shape.Height),
                    })
                {
                    HorizontalFlip = shape.Width < 0,
                    VerticalFlip = shape.Height < 0,
                }
            },
        });
    }



}
