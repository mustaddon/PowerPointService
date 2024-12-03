using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using PowerPointTool._internal;
using System;
using System.Drawing;
using System.IO;
using System.Linq;

namespace PowerPointTool;

public partial class PPTool
{
    internal void AddImage(SlidePart slidePart, Stream image, string type, Rectangle shapeSource, bool fit)
    {
        if (image == null || image.Length == 0)
            return;

        var imagePart = slidePart.AddImagePart(string.IsNullOrEmpty(type) ? "image/*" : type);
        var imageWithInfo = new ImgInfoStream(image);
        imagePart.FeedData(imageWithInfo);

        var (rotation, shape) = ApplyInfo(shapeSource, imageWithInfo, fit);
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
                    Embed = slidePart.GetIdOfPart(imagePart),
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
                        X = shape.X,
                        Y = shape.Y,

                    },
                    new DocumentFormat.OpenXml.Drawing.Extents
                    {
                        Cx = Math.Abs(shape.Width),
                        Cy = Math.Abs(shape.Height),
                    })
                {
                    HorizontalFlip = shape.Width < 0,
                    VerticalFlip = shape.Height < 0,
                    Rotation = rotation * 60000,
                }
            },
        });
    }

    static (int, Rectangle) ApplyInfo(Rectangle shape, ImgInfoStream imageInfo, bool fit)
    {
        var rotate = 0;

        if (imageInfo.Orientation.HasValue)
        {
            var orientation = imageInfo.Orientation.Value;

            rotate = orientation.HasFlag(ImgOrientation.Rotate90CW) ? 90
                : orientation.HasFlag(ImgOrientation.Rotate180) ? 180
                : orientation.HasFlag(ImgOrientation.Rotate270CW) ? 270
                : 0;

            if (orientation.HasFlag(ImgOrientation.Rotate90CW) || orientation.HasFlag(ImgOrientation.Rotate270CW))
            {
                var halfDelta = (Math.Abs(shape.Width) - Math.Abs(shape.Height)) / 2;
                shape = new Rectangle(shape.X + halfDelta, shape.Y - halfDelta, shape.Height, shape.Width);
            }

            if (orientation.HasFlag(ImgOrientation.MirrorVertical) == true)
                shape = new Rectangle(shape.X, shape.Y, shape.Width, -shape.Height);

            if (orientation.HasFlag(ImgOrientation.MirrorHorizontal) == true)
                shape = new Rectangle(shape.X, shape.Y, -shape.Width, shape.Height);
        }

        if (!fit && imageInfo.Width.HasValue && imageInfo.Height.HasValue)
        {
            var imgAR = 1d * imageInfo.Width.Value / imageInfo.Height.Value;

            var width = shape.Width;
            var height = (int)Math.Abs(shape.Width / imgAR) * Sign(shape.Height);

            if (Math.Abs(height) > Math.Abs(shape.Height))
            {
                height = shape.Height;
                width = (int)Math.Abs(shape.Height * imgAR) * Sign(shape.Width);
            }

            shape = new Rectangle(
                Math.Abs(shape.Width - width) / 2 + shape.X,
                Math.Abs(shape.Height - height) / 2 + shape.Y,
                width,
                height);
        }

        return (rotate, shape);
    }

    static int Sign(int val) => val < 0 ? -1 : 1;

}
