using System;

namespace PowerPointTool._internal;

interface IImgInfoReader
{
    ImgFormat Format { get; }
    ImgOrientation? Orientation { get; }
    uint? Width { get; }
    uint? Height { get; }

    void OnRead(byte[] buffer, int offset, int count);
}

enum ImgFormat
{
    Jpeg = 1,
    Png = 2,
}


[Flags]
enum ImgOrientation
{
    MirrorHorizontal = 1,
    MirrorVertical = 2,
    Rotate90CW = 4,
    Rotate180 = 8,
    Rotate270CW = 16,
}


