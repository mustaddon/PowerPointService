using System;
using System.Text;

namespace PowerPointTool._internal;


abstract class BaseInfoReader : IImgInfoReader
{
    public BaseInfoReader(int stackSize)
    {
        _processor = DefaultProcess;
        _stack = new(stackSize);
    }

    public ImgOrientation? Orientation { get; protected set; }
    public uint? Width { get; protected set; }
    public uint? Height { get; protected set; }

    protected readonly DropOutStack<byte> _stack;
    protected Action<byte> _processor;
    protected int _position = 0;
    protected int _positionWait = 0;


    public abstract ImgFormat Format { get; }
    protected abstract void DefaultProcess(byte val);

    public void OnRead(byte[] buffer, int offset, int count)
    {
        if (_processor == null)
            return;

        var skip = Math.Min(count, Math.Max(0, _positionWait - _position - _stack.Capacity));
        _position += skip;

        var i = skip + offset;
        var end = count + offset;
        byte current;
        var stackSize = _stack.Capacity;

        while (i < end)
        {
            _stack.Push(current = buffer[i]);

            var step = 1;

            if (_positionWait <= _position)
            {
                _processor(current);

                if (_processor == null)
                    return;

                step = Math.Min(end - i, Math.Max(1, _positionWait - _position - stackSize));
            }

            _position += step;
            i += step;
        }
    }

    protected void WaitBytes(int n)
    {
        _positionWait = _position + n;
    }

    protected string GetAscii(int length, int skip = 0)
    {
        var bytes = new byte[length];

        for (var i = 0; i < length; i++)
            bytes[length - i - 1] = _stack[i + skip];

        return Encoding.ASCII.GetString(bytes);
    }

    protected ushort GetUInt16(bool isLE = false, int skip = 0)
    {
        return ToUInt16(_stack[skip + 1], _stack[skip], isLE);
    }

    protected uint GetUInt32(bool isLE = false, int skip = 0)
    {
        return ToUInt32(_stack[skip + 3], _stack[skip + 2], _stack[skip + 1], _stack[skip], isLE);
    }

    static ushort ToUInt16(byte a, byte b, bool isLE = false)
    {
        return isLE
            ? (ushort)((b << 8) | a)
            : (ushort)((a << 8) | b);
    }

    static uint ToUInt32(byte a, byte b, byte c, byte d, bool isLE = false)
    {
        return isLE
            ? (uint)((d << 24) | (c << 16) | (b << 8) | a)
            : (uint)((a << 24) | (b << 16) | (c << 8) | d);
    }

    protected void SetOrientation(ushort exifValue)
    {
        Orientation = exifValue switch
        {
            2 => ImgOrientation.MirrorHorizontal,
            3 => ImgOrientation.Rotate180,
            4 => ImgOrientation.MirrorVertical,
            5 => ImgOrientation.MirrorHorizontal | ImgOrientation.Rotate270CW,
            6 => ImgOrientation.Rotate90CW,
            7 => ImgOrientation.MirrorHorizontal | ImgOrientation.Rotate90CW,
            8 => ImgOrientation.Rotate270CW,
            _ => (ImgOrientation)0,
        };
    }
}



//enum ExifOrientation
//{
//    Normal = 1,
//    MirrorHorizontal = 2,
//    Rotate180 = 3,
//    MirrorVertical = 4,
//    MirrorHorizontalRotate270CW = 5,
//    Rotate90CW = 6,
//    MirrorHorizontalRotate90CW = 7,
//    Rotate270CW = 8,
//}