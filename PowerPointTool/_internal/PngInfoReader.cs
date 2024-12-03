namespace PowerPointTool._internal;


class PngInfoReader() : BaseInfoReader(8)
{
    public override ImgFormat Format => ImgFormat.Png;

    public static byte[] Magic = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

    int _chunkNext;

    protected override void DefaultProcess(byte val)
    {
        _processor = ChunkProcess;
        WaitBytes(3 + 4);
    }

    void ChunkProcess(byte val)
    {
        var type = GetAscii(4);
        var length = GetUInt32(false, 4);
        _chunkNext = (int)(_position + length + 12);

        switch (type)
        {
            case "IHDR":
                _processor = ResolutionProcess; WaitBytes(4 + 4);
                return;

            //case "eXIf":
            //    _processor = ExifProcess; WaitBytes(2);
            //    return;

            default:
                ProcessNextChunk();
                break;
        }
    }

    void ProcessNextChunk()
    {
        if (Width.HasValue && Height.HasValue)
        {
            _processor = null;
            return;
        }

        _positionWait = _chunkNext;
        _processor = ChunkProcess;
    }

    void ResolutionProcess(byte val)
    {
        Height = GetUInt32();
        Width = GetUInt32(false, 4);
        ProcessNextChunk();
    }
}

