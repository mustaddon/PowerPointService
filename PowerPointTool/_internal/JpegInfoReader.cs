namespace PowerPointTool._internal;


class JpegInfoReader() : BaseInfoReader(EXIF_IDF_LENGTH)
{
    public override ImgFormat Format => ImgFormat.Jpeg;

    public static byte[] Magic = [0xFF, 0xD8];

    const int EXIF_IDF_LENGTH = 2 + 2 + 4 + 4;

    byte _marker;
    int _markerNext;
    int _exifMatch = 0;
    int _exifStart = -1;
    bool _exifLE;
    int _ifdLeft = -1;

    protected override void DefaultProcess(byte val)
    {
        _processor = MarkerProcess;
    }

    void MarkerProcess(byte val)
    {
        if (_stack[1] != 0xFF)
        {
            _processor = null;
            return;
        }

        _marker = val;
        _processor = MarkerSizeProcess;
        WaitBytes(2);
    }

    void MarkerSizeProcess(byte val)
    {
        _markerNext = _position + GetUInt16();

        if (_marker == 0xE1 && !Orientation.HasValue)
        {
            _processor = ExifSearchProcess;
        }
        else if (_marker >= 0xC0 && _marker <= 0xC3 && !Width.HasValue)
        {
            _processor = ResolutionProcess;
            WaitBytes(1 + 2 + 2);
        }
        else
        {
            ProcessNextMarker();
        }
    }

    void ProcessNextMarker()
    {
        if (Orientation.HasValue && Width.HasValue && Height.HasValue)
        {
            _processor = null;
            return;
        }

        _positionWait = _markerNext;
        _processor = MarkerProcess;
    }

    void ResolutionProcess(byte val)
    {
        Width = GetUInt16();
        Height = GetUInt16(false, 2);
        ProcessNextMarker();
    }

    static readonly byte[] _exif = [0x45, 0x78, 0x69, 0x66, 0x00, 0x00];

    void ExifSearchProcess(byte val)
    {
        if (val == _exif[_exifMatch])
        {
            _exifMatch++;

            if (_exifMatch == _exif.Length)
            {
                _exifMatch = 0;
                _exifStart = _position + 1;
                _processor = ExifLeProcess;
                WaitBytes(2);
            }
        }
        else if (_exifMatch != 0)
        {
            _exifMatch = 0;
        }

        if (_markerNext - _position <= 2)
            ProcessNextMarker();
    }

    void ExifLeProcess(byte val)
    {
        _exifLE = val != 0x4D;
        _processor = IdfStartProcess;
        WaitBytes(2 + 4);
    }

    void IdfStartProcess(byte val)
    {
        var ifdOffset = GetUInt32(_exifLE);
        _positionWait = (int)(_exifStart + ifdOffset + 1);
        _processor = IdfCountProcess;
    }

    void IdfCountProcess(byte val)
    {
        _ifdLeft = GetUInt16(_exifLE);
        _processor = IdfProcess;
        WaitBytes(EXIF_IDF_LENGTH);
    }

    void IdfProcess(byte val)
    {
        var ifdId = GetUInt16(_exifLE, 10);

        if (ifdId == 0x0112)
        {
            SetOrientation(GetUInt16(_exifLE, 2));
            ProcessNextMarker();
            return;
        }

        _ifdLeft--;

        if (_ifdLeft > 0)
            WaitBytes(EXIF_IDF_LENGTH);
        else
            ProcessNextMarker();
    }
}

