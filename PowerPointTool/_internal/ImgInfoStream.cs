using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace PowerPointTool._internal;

internal class ImgInfoStream(Stream source) : Stream
{
    public override bool CanRead => source.CanRead;
    public override bool CanSeek => source.CanSeek;
    public override bool CanWrite => source.CanWrite;
    public override long Length => source.Length;
    public override long Position { get => source.Position; set => source.Position = value; }


    public override void Flush() => source.Flush();

    public override long Seek(long offset, SeekOrigin origin) => source.Seek(offset, origin);

    public override void SetLength(long value) => source.SetLength(value);

    public override void Write(byte[] buffer, int offset, int count) => source.Write(buffer, offset, count);

    public override Task WriteAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) => source.WriteAsync(buffer, offset, count, cancellationToken);

    public override int Read(byte[] buffer, int offset, int count) => OnReaded(buffer, offset, source.Read(buffer, offset, count));

    public override async Task<int> ReadAsync(byte[] buffer, int offset, int count, CancellationToken cancellationToken) => OnReaded(buffer, offset, await source.ReadAsync(buffer, offset, count, cancellationToken));


    public ImgFormat? Format => _info?.Format;
    public ImgOrientation? Orientation => _info?.Orientation;
    public uint? Width => _info?.Width;
    public uint? Height => _info?.Height;


    IImgInfoReader _info;
    byte[] _formatBuffer = new byte[_formatBufferSize];
    int _formatCount = 0;

    int OnReaded(byte[] buffer, int offset, int count)
    {
        if (_info != null)
        {
            _info.OnRead(buffer, offset, count);
        }
        else if (_formatCount < _formatBuffer.Length)
        {
            Array.Copy(buffer, offset, _formatBuffer, _formatCount, Math.Min(_formatBuffer.Length - _formatCount, count));

            var readead = _formatCount + count;
            var format = _imgFormats.FirstOrDefault(kvp => kvp.Magic.Length <= readead && kvp.Magic.SequenceEqual(_formatBuffer.Take(kvp.Magic.Length)));

            if (format.Type != null)
            {
                _info = (IImgInfoReader)Activator.CreateInstance(format.Type);
                var extraOffset = format.Magic.Length - _formatCount;
                _info.OnRead(buffer, offset + extraOffset, count - extraOffset);
            }

            _formatCount += count;
        }

        return count;
    }

    static readonly List<(Type Type, byte[] Magic)> _imgFormats = typeof(IImgInfoReader).Assembly.GetTypes()
        .Where(x => x.IsClass && !x.IsAbstract && typeof(IImgInfoReader).IsAssignableFrom(x))
        .Select(x => (x, x.GetField(nameof(JpegInfoReader.Magic), BindingFlags.Public | BindingFlags.Static)?.GetValue(null) as byte[]))
        .ToList();

    static readonly int _formatBufferSize = _imgFormats.Select(x => x.Magic.Length).Max();
}

