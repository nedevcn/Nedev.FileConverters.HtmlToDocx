using System.IO.Compression;

namespace Nedev.FileConverters.HtmlToDocx.Core.Docx;

public sealed class ZipArchiveHelper : IDisposable
{
    private readonly MemoryStream _stream;
    private readonly ZipArchive _archive;

    public ZipArchiveHelper()
    {
        _stream = new MemoryStream();
        _archive = new ZipArchive(_stream, ZipArchiveMode.Create, true);
    }

    public void AddEntry(string entryName, byte[] content)
    {
        var entry = _archive.CreateEntry(entryName, CompressionLevel.Optimal);
        using var entryStream = entry.Open();
        entryStream.Write(content, 0, content.Length);
    }

    public void AddEntry(string entryName, string content)
    {
        AddEntry(entryName, System.Text.Encoding.UTF8.GetBytes(content));
    }

    public byte[] ToArray()
    {
        _archive.Dispose();
        return _stream.ToArray();
    }

    public void Dispose()
    {
        _archive.Dispose();
        _stream.Dispose();
    }
}
