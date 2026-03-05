using Nedev.HtmlToDocx.Core.Conversion;
using Nedev.HtmlToDocx.Core.Docx;
using Nedev.HtmlToDocx.Core.Models;

namespace Nedev.HtmlToDocx;

public sealed class HtmlToDocxService : IDisposable
{
    private readonly ConverterOptions _options;
    private HtmlToDocxConverter? _converter;
    private bool _disposed;

    public HtmlToDocxService()
    {
        _options = new ConverterOptions();
    }

    public HtmlToDocxService(ConverterOptions options)
    {
        _options = options ?? new ConverterOptions();
    }

    public byte[] Convert(string html)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        _converter ??= new HtmlToDocxConverter(_options);
        return _converter.Convert(html);
    }

    public async Task<byte[]> ConvertAsync(string html, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        _converter ??= new HtmlToDocxConverter(_options);
        return await _converter.ConvertAsync(html, cancellationToken);
    }

    public void ConvertFile(string inputPath, string outputPath)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var html = File.ReadAllText(inputPath);
        var docxBytes = Convert(html);
        File.WriteAllBytes(outputPath, docxBytes);
    }

    public async Task ConvertFileAsync(string inputPath, string outputPath, CancellationToken cancellationToken = default)
    {
        ObjectDisposedException.ThrowIf(_disposed, this);
        var html = await File.ReadAllTextAsync(inputPath, cancellationToken);
        var docxBytes = await ConvertAsync(html, cancellationToken);
        await File.WriteAllBytesAsync(outputPath, docxBytes, cancellationToken);
    }

    public void Dispose()
    {
        if (!_disposed)
        {
            _disposed = true;
            _converter?.Dispose();
        }
    }
}

public static class HtmlToDocxExtensions
{
    public static byte[] ToDocx(this string html, ConverterOptions? options = null)
    {
        using var service = new HtmlToDocxService(options ?? new ConverterOptions());
        return service.Convert(html);
    }

    public static async Task<byte[]> ToDocxAsync(this string html, ConverterOptions? options = null, CancellationToken cancellationToken = default)
    {
        using var service = new HtmlToDocxService(options ?? new ConverterOptions());
        return await service.ConvertAsync(html, cancellationToken);
    }
}
