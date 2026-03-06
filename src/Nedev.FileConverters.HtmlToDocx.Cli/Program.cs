using Nedev.FileConverters.HtmlToDocx;

namespace Nedev.FileConverters.HtmlToDocx.Cli;

class Program
{
    static async Task<int> Main(string[] args)
    {
        Console.WriteLine("Nedev.FileConverters.HtmlToDocx - High Performance HTML to DOCX Converter");
        Console.WriteLine("==========================================================");
        Console.WriteLine();

        if (args.Length < 2)
        {
            Console.WriteLine("Usage:");
            Console.WriteLine("  htmltodocx <input.html> <output.docx>    Convert single file");
            Console.WriteLine();
            return 1;
        }

        try
        {
            var inputPath = args[0];
            var outputPath = args[1];

            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found: {inputPath}");
                return 1;
            }

            Console.WriteLine($"Converting: {inputPath}");
            Console.WriteLine($"Output: {outputPath}");

            var html = await File.ReadAllTextAsync(inputPath);
            Console.WriteLine($"HTML size: {html.Length:N0} characters");
            Console.WriteLine();

            var stopwatch = System.Diagnostics.Stopwatch.StartNew();

            using var service = new HtmlToDocxService();
            var docxBytes = await service.ConvertAsync(html);

            stopwatch.Stop();

            await File.WriteAllBytesAsync(outputPath, docxBytes);

            Console.WriteLine("Conversion completed successfully!");
            Console.WriteLine($"Time: {stopwatch.ElapsedMilliseconds} ms");
            Console.WriteLine($"Output size: {docxBytes.Length:N0} bytes");
            Console.WriteLine($"Throughput: {html.Length / Math.Max(1, stopwatch.ElapsedMilliseconds):N0} chars/ms");

            return 0;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
            return 1;
        }
    }
}
