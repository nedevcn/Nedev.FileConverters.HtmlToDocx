using Xunit;
using System;
using System.Reflection;
using Nedev.FileConverters.HtmlToDocx.Core.Conversion;

namespace Nedev.FileConverters.HtmlToDocx.Tests
{
    public class ImageTests
    {
        [Fact]
        public void ProcessImage_ReturnsActualDimensionsFromDataUri()
        {
            // create a 10x20 bitmap in memory
            using var bmp = new System.Drawing.Bitmap(10,20);
            using var g = System.Drawing.Graphics.FromImage(bmp);
            g.Clear(System.Drawing.Color.Blue);
            using var ms = new System.IO.MemoryStream();
            bmp.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
            var bytes = ms.ToArray();
            var dataUri = "data:image/png;base64," + Convert.ToBase64String(bytes);

            var method = typeof(HtmlToDocxConverter).GetMethod("ProcessImage", BindingFlags.NonPublic | BindingFlags.Instance);
            var converter = new HtmlToDocxConverter();
            var taskObj = method!.Invoke(converter, new object[] { dataUri });
            var task = (System.Threading.Tasks.Task)taskObj!;
            task.Wait();
            var result = task.GetType().GetProperty("Result")!.GetValue(task);
            Assert.NotNull(result);
            int width = (int)result.GetType().GetProperty("Width")!.GetValue(result)!;
            int height = (int)result.GetType().GetProperty("Height")!.GetValue(result)!;
            Assert.Equal(10, width);
            Assert.Equal(20, height);
        }
    }
}