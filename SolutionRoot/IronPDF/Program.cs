using System;
using IronPdf;

namespace IronPDF
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            // Instantiate Renderer
            var Renderer = new IronPdf.ChromePdfRenderer();

            // Create a PDF from a URL or local file path
            var pdf = Renderer.RenderUrlAsPdf("https://ironpdf.com/");

            // Export to a file or Stream
            pdf.SaveAs("url.pdf");
        }
    }
}
