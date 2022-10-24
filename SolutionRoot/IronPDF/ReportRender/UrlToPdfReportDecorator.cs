using IronPdf;
using IronPDFProject.ReportEntity;
using IronPDFProject.ReportMain;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IronPDFProject.ReportRender
{
    public class UrlToPdfReportDecorator : IronPdfDecorator
    {
        protected ChromePdfRenderer renderer;
        protected PdfDocument pdfDocument;

        public UrlToPdfReportDecorator() : base()
        {
            this.renderer = new ChromePdfRenderer();
        }
        public UrlToPdfReportDecorator(UrlToPdfReportEntity _reportEntity, string _filename, string urlPath) : base(_reportEntity, _filename = "")
        {
            // check filename

            // check urlPath

            // Instantiate Renderer
            this.renderer = new ChromePdfRenderer();

            // Create a PDF from a URL or local file path
            this.pdfDocument = this.renderer.RenderUrlAsPdf(urlPath);
        }


        public override void SaveFile()
        {
            string pdfFilePath = Path.Combine(
                this.ironRenderFolder,
                this.filename + ".pdf");

            // Export to a file or Stream
            this.pdfDocument.SaveAs(pdfFilePath);

        }
    }
}
