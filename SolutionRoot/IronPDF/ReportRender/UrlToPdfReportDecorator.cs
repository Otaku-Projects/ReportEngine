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
        protected List<PdfDocument> PdfDocumentList;
        public UrlToPdfReportDecorator() : base()
        {
        }
        public UrlToPdfReportDecorator(UrlToPdfReportEntity _reportEntity, string _filename) : base(_reportEntity, _filename = "")
        {
            // check filename

            // check urlPath

            // Create a PDF from a URL or local file path
            var urlPathList = _reportEntity.GetUrlPath();
            foreach(string urlPath in urlPathList)
            {
                this.PdfDocumentList.Add(this.renderer.RenderUrlAsPdf(urlPath));
            }
        }

        public override void Initialize()
        {
            base.Initialize();
            this.PdfDocumentList = new List<PdfDocument>();
        }


        public override void SaveFile()
        {
            string pdfFilePath = string.Empty;

            // Export to a file or Stream
            foreach(PdfDocument pdfDocument in PdfDocumentList)
            {
                pdfFilePath = Path.Combine(
                this.ironRenderFolder,
                this.filename + ".pdf");

                pdfDocument.SaveAs(pdfFilePath);

                this.ReGenFilename();
            }
        }
    }
}
