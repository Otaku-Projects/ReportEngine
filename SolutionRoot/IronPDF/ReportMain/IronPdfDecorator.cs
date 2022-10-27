using CoreReport;
using IronPdf;
using IronPDFProject.ReportEntity;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IronPDFProject.ReportMain
{
    public class IronPdfDecorator : VisualizationDecorator
    {
        protected ChromePdfRenderer renderer;
        protected PdfDocument pdfDocument;

        protected string createdBy;
        protected DateTime createdDate;
        protected DateTime printedDate;
        protected string filename;

        protected IronPdfReportEntity reportEntity;

        protected IDictionary<string, object> dataSetObj;
        protected string ironRenderFolder;

        public List<string> _fonts;

        protected string report_instance_dir;
        protected string report_template_dir;
        protected string fonts_folder;

        public IronPdfDecorator()
        {
            this.Initialize();
        }
        public IronPdfDecorator(IronPdfReportEntity _reportEntity, string _filename = "")
        {
            if (string.IsNullOrEmpty(_filename))
            {
                _filename = this.ReGenFilename();
            }

            this.dataSetObj = _reportEntity.GetDataSetObj();

            this.filename = _filename;

            this.reportEntity = _reportEntity;

            this.createdBy = "CoreSystem";
            this.createdDate = new DateTime();

            this.report_instance_dir = string.Empty;
            this.report_template_dir = string.Empty;

            this.Initialize();
        }

        public virtual void Initialize()
        {
            // Instantiate Renderer
            this.renderer = new ChromePdfRenderer();

            this.ironRenderFolder = this.tempRenderFolder;
        }
        public string ReGenFilename()
        {
            Guid obj = Guid.NewGuid();
            string _filename = obj.ToString();
            this.filename = _filename;

            return _filename;
        }
    }
}
