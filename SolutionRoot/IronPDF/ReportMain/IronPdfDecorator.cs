using CoreReport;
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
                Guid obj = Guid.NewGuid();
                _filename = obj.ToString();
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

        public void Initialize()
        {
            this.ironRenderFolder = this.tempRenderFolder;
        }
    }
}
