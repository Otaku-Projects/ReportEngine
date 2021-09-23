using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using jsreport.Binary;
using jsreport.Local;
using jsreport.Shared;
using jsreport.Types;

namespace JasperReport.ReportEntity
{
    public class BaseReportEntity
    {
        protected IRenderService rs;
        protected DataSet dataSet;
        protected IDictionary<string, object> dataSetObj;

        protected string rptFilesFolder;

        protected string templateBaseDirectory;
        protected string templateReportFileDirectory;
        protected string templateReportFileLocation;

        public BaseReportEntity()
        {
            this.templateBaseDirectory = @"D:\Documents\ReportEngine\SolutionRoot\JasperReport\ReportTemplate";
            this.templateBaseDirectory = Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate");

            /*
            this.rs = new LocalReporting()
                .UseBinary(JsReportBinary.GetBinary())
                .AsUtility()
                .Create();
            */
            this.rs = new LocalReporting()
                .RunInDirectory(Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate"))
                .KillRunningJsReportProcesses()
                .UseBinary(JsReportBinary.GetBinary())
                .Configure(cfg => cfg.AllowedLocalFilesAccess().FileSystemStore().BaseUrlAsWorkingDirectory())
                .AsUtility()
                .Create();
        }

        protected void SetRenderService(IRenderService rs)
        {
            this.rs = rs;
        }

        public IRenderService GetRenderService()
        {
            return this.rs;
        }

        protected void SetDataSet(DataSet _dataSet)
        {
            this.dataSet = _dataSet;
        }

        public DataSet GetDataSet()
        {
            return this.dataSet;
        }

        protected void SetDataSetObj(IDictionary<string, object> _dataSetObj)
        {
            this.dataSetObj = _dataSetObj;
        }

        public IDictionary<string, object> GetDataSetObj()
        {
            return this.dataSetObj;
        }

        public string GetTemplateFilePath()
        {
            return this.templateReportFileLocation;
        }

        public string GetTemplateFileDirectory()
        {
            return this.templateReportFileDirectory;
        }
    }
}
