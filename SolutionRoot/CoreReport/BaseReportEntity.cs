using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreReport
{
    public abstract class BaseReportEntity
    {
        protected DataSet dataSet;
        protected IDictionary<string, object> dataSetObj;

        protected string rptFilesFolder;

        protected string templateBaseDirectory;
        protected string templateReportFileDirectory;
        //protected string templateReportFileLocation;
        //protected string templateReportHeaderLocation;
        //protected string templateReportFooterLocation;

        protected string xlsxTemplateFileName;
        protected string pdfTemplateFileName;

        public abstract void InitializateMetaData();
        public abstract void InitializateMainContent();
        public abstract void InitializateHeaderFooter();
        public abstract void InitializateDataGrid();

        protected HeaderFooterOptions headerFooterOption;

        public BaseReportEntity()
        {
            this.templateBaseDirectory = @"D:\Documents\ReportEngine\SolutionRoot\JasperReport\ReportTemplate";
            // this return the start up project directory
            // e.g: "D:\\Documents\\CoreSystem\\WebApi\\" + \ReportTemplate
            this.templateBaseDirectory = Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate");
            // this return the program running directory
            // e.g: "D:\\Documents\\CoreSystem\\WebApi\\bin\\Debug\\net5.0" + \ReportTemplate
            this.templateBaseDirectory = Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ReportTemplate");

            this.xlsxTemplateFileName = string.Empty;
            this.pdfTemplateFileName = string.Empty;

            this.Initializate();
        }
        protected virtual void Initializate()
        {
            this.InitializateMetaData();
            this.InitializateMainContent();
            this.InitializateHeaderFooter();

            this.InitializateDataGrid();
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
            if (this.dataSetObj == null)
            {
                return this.dataSetObj;
            }
            // renew date, time on each get
            string _tableName = "ReportMetaData";
            this.dataSetObj.Remove(_tableName);
            dynamic _obj = new ExpandoObject();
            _obj = new
            {
                DateTime = DateTime.Now.ToString("dd MMMM yyyy HH:mm")
            };

            this.dataSetObj.Add(_tableName, _obj);

            return this.dataSetObj;
        }

        public string GetTemplateFileDirectory()
        {
            return this.templateReportFileDirectory;
        }
        public string GetXlsxTemplateFilePath()
        {
            return Path.Combine(this.templateReportFileDirectory, this.xlsxTemplateFileName);
        }
        public string GetPdfTemplateFilePath()
        {
            return Path.Combine(this.templateReportFileDirectory, this.pdfTemplateFileName);
        }
        public enum HeaderFooterOptions
        {
            None = 0,
            Header = 11,
            Footer = 12,
            HeaderFooterInSingleFile = 21,
            HeaderFooterInSeparateFile = 22
        }

        public enum PageNature
        {
            None = 0,
            Header = 11,
            Footer = 12,
            HeaderAndFooter = 13,
            MainContent = 14
        }

        public HeaderFooterOptions GetHeaderFooterOption()
        {
            return this.headerFooterOption;
        }

        protected void SetXlsxTemplateFileName(string _xlsxTemplateFileName)
        {
            this.xlsxTemplateFileName = _xlsxTemplateFileName;
        }
        public string GetXlsxTemplateFileName()
        {
            return this.xlsxTemplateFileName;
        }

        protected void SetPdfTemplateFileName(string _xlsxTemplateFileName)
        {
            this.pdfTemplateFileName = _xlsxTemplateFileName;
        }
        public string GetPdfTemplateFileName()
        {
            return this.pdfTemplateFileName;
        }
    }
}
