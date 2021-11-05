
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace Puppeteer.ReportEntity
{
    public class ReportReference1 : PuppeteerReportEntity
    {
        public ReportReference1(DataSet _dataSet)
        {
            Console.WriteLine("Said \"Hello World!\" from ReportReference1");
            //this.dataSet = _dataSet;
            this.dataSet = _dataSet;
        }

        public ReportReference1(IDictionary<string, object> _dataSetObj)
        {
            Console.WriteLine("Said \"Hello World!\" from ReportReference1");
            //this.dataSet = _dataSet;
            this.dataSetObj = _dataSetObj;
        }
        public override void InitializateMetaData()
        {
            string _templateDirectory = string.Empty;
            string _contentFilePath = string.Empty;
            string _templateScriptLocation = string.Empty;
            _templateDirectory = Path.Combine(this.templateBaseDirectory, @"ReportReference1");

            this.templateReportFileDirectory = _templateDirectory;
            this.SetPdfTemplateFileName("index.html");
        }
        public override void InitializateMainContent()
        {
        }

        public override void InitializateDataGrid()
        {
        }

        public override void InitializateHeaderFooter()
        {
        }
    }
}
