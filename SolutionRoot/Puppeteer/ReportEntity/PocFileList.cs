
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
    public class PocFileList : PuppeteerReportEntity
    {
        public PocFileList(DataSet _dataSet)
        {
            Console.WriteLine("Said \"Hello World!\" from PocFileList");
            //this.dataSet = _dataSet;
            this.dataSet = _dataSet;
        }

        public PocFileList(IDictionary<string, object> _dataSetObj)
        {
            Console.WriteLine("Said \"Hello World!\" from PocFileList");
            //this.dataSet = _dataSet;
            this.dataSetObj = _dataSetObj;
        }
        public override void InitializateMetaData()
        {
        }
        public override void InitializateMainContent()
        {
            string _templateDirectory = string.Empty;
            string _contentFilePath = string.Empty;
            string _templateScriptLocation = string.Empty;
            _templateDirectory = Path.Combine(this.templateBaseDirectory, @"PocFileList");

            this.templateReportFileDirectory = _templateDirectory;
            this.SetPdfTemplateFileName("index.html");
        }

        public override void InitializateDataGrid()
        {
        }

        public override void InitializateHeaderFooter()
        {
        }
    }
}
