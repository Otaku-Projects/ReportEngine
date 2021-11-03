using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using jsreport.Binary;
using jsreport.Local;
using jsreport.Shared;
using jsreport.Types;

namespace JasperReport.ReportEntity
{
    public class InvoiceReport2 : JasperReportEntity
    {
        public InvoiceReport2(DataSet _dataSet) { }

        public InvoiceReport2(IDictionary<string, object> _dataSetObj)
        {
            Console.WriteLine("Said \"Hello World!\" from InvoiceReport2");
            //this.dataSet = _dataSet;
            this.dataSetObj = _dataSetObj;
        }
        public override void InitializateMetaData() {
            this.headerFooterOption = HeaderFooterOptions.HeaderFooterInSingleFile;
        }
        public override void InitializateMainContent()
        {
            string _templateDirectory = string.Empty;
            string _contentFilePath = string.Empty;
            string _templateScriptLocation = string.Empty;
            _templateDirectory = Path.Combine(this.templateBaseDirectory, @"InvoiceReport");

            if (File.Exists(Path.Combine(_templateDirectory, @"index.html")))
            {
                _contentFilePath = Path.Combine(_templateDirectory, @"index.html");
            }
            else if (File.Exists(Path.Combine(_templateDirectory, @"index.htm")))
            {
                _contentFilePath = Path.Combine(_templateDirectory, @"index.htm");
            }

            if (File.Exists(Path.Combine(_templateDirectory, @"helper.js")))
            {
                _templateScriptLocation = Path.Combine(_templateDirectory, @"helper.js");
            }

            this.templateReportFileDirectory = _templateDirectory;

            PageComponent _pageMainContent = new PageComponent();
            _pageMainContent.SetDirectory(_templateDirectory);
            _pageMainContent.SetHtmlPath(_contentFilePath);
            _pageMainContent.SetScriptPath(_templateScriptLocation);

            this.AddPageContent(_pageMainContent);
        }

        public override void InitializateHeaderFooter()
        {
            string _templateDirectory = this.templateReportFileDirectory;

            string _headerFooterFilePath = string.Empty;

            if (File.Exists(Path.Combine(_templateDirectory, @"header-footer.html")))
            {
                _headerFooterFilePath = Path.Combine(_templateDirectory, @"header-footer.html");
            }
            else if (File.Exists(Path.Combine(_templateDirectory, @"header-footer.htm")))
            {
                _headerFooterFilePath = Path.Combine(_templateDirectory, @"header-footer.htm");
            }

            PageComponent _pageHeaderFooter = new PageComponent();
            _pageHeaderFooter.SetDirectory(_templateDirectory);
            _pageHeaderFooter.SetHtmlPath(_headerFooterFilePath);
            _pageHeaderFooter.SetScriptPath(Path.Combine(_templateDirectory, @"header-footer.js"));

            this.AddPageFooter(_pageHeaderFooter);
        }

    }
}
