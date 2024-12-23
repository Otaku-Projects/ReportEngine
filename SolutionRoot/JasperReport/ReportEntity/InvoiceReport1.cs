﻿using System;
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
    public class InvoiceReport1 : JasperReportEntity
    {
        public InvoiceReport1(DataSet _dataSet) { }

        public InvoiceReport1(IDictionary<string, object> _dataSetObj)
        {
            Console.WriteLine("Said \"Hello World!\" from InvoiceReport1");
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

            string _headerFilePath = string.Empty;
            string _footerFilePath = string.Empty;

            if (File.Exists(Path.Combine(_templateDirectory, @"header.html")))
            {
                _headerFilePath = Path.Combine(_templateDirectory, @"header.html");
            }
            else if (File.Exists(Path.Combine(_templateDirectory, @"header.htm")))
            {
                _headerFilePath = Path.Combine(_templateDirectory, @"header.htm");
            }
            if (File.Exists(Path.Combine(_templateDirectory, @"footer.html")))
            {
                _footerFilePath = Path.Combine(_templateDirectory, @"footer.html");
            }
            else if (File.Exists(Path.Combine(_templateDirectory, @"footer.htm")))
            {
                _footerFilePath = Path.Combine(_templateDirectory, @"footer.htm");
            }

            PageComponent _pageHeader = new PageComponent();
            _pageHeader.SetDirectory(_templateDirectory);
            _pageHeader.SetHtmlPath(_headerFilePath);
            _pageHeader.SetScriptPath(Path.Combine(_templateDirectory, @"header.js"));

            PageComponent _pageFooter = new PageComponent();
            _pageFooter.SetDirectory(_templateDirectory);
            _pageFooter.SetHtmlPath(_footerFilePath);
            _pageFooter.SetScriptPath(Path.Combine(_templateDirectory, @"footer.js"));

            this.AddPageHeader(_pageHeader);
            this.AddPageFooter(_pageFooter);
        }

    }
}
