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
    public class HitRateReport : BaseReportEntity
    {
        public HitRateReport(DataSet _dataSet) { }

        public HitRateReport(IDictionary<string, object> _dataSetObj)
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateReport");
            //this.dataSet = _dataSet;
            this.dataSetObj = _dataSetObj;
        }
        public override void InitializateMetaData() {
            this.headerFooterOption = HeaderFooterOptions.HeaderFooterInSingleFile;
        }
        public override void InitializateMainContent()
        {
            string _templateDirectory = string.Empty;
            string _templateFileLocation = string.Empty;
            string _templateScriptLocation = string.Empty;
            _templateDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"HitRateReport");
            _templateDirectory = Path.Combine(this.templateBaseDirectory, @"HitRateReport");

            if (File.Exists(Path.Combine(_templateDirectory, @"index.html")))
            {
                _templateFileLocation = Path.Combine(_templateDirectory, @"index.html");
            }
            else if (File.Exists(Path.Combine(_templateDirectory, @"index.htm")))
            {
                _templateFileLocation = Path.Combine(_templateDirectory, @"index.htm");
            }

            if (File.Exists(Path.Combine(_templateDirectory, @"helper.js")))
            {
                _templateScriptLocation = Path.Combine(_templateDirectory, @"helper.js");
            }

            this.templateReportFileDirectory = _templateDirectory;
            this.templateReportFileLocation = _templateFileLocation;

            PageComponent _pageMainContent = new PageComponent();
            _pageMainContent.SetDirectory(this.templateReportFileDirectory);
            _pageMainContent.SetHtmlPath(this.templateReportFileLocation);
            _pageMainContent.SetScriptPath(_templateScriptLocation);

            this.AddPageContent(_pageMainContent);
        }
        public override void InitializateHeaderFooter()
        {
            string _templateDirectory = this.templateReportFileDirectory;

            string _templateHeaderLocation = string.Empty;
            string _templateFooterLocation = string.Empty;

            /*
            if (this.headerFooterOption == HeaderFooterOptions.Header)
            {
                if (File.Exists(Path.Combine(_templateDirectory, @"header.html")))
                {
                    _templateHeaderLocation = Path.Combine(_templateDirectory, @"header.html");
                }
                else if (File.Exists(Path.Combine(_templateDirectory, @"header.htm")))
                {
                    _templateHeaderLocation = Path.Combine(_templateDirectory, @"header.htm");
                }
            }
            else if (this.headerFooterOption == HeaderFooterOptions.Footer)
            {
                if (File.Exists(Path.Combine(_templateDirectory, @"footer.html")))
                {
                    _templateFooterLocation = Path.Combine(_templateDirectory, @"footer.html");
                }
                else if (File.Exists(Path.Combine(_templateDirectory, @"footer.htm")))
                {
                    _templateFooterLocation = Path.Combine(_templateDirectory, @"footer.htm");
                }
            }
            else if (this.headerFooterOption == HeaderFooterOptions.HeaderFooterInSingleFile)
            {
                if (File.Exists(Path.Combine(_templateDirectory, @"header-footer.html")))
                {
                    _templateHeaderLocation = Path.Combine(_templateDirectory, @"header-footer.html");
                }
                else if (File.Exists(Path.Combine(_templateDirectory, @"header-footer.htm")))
                {
                    _templateHeaderLocation = Path.Combine(_templateDirectory, @"header-footer.htm");
                }
            }
            */

            if (File.Exists(Path.Combine(_templateDirectory, @"header.html")))
            {
                _templateHeaderLocation = Path.Combine(_templateDirectory, @"header.html");
            }
            else if (File.Exists(Path.Combine(_templateDirectory, @"header.htm")))
            {
                _templateHeaderLocation = Path.Combine(_templateDirectory, @"header.htm");
            }
            if (File.Exists(Path.Combine(_templateDirectory, @"footer.html")))
            {
                _templateFooterLocation = Path.Combine(_templateDirectory, @"footer.html");
            }
            else if (File.Exists(Path.Combine(_templateDirectory, @"footer.htm")))
            {
                _templateFooterLocation = Path.Combine(_templateDirectory, @"footer.htm");
            }

            if (File.Exists(Path.Combine(_templateDirectory, @"header-footer.html")))
            {
                _templateHeaderLocation = Path.Combine(_templateDirectory, @"header-footer.html");
            }
            else if (File.Exists(Path.Combine(_templateDirectory, @"header-footer.htm")))
            {
                _templateHeaderLocation = Path.Combine(_templateDirectory, @"header-footer.htm");
            }

            this.templateReportHeaderLocation = _templateHeaderLocation;
            this.templateReportFooterLocation = _templateFooterLocation;

            PageComponent _pageHeader = new PageComponent();
            _pageHeader.SetDirectory(this.templateReportFileDirectory);
            _pageHeader.SetHtmlPath(this.templateReportHeaderLocation);
            _pageHeader.SetScriptPath(Path.Combine(_templateDirectory, @"header.js"));

            PageComponent _pageFooter = new PageComponent();
            _pageFooter.SetDirectory(this.templateReportFileDirectory);
            _pageFooter.SetHtmlPath(this.templateReportFooterLocation);
            _pageFooter.SetScriptPath(Path.Combine(_templateDirectory, @"footer.js"));

            PageComponent _pageHeaderFooter = new PageComponent();
            _pageHeaderFooter.SetDirectory(this.templateReportFileDirectory);
            _pageHeaderFooter.SetHtmlPath(this.templateReportHeaderLocation);
            _pageHeaderFooter.SetScriptPath(Path.Combine(_templateDirectory, @"header-footer.js"));

            //this.AddPageHeader(_pageHeader);
            //this.AddPageFooter(_pageFooter);
            this.AddPageHeaderFooter(_pageHeaderFooter);

            //if (this.headerFooterOption == HeaderFooterOptions.HeaderFooterInSingleFile)
            //{
            //}
        }

    }
}
