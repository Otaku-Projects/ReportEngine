
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
using OfficeOpenXml;

namespace EPPlus5Report.ReportEntity
{
    public class HitRateReport5 : BaseReportEntity
    {
        public HitRateReport5(DataSet _dataSet)
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateReport5");
            //this.dataSet = _dataSet;
            this.dataSet = _dataSet;
        }

        public HitRateReport5(IDictionary<string, object> _dataSetObj)
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateReport5");
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
            _templateDirectory = Path.Combine(this.templateBaseDirectory, @"HitRateReport5");

            this.templateReportFileDirectory = _templateDirectory;
            this.SetXlsxTemplateFileName("HitRateReport5Template.xlsx");
        }

        public override void InitializateDataGrid()
        {
            // define the page setup - header
            // define the page setup - footer

            // define the sheet - rows to repeat at top
            ExcelDataGrid _dataGrid = null;
            _dataGrid = new ExcelDataGrid("Sheet1");
            _dataGrid.SetHeaderRange(new ExcelDataGridSection("", "", ""));
            _dataGrid.SetBodyRange(new ExcelDataGridSection("T1B", "17:19", "20:20"));
            _dataGrid.SetFooterRange(new ExcelDataGridSection("T1F", "21:21", "22:22"));
            this.AddDataGrid(_dataGrid);

            _dataGrid = new ExcelDataGrid("Sheet1");
            _dataGrid.SetHeaderRange(new ExcelDataGridSection("", "", ""));
            _dataGrid.SetBodyRange(new ExcelDataGridSection("", "21:21", "22:22"));
            _dataGrid.SetFooterRange(new ExcelDataGridSection("", "", ""));
            //this.AddDataGrid(_dataGrid);

            // define the sheet - rows to column at left
        }

        public override void InitializateHeaderFooter()
        {
            string _templateDirectory = this.templateReportFileDirectory;

            string _headerFilePath = string.Empty;
            string _footerFilePath = string.Empty;
            string _headerFooterFilePath = string.Empty;

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

            if (File.Exists(Path.Combine(_templateDirectory, @"header-footer.html")))
            {
                _headerFooterFilePath = Path.Combine(_templateDirectory, @"header-footer.html");
            }
            else if (File.Exists(Path.Combine(_templateDirectory, @"header-footer.htm")))
            {
                _headerFooterFilePath = Path.Combine(_templateDirectory, @"header-footer.htm");
            }

            PageComponent _pageHeader = new PageComponent();
            _pageHeader.SetDirectory(_templateDirectory);
            _pageHeader.SetHtmlPath(_headerFilePath);
            _pageHeader.SetScriptPath(Path.Combine(_templateDirectory, @"header.js"));

            PageComponent _pageFooter = new PageComponent();
            _pageFooter.SetDirectory(_templateDirectory);
            _pageFooter.SetHtmlPath(_footerFilePath);
            _pageFooter.SetScriptPath(Path.Combine(_templateDirectory, @"footer.js"));

            PageComponent _pageHeaderFooter = new PageComponent();
            _pageHeaderFooter.SetDirectory(_templateDirectory);
            _pageHeaderFooter.SetHtmlPath(_headerFooterFilePath);
            _pageHeaderFooter.SetScriptPath(Path.Combine(_templateDirectory, @"header-footer.js"));

            this.AddPageFooter(_pageHeaderFooter);
        }

    }
}
