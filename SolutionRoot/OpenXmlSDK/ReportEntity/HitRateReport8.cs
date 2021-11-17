
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

namespace OpenXmlSDK.ReportEntity
{
    public class HitRateReport8 : OpenXmlSDKReportEntity
    {
        public HitRateReport8(DataSet _dataSet)
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateReport6");
            //this.dataSet = _dataSet;
            this.dataSet = _dataSet;
        }

        public HitRateReport8(IDictionary<string, object> _dataSetObj)
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateReport6");
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
            _templateDirectory = Path.Combine(this.templateBaseDirectory, @"HitRateReport6");

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
            _dataGrid.SetDynamicRange(new ExcelDataSection("T1B", "17:19", "20:20"));
            this.AddDataGrid(_dataGrid);

            _dataGrid = new ExcelDataGrid("Sheet1");
            _dataGrid.SetDynamicRange(new ExcelDataSection("", "21:21", "22:22"));
            this.AddDataGrid(_dataGrid);

            // define the sheet - rows to column at left
        }

        public override void InitializateHeaderFooter()
        {
            string _templateDirectory = this.templateReportFileDirectory;

            string _headerFilePath = string.Empty;
            string _footerFilePath = string.Empty;
            string _headerFooterFilePath = string.Empty;

            // define excel header, footer
            // Page Setup, Header/Footer

            // Custom Header
            // Left section, Center section, Right section

            // Custom Footer
            // Left section, Center section, Right section
        }

    }
}
