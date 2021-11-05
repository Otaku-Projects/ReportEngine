
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

namespace ITextGroupNV.ReportEntity
{
    public class PocFileList : ITextReportEntity
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
            // define the page setup - header
            // define the page setup - footer

            // define the sheet - rows to repeat at top
            ExcelDataGrid _dataGrid = null;
            _dataGrid = new ExcelDataGrid("Sheet1");
            _dataGrid.SetHeaderRange(new ExcelDataGridSection("", "", ""));
            _dataGrid.SetBodyRange(new ExcelDataGridSection("T1B", "17:18", "19:19"));
            _dataGrid.SetFooterRange(new ExcelDataGridSection("", "", ""));
            this.AddDataGrid(_dataGrid);

            //_dataGrid = new ExcelDataGrid("Sheet1");
            //_dataGrid.SetHeaderRange(new ExcelDataGridSection("", "", ""));
            //_dataGrid.SetBodyRange(new ExcelDataGridSection("", "21:21", "22:22"));
            //_dataGrid.SetFooterRange(new ExcelDataGridSection("", "", ""));

            //this.AddDataGrid(_dataGrid);

            // define the sheet - rows to column at left
        }

        public override void InitializateHeaderFooter()
        {
        }
    }
}
