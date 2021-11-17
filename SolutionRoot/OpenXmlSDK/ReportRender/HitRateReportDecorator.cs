using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreReport;
using System.Reflection;
using System.Drawing;
using System.Diagnostics;
using System.Net;
using OpenXmlSDK.ReportEntity;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CoreReport.OpenXmlSDK
{
    public class HitRateReportDecorator: OpenXmlSDKDecorator
    {

        public HitRateReportDecorator() : base()
        {
        }
        public HitRateReportDecorator(OpenXmlSDKReportEntity _reportEntity, string _filename = "") : base(_reportEntity, _filename = "")
        {
        }

        protected override SpreadsheetDocument RenderBodyAndMergeToTemplate(SpreadsheetDocument _spreadsheetDocument)
        {
            DataSet _dataSet = this.reportEntity.GetDataSet();
            IDictionary<string, object> _dataSetObj = this.reportEntity.GetDataSetObj();
            //ExcelWorksheet worksheet = _excelPackage.Workbook.Worksheets[0];
            //ExcelWorksheet activeSheet = _excelPackage.Workbook.Worksheets.FirstOrDefault(f => f.View.TabSelected);

            Sheet activeSheet = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Sheet1");

            DataView dataView = _dataSet.Tables["GeneralView"].DefaultView;
            dataView.Sort = "OfficeName";
            DataTable sortedDataTable = dataView.ToTable();

            string officeName = "";
            var officeNameList = (from DataRow dr in sortedDataTable.Rows
                                  select (string)dr["OfficeName"]).Distinct();

            foreach (string _officeName in officeNameList)
            {
                //DataRow distinctOfficeName = sortedDataTable.Select($"OfficeName = '{_officeName}'");
                var distinctOfficeName = from DataRow row in sortedDataTable.Rows
                                         where row["OfficeName"] == _officeName
                                         select row;

                DataRow officeRow = null;
                foreach (DataRow ofDepartmentRow in distinctOfficeName)
                {
                    officeRow = sortedDataTable.NewRow();
                    officeRow.ItemArray = ofDepartmentRow.ItemArray.Clone() as object[];
                    this.MergeDataRow(activeSheet, "T1B", ofDepartmentRow);
                }
                this.MergeDataRow(activeSheet, "T1F", officeRow);
                this.PrintSectionSeparateLine(activeSheet, "T1B", "T1F");
            }

            _spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.ForceFullCalculation = true;
            _spreadsheetDocument.WorkbookPart.Workbook.CalculationProperties.FullCalculationOnLoad = true;

            return _spreadsheetDocument;
        }

    }
}