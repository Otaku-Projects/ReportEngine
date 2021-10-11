using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreReport;
using EPPlus5Report.ReportEntity;
using System.Reflection;
using OfficeOpenXml.Style;
using System.Drawing;
using QRCoder;
using System.Diagnostics;

namespace CoreReport.EPPlus5Report
{
    public class HitRateReportDecorator: EPPlus5Decorator
    {

        public HitRateReportDecorator() : base()
        {
        }
        public HitRateReportDecorator(BaseReportEntity _reportEntity, string _filename = "") : base(_reportEntity, _filename = "")
        {
        }

        protected virtual ExcelPackage RenderBodyAndMergeToTemplate(ExcelPackage _excelPackage)
        {
            DataSet _dataSet = this.reportEntity.GetDataSet();
            IDictionary<string, object> _dataSetObj = this.reportEntity.GetDataSetObj();
            //ExcelWorksheet worksheet = _excelPackage.Workbook.Worksheets[0];
            //ExcelWorksheet activeSheet = _excelPackage.Workbook.Worksheets.FirstOrDefault(f => f.View.TabSelected);

            ExcelWorksheet activeSheet = _excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Sheet1");

            //foreach (dynamic tupleList in (List<object>)_dataSetObj["GeneralView"])
            //{
            //    //this.MergeDataRow(activeSheet, "G", tuple);
            //    //this.MergeDataRows(activeSheet, "G", tupleList);
            //    Console.WriteLine(tupleList);
            //}
            //this.MergeDataRows(activeSheet, "T1B", (List<Object>)_dataSetObj["GeneralView"]);
            // refresh and re-calculate the formula result, e.g. sum(...)

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
                foreach(DataRow ofDepartmentRow in distinctOfficeName)
                {
                    officeRow = sortedDataTable.NewRow();
                    officeRow.ItemArray = ofDepartmentRow.ItemArray.Clone() as object[];
                    this.MergeDataRow(activeSheet, "T1B", ofDepartmentRow);
                }
                this.MergeDataRow(activeSheet, "T1F", officeRow);
                this.PrintSectionSeparateLine(activeSheet, "T1B", "T1F");
            }

            //foreach (DataRow dRow in sortedDataTable.Rows)
            //{
            //    if (!string.IsNullOrEmpty(officeName) && officeName != dRow["OfficeName"].ToString())
            //    {
            //        this.MergeDataRow(activeSheet, "T1F", dRow);
            //        this.PrintSectionSeparateLine(activeSheet, "T1B", "T1F");
            //        break;
            //    }
            //        this.MergeDataRow(activeSheet, "T1B", dRow);
            //    officeName = dRow["OfficeName"].ToString();
            //}

            activeSheet.Calculate();

            return _excelPackage;
        }
        protected override void CustomPostMergeCellExpression(ExcelWorksheet _worksheet, string cellAddress, IDictionary<string, Object> _tuple)
        {
            ExcelRange _cell = _worksheet.Cells[cellAddress];

            // 1.1 get cell value
            string cellVal = _cell.GetValue<string>();
            // 1.2 skip if cell value is empty or null
            if (string.IsNullOrEmpty(cellVal)) return;

            // 2.1 if cell value is Image
            string matchExpression = string.Empty;
            string mergedValue = cellVal;
            int pixelTop = 88;
            int pixelLeft = 129;
            int Height = 150;
            int Width = 112;
            Guid obj = Guid.NewGuid();
            string imgID = obj.ToString();
            if (cellVal.IndexOf("{{Image}}") > -1)
            {
                Image img = Image.FromFile(@"D:\Documents\ReportEngine\SolutionRoot\EPPlus5\ReportTemplate\HitRateReport5\man-4367499_480.png");
                ExcelPicture pic = _worksheet.Drawings.AddPicture(imgID, img);
                pic.SetPosition(_cell.Start.Row-1, 0, _cell.Start.Column-1, 0);
                //pic.SetPosition(PixelTop, PixelLeft);  
                pic.SetSize(Height, Width);
                //pic.SetSize(40);
                //pic.EditAs = eEditAs.TwoCell;
                pic.ChangeCellAnchor(eEditAs.TwoCell);

                _cell.Clear();
            }

            if (cellVal.IndexOf("{{QRCode}}") > -1)
            {
                string qrCodeText = string.Empty;
                //PropertyInfo prop = _tuple.GetType().GetProperty("Department");
                //qrCodeText = prop.GetValue(_tuple, null).ToString();

                qrCodeText = _tuple["Department"].ToString();

                QRCodeGenerator qrGenerator = new QRCodeGenerator();
                QRCodeData qrCodeData = qrGenerator.CreateQrCode(qrCodeText, QRCodeGenerator.ECCLevel.Q);
                QRCode qrCode = new QRCode(qrCodeData);
                Bitmap qrCodeImage = qrCode.GetGraphic(20);

                ExcelPicture pic2 = _worksheet.Drawings.AddPicture(imgID+"2", qrCodeImage);
                pic2.SetPosition(_cell.Start.Row - 1, 0, _cell.Start.Column - 1, 0);
                pic2.SetSize(96, 96);
                pic2.ChangeCellAnchor(eEditAs.TwoCell);

                _cell.Clear();
            }
        }
        protected virtual ExcelPackage RenderHeaderAndMergeToTemplate(ExcelPackage _excelPackage)
        {
            return _excelPackage;
        }
        protected virtual ExcelPackage RenderFooterAndMergeToTemplate(ExcelPackage _excelPackage)
        {

            return _excelPackage;
        }
        public override ExcelPackage RenderDataAndMergeToTemplate(ExcelPackage _excelPackage)
        {
            //ExcelPackage _excelPackage = this.GetXlsxTemplateInstance();

            string excelRange = string.Empty;
            string _indicator = string.Empty;
            //ExcelWorksheet templateSheet = _excelPackage.Workbook.Worksheets.FirstOrDefault(f => f.View.TabSelected);
            //ExcelWorksheet templateSheet = _excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Template");
            ExcelWorksheet templateSheet = _excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Template");

            // update indicator for all dataGridList
            List<ExcelDataGrid> _excelDataGridList = this.reportEntity.GetDataGrid();
            foreach (ExcelDataGrid excelDataGrid in _excelDataGridList)
            {
                // update header indicator
                excelRange = excelDataGrid.GetHeaderRange().GetTemplateRange();
                if (string.IsNullOrEmpty(excelRange)) continue;

                _indicator = templateSheet.Cells["A" + excelDataGrid.GetHeaderRange().TemplateFromRow].GetValue<string>();
                excelDataGrid.GetHeaderRange().Indicator = _indicator;

                // update body indicator
                excelRange = excelDataGrid.GetBodyRange().GetTemplateRange();
                if (string.IsNullOrEmpty(excelRange)) continue;

                _indicator = templateSheet.Cells["A" + excelDataGrid.GetBodyRange().TemplateFromRow].GetValue<string>();
                excelDataGrid.GetBodyRange().Indicator = _indicator;
                 
                // update footer indicator
                excelRange = excelDataGrid.GetFooterRange().GetTemplateRange();
                if (string.IsNullOrEmpty(excelRange)) continue;

                _indicator = templateSheet.Cells["A" + excelDataGrid.GetFooterRange().TemplateFromRow].GetValue<string>();
                excelDataGrid.GetFooterRange().Indicator = _indicator;
            }

            this.RenderBodyAndMergeToTemplate(_excelPackage);
            this.RenderHeaderAndMergeToTemplate(_excelPackage);
            this.RenderFooterAndMergeToTemplate(_excelPackage);

            List<ExpandoObject> tupleExpandoObjectList = new List<ExpandoObject>();
            ExpandoObject expandoObject = new ExpandoObject();

            //var a = _dataSetObj["GeneralView"].Select(i=>i.Value).;

            /*
            foreach (KeyValuePair<string, Object> _dataView in _dataSetObj)
            {
                string tableName = _dataView.Key;
                if (tableName.ToLower().IndexOf("view") == -1) continue;

                tupleExpandoObjectList = new List<ExpandoObject>();
                foreach (object tuple in (List<object>)_dataView.Value)
                {
                    expandoObject = new ExpandoObject();
                    foreach (var property in tuple.GetType().GetProperties())
                    {
                        ((IDictionary<string, object>)expandoObject).Add(property.Name, property.GetValue(tuple));
                    }
                    tupleExpandoObjectList.Add(expandoObject);
                }

                var sheet = _excelPackage.Workbook.Worksheets.Add(tableName);
                sheet.Cells["A1"].LoadFromDictionaries(tupleExpandoObjectList, c =>
                {
                    // Print headers using the property names
                    c.PrintHeaders = true;
                    // insert a space before each capital letter in the header
                    c.HeaderParsingType = HeaderParsingTypes.CamelCaseToSpace;
                    // when TableStyle is not TableStyles.None the data will be loaded into a table with the 
                    // selected style.
                    c.TableStyle = TableStyles.Medium6;

                    // SetKeys takes a params string[] - you can add any number of
                    // keys as arguments to this function.
                    //c.SetKeys("name", "price");
                });
            }
            */

            return _excelPackage;
        }
        public override void RenderTemplateAndSaveAsPdf(string _fileName = "")
        {
            if (string.IsNullOrEmpty(_fileName))
            {
                Guid obj = Guid.NewGuid();
                _fileName = obj.ToString();
            }
            string xlsxFilePath = Path.Combine(
                this.epplusReportRenderFolder,
                _fileName + ".xlsx");
            string pdfFilePath = Path.Combine(
                this.epplusReportRenderFolder,
                _fileName + ".pdf");

            try
            {
                using (ExcelPackage _excelPackage = this.StartRenderDataAndMergeToTemplate())
                {
                    this.RenderDataAndMergeToTemplate(_excelPackage);
                    //this.RemoveTemplateRowsForPdf(_excelPackage);
                    this.RemoveTemplateRows(_excelPackage);

                    // SaveAs Method2
                    //Instead of converting to bytes, you could also use FileInfo
                    FileInfo fi = new FileInfo(xlsxFilePath);
                    _excelPackage.SaveAs(fi);
                }

                Assembly[] aList = AppDomain.CurrentDomain.GetAssemblies();
                IEnumerable<Assembly> assemblies = AppDomain.CurrentDomain.GetAssemblies().Where(a => a.FullName.Contains("OfficeToPDF"));

                string exePath = Path.Combine(
                    Directory.GetCurrentDirectory(),
                    "OfficeToPDF-1.9.0.2.exe");
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = exePath;
                //startInfo.Arguments = $"/hidden /readonly /excel_active_sheet {xlsxFilePath} {pdfFilePath}";
                startInfo.Arguments = $"/hidden /readonly /excel_worksheet 1 {xlsxFilePath} {pdfFilePath}";
                // convert xlsx to pdf
                using (Process exeProcess = Process.Start(startInfo))
                {

                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public override void RenderTemplateAndSaveAsXlsx(string _fileName = "")
        {
            if (string.IsNullOrEmpty(_fileName))
            {
                Guid obj = Guid.NewGuid();
                _fileName = obj.ToString();
            }
            string filePath = Path.Combine(
                this.epplusReportRenderFolder,
                _fileName+".xlsx");
            //this.RenderDataAndMergeToTemplate();

            try
            {
                using (ExcelPackage _excelPackage = this.StartRenderDataAndMergeToTemplate())
                {
                    this.RenderDataAndMergeToTemplate(_excelPackage);
                    //this.RemoveTemplateRowsForXlsx(_excelPackage);
                    this.RemoveTemplateRows(_excelPackage);
                    /*
                    // SaveAs Method1
                    //convert the excel package to a byte array
                    byte[] bin = _excelPackage.GetAsByteArray();
                    //the path of the file
                    //write the file to the disk
                    File.WriteAllBytes(filePath, bin);
                    */

                    // SaveAs Method2
                    //Instead of converting to bytes, you could also use FileInfo
                    FileInfo fi = new FileInfo(filePath);
                    _excelPackage.SaveAs(fi);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
    }
}