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
using System.Collections;
using System.Diagnostics;
using OpenXmlSDK.ReportEntity;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;

namespace CoreReport.OpenXmlSDK
{
    public class OpenXmlSDKDecorator : VisualizationDecorator
    {
        protected string createdBy;
        protected DateTime createdDate;
        protected DateTime printedDate;
        protected string filename;

        protected OpenXmlSDKReportEntity reportEntity;

        protected DataSet dataSet;
        protected IDictionary<string, object> dataSetObj;
        protected string openXmlSDKRenderFolder;

        protected SpreadsheetDocument sourceSpreadsheetDocument; // the template spreadsheet
        protected SpreadsheetDocument targetSpreadsheetDocument; // the copy destination spreadsheet

        public List<string> _fonts;

        protected string report_instance_dir;
        protected string report_template_dir;
        protected string fonts_folder;

        public OpenXmlSDKDecorator()
        {
            this.Initialize();
        }
        public OpenXmlSDKDecorator(OpenXmlSDKReportEntity _reportEntity, string _filename = "")
        {
            if (string.IsNullOrEmpty(_filename))
            {
                Guid obj = Guid.NewGuid();
                _filename = obj.ToString();
            }

            this.dataSet = _reportEntity.GetDataSet();
            this.dataSetObj = _reportEntity.GetDataSetObj();

            this.filename = _filename;

            this.reportEntity = _reportEntity;

            this.createdBy = "CoreSystem";
            this.createdDate = new DateTime();

            this.report_instance_dir = string.Empty;
            this.report_template_dir = string.Empty;

            this.Initialize();
        }

        public void Initialize()
        {
            this._fonts = new List<string>();
            this._fonts.Add("NotoSansCJKjp-Regular.otf");
            this._fonts.Add("NotoSansCJKkr-Regular.otf");
            this._fonts.Add("NotoSansCJKsc-Regular.otf");
            this._fonts.Add("NotoSansCJKtc-Regular.otf");

            this.report_instance_dir = this.reportEntity.GetTemplateFileDirectory();
            this.report_template_dir = System.IO.Directory.GetParent(this.report_instance_dir).ToString();
            this.fonts_folder = Path.Combine(report_template_dir, "General", "fonts");

            this.openXmlSDKRenderFolder = this.tempRenderFolder;
        }

        public void RefreshPrintDate()
        {
            this.printedDate = new DateTime();
        }

        public override void Display()
        {
            throw new NotImplementedException();
        }

        protected SpreadsheetDocument CreateXlsxTemplateInstance()
        {
            string _xlsxTemplateFilePath = this.reportEntity.GetXlsxTemplateFilePath();
            if (!File.Exists(_xlsxTemplateFilePath))
            {
                throw new FileNotFoundException($"Excel template not found at {_xlsxTemplateFilePath}");
            }
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_xlsxTemplateFilePath, false);
            return spreadsheetDocument;
        }
        public SpreadsheetDocument GetXlsxTemplateInstance()
        {
            return this.CreateXlsxTemplateInstance();
        }

        protected void CreatePdfTemplatePropertiesInstance()
        {
            string _pdfTemplateFilePath = this.reportEntity.GetPdfTemplateFilePath();
            if (!File.Exists(_pdfTemplateFilePath))
            {
                throw new FileNotFoundException($"PDF template (HTML file) not found at {_pdfTemplateFilePath}");
            }
        }
        public void GetPdfTemplatePropertiesInstance()
        {
            this.CreatePdfTemplatePropertiesInstance();
        }
        public virtual void RenderTemplateAndSaveAsXlsx(string _fileName = "")
        {
            // you should not call into here, please inherit the decorator and override this function
            throw new NotImplementedException();

            if (string.IsNullOrEmpty(_fileName))
            {
                Guid obj = Guid.NewGuid();
                _fileName = obj.ToString();
            }
            string filePath = Path.Combine(
                this.openXmlSDKRenderFolder,
                _fileName + ".xlsx");
            //this.RenderDataAndMergeToTemplate();

            try
            {
                using (SpreadsheetDocument _spreadsheetDocument = this.StartRenderDataAndMergeToTemplate())
                {
                    this.RenderDataAndMergeToTemplate(_spreadsheetDocument);
                    //this.RemoveTemplateRowsForXlsx(_excelPackage);
                    this.RemoveTemplateRows(_spreadsheetDocument);
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
                    //FileInfo fi = new FileInfo(filePath);
                    //_excelPackage.SaveAs(fi);

                    _spreadsheetDocument.SaveAs(filePath);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        public virtual FileStream RenderTemplateAndSaveAsPdf(string _fileName = "")
        {
            if (string.IsNullOrEmpty(_fileName))
            {
                Guid obj = Guid.NewGuid();
                _fileName = obj.ToString();
            }

            string xlsxFilePath = Path.Combine(
                this.openXmlSDKRenderFolder,
                _fileName + ".xlsx");
            string pdfFilePath = Path.Combine(
                this.openXmlSDKRenderFolder,
                _fileName + ".pdf");

            string _report_instance_dir = this.report_instance_dir;
            string _report_template_dir = this.report_template_dir;
            string _fonts_folder = this.fonts_folder;

            try
            {
                IDictionary<string, object> _dataSetObj = this.reportEntity.GetDataSetObj();
                var newDict = new Dictionary<string, object>(_dataSetObj);

                // Convert IDictionary/Dictionary<string, object> To Anonymous Object
                var _expandoObject = new ExpandoObject();
                var eoColl = (ICollection<KeyValuePair<string, object>>)_expandoObject;
                foreach (var kvp in _dataSetObj)
                {
                    eoColl.Add(kvp);
                }
                dynamic eoDynamic = eoColl;

                //eoDynamic.ReportTemplate_Root = Path.Combine("file:///", System.IO.Directory.GetParent(_report_instance_dir).ToString());
                //eoDynamic.ReportInstance_Folder = Path.Combine("file:///", _report_instance_dir);
                eoDynamic.meta_PrintDateTime = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                eoDynamic.meta_DateTime_yyyy_mm_dd = DateTime.Now.ToString("yyyy-MM-dd");
                eoDynamic.meta_DateTime_yyyy_mm_dd_hh_mm_ss = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

                string htmlRenderResult = string.Empty;
                //dynamic eo = _dataSetObj.Aggregate(new ExpandoObject() as IDictionary<string, Object>,
                //            (a, p) => { a.Add(p.Key, p.Value); return a; });

                # region read report template content
                string _pdfTemplateFilePath = this.reportEntity.GetPdfTemplateFilePath();
                //FileStream htmlTemplateStream = File.Open(_pdfTemplateFilePath, FileMode.Open, FileAccess.Read);

                using var fs = new FileStream(_pdfTemplateFilePath, FileMode.Open, FileAccess.Read);
                using var sr = new StreamReader(fs, Encoding.UTF8);
                string htmlTemplateSource = sr.ReadToEnd();
                #endregion

                string exePath = Path.Combine(
                    Directory.GetCurrentDirectory(),
                    "OfficeToPDF-1.9.0.2.exe");
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = exePath;
                //startInfo.Arguments = $"/hidden /readonly /excel_active_sheet {xlsxFilePath} {pdfFilePath}";
                startInfo.Arguments = $"/hidden /readonly /excel_worksheet 1 {xlsxFilePath} {pdfFilePath}";
                // convert xlsx to pdf
                Process exeProcess = Process.Start(startInfo);

                //Set a time-out value.
                int timeOut = 15000;

                // wait until it's done or time out.
                exeProcess.WaitForExit(timeOut);

                // Alternatively, if it's an application with a UI that you are waiting to enter into a message loop
                //exeProcess.WaitForInputIdle();

                //Check to see if the process is still running.
                if (exeProcess.HasExited == false)
                    //Process is still running.
                    //Test to see if the process is hung up.
                    if (exeProcess.Responding)
                        //Process was responding; close the main window.
                        exeProcess.CloseMainWindow();
                    else
                        //Process was not responding; force the process to close.
                        exeProcess.Kill();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            FileStream _fileStream = new FileStream(pdfFilePath, FileMode.Open, FileAccess.Read, FileShare.None);
            return _fileStream;
        }

        protected virtual SpreadsheetDocument StartRenderDataAndMergeToTemplate()
        {
            SpreadsheetDocument _spreadsheetDocument = this.GetXlsxTemplateInstance();

            // if sheet1 exists, delete it
            //Sheet sheet1 = _excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Sheet1");
            Sheet sheet1 = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Sheet1");
            //Worksheet sheet1 = _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheet>().First(s => sheetName.Equals(s.NamespaceUri)).Id;

            if (sheet1 != null)
            {
                sheet1.Remove();
            }

            // clone template sheet to sheet1
            Sheet sheet2 = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Template");
            Sheet sheetTemplate = (Sheet)sheet2.Clone();
            sheetTemplate.Name = "Sheet1";
            Sheets sheets = _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            sheets.Append(sheetTemplate);

            //_excelPackage.Workbook.Worksheets["Template"].View.TabSelected = false;

            // backup the data grid 
            this.reportEntity.BackupDataGridSetting();

            // move sheet1 to start
            //_excelPackage.Workbook.Worksheets.MoveToStart("Sheet1");

            // select sheet1 as the default sheet
            //sheet1.View.SetTabSelected();

            this.sourceSpreadsheetDocument = _spreadsheetDocument;
            return _spreadsheetDocument;
        }

        protected virtual void MergeDataRow(Sheet _worksheet, string _indicator, DataRow _dataRow)
        {
            var expObj = new ExpandoObject() as IDictionary<string, Object>;

            foreach (DataColumn dc in _dataRow.Table.Columns)
            {
                expObj.Add(dc.ColumnName, _dataRow[dc]);
            }

            this.MergeDataRow(_worksheet, _indicator, expObj);
        }
        public int ExcelColumnNameToNumber(string columnName)
        {
            if (string.IsNullOrEmpty(columnName)) throw new ArgumentNullException("columnName");

            columnName = columnName.ToUpperInvariant();

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }
        public string ExcelGetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }

        protected virtual void MergeDataRow(Sheet _worksheet, string _indicator, IDictionary<string, Object> _tuple)
        {
            List<ExcelDataGrid> excelDataGridList = this.reportEntity.GetDataGrid();

            WorkbookPart workbookPart = this.sourceSpreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.GetPartById(_worksheet.Id) as WorksheetPart;
            Console.WriteLine(worksheetPart.Worksheet.SheetDimension.Reference);

            string usedDimension = worksheetPart.Worksheet.SheetDimension.Reference;
            usedDimension = usedDimension.ToUpper();
            string numericPart = new String(usedDimension.Where(Char.IsDigit).ToArray());
            string alphabetsPart = Regex.Replace(usedDimension, @"[\d-]", string.Empty);

            int rowCount = Convert.ToInt32(numericPart);
            int columnCount = this.ExcelColumnNameToNumber(alphabetsPart);

            string fromRange = string.Empty;
            string destinationRange = string.Empty;
            string newAppendToRange = string.Empty;

            string templateStartColLetter = string.Empty;
            int templateStartRowIndex = -1;
            string templateEndColLetter = string.Empty;
            int templateEndRowIndex = -1;

            string appendStartColLetter = string.Empty;
            int appendStartRowIndex = -1;
            string appendEndColLetter = string.Empty;
            int appendEndRowIndex = -1;

            // 1.1 find the indicator data grid
            ExcelDataSection affectSection = null;
            foreach (ExcelDataGrid dataGrid in excelDataGridList)
            {
                if (dataGrid.GetDynamicRange().Indicator == _indicator)
                {
                    affectSection = dataGrid.GetDynamicRange();
                    break;
                }
            }

            if (affectSection == null)
            {
                // the _indicator was not found in excel template
                return;
            }

            // 1.2 get data grid range
            //startColLetter = OfficeOpenXml.ExcelCellAddress.GetColumnLetter(1);
            templateEndColLetter = alphabetsPart;

            templateStartColLetter = affectSection.TemplateFromCol;
            templateStartRowIndex = affectSection.TemplateFromRow;
            templateEndColLetter = affectSection.TemplateToCol;
            templateEndRowIndex = affectSection.TemplateToRow;

            appendStartColLetter = affectSection.AppendFromCol;
            appendStartRowIndex = affectSection.AppendFromRow;
            appendEndColLetter = affectSection.AppendToCol;
            appendEndRowIndex = affectSection.AppendToRow;

            // 2.1 find how many rows needed to insert
            int dataGridRowCount = templateEndRowIndex - templateStartRowIndex + 1;

            // 2.2 insert the new row(s)
            //_worksheet.InsertRow(appendStartRowIndex, dataGridRowCount, templateStartRowIndex);
            //this.InsertRow(spreadsheetDocument, "Sheet1", templateStartRowIndex);
            for (int _rowCount = 0; _rowCount <= dataGridRowCount; _rowCount++)
            {
                this.InsertRow(_worksheet, templateStartRowIndex);
            }

            // 3.0 copy value, style to cell/column/row
            // 3.1 calculate the copy from range(fromRange), copy to range( destinationRange), and the new append to Range (newAppendToRange) 
            fromRange = templateStartColLetter + templateStartRowIndex + ":" + templateEndColLetter + templateEndRowIndex;
            destinationRange = affectSection.GetAppendRange();
            destinationRange = templateStartColLetter + appendStartRowIndex + ":" + templateEndColLetter + (appendStartRowIndex + dataGridRowCount - 1);
            newAppendToRange = templateStartColLetter + (appendEndRowIndex + dataGridRowCount) + ":" + templateEndColLetter + (appendEndRowIndex + dataGridRowCount);

            // 3.2 copy template cell value to newly inserted row
            //ExcelRange[int FromRow, int FromCol, int ToRow, int ToCol]
            //_worksheet.Cells[fromRange].Copy(_worksheet.Cells[destinationRange]);
            this.CopyRange(this.sourceSpreadsheetDocument, _worksheet.Name, templateStartRowIndex, templateEndRowIndex, appendStartRowIndex);
            // 3.3 copy row height
            //for (int start = templateStartRowIndex; start <= templateEndRowIndex; start++)
            //{
            //    _worksheet.Row(appendStartRowIndex + (start - templateStartRowIndex)).Height = _worksheet.Row(start).Height;
            //}

            // 3.4 copy row data validation
            // 3.5 copy row conditional formatting

            // 4.0 update the new append to Range in DataGridSection
            // 4.1
            // for others data grid section, update the template range, and append range if its place lower then the inserted position
            /*
                e.g. new ExcelDataSection("T1B", "17:19", "20:20")
                e.g. new ExcelDataSection("T1F", "21:21", "22:22")
                
                if inserted three new rows at 20th for repeating T1B
                then the T1F template, append range should be shifted lower
             */
            foreach (ExcelDataGrid dataGrid in excelDataGridList)
            {
                ExcelDataSection _bodySection = dataGrid.GetDynamicRange();
                //ExcelDataSection _footerSection = dataGrid.GetStaticRange();

                string updateTemplateToRange = string.Empty;
                string updateAppendToRange = string.Empty;
                if (!affectSection.Equals(_bodySection)
                    && _bodySection.AppendFromRow > affectSection.AppendFromRow)
                {
                    updateTemplateToRange = _bodySection.TemplateFromCol + (_bodySection.TemplateFromRow + dataGridRowCount) + ":" + _bodySection.TemplateToCol + (_bodySection.TemplateToRow + dataGridRowCount);
                    updateAppendToRange = _bodySection.AppendFromCol + (_bodySection.AppendFromRow + dataGridRowCount) + ":" + _bodySection.AppendToCol + (_bodySection.AppendToRow + dataGridRowCount);
                    _bodySection.SetTemplateRange(updateTemplateToRange);
                    _bodySection.SetAppendToRange(updateAppendToRange);
                }
            }
            // 4.2
            // for current data grid section, update the append to row range
            /*
                e.g. new ExcelDataSection("T1B", "17:19", "20:20")
                3 rows 17,18,19 will insert at row 20
                the new rows are 20, 21, 22
                and the append to row 20 will shifted to 23
             */
            affectSection.SetAppendToRange(newAppendToRange);

            // 20.1 merge value into newly inserted rows
            string colLetterStart = string.Empty;
            string colLetterEnd = string.Empty;
            colLetterStart = String.IsNullOrEmpty(templateStartColLetter) ? "B" : templateStartColLetter;
            colLetterEnd = String.IsNullOrEmpty(templateEndColLetter) ? this.ExcelGetColumnName(columnCount) : templateEndColLetter;
            int colIndexStart = 0, colIndexEnd = 0, pow = 0;
            pow = 1;
            for (var i = colLetterStart.Length - 1; i >= 0; i--)
            {
                colIndexStart += (colLetterStart[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            pow = 1;
            for (var i = colLetterEnd.Length - 1; i >= 0; i--)
            {
                colIndexEnd += (colLetterEnd[i] - 'A' + 1) * pow;
                pow *= 26;
            }

            // if the grid append direction is FromTopToBottom
            // then the summary will be calculated by column base
            // the render sequence should be
            // column A => cell A1, A2, A3
            // column B => cell B1, B2, B3
            for (int colIndex = colIndexStart; colIndex <= colIndexEnd; colIndex++)
            {
                for (int rowIndex = appendStartRowIndex; rowIndex <= (appendStartRowIndex + dataGridRowCount - 1); rowIndex++)
                {
                    this.DefaultMergeCellExpression(_worksheet, this.ExcelGetColumnName(colIndex) + rowIndex, _tuple);
                    this.CustomPostMergeCellExpression(_worksheet, this.ExcelGetColumnName(colIndex) + rowIndex, _tuple);
                }
            }
        }


        protected virtual void DefaultMergeCellExpression(Sheet _worksheet, string cellAddress, IDictionary<string, Object> _tuple)
        {
            string numericPart = new String(cellAddress.Where(Char.IsDigit).ToArray());
            string alphabetsPart = Regex.Replace(cellAddress, @"[\d-]", string.Empty);

            int rowCount = Convert.ToInt32(numericPart);
            int columnCount = this.ExcelColumnNameToNumber(alphabetsPart);

            WorksheetPart worksheetPart =
                      this.GetWorksheetPartByName(this.sourceSpreadsheetDocument, _worksheet.Name);
            Cell cell = this.GetCell(worksheetPart.Worksheet, alphabetsPart, rowCount);


            // 1.1 get cell value
            CellValue cellVal = cell.CellValue;
            string cellValInStr = cellVal.ToString();

            // 1.2 skip if cell value is empty or null
            if (string.IsNullOrEmpty(cellValInStr)) return;

            // 2.1 match expression between dataRow and cell
            Boolean isMerge = false;
            string matchExpression = string.Empty;
            string mergedValue = cellValInStr;

            foreach (KeyValuePair<string, object> kvp in _tuple)
            {
                matchExpression = "{{" + kvp.Key + "}}";
                if (cellValInStr.IndexOf(matchExpression) > -1)
                {
                    isMerge = true;
                    mergedValue = mergedValue.Replace(matchExpression, kvp.Value.ToString());
                }
            }

            // 2.2 
            if (isMerge)
            {
                // https://stackoverflow.com/questions/527028/open-xml-sdk-2-0-how-to-update-a-cell-in-a-spreadsheet
                cell.CellValue = new CellValue(mergedValue);
                //cell.DataType =
                //        new EnumValue<CellValues>(CellValues.Number);


                //ExcelStyle cellStyle = _cell.Style;
                //string _cellFormat = _cell.Style.Numberformat.Format;
                //// format reference
                //// https://stackoverflow.com/questions/40209636/epplus-number-format/40214134
                //if (_cellFormat.IndexOfAny("%".ToCharArray()) > -1)
                //{
                //    _cell.Value = Convert.ToDecimal(mergedValue);
                //}
                //else if (_cellFormat.IndexOfAny("dMyHmAP".ToCharArray()) > -1)
                //{
                //    _cell.Value = Convert.ToDateTime(mergedValue);
                //}
                //else if (_cellFormat.IndexOfAny("€#,0._$*".ToCharArray()) > -1)
                //{
                //    _cell.Value = Convert.ToDecimal(mergedValue);
                //}
                //else
                //{
                //    _cell.Value = Convert.ToString(mergedValue);
                //}
            }
        }

        protected virtual void CustomPostMergeCellExpression(Sheet _worksheet, string cellAddress, IDictionary<string, Object> _tuple)
        {
            throw new NotImplementedException();
        }

        protected virtual void PrintSectionSeparateLine(Sheet _worksheet, params string[] _indicators)
        {
            // 1. check indicators array, is all are valid (exists in the template)
            List<string> indicatorArray = new List<string>();
            List<ExcelDataGrid> allExcelDataGridList = this.reportEntity.GetDataGrid();
            List<ExcelDataGrid> affectingDataGridList = new List<ExcelDataGrid>();
            List<ExcelDataSection> affectingDataGridSectionList = new List<ExcelDataSection>();
            foreach (string _indicator in _indicators)
            {
                foreach (ExcelDataGrid dataGrid in allExcelDataGridList)
                {
                    ExcelDataSection _bodySection = dataGrid.GetDynamicRange();
                    ExcelDataSection updateDataGridSheet1 = null;
                    if (!string.IsNullOrEmpty(_bodySection.Indicator)
                        && _bodySection.Indicator == _indicator)
                    {
                        indicatorArray.Add(_indicator);
                        if (!affectingDataGridList.Contains(dataGrid)) affectingDataGridList.Add(dataGrid);
                        updateDataGridSheet1 = _bodySection;
                    }

                    if (updateDataGridSheet1 != null)
                    {
                        affectingDataGridSectionList.Add(updateDataGridSheet1);
                    }
                }
            }

            // 2. locate the proper posiition for insert all section appendToRange
            // 2.1 find the most bottom appendToRange
            int mostBottomAppendPosition = -1;
            int mostBottomInsertPosition = -1;
            foreach (ExcelDataGrid dataGrid in affectingDataGridList)
            {
                ExcelDataSection _bodySection = dataGrid.GetDynamicRange();
                if (!string.IsNullOrEmpty(_bodySection.Indicator)
                        && _bodySection.AppendFromRow > mostBottomAppendPosition)
                {
                    mostBottomAppendPosition = _bodySection.AppendFromRow;
                }
            }

            // 3.0 copy data grid section from Template sheet
            // 3.1 copy after the most bottom appendTo position (mostBottomInsertPosition)
            mostBottomInsertPosition = mostBottomAppendPosition + 1;
            // 3.2 find the section from template
            List<ExcelDataGrid> templateDataGridList = this.reportEntity.GetBackupTemplateDataGrid();
            List<ExcelDataSection> targetToCloneGridList = new List<ExcelDataSection>();
            foreach (ExcelDataGrid dataGrid in templateDataGridList)
            {
                List<ExcelDataSection> rangeList = dataGrid.GetRangeList();
                foreach (ExcelDataSection gridSection in rangeList)
                {
                    if (_indicators.Contains(gridSection.Indicator))
                    {
                        targetToCloneGridList.Add(gridSection);
                    }
                }
            }
            // 3.3 order the copy sequence
            //targetToCloneGridList.Sort();

            // 4.0 delete old templateRange, appendToRange from the Sheet1
            // 4.1 delete grids from bottom to top
            //targetToCloneGridList.Reverse();
            foreach (ExcelDataSection dataGridTemplate in affectingDataGridSectionList.Reverse<ExcelDataSection>())
            {
                int rowCountForTemplateRange = (dataGridTemplate.TemplateToRow - dataGridTemplate.TemplateFromRow) + 1;
                int rowCountForAppendRange = (dataGridTemplate.AppendToRow - dataGridTemplate.AppendFromRow) + 1;

                // remove the old appendToRange
                //_worksheet.DeleteRow(dataGridTemplate.AppendFromRow, rowCountForAppendRange);
                this.RemoveRow(this.targetSpreadsheetDocument, _worksheet.Name, dataGridTemplate.AppendFromRow, rowCountForAppendRange);
                mostBottomInsertPosition -= (rowCountForAppendRange);
                // remove the old templateRanageRows
                //_worksheet.DeleteRow(dataGridTemplate.TemplateFromRow, rowCountForTemplateRange);
                this.RemoveRow(this.targetSpreadsheetDocument, _worksheet.Name, dataGridTemplate.TemplateFromRow, rowCountForTemplateRange);
                mostBottomInsertPosition -= (rowCountForTemplateRange);
            }

            // 5.0 insert templateRange, appendToRange to the bottom row (mostBottomInsertPosition)
            int totalShiftedTemplateRow = 0;
            int totalShiftedAppendRow = 0;

            //Sheet templateSheet = this.excelPackage.Workbook.Worksheets["Template"];
            Sheet templateSheet = this.sourceSpreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Template");

            targetToCloneGridList.Sort();
            foreach (ExcelDataSection dataGridTemplate in targetToCloneGridList)
            {
                // insert new rows
                int rowCountForTemplateRange = (dataGridTemplate.TemplateToRow - dataGridTemplate.TemplateFromRow) + 1;
                int rowCountForAppendRange = (dataGridTemplate.AppendToRow - dataGridTemplate.AppendFromRow) + 1;

                // copy templateRange from template sheet
                string copyTemplateDestinationRange = dataGridTemplate.TemplateFromCol + (mostBottomInsertPosition) + ":" + dataGridTemplate.TemplateToCol + (mostBottomInsertPosition + rowCountForTemplateRange - 1);
                //_worksheet.InsertRow(mostBottomInsertPosition, rowCountForTemplateRange);
                this.InsertRows(_worksheet, mostBottomInsertPosition, rowCountForTemplateRange);

                //templateSheet.Cells[dataGridTemplate.GetTemplateRange()].Copy(_worksheet.Cells[copyTemplateDestinationRange]);
                WorkbookPart workbookPart = this.targetSpreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                this.CopyRange(this.sourceSpreadsheetDocument, "Template", this.targetSpreadsheetDocument, _worksheet.Name, dataGridTemplate.GetTemplateRange(), copyTemplateDestinationRange);

                // copy templateRange row height from template sheet
                //for (int start = dataGridTemplate.TemplateFromRow; start <= dataGridTemplate.TemplateToRow; start++)
                //{
                //    _worksheet.Row(mostBottomInsertPosition + (start - dataGridTemplate.TemplateFromRow)).Height = templateSheet.Row(start).Height;
                //}

                // update most bottom insert position
                mostBottomInsertPosition += rowCountForTemplateRange;

                // copy appendtoRange from template sheet
                string copyAppendToDestinationRange = dataGridTemplate.AppendFromCol + (mostBottomInsertPosition) + ":" + dataGridTemplate.AppendToCol + (mostBottomInsertPosition + rowCountForAppendRange - 1);
                //_worksheet.InsertRow(mostBottomInsertPosition, rowCountForAppendRange);
                this.InsertRows(_worksheet, mostBottomInsertPosition, rowCountForAppendRange);
                //templateSheet.Cells[dataGridTemplate.GetAppendRange()].Copy(_worksheet.Cells[copyAppendToDestinationRange]);
                this.CopyRange(this.sourceSpreadsheetDocument, "Template", this.targetSpreadsheetDocument, _worksheet.Name, dataGridTemplate.GetAppendRange(), copyAppendToDestinationRange);

                // copy appendtoRange row height from template sheet
                //for (int start = dataGridTemplate.AppendFromRow; start <= dataGridTemplate.AppendToRow; start++)
                //{
                //    _worksheet.Row(mostBottomInsertPosition + (start - dataGridTemplate.AppendFromRow)).Height = templateSheet.Row(start).Height;
                //}

                // update most bottom insert position
                mostBottomInsertPosition += rowCountForAppendRange;

                foreach (ExcelDataSection updateDataGridSheet1 in affectingDataGridSectionList)
                {
                    if (dataGridTemplate.Indicator != updateDataGridSheet1.Indicator)
                    {
                        continue;
                    }
                    // update templateRange, appendToRange new position
                    updateDataGridSheet1.SetTemplateRange(copyTemplateDestinationRange);
                    updateDataGridSheet1.SetAppendToRange(copyAppendToDestinationRange);

                }

            }

        }



        public virtual SpreadsheetDocument RenderDataAndMergeToTemplate(SpreadsheetDocument _spreadsheetDocument)
        {
            //ExcelPackage _excelPackage = this.GetXlsxTemplateInstance();

            string excelRange = string.Empty;
            string _indicator = string.Empty;

            //Sheet templateSheet = _excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Template");

            //// update indicator for all dataGridList
            //List<ExcelDataGrid> _excelDataGridList = this.reportEntity.GetDataGrid();
            //foreach (ExcelDataGrid excelDataGrid in _excelDataGridList)
            //{
            //    // update header indicator
            //    excelRange = excelDataGrid.GetHeaderRange().GetTemplateRange();
            //    if (string.IsNullOrEmpty(excelRange)) continue;

            //    _indicator = templateSheet.Cells["A" + excelDataGrid.GetHeaderRange().TemplateFromRow].GetValue<string>();
            //    excelDataGrid.GetHeaderRange().Indicator = _indicator;

            //    // update body indicator
            //    excelRange = excelDataGrid.GetBodyRange().GetTemplateRange();
            //    if (string.IsNullOrEmpty(excelRange)) continue;

            //    _indicator = templateSheet.Cells["A" + excelDataGrid.GetBodyRange().TemplateFromRow].GetValue<string>();
            //    excelDataGrid.GetBodyRange().Indicator = _indicator;

            //    // update footer indicator
            //    excelRange = excelDataGrid.GetFooterRange().GetTemplateRange();
            //    if (string.IsNullOrEmpty(excelRange)) continue;

            //    _indicator = templateSheet.Cells["A" + excelDataGrid.GetFooterRange().TemplateFromRow].GetValue<string>();
            //    excelDataGrid.GetFooterRange().Indicator = _indicator;
            //}

            this.RenderBodyAndMergeToTemplate(_spreadsheetDocument);
            this.RenderHeaderAndMergeToTemplate(_spreadsheetDocument);
            this.RenderFooterAndMergeToTemplate(_spreadsheetDocument);

            return _spreadsheetDocument;
        }

        protected virtual SpreadsheetDocument RenderHeaderAndMergeToTemplate(SpreadsheetDocument _spreadsheetDocument)
        {
            return _spreadsheetDocument;
        }
        protected virtual SpreadsheetDocument RenderBodyAndMergeToTemplate(SpreadsheetDocument _spreadsheetDocument)
        {
            return _spreadsheetDocument;
        }
        protected virtual SpreadsheetDocument RenderFooterAndMergeToTemplate(SpreadsheetDocument _spreadsheetDocument)
        {
            return _spreadsheetDocument;
        }
        public virtual void RemoveTemplateRows(SpreadsheetDocument _spreadsheetDocument)
        {
            List<ExcelDataSection> allDataGridSectionList = new List<ExcelDataSection>();
            // find all data grid section
            List<ExcelDataGrid> _excelDataGridList = this.reportEntity.GetDataGrid();
            foreach (ExcelDataGrid dataGrid in _excelDataGridList)
            {
                List<ExcelDataSection> rangeList = dataGrid.GetRangeList();
                foreach (ExcelDataSection gridSection in rangeList)
                {
                    allDataGridSectionList.Add(gridSection);
                }
            }
            allDataGridSectionList.Sort();

            // remove appendRange, templateRange from bottom to top
            Sheet sheet1 = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Sheet1");
            Sheet template = _spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().FirstOrDefault(s => s.Name == "Template");
            string relationshipId = sheet1.Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)_spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
            IEnumerable<Row> rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

            foreach (ExcelDataSection dataGridSection in allDataGridSectionList.Reverse<ExcelDataSection>())
            {
                int deleteAppendRange = dataGridSection.AppendToRow - dataGridSection.AppendFromRow + 1;
                int deleteTemplateRange = dataGridSection.TemplateToRow - dataGridSection.TemplateFromRow + 1;
                //sheet1.DeleteRow(dataGridSection.AppendFromRow, deleteAppendRange);
                //sheet1.DeleteRow(dataGridSection.TemplateFromRow, deleteTemplateRange);

                //rows.ElementAt<Row>(dataGridSection.AppendFromRow).Remove();
                //rows.ElementAt<Row>(dataGridSection.TemplateFromRow).Remove();
                this.RemoveRow(_spreadsheetDocument, "Sheet1", dataGridSection.AppendToRow);
                this.RemoveRow(_spreadsheetDocument, "Sheet1", dataGridSection.TemplateToRow);
            }

            // set sheet1 as default
            //sheet1.View.SetTabSelected();

            // remove template sheet
            //_excelPackage.Workbook.Worksheets.Delete(template);

            // remove column A in sheet1
            //sheet1.DeleteColumn(1);
            // hide column A
            //sheet1.Column(1).Hidden = true;


            // auto fid the columns
            //sheet1.Cells.AutoFitColumns();
        }

        public override void SaveAndDownloadAsBase64()
        {
            this.RefreshPrintDate();
        }

        public override void SaveFile()
        {
            this.RefreshPrintDate();
        }

        public virtual void SavePdf(string _templateFile="", string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            try
            {
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public void InsertText(SpreadsheetDocument _spreadsheetDocument, string docName, string text)
        {
            // Get the SharedStringTablePart. If it does not exist, create a new one.
            SharedStringTablePart shareStringPart;
            if (_spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
            {
                shareStringPart = _spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
            }
            else
            {
                shareStringPart = _spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
            }

            // Insert the text into the SharedStringTablePart.
            int index = this.InsertSharedStringItem(text, shareStringPart);

            // Insert a new worksheet.
            WorksheetPart worksheetPart = this.InsertWorksheet(_spreadsheetDocument.WorkbookPart);

            // Insert cell A1 into the new worksheet.
            Cell cell = InsertCellInWorksheet("A", 1, worksheetPart);

            // Set the value of cell A1.
            cell.CellValue = new CellValue(index.ToString());
            cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

            //// Save the new worksheet.
            //worksheetPart.Worksheet.Save();
        }

        // Given text and a SharedStringTablePart, creates a SharedStringItem with the specified text 
        // and inserts it into the SharedStringTablePart. If the item already exists, returns its index.
        private int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        // Given a WorkbookPart, inserts a new worksheet.
        private WorksheetPart InsertWorksheet(WorkbookPart workbookPart)
        {
            // Add a new worksheet part to the workbook.
            WorksheetPart newWorksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = workbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new sheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0)
            {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            workbookPart.Workbook.Save();

            return newWorksheetPart;
        }

        // Given a column name, a row index, and a WorksheetPart, inserts a cell into the worksheet. 
        // If the cell already exists, returns it. 
        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            Worksheet worksheet = worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            string cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            else
            {
                // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
                Cell refCell = null;
                foreach (Cell cell in row.Elements<Cell>())
                {
                    if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                    {
                        refCell = cell;
                        break;
                    }
                }

                Cell newCell = new Cell() { CellReference = cellReference };
                row.InsertBefore(newCell, refCell);

                worksheet.Save();
                return newCell;
            }
        }


        public bool RemoveColumn(SpreadsheetDocument _spreadsheetDocument, string sheetName, string colName)
        {
            // Open the _spreadsheetDocument for editing.
            IEnumerable<Sheet> sheets = _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                //return sheetName + "doesn't exist";
                return false;
            }

            string relationshipId = sheets.First().Id.Value;

            WorksheetPart worksheetPart = (WorksheetPart)_spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);

            // Get the Total Rows
            IEnumerable<Row> rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

            if (rows.Count() == 0)
            {
                //return "Rows doesn't exist";
                return false;
            }

            // Loop through the rows and adjust Cell Index
            foreach (Row row in rows)
            {
                int index = (int)row.RowIndex.Value;

                IEnumerable<Cell> cells = row.Elements<Cell>();

                IEnumerable<Cell> cellToDelete = cells.Where(c => string.Compare(c.CellReference.Value, colName + index, true) == 0);

                if (cellToDelete.Count() > 0)
                {
                    cellToDelete.First().Remove();
                }
            }
            worksheetPart.Worksheet.Save();
            //return "Removed Column";
            return true;
        }

        public bool InsertRows(SpreadsheetDocument _spreadsheetDocument, string sheetName, int rowIndex, int rowCount)
        {
            bool isRemoved = true;
            if (rowIndex < 0 || rowCount <=0) return false;

            for (int i = 0; i < rowCount; i++)
            {
                isRemoved = isRemoved && this.InsertRow(_spreadsheetDocument, sheetName, Convert.ToUInt32(rowIndex));
            }
            return isRemoved;
        }
        public bool InsertRows(Sheet sheet, int rowIndex, int rowCount)
        {
            bool isRemoved = true;
            if (rowIndex<0 || rowCount <=0) return false;

            for (int i = 0; i < rowCount; i++)
            {
                isRemoved = isRemoved && this.InsertRow(sheet, Convert.ToUInt32(rowIndex));
            }

            return isRemoved;
        }

        public bool InsertRow(SpreadsheetDocument _spreadsheetDocument, string sheetName, int rowIndex)
        {
            if (rowIndex < 0) return false;

            return this.InsertRow(_spreadsheetDocument, sheetName, Convert.ToUInt32(rowIndex));
        }

        public bool InsertRow(SpreadsheetDocument _spreadsheetDocument, string sheetName, uint rowIndex)
        {
            IEnumerable<Sheet> sheets = _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);
            WorkbookPart wbPart = _spreadsheetDocument.WorkbookPart;

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                //return sheetName + "doesn't exist";
                return false;
            }

            string relationshipId = sheets.First().Id.Value;

            WorksheetPart worksheetPart = (WorksheetPart)_spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);

            Sheet sheet = wbPart.Workbook.Descendants<Sheet>().Where((s) => s.Name == sheetName).FirstOrDefault();

            if (sheet != null)
            {
                Worksheet ws = ((WorksheetPart)(wbPart.GetPartById(sheet.Id))).Worksheet;
                SheetData sheetData = ws.WorksheetPart.Worksheet.GetFirstChild<SheetData>();
                Row refRow = GetRow(sheetData, rowIndex);
                ++rowIndex;

                Cell cell1 = new Cell() { CellReference = "A" + rowIndex };
                CellValue cellValue1 = new CellValue();
                cellValue1.Text = "";
                cell1.Append(cellValue1);
                Row newRow = new Row()
                {
                    RowIndex = rowIndex
                };
                newRow.Append(cell1);
                for (int i = (int)rowIndex; i <= sheetData.Elements<Row>().Count(); i++)
                {
                    var row = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == i).FirstOrDefault();
                    row.RowIndex++;
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        string refer = c.CellReference.Value;
                        int num = Convert.ToInt32(Regex.Replace(refer, @"[^\d]*", ""));
                        num++;
                        string letters = Regex.Replace(refer, @"[^A-Z]*", "");
                        c.CellReference.Value = letters + num;
                    }
                }
                sheetData.InsertAfter(newRow, refRow);
                //ws.Save();
                return true;
            }
            return false;
        }


        public bool InsertRow(Sheet sheet, int rowIndex)
        {
            if (rowIndex < 0) return false;

            return this.InsertRow(sheet, Convert.ToUInt32(rowIndex));
        }

        public bool InsertRow(Sheet sheet, uint rowIndex)
        {
            if (sheet != null)
            {
                SheetData sheetData = sheet.GetFirstChild<SheetData>();
                Row refRow = GetRow(sheetData, rowIndex);
                ++rowIndex;

                Cell cell1 = new Cell() { CellReference = "A" + rowIndex };
                CellValue cellValue1 = new CellValue();
                cellValue1.Text = "";
                cell1.Append(cellValue1);
                Row newRow = new Row()
                {
                    RowIndex = rowIndex
                };
                newRow.Append(cell1);
                for (int i = (int)rowIndex; i <= sheetData.Elements<Row>().Count(); i++)
                {
                    var row = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == i).FirstOrDefault();
                    row.RowIndex++;
                    foreach (Cell c in row.Elements<Cell>())
                    {
                        string refer = c.CellReference.Value;
                        int num = Convert.ToInt32(Regex.Replace(refer, @"[^\d]*", ""));
                        num++;
                        string letters = Regex.Replace(refer, @"[^A-Z]*", "");
                        c.CellReference.Value = letters + num;
                    }
                }
                sheetData.InsertAfter(newRow, refRow);
                //ws.Save();
                return true;
            }
            return false;
        }

        public Row GetRow(SheetData wsData, UInt32 rowIndex)
        {
            var row = wsData.Elements<Row>().
            Where(r => r.RowIndex.Value == rowIndex).FirstOrDefault();
            if (row == null)
            {
                row = new Row();
                row.RowIndex = rowIndex;
                wsData.Append(row);
            }
            return row;
        }

        public WorksheetPart
             GetWorksheetPartByName(SpreadsheetDocument document,
             string sheetName)
        {
            IEnumerable<Sheet> sheets =
               document.WorkbookPart.Workbook.GetFirstChild<Sheets>().
               Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.

                return null;
            }

            string relationshipId = sheets.First().Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)
                 document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;

        }

        public Cell GetCell(Worksheet worksheet, string columnName, int rowIndex)
        {
            if (rowIndex <0) return null;

            return this.GetCell(worksheet, columnName, Convert.ToUInt32(rowIndex));
        }
        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and 
        public Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = this.GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare
                   (c.CellReference.Value, columnName +
                   rowIndex, true) == 0).First();

            // or
            /*
             * 
            string cellName = columnIndex + newRowIndex;
            Row row = worksheet.GetFirstChild<SheetData>().Descendants<Row>().FirstOrDefault(r => r.RowIndex.Value == newRowIndex);
            return row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellName);
             */
        }

        public Cell GetCell(SpreadsheetDocument _spreadsheetDocument, Sheet sheet, string columnName, int rowIndex)
        {
            Cell _cell = null;
            // Retrieve a reference to the workbook part.
            WorkbookPart wbPart = _spreadsheetDocument.WorkbookPart;
            // Retrieve a reference to the worksheet part.
            //WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));
            string addressName = columnName + rowIndex;

            if (sheet != null)
            {
                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(sheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                _cell = wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).FirstOrDefault();
            }
            return _cell;
        }

        public CellValue GetCellValue(SpreadsheetDocument _spreadsheetDocument, Sheet theSheet, string columnName, int rowIndex)
        {
            Cell _cell = this.GetCell(_spreadsheetDocument, theSheet, columnName, rowIndex);
            if (_cell == null) return null;

            return _cell.CellValue;
        }

        public string GetCellValue(SpreadsheetDocument _spreadsheetDocument, Sheet theSheet, string addressName)
        {
            string value = null;

            // Retrieve a reference to the workbook part.
            WorkbookPart wbPart = _spreadsheetDocument.WorkbookPart;

            // Throw an exception if there is no sheet.
            if (theSheet == null)
            {
                throw new ArgumentException("sheetName");
            }

            // Retrieve a reference to the worksheet part.
            WorksheetPart wsPart =
                (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

            // Use its Worksheet property to get a reference to the cell 
            // whose address matches the address you supplied.
            Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                Where(c => c.CellReference == addressName).FirstOrDefault();

            // If the cell does not exist, return an empty string.
            if (theCell.InnerText.Length > 0)
            {
                value = theCell.InnerText;

                // If the cell represents an integer number, you are done. 
                // For dates, this code returns the serialized value that 
                // represents the date. The code handles strings and 
                // Booleans individually. For shared strings, the code 
                // looks up the corresponding value in the shared string 
                // table. For Booleans, the code converts the value into 
                // the words TRUE or FALSE.
                if (theCell.DataType != null)
                {
                    switch (theCell.DataType.Value)
                    {
                        case CellValues.Boolean:
                            switch (value)
                            {
                                case "0":
                                    value = "FALSE";
                                    break;
                                default:
                                    value = "TRUE";
                                    break;
                            }
                            break;
                        case CellValues.Date:
                            break;
                        case CellValues.Error:
                            break;
                        case CellValues.InlineString:
                            break;
                        case CellValues.Number:
                            break;
                        case CellValues.SharedString:

                            // For shared strings, look up the value in the
                            // shared strings table.
                            var stringTable =
                                wbPart.GetPartsOfType<SharedStringTablePart>()
                                .FirstOrDefault();

                            // If the shared string table is missing, something 
                            // is wrong. Return the index that is in
                            // the cell. Otherwise, look up the correct text in 
                            // the table.
                            if (stringTable != null)
                            {
                                value =
                                    stringTable.SharedStringTable
                                    .ElementAt(int.Parse(value)).InnerText;
                            }
                            break;
                        case CellValues.String:
                            break;
                    }
                }
            }
            return value;
        }

        // Retrieve the value of a cell, given a file name, sheet name, 
        // and address name.
        public string GetCellValue(SpreadsheetDocument _spreadsheetDocument, string sheetName, string addressName)
        {
            // Retrieve a reference to the workbook part.
            WorkbookPart wbPart = _spreadsheetDocument.WorkbookPart;

            // Find the sheet with the supplied name, and then use that 
            // Sheet object to retrieve a reference to the first worksheet.
            Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                Where(s => s.Name == sheetName).FirstOrDefault();

            return this.GetCellValue(_spreadsheetDocument, theSheet, addressName);
        }


        // Given a worksheet and a row index, return the row.
        public Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
              Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        public void CopyRange(SpreadsheetDocument sourceDocument, string sourceSheetName, SpreadsheetDocument targetDocument, string targetSheetName, string sourceRange, string targetRange)
        {

            // Retrieve a reference to the workbook part.
            WorkbookPart sourceWbPart = sourceDocument.WorkbookPart;
            WorkbookPart targetWbPart = targetDocument.WorkbookPart;

            // Find the sheet with the supplied name, and then use that 
            // Sheet object to retrieve a reference to the first worksheet.
            Sheet sourceSheet = sourceWbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == sourceSheetName).FirstOrDefault();
            Sheet tagetSheet = targetWbPart.Workbook.Descendants<Sheet>().
              Where(s => s.Name == targetSheetName).FirstOrDefault();

            // Throw an exception if there is no sheet.
            if (sourceSheet == null)
            {
                throw new ArgumentException("sourceSheetName");
            }
            // Throw an exception if there is no sheet.
            if (tagetSheet == null)
            {
                throw new ArgumentException("targetSheetName");
            }

            // Retrieve a reference to the worksheet part.
            WorksheetPart sourceWsPart =
                (WorksheetPart)(sourceWbPart.GetPartById(sourceSheet.Id));
            WorksheetPart targetWsPart =
                (WorksheetPart)(targetWbPart.GetPartById(tagetSheet.Id));


            // loop source range - in terms of row
            // insert new row/column at targetRange - in terms of append direction
            // loop source range - in terms of column
            // read cell value from source
            // copy cell value to target
        }

        #region Copy Range
        public void CopyRange(SpreadsheetDocument document, string sheetName, int srcRowFrom, int srcRowTo, int destRowFrom)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            if (srcRowTo < srcRowFrom || destRowFrom < srcRowFrom) return;
            int destRowFromBase = destRowFrom;

            WorksheetPart worksheetPart = this.GetWorksheetPartByName(document, sheetName);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            IList<Cell> cells = sheetData.Descendants<Cell>().Where(c =>
                GetRowIndex(c.CellReference) >= srcRowFrom &&
                GetRowIndex(c.CellReference) <= srcRowTo).ToList<Cell>();

            if (cells.Count() == 0) return;

            int copiedRowCount = srcRowTo - srcRowFrom + 1;

            MoveRowIndex(document, sheetName, destRowFrom - 1, srcRowTo, srcRowFrom);

            IDictionary<int, IList<Cell>> clonedCells = null;

            IList<Cell> formulaCells = new List<Cell>();

            IList<Row> cloneRelatedRows = new List<Row>();

            destRowFrom = destRowFromBase;
            int changedRowsCount = destRowFrom - srcRowFrom;

            formulaCells.Clear();

            clonedCells = new Dictionary<int, IList<Cell>>();

            foreach (Cell cell in cells)
            {
                Cell newCell = (Cell)cell.CloneNode(true);
                int index = Convert.ToInt32(GetRowIndex(cell.CellReference));

                int rowIndex = index - changedRowsCount;
                newCell.CellReference = GetColumnName(cell.CellReference) + rowIndex.ToString();

                IList<Cell> rowCells = null;
                if (clonedCells.ContainsKey(rowIndex))
                    rowCells = clonedCells[rowIndex];
                else
                {
                    rowCells = new List<Cell>();
                    clonedCells.Add(rowIndex, rowCells);
                }
                rowCells.Add(newCell);

                if (newCell.CellFormula != null && newCell.CellFormula.Text.Length > 0)
                {
                    formulaCells.Add(newCell);
                }
            }

            foreach (int rowIndex in clonedCells.Keys)
            {
                Row row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).FirstOrDefault();
                if (row == null)
                {
                    row = new Row() { RowIndex = (uint)rowIndex };

                    Row refRow = sheetData.Elements<Row>().Where(r => r.RowIndex > rowIndex).OrderBy(r => r.RowIndex).FirstOrDefault();
                    if (refRow == null)
                        sheetData.AppendChild<Row>(row);
                    else
                        sheetData.InsertBefore<Row>(row, refRow);
                }
                row.Append(clonedCells[rowIndex].ToArray());

                cloneRelatedRows.Add(row);
            }

            ChangeFormulaRowNumber(worksheetPart.Worksheet, formulaCells, changedRowsCount);

            foreach (Row row in cloneRelatedRows)
            {
                IList<Cell> cs = row.Elements<Cell>().OrderBy(c => c.CellReference.Value).ToList<Cell>();
                row.RemoveAllChildren();
                row.Append(cs.ToArray());
            }

            MergeCells mcells = worksheetPart.Worksheet.GetFirstChild<MergeCells>();
            if (mcells != null)
            {
                IList<MergeCell> newMergeCells = new List<MergeCell>();
                IEnumerable<MergeCell> clonedMergeCells = mcells.Elements<MergeCell>().
                    Where(m => MergeCellInRange(m, srcRowFrom, srcRowTo)).ToList<MergeCell>();
                foreach (MergeCell cmCell in clonedMergeCells)
                {
                    MergeCell newMergeCell = CreateChangedRowMergeCell(worksheetPart.Worksheet, cmCell, changedRowsCount);
                    newMergeCells.Add(newMergeCell);
                }
                uint count = mcells.Count.Value;
                mcells.Count = new UInt32Value(count + (uint)newMergeCells.Count);
                mcells.Append(newMergeCells.ToArray());
            }
        }

        private void MoveRowIndex(SpreadsheetDocument document, string sheetName, int destRowFrom, int srcRowTo, int srcRowFrom)
        {
            WorksheetPart worksheetPart = this.GetWorksheetPartByName(document, sheetName);
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            uint newRowIndex;

            IEnumerable<Row> rows = sheetData.Descendants<Row>().Where(r => r.RowIndex.Value >= srcRowFrom && r.RowIndex.Value <= srcRowTo);
            foreach (Row row in rows)
            {
                newRowIndex = Convert.ToUInt32(destRowFrom + 1);

                foreach (Cell cell in row.Elements<Cell>())
                {
                    string cellReference = cell.CellReference.Value;
                    cell.CellReference = new StringValue(cellReference.Replace(row.RowIndex.Value.ToString(), newRowIndex.ToString()));
                }
                row.RowIndex = new UInt32Value(newRowIndex);

                destRowFrom++;
            }

        }

        private void ChangeFormulaRowNumber(Worksheet worksheet, IList<Cell> formulaCells, int changedRowsCount)
        {
            foreach (Cell formulaCell in formulaCells)
            {
                Regex regex = new Regex(@"\d+");
                var rowIndex = Convert.ToInt32(regex.Match(formulaCell.CellReference).Value);

                Regex regex2 = new Regex("[A-Za-z]+");
                var columnIndex = regex2.Match(formulaCell.CellReference).Value;

                int newRowIndex = rowIndex + changedRowsCount;
                Cell cell = this.GetCell(worksheet, columnIndex, newRowIndex);
                cell.CellFormula = new CellFormula(cell.CellFormula.Text.Replace($"{rowIndex}", $"{newRowIndex}"));
            }
        }

        private static MergeCell CreateChangedRowMergeCell(Worksheet worksheet, MergeCell cmCell, int changedRows)
        {
            string[] cells = cmCell.Reference.Value.Split(':', 2, StringSplitOptions.RemoveEmptyEntries);

            Regex regex = new Regex(@"\d+");
            var rowIndex1 = Convert.ToInt32(regex.Match(cells[0]).Value);
            var rowIndex2 = Convert.ToInt32(regex.Match(cells[1]).Value);

            Regex regex2 = new Regex("[A-Za-z]+");
            var columnIndex1 = regex2.Match(cells[0]).Value;
            var columnIndex2 = regex2.Match(cells[1]).Value;

            var cell1Name = $"{columnIndex1}{rowIndex1 + changedRows}";
            var cell2Name = $"{columnIndex2}{rowIndex2 + changedRows}";

            CreateSpreadsheetCellIfNotExist(worksheet, cell1Name);
            CreateSpreadsheetCellIfNotExist(worksheet, cell2Name);

            return new MergeCell() { Reference = new StringValue(cell1Name + ":" + cell2Name) };
        }

        private static bool MergeCellInRange(MergeCell mergeCell, int srcRowFrom, int srcRowTo)
        {
            string[] cells = mergeCell.Reference.Value.Split(':', 2, StringSplitOptions.RemoveEmptyEntries);

            Regex regex = new Regex(@"\d+");
            var cellIndex1 = Convert.ToInt32(regex.Match(cells[0]).Value);
            var cellIndex2 = Convert.ToInt32(regex.Match(cells[1]).Value);

            if (srcRowFrom <= cellIndex1 && cellIndex1 <= srcRowTo &&
                srcRowFrom <= cellIndex2 && cellIndex2 <= srcRowTo)
                return true;
            else
                return false;
        }

        private static uint GetRowIndex(string cellName)
        {
            Regex regex = new Regex(@"\d+");
            Match match = regex.Match(cellName);

            return uint.Parse(match.Value);
        }

        private static string GetColumnName(string cellName)
        {
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellName);

            return match.Value;
        }

        private static void CreateSpreadsheetCellIfNotExist(Worksheet worksheet, string cellName)
        {
            string columnName = GetColumnName(cellName);
            uint rowIndex = GetRowIndex(cellName);

            IEnumerable<Row> rows = worksheet.Descendants<Row>().Where(r => r.RowIndex.Value == rowIndex);

            if (rows.Count() == 0)
            {
                Row row = new Row() { RowIndex = new UInt32Value(rowIndex) };
                Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                row.Append(cell);
                worksheet.Descendants<SheetData>().First().Append(row);
                worksheet.Save();
            }
            else
            {
                Row row = rows.First();

                IEnumerable<Cell> cells = row.Elements<Cell>().Where(c => c.CellReference.Value == cellName);

                if (cells.Count() == 0)
                {
                    Cell cell = new Cell() { CellReference = new StringValue(cellName) };
                    row.Append(cell);
                    worksheet.Save();
                }
            }
        }
        #endregion

        public bool RemoveRow(SpreadsheetDocument _spreadsheetDocument, string sheetName, int rowIndex, int rowCount)
        {
            bool isRemoved = true;
            for (int i = 0; i < rowCount; i++) {
                isRemoved = isRemoved && this.RemoveRow(_spreadsheetDocument, sheetName, rowIndex);
            }
            return isRemoved;
        }

        public bool RemoveRow(SpreadsheetDocument _spreadsheetDocument, string sheetName, int rowIndex)
        {
            if (rowIndex < 0) return false;

            return this.RemoveRow(_spreadsheetDocument, sheetName, Convert.ToUInt32(rowIndex));
        }

        public bool RemoveRow(SpreadsheetDocument _spreadsheetDocument, string sheetName, uint rowIndex)
        {
            IEnumerable<Sheet> sheets = _spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                // The specified worksheet does not exist.
                //return sheetName + "doesn't exist";
                return false;
            }

            string relationshipId = sheets.First().Id.Value;

            WorksheetPart worksheetPart = (WorksheetPart)_spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);

            // Get the Total Rows
            IEnumerable<Row> rows = worksheetPart.Worksheet.GetFirstChild<SheetData>().Elements<Row>();

            if (rows.Count() == 0)
            {
                //return "Rows doesn't exist";
                return false;
            }

            // Loop through the rows and adjust Cell Index
            foreach (Row row in rows)
            {
                int index = (int)row.RowIndex.Value;

                if (rowIndex == index)
                {
                    row.Remove();
                }
            }
            //worksheetPart.Worksheet.Save();
            //return "Removed Row";
            return true;
        }

        protected object ConvertDataSetToObject(DataSet _dataSet)
        {
        var _obj = new ExpandoObject() as IDictionary<string, object>;
            if (_dataSet == null || _dataSet.Tables.Count == 0) return _obj;

            foreach (DataTable _table in _dataSet.Tables)
            {
                List<dynamic> rowList = new List<dynamic>();
                _obj.Add(_table.TableName, rowList);
                foreach (DataRow _row in _table.Rows)
                {
                    var expandoDict = new ExpandoObject() as IDictionary<String, Object>;
                    foreach (DataColumn col in _table.Columns)
                    {
                        //put every column of this row into the new dictionary
                        expandoDict.Add(col.ColumnName, _row[col.ColumnName]);
                    }
                    rowList.Add(expandoDict);
                }
            }

            return _obj;
        }
    }
}