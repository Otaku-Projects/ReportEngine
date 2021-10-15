using OfficeOpenXml;
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

namespace CoreReport.EPPlus5Report
{
    public class EPPlus5Decorator : VisualizationDecorator
    {
        protected string createdBy;
        protected DateTime createdDate;
        protected DateTime printedDate;
        protected string filename;

        protected BaseReportEntity reportEntity;

        protected DataSet dataSet;
        protected IDictionary<string, object> dataSetObj;
        protected string epplusReportRenderFolder;

        protected ExcelPackage excelPackage;

        public EPPlus5Decorator()
        {
            this.epplusReportRenderFolder = this.tempRenderFolder;

            FileOutputUtil.OutputDir = new DirectoryInfo(@"D:\Temp");
            FileOutputUtil fileOutputUtil = new FileOutputUtil();

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }
        public EPPlus5Decorator(BaseReportEntity _reportEntity, string _filename = "")
        {
            if (string.IsNullOrEmpty(_filename))
            {
                Guid obj = Guid.NewGuid();
                _filename = obj.ToString();
            }

            this.dataSet = _reportEntity.GetDataSet();
            this.dataSetObj = _reportEntity.GetDataSetObj();
            this.epplusReportRenderFolder = this.tempRenderFolder;

            this.filename = _filename;

            this.reportEntity = _reportEntity;

            this.createdBy = "CoreSystem";
            this.createdDate = new DateTime();

            FileOutputUtil.OutputDir = new DirectoryInfo(this.epplusReportRenderFolder);
            FileOutputUtil fileOutputUtil = new FileOutputUtil();

            // If you use EPPlus in a noncommercial context
            // according to the Polyform Noncommercial license:
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
        }

        public void RefreshPrintDate()
        {
            this.printedDate = new DateTime();
        }

        public override void Display()
        {
            throw new NotImplementedException();
        }

        protected ExcelPackage CreateXlsxTemplateInstance()
        {
            string _xlsxTemplateFilePath = this.reportEntity.GetXlsxTemplateFilePath();
            if (!File.Exists(_xlsxTemplateFilePath))
            {
                throw new FileNotFoundException($"Excel template not found at {_xlsxTemplateFilePath}");
            }
            ExcelPackage _excelPackage = new ExcelPackage(new FileInfo(_xlsxTemplateFilePath));
            return _excelPackage;
        }
        public ExcelPackage GetXlsxTemplateInstance()
        {
            return this.CreateXlsxTemplateInstance();
        }
        public virtual void RenderTemplateAndSaveAsXlsx(string _fileName = "")
        {
            // you should not call into here, please inherit the decorator and override this function
            throw new NotImplementedException();
        }
        public virtual void RenderTemplateAndSaveAsPdf(string _fileName = "")
        {
            // you should not call into here, please inherit the decorator and override this function
            throw new NotImplementedException();
        }
        public virtual ExcelPackage RenderDataAndMergeToTemplate(ExcelPackage _excelPackage)
        {
            return this.GetXlsxTemplateInstance();
        }
        protected virtual ExcelPackage StartRenderDataAndMergeToTemplate()
        {
            ExcelPackage _excelPackage = this.GetXlsxTemplateInstance();

            // if sheet1 exists, delete it
            //ExcelWorksheet sheet1 = _excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Sheet1");
            ExcelWorksheet sheet1 = _excelPackage.Workbook.Worksheets["Sheet1"];
            if (sheet1 != null)
            {
                _excelPackage.Workbook.Worksheets.Delete("Sheet1");
            }

            // clone template sheet to sheet1
            //ExcelWorksheet clonedSheet = _excelPackage.Workbook.Worksheets.Copy("Template", "Sheet1");
            _excelPackage.Workbook.Worksheets.Copy("Template", "Sheet1");
            sheet1 = _excelPackage.Workbook.Worksheets["Sheet1"];
            //_excelPackage.Workbook.Worksheets["Template"].View.TabSelected = false;

            // backup the data grid 
            this.reportEntity.BackupDataGridSetting();

            // move sheet1 to start
            _excelPackage.Workbook.Worksheets.MoveToStart("Sheet1");

            // select sheet1 as the default sheet
            sheet1.View.SetTabSelected();

            this.excelPackage = _excelPackage;
            return _excelPackage;
        }
        protected virtual void MergeDataRows(ExcelWorksheet _worksheet, string _indicator, DataTable _dt)
        {
            int i = 1;
            foreach (DataRow _dtRow in _dt.Rows)
            {
                i++;
                this.MergeDataRow(_worksheet, _indicator, _dtRow);
                //if(i==3)
                //break;
            }
        }

        protected virtual void MergeDataRow(ExcelWorksheet _worksheet, string _indicator, DataRow _dataRow)
        {
            var expObj = new ExpandoObject() as IDictionary<string, Object>;

            foreach (DataColumn dc in _dataRow.Table.Columns)
            {
                expObj.Add(dc.ColumnName, _dataRow[dc]);
            }

            this.MergeDataRow(_worksheet, _indicator, expObj);
        }

        protected virtual void MergeDataRow(ExcelWorksheet _worksheet, string _indicator, IDictionary<string, Object> _tuple)
        {
            List<ExcelDataGrid> excelDataGridList = this.reportEntity.GetDataGrid();

            int rowCount = _worksheet.Dimension.Rows;
            int columnCount = _worksheet.Dimension.Columns;

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
            ExcelDataGridSection affectSection = null;
            foreach (ExcelDataGrid dataGrid in excelDataGridList)
            {
                if (dataGrid.GetHeaderRange().Indicator == _indicator)
                {
                    affectSection = dataGrid.GetHeaderRange();
                    break;
                }
                else if (dataGrid.GetBodyRange().Indicator == _indicator)
                {
                    affectSection = dataGrid.GetBodyRange();
                    break;
                }
                else if (dataGrid.GetFooterRange().Indicator == _indicator)
                {
                    affectSection = dataGrid.GetFooterRange();
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
            templateEndColLetter = OfficeOpenXml.ExcelCellAddress.GetColumnLetter(columnCount);

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
            _worksheet.InsertRow(appendStartRowIndex, dataGridRowCount, templateStartRowIndex);

            // 3.0 copy value, style to cell/column/row
            // 3.1 calculate the copy from range(fromRange), copy to range( destinationRange), and the new append to Range (newAppendToRange) 
            fromRange = templateStartColLetter + templateStartRowIndex + ":" + templateEndColLetter + templateEndRowIndex;
            destinationRange = affectSection.GetAppendRange();
            destinationRange = templateStartColLetter + appendStartRowIndex + ":" + templateEndColLetter + (appendStartRowIndex + dataGridRowCount - 1);
            newAppendToRange = templateStartColLetter + (appendEndRowIndex + dataGridRowCount) + ":" + templateEndColLetter + (appendEndRowIndex + dataGridRowCount);

            // 3.2 copy template cell value to newly inserted row
            //ExcelRange[int FromRow, int FromCol, int ToRow, int ToCol]
            _worksheet.Cells[fromRange].Copy(_worksheet.Cells[destinationRange]);
            // 3.3 copy row height
            for (int start = templateStartRowIndex; start <= templateEndRowIndex; start++)
            {
                _worksheet.Row(appendStartRowIndex + (start - templateStartRowIndex)).Height = _worksheet.Row(start).Height;
            }
            // 3.4 copy row data validation
            // 3.5 copy row conditional formatting

            // 4.0 update the new append to Range in DataGridSection
            // 4.1
            // for others data grid section, update the template range, and append range if its place lower then the inserted position
            /*
                e.g. new ExcelDataGridSection("T1B", "17:19", "20:20")
                e.g. new ExcelDataGridSection("T1F", "21:21", "22:22")
                
                if inserted three new rows at 20th for repeating T1B
                then the T1F template, append range should be shifted lower
             */
            foreach (ExcelDataGrid dataGrid in excelDataGridList)
            {
                ExcelDataGridSection _headerSection = dataGrid.GetHeaderRange();
                ExcelDataGridSection _bodySection = dataGrid.GetBodyRange();
                ExcelDataGridSection _footerSection = dataGrid.GetFooterRange();

                string updateTemplateToRange = string.Empty;
                string updateAppendToRange = string.Empty;
                if (!affectSection.Equals(_headerSection)
                    && _headerSection.AppendFromRow > affectSection.AppendFromRow)
                {
                    updateTemplateToRange = _headerSection.TemplateFromCol + (_headerSection.TemplateFromRow + dataGridRowCount) + ":" + _headerSection.TemplateToCol + (_headerSection.TemplateToRow + dataGridRowCount);
                    updateAppendToRange = _headerSection.AppendFromCol + (_headerSection.AppendFromRow + dataGridRowCount) + ":" + _headerSection.AppendToCol + (_headerSection.AppendToRow + dataGridRowCount);
                    _headerSection.SetTemplateRange(updateTemplateToRange);
                    _headerSection.SetAppendToRange(updateAppendToRange);
                }
                if (!affectSection.Equals(_bodySection)
                    && _bodySection.AppendFromRow > affectSection.AppendFromRow)
                {
                    updateTemplateToRange = _bodySection.TemplateFromCol + (_bodySection.TemplateFromRow + dataGridRowCount) + ":" + _bodySection.TemplateToCol + (_bodySection.TemplateToRow + dataGridRowCount);
                    updateAppendToRange = _bodySection.AppendFromCol + (_bodySection.AppendFromRow + dataGridRowCount) + ":" + _bodySection.AppendToCol + (_bodySection.AppendToRow + dataGridRowCount);
                    _bodySection.SetTemplateRange(updateTemplateToRange);
                    _bodySection.SetAppendToRange(updateAppendToRange);
                }
                if (!affectSection.Equals(_footerSection)
                    && _footerSection.AppendFromRow > affectSection.AppendFromRow)
                {
                    updateTemplateToRange = _footerSection.TemplateFromCol + (_footerSection.TemplateFromRow + dataGridRowCount) + ":" + _footerSection.TemplateToCol + (_footerSection.TemplateToRow + dataGridRowCount);
                    updateAppendToRange = _footerSection.AppendFromCol + (_footerSection.AppendFromRow + dataGridRowCount) + ":" + _footerSection.AppendToCol + (_footerSection.AppendToRow + dataGridRowCount);
                    _footerSection.SetTemplateRange(updateTemplateToRange);
                    _footerSection.SetAppendToRange(updateAppendToRange);
                }
            }
            // 4.2
            // for current data grid section, update the append to row range
            /*
                e.g. new ExcelDataGridSection("T1B", "17:19", "20:20")
                3 rows 17,18,19 will insert at row 20
                the new rows are 20, 21, 22
                and the append to row 20 will shifted to 23
             */
            affectSection.SetAppendToRange(newAppendToRange);

            // 20.1 merge value into newly inserted rows
            string colLetterStart = string.Empty;
            string colLetterEnd = string.Empty;
            colLetterStart = String.IsNullOrEmpty(templateStartColLetter) ? "B" : templateStartColLetter;
            colLetterEnd = String.IsNullOrEmpty(templateEndColLetter) ? OfficeOpenXml.ExcelCellAddress.GetColumnLetter(columnCount) : templateEndColLetter;
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
                    this.DefaultMergeCellExpression(_worksheet, OfficeOpenXml.ExcelCellAddress.GetColumnLetter(colIndex) + rowIndex, _tuple);
                    this.CustomPostMergeCellExpression(_worksheet, OfficeOpenXml.ExcelCellAddress.GetColumnLetter(colIndex) + rowIndex, _tuple);
                }
            }
        }
        protected virtual void DefaultMergeCellExpression(ExcelWorksheet _worksheet, string cellAddress, IDictionary<string, Object> _tuple)
        {
            ExcelRange _cell = _worksheet.Cells[cellAddress];

            // 1.1 get cell value
            string cellVal = _cell.GetValue<string>();
            // 1.2 skip if cell value is empty or null
            if (string.IsNullOrEmpty(cellVal)) return;

            // 2.1 match expression between dataRow and cell
            Boolean isMerge = false;
            string matchExpression = string.Empty;
            string mergedValue = cellVal;

            foreach (KeyValuePair<string, object> kvp in _tuple)
            {
                matchExpression = "{{" + kvp.Key + "}}";
                if (cellVal.IndexOf(matchExpression) > -1)
                {
                    isMerge = true;
                    mergedValue = mergedValue.Replace(matchExpression, kvp.Value.ToString());
                }
            }

            // 2.2 
            if (isMerge)
            {
                ExcelStyle cellStyle = _cell.Style;
                string _cellFormat = _cell.Style.Numberformat.Format;
                // format reference
                // https://stackoverflow.com/questions/40209636/epplus-number-format/40214134
                if (_cellFormat.IndexOfAny("%".ToCharArray()) > -1)
                {
                    _cell.Value = Convert.ToDecimal(mergedValue);
                }
                else if (_cellFormat.IndexOfAny("dMyHmAP".ToCharArray()) > -1)
                {
                    _cell.Value = Convert.ToDateTime(mergedValue);
                }
                else if (_cellFormat.IndexOfAny("€#,0._$*".ToCharArray()) > -1)
                {
                    _cell.Value = Convert.ToDecimal(mergedValue);
                }
                else
                {
                    _cell.Value = Convert.ToString(mergedValue);
                }
            }
        }

        protected virtual void CustomPostMergeCellExpression(ExcelWorksheet _worksheet, string cellAddress, IDictionary<string, Object> _tuple)
        {
            throw new NotImplementedException();
        }

        protected virtual void PrintSectionSeparateLine(ExcelWorksheet _worksheet, params string[] _indicators)
        {
            // 1. check indicators array, is all are valid (exists in the template)
            List<string> indicatorArray = new List<string>();
            List<ExcelDataGrid> allExcelDataGridList = this.reportEntity.GetDataGrid();
            List<ExcelDataGrid> affectingDataGridList = new List<ExcelDataGrid>();
            List<ExcelDataGridSection> affectingDataGridSectionList = new List<ExcelDataGridSection>();
            foreach (string _indicator in _indicators) {
                foreach (ExcelDataGrid dataGrid in allExcelDataGridList)
                {
                    ExcelDataGridSection _headerSection = dataGrid.GetHeaderRange();
                    ExcelDataGridSection _bodySection = dataGrid.GetBodyRange();
                    ExcelDataGridSection _footerSection = dataGrid.GetFooterRange();
                    ExcelDataGridSection updateDataGridSheet1 = null;
                    if (!string.IsNullOrEmpty(_headerSection.Indicator)
                        && _headerSection.Indicator == _indicator)
                    {
                        indicatorArray.Add(_indicator);
                        if(!affectingDataGridList.Contains(dataGrid)) affectingDataGridList.Add(dataGrid);
                        updateDataGridSheet1 = _headerSection;
                    }
                    else if (!string.IsNullOrEmpty(_bodySection.Indicator)
                        && _bodySection.Indicator == _indicator)
                    {
                        indicatorArray.Add(_indicator);
                        if (!affectingDataGridList.Contains(dataGrid)) affectingDataGridList.Add(dataGrid);
                        updateDataGridSheet1 = _bodySection;
                    }
                    else if (!string.IsNullOrEmpty(_footerSection.Indicator)
                        && _footerSection.Indicator == _indicator)
                    {
                        indicatorArray.Add(_indicator);
                        if (!affectingDataGridList.Contains(dataGrid)) affectingDataGridList.Add(dataGrid);
                        updateDataGridSheet1 = _footerSection;
                    }

                    if(updateDataGridSheet1!= null)
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
                ExcelDataGridSection _headerSection = dataGrid.GetHeaderRange();
                ExcelDataGridSection _bodySection = dataGrid.GetBodyRange();
                ExcelDataGridSection _footerSection = dataGrid.GetFooterRange();
                if (!string.IsNullOrEmpty(_headerSection.Indicator)
                        && _headerSection.AppendFromRow > mostBottomAppendPosition) 
                {
                    mostBottomAppendPosition = _headerSection.AppendFromRow;
                }
                if (!string.IsNullOrEmpty(_bodySection.Indicator)
                        && _bodySection.AppendFromRow > mostBottomAppendPosition)
                {
                    mostBottomAppendPosition = _bodySection.AppendFromRow;
                }
                if (!string.IsNullOrEmpty(_footerSection.Indicator)
                        && _footerSection.AppendFromRow > mostBottomAppendPosition)
                {
                    mostBottomAppendPosition = _footerSection.AppendFromRow;
                }
            }

            // 3.0 copy data grid section from Template sheet
            // 3.1 copy after the most bottom appendTo position (mostBottomInsertPosition)
            mostBottomInsertPosition = mostBottomAppendPosition + 1;
            // 3.2 find the section from template
            List<ExcelDataGrid> templateDataGridList = this.reportEntity.GetBackupTemplateDataGrid();
            List<ExcelDataGridSection> targetToCloneGridList = new List<ExcelDataGridSection>();
            foreach (ExcelDataGrid dataGrid in templateDataGridList)
            {
                List<ExcelDataGridSection> rangeList = dataGrid.GetRangeList();
                foreach(ExcelDataGridSection gridSection in rangeList)
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
            foreach (ExcelDataGridSection dataGridTemplate in affectingDataGridSectionList.Reverse<ExcelDataGridSection>()) {
                int rowCountForTemplateRange = (dataGridTemplate.TemplateToRow - dataGridTemplate.TemplateFromRow) + 1;
                int rowCountForAppendRange = (dataGridTemplate.AppendToRow - dataGridTemplate.AppendFromRow) + 1;

                // remove the old appendToRange
                _worksheet.DeleteRow(dataGridTemplate.AppendFromRow, rowCountForAppendRange);
                mostBottomInsertPosition -= (rowCountForAppendRange);
                // remove the old templateRanageRows
                _worksheet.DeleteRow(dataGridTemplate.TemplateFromRow, rowCountForTemplateRange);
                mostBottomInsertPosition -= (rowCountForTemplateRange);
            }

            // 5.0 insert templateRange, appendToRange to the bottom row (mostBottomInsertPosition)
            int totalShiftedTemplateRow = 0;
            int totalShiftedAppendRow = 0;
            //ExcelWorksheet templateSheet = this.excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Template");
            ExcelWorksheet templateSheet = this.excelPackage.Workbook.Worksheets["Template"];
            targetToCloneGridList.Sort();
            foreach (ExcelDataGridSection dataGridTemplate in targetToCloneGridList)
            {
                // insert new rows
                int rowCountForTemplateRange = (dataGridTemplate.TemplateToRow - dataGridTemplate.TemplateFromRow) + 1;
                int rowCountForAppendRange = (dataGridTemplate.AppendToRow - dataGridTemplate.AppendFromRow) + 1;

                // copy templateRange from template sheet
                string copyTemplateDestinationRange = dataGridTemplate.TemplateFromCol + (mostBottomInsertPosition) + ":" + dataGridTemplate.TemplateToCol + (mostBottomInsertPosition + rowCountForTemplateRange - 1);
                _worksheet.InsertRow(mostBottomInsertPosition, rowCountForTemplateRange);
                templateSheet.Cells[dataGridTemplate.GetTemplateRange()].Copy(_worksheet.Cells[copyTemplateDestinationRange]);

                // copy templateRange row height from template sheet
                for (int start = dataGridTemplate.TemplateFromRow; start <= dataGridTemplate.TemplateToRow; start++)
                {
                    _worksheet.Row(mostBottomInsertPosition + (start - dataGridTemplate.TemplateFromRow)).Height = templateSheet.Row(start).Height;
                }
                mostBottomInsertPosition += rowCountForTemplateRange;

                // copy appendtoRange from template sheet
                string copyAppendToDestinationRange = dataGridTemplate.AppendFromCol + (mostBottomInsertPosition) + ":" + dataGridTemplate.AppendToCol + (mostBottomInsertPosition + rowCountForAppendRange - 1);
                _worksheet.InsertRow(mostBottomInsertPosition, rowCountForAppendRange);
                templateSheet.Cells[dataGridTemplate.GetAppendRange()].Copy(_worksheet.Cells[copyAppendToDestinationRange]);
                // copy appendtoRange row height from template sheet
                for (int start = dataGridTemplate.AppendFromRow; start <= dataGridTemplate.AppendToRow; start++)
                {
                    _worksheet.Row(mostBottomInsertPosition + (start - dataGridTemplate.AppendFromRow)).Height = templateSheet.Row(start).Height;
                }
                mostBottomInsertPosition += rowCountForAppendRange;

                foreach(ExcelDataGridSection updateDataGridSheet1 in affectingDataGridSectionList)
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


        public virtual void RemoveTemplateRows(ExcelPackage _excelPackage)
        {
            List<ExcelDataGridSection> allDataGridSectionList = new List<ExcelDataGridSection>();
            // find all data grid section
            List<ExcelDataGrid> _excelDataGridList = this.reportEntity.GetDataGrid();
            foreach (ExcelDataGrid dataGrid in _excelDataGridList)
            {
                List<ExcelDataGridSection> rangeList = dataGrid.GetRangeList();
                foreach (ExcelDataGridSection gridSection in rangeList)
                {
                    allDataGridSectionList.Add(gridSection);
                }
            }
            allDataGridSectionList.Sort();

            // remove appendRange, templateRange from bottom to top
            //ExcelWorksheet template = this.excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Template");
            //ExcelWorksheet sheet1 = this.excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Sheet1");
            ExcelWorksheet template = this.excelPackage.Workbook.Worksheets["Template"];
            ExcelWorksheet sheet1 = this.excelPackage.Workbook.Worksheets["Sheet1"];
            foreach (ExcelDataGridSection dataGridSection in allDataGridSectionList.Reverse<ExcelDataGridSection>())
            {
                int deleteAppendRange = dataGridSection.AppendToRow - dataGridSection.AppendFromRow + 1;
                int deleteTemplateRange = dataGridSection.TemplateToRow - dataGridSection.TemplateFromRow + 1;
                sheet1.DeleteRow(dataGridSection.AppendFromRow, deleteAppendRange);
                sheet1.DeleteRow(dataGridSection.TemplateFromRow, deleteTemplateRange);
            }

            // set sheet1 as default
            //sheet1.View.SetTabSelected();

            // remove template sheet
            //_excelPackage.Workbook.Worksheets.Delete(template);

            // remove column A in sheet1
            //sheet1.DeleteColumn(1);
            // hide column A
            sheet1.Column(1).Hidden = true;

            // auto fid the columns
            //sheet1.Cells.AutoFitColumns();
        }

        public virtual void RemoveTemplateRowsForXlsx(ExcelPackage _excelPackage)
        {
            List<ExcelDataGridSection> allDataGridSectionList = new List<ExcelDataGridSection>();
            // find all data grid section
            List<ExcelDataGrid> _excelDataGridList = this.reportEntity.GetDataGrid();
            foreach (ExcelDataGrid dataGrid in _excelDataGridList)
            {
                List<ExcelDataGridSection> rangeList = dataGrid.GetRangeList();
                foreach (ExcelDataGridSection gridSection in rangeList)
                {
                    allDataGridSectionList.Add(gridSection);
                }
            }
            allDataGridSectionList.Sort();

            // remove appendRange, templateRange from bottom to top
            //ExcelWorksheet template = this.excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Template");
            //ExcelWorksheet sheet1 = this.excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Sheet1");
            ExcelWorksheet template = this.excelPackage.Workbook.Worksheets["Template"];
            ExcelWorksheet sheet1 = this.excelPackage.Workbook.Worksheets["Sheet1"];
            foreach (ExcelDataGridSection dataGridSection in allDataGridSectionList.Reverse<ExcelDataGridSection>())
            {
                int deleteAppendRange = dataGridSection.AppendToRow - dataGridSection.AppendFromRow + 1;
                int deleteTemplateRange = dataGridSection.TemplateToRow - dataGridSection.TemplateFromRow + 1;
                sheet1.DeleteRow(dataGridSection.AppendFromRow, deleteAppendRange);
                sheet1.DeleteRow(dataGridSection.TemplateFromRow, deleteTemplateRange);
            }

            //_excelPackage.Workbook.Worksheets.Delete(template);

            // remove column A
            sheet1.DeleteColumn(1);
            //sheet1.Column(1).Hidden = true;
        }

        public virtual void RemoveTemplateRowsForPdf(ExcelPackage _excelPackage)
        {
            List<ExcelDataGridSection> allDataGridSectionList = new List<ExcelDataGridSection>();
            // find all data grid section
            List<ExcelDataGrid> _excelDataGridList = this.reportEntity.GetDataGrid();
            foreach (ExcelDataGrid dataGrid in _excelDataGridList)
            {
                List<ExcelDataGridSection> rangeList = dataGrid.GetRangeList();
                foreach (ExcelDataGridSection gridSection in rangeList)
                {
                    allDataGridSectionList.Add(gridSection);
                }
            }
            allDataGridSectionList.Sort();

            // remove appendRange, templateRange from bottom to top
            //ExcelWorksheet template = this.excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Template");
            //ExcelWorksheet sheet1 = this.excelPackage.Workbook.Worksheets.First(worksheet => worksheet.Name == "Sheet1");
            ExcelWorksheet template = this.excelPackage.Workbook.Worksheets["Template"];
            ExcelWorksheet sheet1 = this.excelPackage.Workbook.Worksheets["Sheet1"];
            foreach (ExcelDataGridSection dataGridSection in allDataGridSectionList.Reverse<ExcelDataGridSection>())
            {
                int deleteAppendRange = dataGridSection.AppendToRow - dataGridSection.AppendFromRow + 1;
                int deleteTemplateRange = dataGridSection.TemplateToRow - dataGridSection.TemplateFromRow + 1;
                sheet1.DeleteRow(dataGridSection.AppendFromRow, deleteAppendRange);
                sheet1.DeleteRow(dataGridSection.TemplateFromRow, deleteTemplateRange);
            }

            template.View.TabSelected = false;
            //sheet1.View.SetTabSelected();

            //_excelPackage.Workbook.Worksheets.Delete(template);

            // remove column A
            //sheet1.DeleteColumn(1);
            sheet1.Column(1).Hidden = true;
        }

        public override void SaveAndDownloadAsBase64()
        {
            this.RefreshPrintDate();
        }

        public override void SaveFile()
        {
            this.RefreshPrintDate();
        }

        public virtual void SaveExcel(string _fileName = "")
        {
            this.SaveXlsx(_fileName);
        }

        public virtual void SaveRtf(string _fileName="")
        {

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            try
            {
            }
            catch (Exception ex)
            {
            }
        }

        public virtual void SaveTemplateAsXlsx(ExcelPackage _excelPackage, string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }
            string filePath = _fileName + ".xlsx";

            IDictionary<string, object> _dataSetObj = this.reportEntity.GetDataSetObj();
            DataSet _dataSet = this.reportEntity.GetDataSet();
            string _xlsxTemplateFilePath = this.reportEntity.GetXlsxTemplateFilePath();

            try
            {
                List<ExpandoObject> tupleExpandoObjectList = new List<ExpandoObject>();
                ExpandoObject expandoObject = new ExpandoObject();

                using (var package = new ExcelPackage(FileOutputUtil.GetFileInfo(_xlsxTemplateFilePath)))
                {
                    string tableName = string.Empty;
                    if (_dataSet != null)
                    {
                        foreach (DataTable _dataTable in _dataSet.Tables)
                        {
                            tableName = _dataTable.TableName;
                            if(tableName.ToLower().IndexOf("view") == -1) continue;

                            var sheet = package.Workbook.Worksheets.Add(tableName);
                            sheet.Cells["A1"].LoadFromDataTable(_dataTable, true, TableStyles.Medium9);
                        }
                    }
                    else if (_dataSetObj != null)
                    {
                        foreach (KeyValuePair<string, Object> _dataView in _dataSetObj)
                        {
                            tableName = _dataView.Key;
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

                            var sheet = package.Workbook.Worksheets.Add(tableName);
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
                    }

                    // SaveAs Method1
                    /*
                    //convert the excel package to a byte array
                    byte[] bin = package.GetAsByteArray();
                    //the path of the file
                    //write the file to the disk
                    File.WriteAllBytes(filePath, bin);
                    */

                    // SaveAs Method2
                    //Instead of converting to bytes, you could also use FileInfo
                    FileInfo fi = new FileInfo(filePath);
                    package.SaveAs(fi);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public virtual void SaveTemplateAsXlsxInMasterDataList(string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }
            string filePath = _fileName + ".xlsx";

            IDictionary<string, object> _dataSetObj = this.reportEntity.GetDataSetObj();
            DataSet _dataSet = this.reportEntity.GetDataSet();
            string _xlsxTemplateFilePath = this.reportEntity.GetXlsxTemplateFilePath();

            try
            {
                List<ExpandoObject> tupleExpandoObjectList = new List<ExpandoObject>();
                ExpandoObject expandoObject = new ExpandoObject();

                using (var package = new ExcelPackage(FileOutputUtil.GetFileInfo(_xlsxTemplateFilePath)))
                {
                    string tableName = string.Empty;
                    if (_dataSet != null)
                    {
                        foreach (DataTable _dataTable in _dataSet.Tables)
                        {
                            tableName = _dataTable.TableName;
                            if (tableName.ToLower().IndexOf("view") == -1) continue;

                            var sheet = package.Workbook.Worksheets.Add(tableName);
                            sheet.Cells["A1"].LoadFromDataTable(_dataTable, true, TableStyles.Medium9);
                        }
                    }
                    else if (_dataSetObj != null)
                    {
                        foreach (KeyValuePair<string, Object> _dataView in _dataSetObj)
                        {
                            tableName = _dataView.Key;
                            if (tableName.ToLower().IndexOf("view") == -1) continue;

                            tupleExpandoObjectList = new List<ExpandoObject>();
                            //tupleObjList = (List<Object>)_dataSetObj["GeneralView"];
                            foreach (object tuple in (List<object>)_dataView.Value)
                            {
                                expandoObject = new ExpandoObject();
                                foreach (var property in tuple.GetType().GetProperties())
                                {
                                    ((IDictionary<string, object>)expandoObject).Add(property.Name, property.GetValue(tuple));
                                }
                                tupleExpandoObjectList.Add(expandoObject);
                            }

                            var sheet = package.Workbook.Worksheets.Add(tableName);
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
                    }

                    // SaveAs Method1
                    /*
                    //convert the excel package to a byte array
                    byte[] bin = package.GetAsByteArray();
                    //the path of the file
                    //write the file to the disk
                    File.WriteAllBytes(filePath, bin);
                    */

                    // SaveAs Method2
                    //Instead of converting to bytes, you could also use FileInfo
                    FileInfo fi = new FileInfo(filePath);
                    package.SaveAs(fi);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public virtual void SaveXlsxInMasterDataList(string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }
            IDictionary<string, object> _dataSetObj = this.reportEntity.GetDataSetObj();

            try
            {
                List<ExpandoObject> tupleExpandoObjectList = new List<ExpandoObject>();

                List<Object> tupleObjList = (List<Object>)_dataSetObj["GeneralView"];

                ExpandoObject expandoObject = new ExpandoObject();

                using (var package = new ExcelPackage(FileOutputUtil.GetFileInfo(_fileName + ".xlsx")))
                {
                    foreach (KeyValuePair<string, Object> _dataView in _dataSetObj)
                    {
                        string tableName = _dataView.Key;
                        if (tableName.ToLower().IndexOf("view")==-1) continue;

                        tupleExpandoObjectList = new List<ExpandoObject>();
                        //tupleObjList = (List<Object>)_dataSetObj["GeneralView"];
                        foreach (object tuple in (List<object>)_dataView.Value)
                        {
                            expandoObject = new ExpandoObject();
                            foreach (var property in tuple.GetType().GetProperties())
                            {
                                ((IDictionary<string, object>)expandoObject).Add(property.Name, property.GetValue(tuple));
                            }
                            tupleExpandoObjectList.Add(expandoObject);
                        }

                        var sheet = package.Workbook.Worksheets.Add(tableName);
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
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public virtual void SaveXlsx(string _fileName = "")
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
            }
        }

        public virtual void SaveXls(string _fileName = "")
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
            }
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