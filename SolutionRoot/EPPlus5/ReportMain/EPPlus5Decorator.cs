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
        public virtual ExcelPackage RenderDataAndMergeToTemplate()
        {
            return this.GetXlsxTemplateInstance();
        }
        protected virtual void MergeDataRows(ExcelWorksheet _worksheet, string _indicator, List<Object> _tupleList)
        {
            foreach (Object _tuple in _tupleList)
            {
                this.MergeDataRow(_worksheet, _indicator, _tuple);
            }
        }

        protected virtual void MergeDataRow(ExcelWorksheet _worksheet, string _indicator, Object _tuple)
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

            // 4.1 update the new append to Range(newAppendToRange) to DataGridSection
            // no need to update the current data grid section
            /*
                e.g. new ExcelDataGridSection("T1B", "17:19", "20:20")
                3 rows 17,18,19 will insert at row 20
                the new rows are 20, 21, 22
                and the append to row 20 will shifted to 23
             */
            // the excel will get a lot of empty rows if update the appendToRange
            affectSection.SetAppendToRange(newAppendToRange);
            // but need to update other data grid section which lower that it
            /*
                e.g. new ExcelDataGridSection("T1B", "17:19", "20:20")
                e.g. new ExcelDataGridSection("T1F", "21:21", "22:22")
                
                when render T1B, the T1F will shifted down, need to update the address
             */
            foreach (ExcelDataGrid dataGrid in excelDataGridList)
            {
                ExcelDataGridSection _headerSection = dataGrid.GetHeaderRange();
                ExcelDataGridSection _bodySection = dataGrid.GetBodyRange();
                ExcelDataGridSection _footerSection = dataGrid.GetFooterRange();
                // skip appending section
                if (_headerSection.Indicator != affectSection.Indicator)
                {

                }
                if (_bodySection.Indicator != affectSection.Indicator)
                {

                }
                if (_footerSection.Indicator != affectSection.Indicator)
                {

                }
            }

                /*
                // 10.1 insert and copy value, style
                for (int start = templateStartRowIndex; start < templateEndRowIndex; start++)
                {
                    _worksheet.InsertRow(appendStartRowIndex + (start - templateStartRowIndex), 1, start);
                    fromRange = templateStartColLetter + start + ":" + templateEndColLetter + start;
                    destinationRange = affectSection.GetAppendRange();
                    _worksheet.Cells[fromRange].Copy(_worksheet.Cells[destinationRange]);

                    newAppendToRange = templateStartColLetter + (appendEndRowIndex + 1) + ":" + templateEndColLetter + (appendEndRowIndex + 1);
                    affectSection.SetAppendToRange(newAppendToRange);
                }
                */

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
            for (int colIndex = colIndexStart; colIndex < colIndexEnd; colIndex++)
            {
                for (int rowIndex = appendStartRowIndex; rowIndex <= (appendStartRowIndex+ dataGridRowCount-1); rowIndex++)
                {
                    this.DefaultMergeCellExpression(_worksheet, OfficeOpenXml.ExcelCellAddress.GetColumnLetter(colIndex) + rowIndex, _tuple);
                    this.CustomPostMergeCellExpression(_worksheet, OfficeOpenXml.ExcelCellAddress.GetColumnLetter(colIndex) + rowIndex, _tuple);
                }
            }
        }
        protected virtual void DefaultMergeCellExpression(ExcelWorksheet _worksheet, string cellAddress, Object _tuple)
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
            foreach (PropertyInfo propertyInfo in _tuple.GetType().GetProperties())
            {
                matchExpression = "{{" + propertyInfo.Name + "}}";
                if (cellVal.IndexOf(matchExpression) > -1)
                {
                    isMerge = true;
                    mergedValue = mergedValue.Replace(matchExpression, propertyInfo.GetValue(_tuple).ToString());
                }

                // do stuff here
                //propertyInfo.GetValue(_tuple, null)
            }
            // 2.2 
            if (isMerge)
            {
                ExcelStyle cellStyle = _cell.Style;
                string _cellFormat = _cell.Style.Numberformat.Format;
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

        protected virtual void CustomPostMergeCellExpression(ExcelWorksheet _worksheet, string cellAddress, Object _tuple)
        {
            throw new NotImplementedException();
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
            string _xlsxTemplateFilePath = this.reportEntity.GetXlsxTemplateFilePath();

            try
            {
                List<ExpandoObject> tupleExpandoObjectList = new List<ExpandoObject>();
                ExpandoObject expandoObject = new ExpandoObject();

                //List<Object> tupleObjList = (List<Object>)_dataSetObj["GeneralView"];
                //foreach (object tuple in tupleObjList)
                //{
                //    expandoObject = new ExpandoObject();
                //    foreach (var property in tuple.GetType().GetProperties())
                //    {
                //        ((IDictionary<string, object>)expandoObject).Add(property.Name, property.GetValue(tuple));
                //    }
                //    tupleExpandoObjectList.Add(expandoObject);
                //}

                //FileInfo fi = new FileInfo(filePath);
                //_excelPackage.SaveAs(fi);

                using (var package = new ExcelPackage(FileOutputUtil.GetFileInfo(_xlsxTemplateFilePath)))
                {
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
            string _xlsxTemplateFilePath = this.reportEntity.GetXlsxTemplateFilePath();

            try
            {
                List<ExpandoObject> tupleExpandoObjectList = new List<ExpandoObject>();

                List<Object> tupleObjList = (List<Object>)_dataSetObj["GeneralView"];

                ExpandoObject expandoObject = new ExpandoObject();

                using (var package = new ExcelPackage(FileOutputUtil.GetFileInfo(_xlsxTemplateFilePath)))
                {
                    foreach (KeyValuePair<string, Object> _dataView in _dataSetObj)
                    {
                        string tableName = _dataView.Key;
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