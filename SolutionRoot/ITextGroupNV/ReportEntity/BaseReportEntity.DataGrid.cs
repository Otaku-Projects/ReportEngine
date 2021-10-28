using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace ITextGroupNV.ReportEntity
{
    public class ExcelDataGridSection : Object, IEquatable<ExcelDataGridSection>, IComparable<ExcelDataGridSection>
    {
        private string indicator;
        private string templateRange;
        private string appendToRange;

        private int templateFromRow;
        private string templateFromCol;
        private int templateToRow;
        private string templateToCol;
        private int appendFromRow;
        private string appendFromCol;
        private int appendToRow;
        private string appendToCol;
        private ExcelDataGrid excelDataGrid;
        public ExcelDataGrid ExcelDataGrid { get => excelDataGrid; set => excelDataGrid = value; }

        public string Indicator { get => indicator; set => indicator = value; }
        public int TemplateFromRow { get => templateFromRow; set => templateFromRow = value; }
        public string TemplateFromCol { get => templateFromCol; set => templateFromCol = value; }
        public int TemplateToRow { get => templateToRow; set => templateToRow = value; }
        public string TemplateToCol { get => templateToCol; set => templateToCol = value; }
        public int AppendFromRow { get => appendFromRow; set => appendFromRow = value; }
        public string AppendFromCol { get => appendFromCol; set => appendFromCol = value; }
        public int AppendToRow { get => appendToRow; set => appendToRow = value; }
        public string AppendToCol { get => appendToCol; set => appendToCol = value; }

        protected List<ExcelDataGridSection> historicalCoordinate;
        public List<ExcelDataGridSection> HistoricalCoordinate { get => historicalCoordinate; }

        public ExcelDataGridSection()
        {
            this.indicator = string.Empty;
            this.templateRange = string.Empty;
            this.appendToRange = string.Empty;
            this.excelDataGrid = null;


            this.TemplateFromRow = -1;
            this.TemplateFromCol = string.Empty;
            this.TemplateToRow = -1;
            this.TemplateToCol = string.Empty;

            this.AppendFromRow = -1;
            this.AppendFromCol = string.Empty;
            this.AppendToRow = -1;
            this.AppendToCol = string.Empty;
        }

        //public ExcelDataGridSection(string _indicator, string startColLetter, int startRowIndex, string endColLetter, int endRowIndex) : this()
        //{
        //    this.indicator = _indicator;
        //    this.SetTemplateRange(startColLetter + startRowIndex + ":" + endColLetter + endRowIndex);
        //}
        
        public ExcelDataGridSection(string _indicator, string _templateRange, string _appendToRange) : this()
        {
            if(!string.IsNullOrEmpty(_indicator))
                this.indicator = _indicator;
            if (!string.IsNullOrEmpty(_templateRange))
                this.SetTemplateRange(_templateRange);

            if (!string.IsNullOrEmpty(_appendToRange))
                this.SetAppendToRange(_appendToRange);
        }
        public virtual Boolean IsRangeEmpty()
        {
            return string.IsNullOrEmpty(this.GetTemplateRange()) || string.IsNullOrEmpty(this.GetAppendRange());
        }

        public virtual void SetAppendToRange(string _appendToRange)
        {
            string _fromRow = string.Empty;
            string _fromCol = string.Empty;
            string _toRow = string.Empty;
            string _toCol = string.Empty;
            _appendToRange = _appendToRange.ToUpper();
            if (_appendToRange.IndexOf(":") == -1)
            {
                throw new Exception($"Extracting error on append range '{_appendToRange}', missing ':', please use: 17:19, 20:20, A19:F19 or A19:F21");
            }

            string[] ranges = _appendToRange.Split(':');

            // remove all numeric
            //_fromCol = Regex.Replace(ranges[0], @"[^A-Z]+", String.Empty);
            _fromCol = new string(ranges[0].Where(c => char.IsLetter(c)).ToArray());
            //_fromRow = ranges[0].Replace(_fromCol, String.Empty);
            _fromRow = new string(ranges[0].Where(c => char.IsDigit(c)).ToArray());

            // remove all numeric
            _toCol = new string(ranges[1].Where(c => char.IsLetter(c)).ToArray());
            _toRow = new string(ranges[1].Where(c => char.IsDigit(c)).ToArray());

            if (
                string.IsNullOrEmpty(_fromRow) && string.IsNullOrEmpty(_toRow))
            {
                throw new Exception($"Extracting error on append range '{_appendToRange}'  ");
            }
            if (string.IsNullOrEmpty(_fromCol) != string.IsNullOrEmpty(_toCol))
            {
                throw new Exception($"Extracting error on append range '{_appendToRange}'  ");
            }

            this.AppendFromRow = Int32.Parse(_fromRow);
            this.AppendFromCol = _fromCol;
            this.AppendToRow = Int32.Parse(_toRow);
            this.AppendToCol = _toCol;

            this.appendToRange = _appendToRange;
        }

        public virtual void SetTemplateRange(string _templateRange)
        {
            string _fromRow = string.Empty;
            string _fromCol = string.Empty;
            string _toRow = string.Empty;
            string _toCol = string.Empty;
            _templateRange = _templateRange.ToUpper();
            if (_templateRange.IndexOf(":") == -1)
            {
                throw new Exception($"Extracting error on template range '{_templateRange}', missing ':', please use: 17:19, 20:20, A19:F19 or A19:F21");
            }

            string[] ranges = _templateRange.Split(':');

            // remove all numeric
            //_fromCol = Regex.Replace(ranges[0], @"[^A-Z]+", String.Empty);
            _fromCol = new string(ranges[0].Where(c => char.IsLetter(c)).ToArray());
            //_fromRow = ranges[0].Replace(_fromCol, String.Empty);
            _fromRow = new string(ranges[0].Where(c => char.IsDigit(c)).ToArray());

            // remove all numeric
            _toCol = new string(ranges[1].Where(c => char.IsLetter(c)).ToArray());
            _toRow = new string(ranges[1].Where(c => char.IsDigit(c)).ToArray());

            if (
                string.IsNullOrEmpty(_fromRow) && string.IsNullOrEmpty(_toRow))
            {
                throw new Exception($"Extracting error on template range '{_templateRange}'  ");
            }
            if (string.IsNullOrEmpty(_fromCol) != string.IsNullOrEmpty(_toCol))
            {
                throw new Exception($"Extracting error on template range '{_templateRange}'  ");
            }

            this.TemplateFromRow = Int32.Parse(_fromRow);
            this.TemplateFromCol = _fromCol;
            this.TemplateToRow = Int32.Parse(_toRow);
            this.TemplateToCol = _toCol;

            this.templateRange = _templateRange;
        }

        protected virtual void RefreshExcelRange()
        {

        }

        public virtual string GetTemplateRange()
        {
            return this.templateRange;
        }

        public virtual string GetAppendRange()
        {
            return this.appendToRange;
        }
        public void CopyTemplateRangeToNewLocation(string toRange)
        {

        }
        public void CopyAppendRangeToNewLocation(string toRange)
        {

        }

        // Default comparer for Part type.
        public int CompareTo(ExcelDataGridSection compareSection)
        {
            // A null value means that this object is greater.
            if (compareSection == null)
                return 1;

            else
                return this.templateFromRow.CompareTo(compareSection.templateFromRow);
        }

        public bool Equals(ExcelDataGridSection other)
        {
            return other != null &&
                other.Indicator == this.Indicator &&
                other.GetTemplateRange() == this.GetTemplateRange() &&
                other.GetAppendRange() == this.GetAppendRange();
        }

        //public object Clone()
        //{
        //    ExcelDataGridSection _clonedSection = new ExcelDataGridSection(
        //        this.indicator, this.templateRange, this.appendToRange
        //        );

        //    return _clonedSection;
        //}

        public virtual Boolean Empty()
        {
            return this == null
                || string.IsNullOrEmpty(this.indicator)
                || string.IsNullOrEmpty(this.templateRange)
                || string.IsNullOrEmpty(this.appendToRange);
        }

        public virtual ExcelDataGridSection Clone()
        {
            ExcelDataGridSection _clonedSection = new ExcelDataGridSection(
                this.indicator, this.templateRange, this.appendToRange
                );

            return _clonedSection;
        }
    }
    public class ExcelDataGrid : Object, IEquatable<ExcelDataGrid>
    {
        private string spreadsheetName;

        private string coordinateLeftTop;
        private string coordinateRightBottom;
        private ExcelDataGridSection headerRange;
        private ExcelDataGridSection bodyRange;
        private ExcelDataGridSection footerRange;

        private List<ExcelDataGridSection> rangeList;
        public string SpreadsheetName { get => spreadsheetName; set => spreadsheetName = value; }

        public ExcelDataGrid()
        {
            this.spreadsheetName = string.Empty;

            this.headerRange = new ExcelDataGridSection();
            this.bodyRange = new ExcelDataGridSection();
            this.footerRange = new ExcelDataGridSection();
            this.rangeList = new List<ExcelDataGridSection>();

            this.coordinateLeftTop = string.Empty;
            this.coordinateRightBottom = string.Empty;
        }
        public ExcelDataGrid(ExcelDataGrid _grid)
        {
            this.spreadsheetName = _grid.spreadsheetName;

            this.headerRange = _grid.GetHeaderRange().Clone();
            this.bodyRange = _grid.GetBodyRange().Clone();
            this.footerRange = _grid.GetFooterRange().Clone();

            this.RefreshRangeSequence();
        }
        public ExcelDataGrid(string _spreadSheetName)
        {
            this.spreadsheetName = _spreadSheetName;

            this.headerRange = new ExcelDataGridSection();
            this.bodyRange = new ExcelDataGridSection();
            this.footerRange = new ExcelDataGridSection();
        }


        public virtual void RefreshRangeSequence()
        {
            this.rangeList = new List<ExcelDataGridSection>();
            if (!this.headerRange.Empty()) this.rangeList.Add(this.headerRange);
            if (!this.bodyRange.Empty()) this.rangeList.Add(this.bodyRange);
            if (!this.footerRange.Empty()) this.rangeList.Add(this.footerRange);
        }

        public virtual void SetRange(ExcelDataGridSection _headerRange, ExcelDataGridSection _bodyRange, ExcelDataGridSection _footerRange)
        {
            _headerRange.ExcelDataGrid = this;
            _bodyRange.ExcelDataGrid = this;
            _footerRange.ExcelDataGrid = this;

            this.SetHeaderRange(_headerRange);
            this.SetBodyRange(_bodyRange);
            this.SetFooterRange(_footerRange);
        }

        /// <summary>
        /// Check is data grid contains at least one header/body/footer range, return false if no range is set
        /// </summary>
        /// <returns></returns>
        public Boolean IsValidAddToDataGridList()
        {
            Boolean isValid = true;
            if(this.headerRange.IsRangeEmpty() 
                && this.bodyRange.IsRangeEmpty()
                && this.footerRange.IsRangeEmpty())
                isValid = false;
            return isValid;
        }

        public void SetHeaderRange(ExcelDataGridSection _headerRange)
        {
            _headerRange.ExcelDataGrid = this;
            this.headerRange = _headerRange;

            this.RefreshRangeSequence();
        }
        public void SetBodyRange(ExcelDataGridSection _bodyRange)
        {
            _bodyRange.ExcelDataGrid = this;
            this.bodyRange = _bodyRange;

            this.RefreshRangeSequence();
        }
        public void SetFooterRange(ExcelDataGridSection _footerRange)
        {
            _footerRange.ExcelDataGrid = this;
            this.footerRange = _footerRange;

            this.RefreshRangeSequence();
        }

        public ExcelDataGridSection GetHeaderRange()
        {
            return this.headerRange;
        }
        public ExcelDataGridSection GetBodyRange()
        {
            return this.bodyRange;
        }
        public ExcelDataGridSection GetFooterRange()
        {
            return this.footerRange;
        }
        public List<ExcelDataGridSection> GetRangeList()
        {
            return this.rangeList;
        }
        public bool Equals(ExcelDataGrid other)
        {
            return other != null &&
                other.GetHeaderRange().Equals(this.GetHeaderRange()) &&
                other.GetBodyRange().Equals(this.GetBodyRange()) &&
                other.GetFooterRange().Equals(this.GetFooterRange());
        }

        public virtual ExcelDataGrid Clone()
        {
            ExcelDataGrid _clonedGrid = new ExcelDataGrid();
            _clonedGrid.SetHeaderRange(this.headerRange);
            _clonedGrid.SetBodyRange(this.bodyRange);
            _clonedGrid.SetFooterRange(this.footerRange);

            _clonedGrid.spreadsheetName = this.spreadsheetName;

            return _clonedGrid;
        }

        public enum TupleAppendDirection
        {
            None = 0,
            FixedNotAppend = 11,

            FromTopToBottom = 21,
            FromLeftToRight = 22
        }
    }
    
}
