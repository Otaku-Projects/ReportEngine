using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;


namespace OpenXmlSDK.ReportEntity
{
    public enum TupleAppendDirection
    {
        None = 0,
        FixedNotAppend = 11,

        FromTopToBottom = 21,
        FromLeftToRight = 22
    }

    public class ExcelDataSection : Object, IEquatable<ExcelDataSection>, IComparable<ExcelDataSection>
    {
        private string indicator;

        private string templateRange;
        private string appendToRange;

        private TupleAppendDirection appendDirection;
        public TupleAppendDirection AppendDirection { get => appendDirection; set => appendDirection = value; }

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

        protected List<ExcelDataSection> historicalCoordinate;
        public List<ExcelDataSection> HistoricalCoordinate { get => historicalCoordinate; }

        public ExcelDataSection()
        {
            this.appendDirection = TupleAppendDirection.FixedNotAppend;
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
        public ExcelDataSection(ExcelDataSection _excelDataSection)
        {
            this.AppendDirection = _excelDataSection.AppendDirection;
            this.Indicator = _excelDataSection.Indicator;

            this.SetTemplateRange(_excelDataSection.GetTemplateRange());
            this.SetAppendToRange(_excelDataSection.GetAppendRange());
            this.ExcelDataGrid = _excelDataSection.ExcelDataGrid;
        }

        public ExcelDataSection(string _indicator, string _templateRange, string _appendToRange) : this()
        {
            this.AppendDirection = TupleAppendDirection.FixedNotAppend;

            if (!string.IsNullOrEmpty(_indicator))
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

            // 20211108, make an assumption
            // if append range in a row, append direction will be from top to bottom
            if (_fromRow.Equals(_toRow)) this.AppendDirection = TupleAppendDirection.FromTopToBottom;
            // if append range in a column, append direction will be from left to right
            if (_fromCol.Equals(_toCol)) this.AppendDirection = TupleAppendDirection.FromLeftToRight;

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
        public int CompareTo(ExcelDataSection compareSection)
        {
            // A null value means that this object is greater.
            if (compareSection == null)
                return 1;

            else
                return this.templateFromRow.CompareTo(compareSection.templateFromRow);
        }

        public bool Equals(ExcelDataSection other)
        {
            return other != null &&
                other.Indicator == this.Indicator &&
                other.GetTemplateRange() == this.GetTemplateRange() &&
                other.GetAppendRange() == this.GetAppendRange();
        }

        //public object Clone()
        //{
        //    ExcelDataSection _clonedSection = new ExcelDataSection(
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

        public virtual ExcelDataSection Clone()
        {
            ExcelDataSection _clonedSection = new ExcelDataSection(
                this.indicator, this.templateRange, this.appendToRange
                );

            return _clonedSection;
        }
    }
}
