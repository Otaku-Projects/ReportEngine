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
    public class ExcelDataGrid : Object, IEquatable<ExcelDataGrid>
    {
        private string spreadsheetName;

        private string coordinateLeftTop;
        private string coordinateRightBottom;
        private ExcelDataSection dynamicRenderRange;
        private List<ExcelDataSection> staticRenderRange;

        private List<ExcelDataSection> rangeList;
        public string SpreadsheetName { get => spreadsheetName; set => spreadsheetName = value; }

        public ExcelDataGrid()
        {
            this.spreadsheetName = string.Empty;

            this.dynamicRenderRange = new ExcelDataSection();
            this.staticRenderRange = new List<ExcelDataSection>();
            this.rangeList = new List<ExcelDataSection>();

            this.coordinateLeftTop = string.Empty;
            this.coordinateRightBottom = string.Empty;
        }
        public ExcelDataGrid(ExcelDataGrid _grid)
        {
            this.spreadsheetName = _grid.spreadsheetName;

            this.dynamicRenderRange = _grid.GetDynamicRange().Clone();

            List<ExcelDataSection> oldList = _grid.GetStaticRange();
            List<ExcelDataSection> newList = _grid.GetStaticRange();
            this.staticRenderRange = new List<ExcelDataSection>(oldList.Count);
            oldList.ForEach((item) =>
            {
                newList.Add(new ExcelDataSection(item));
            });

            this.rangeList = newList;

            this.RefreshRangeSequence();
        }
        public ExcelDataGrid(string _spreadSheetName)
        {
            this.spreadsheetName = _spreadSheetName;

            this.dynamicRenderRange = new ExcelDataSection();
            this.staticRenderRange = new List<ExcelDataSection>();
            this.rangeList = new List<ExcelDataSection>();
        }


        public virtual void RefreshRangeSequence()
        {
            this.rangeList = new List<ExcelDataSection>();
            if (!this.dynamicRenderRange.Empty()) this.rangeList.Add(this.dynamicRenderRange);
            if (this.staticRenderRange.Count >0) this.rangeList.AddRange(this.staticRenderRange);
            //if (!this.rangeList.Empty()) this.rangeList.Add(this.footerRange);
        }

        public virtual void SetRange(ExcelDataSection _dynamicRange, List<ExcelDataSection> _staticRange)
        {
            this.SetDynamicRange(_dynamicRange);
            this.SetStaticRange(_staticRange);
        }

        /// <summary>
        /// Check is data grid contains at least one header/body/footer range, return false if no range is set
        /// </summary>
        /// <returns></returns>
        public Boolean IsValidAddToDataGridList()
        {
            Boolean isValid = true;
            if(this.GetDynamicRange().IsRangeEmpty() 
                && this.GetStaticRange().Count<=0)
                isValid = false;
            return isValid;
        }

        public void SetDynamicRange(ExcelDataSection _dynamicRange)
        {
            _dynamicRange.ExcelDataGrid = this;
            this.dynamicRenderRange = _dynamicRange;

            this.RefreshRangeSequence();
        }

        public void SetStaticRange(List<ExcelDataSection> _staticRange)
        {
            //_dynamicRange.ExcelDataGrid = this;
            if (_staticRange.Count <= 0) return;

            foreach(ExcelDataSection dataSection in _staticRange)
            {
                dataSection.ExcelDataGrid = this;
            }

            this.staticRenderRange = _staticRange;

            this.RefreshRangeSequence();
        }

        public ExcelDataSection GetDynamicRange()
        {
            return this.dynamicRenderRange;
        }
        public List<ExcelDataSection> GetStaticRange()
        {
            return this.staticRenderRange;
        }
        public List<ExcelDataSection> GetRangeList()
        {
            return this.rangeList;
        }
        public bool Equals(ExcelDataGrid otherGrid)
        {
            //return other != null &&
            //    other.GetDynamicRange().Equals(this.GetDynamicRange()) &&
            //    other.GetStaticRange().Equals(this.GetStaticRange()) &&
            //    other.GetFooterRange().Equals(this.GetFooterRange());

            return otherGrid != null &&
                otherGrid.GetRangeList().All(this.GetRangeList().Contains);
        }

        public virtual ExcelDataGrid Clone()
        {
            ExcelDataGrid _clonedGrid = new ExcelDataGrid();
            _clonedGrid.SetDynamicRange(this.GetDynamicRange());
            _clonedGrid.SetStaticRange(this.GetStaticRange());

            _clonedGrid.spreadsheetName = this.spreadsheetName;

            return _clonedGrid;
        }
    }
    
}
