using CoreReport;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace ITextGroupNV.ReportEntity
{
    public abstract class ITextReportEntity : BaseReportEntity
    {
        private List<ExcelDataGrid> dataGridList;
        private List<ExcelDataGrid> dataGridTemplateBackupList;

        public ITextReportEntity()
        {
            this.templateBaseDirectory = @"D:\Documents\ReportEngine\SolutionRoot\JasperReport\ReportTemplate";
            // this return the start up project directory
            // e.g: "D:\\Documents\\CoreSystem\\WebApi\\" + \ReportTemplate
            this.templateBaseDirectory = Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate");
            // this return the program running directory
            // e.g: "D:\\Documents\\CoreSystem\\WebApi\\bin\\Debug\\net5.0" + \ReportTemplate
            this.templateBaseDirectory = Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ReportTemplate");

            this.dataGridList = new List<ExcelDataGrid>();
            this.dataGridTemplateBackupList = new List<ExcelDataGrid>();
        }

        protected virtual void AddDataGrid(ExcelDataGrid _dataGrid)
        {
            if (_dataGrid.IsValidAddToDataGridList())
            {
                this.dataGridList.Add(_dataGrid);
            }
        }

        public virtual List<ExcelDataGrid> GetDataGrid()
        {
            return this.dataGridList;
        }

        public virtual void SetDataGrid(List<ExcelDataGrid> _dataGridList)
        {
            this.dataGridList = _dataGridList;
        }

        public virtual void BackupDataGridSetting()
        {
            this.dataGridTemplateBackupList = new List<ExcelDataGrid>();
            this.dataGridList.ForEach((item) =>
            {
                this.dataGridTemplateBackupList.Add(new ExcelDataGrid(item));
            });
        }
        public virtual List<ExcelDataGrid> GetBackupTemplateDataGrid()
        {
            return this.dataGridTemplateBackupList;
        }
    }

}
