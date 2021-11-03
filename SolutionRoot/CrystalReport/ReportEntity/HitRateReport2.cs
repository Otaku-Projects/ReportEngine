using CoreReport.CrystalReport;
using CrystalReport.ReportTemplate;
using CrystalDecisions.CrystalReports.Engine;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CrystalReport.ReportEntity
{
    public class HitRateReport2 : CrystalReportEntity
    {

        public HitRateReport2(DataSet _dataSet)
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateReport2");

            string _rptPath = string.Empty;
            //_rptPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"HitRateTemplateObjCollection");
            _rptPath = Path.Combine(this.rptFilesFolder, "HitRateTemplateObjCollection.rpt");

            this.rptDocument = new HitRateTemplateObjCollection();
            this.rptDocument.Load(_rptPath);

            //this.rptDocument.SetDataSource(_dataSet);
            if (_dataSet.Tables.Contains("GeneralView"))
            {
                this.rptDocument.Database.Tables[0].SetDataSource(_dataSet.Tables["GeneralView"]);
            }
            else
            {
                this.rptDocument.Database.Tables[0].SetDataSource(_dataSet.Tables[0]);
            }
        }

    }
}
