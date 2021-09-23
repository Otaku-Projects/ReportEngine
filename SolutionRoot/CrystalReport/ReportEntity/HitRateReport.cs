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
    public class HitRateReport : BaseReportEntity
    {

        public HitRateReport(DataSet _dataSet)
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateReport");

            string _rptPath = string.Empty;
            _rptPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"HitRateReport");
            _rptPath = Path.Combine(this.rptFilesFolder, "HitRateReport.rpt");

            //this.rptDocument = new ReportDocument();
            this.rptDocument = new HitRateTemplate();
            this.rptDocument.Load(_rptPath);
            //this.hitRate.Load(ReportTemplate)

            //hitRate.Database.Tables["GeneralView"].SetDataSource(_dataSet);
            this.rptDocument.SetDataSource(_dataSet);

            // Create Crystl Report entity
            //this.crystalReportEntity = new CrystalReportEntity(this.rptDocument, _dataSet);
        }

    }
}
