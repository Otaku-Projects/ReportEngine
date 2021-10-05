using CoreReport.CrystalReport;
using CrystalDecisions.CrystalReports.Engine;
using CrystalReport.ReportEntity;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSystemConsoleInNet.ProgramEntity
{
    public class HitRateProgram
    {
        private CrystalReportDecorator crystalReportEntity;
        private ReportDocument hitRate;

        public HitRateProgram()
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateProgram");

            //this.hitRate = new ReportDocument();
            //this.hitRate.Load(ReportTemplate)

            HitRateDataView hitRateDataView = new HitRateDataView();
            DataSet dataSet = hitRateDataView.GetDataSet();

            HitRateReport hitRateReport = new HitRateReport(dataSet);
            this.hitRate = hitRateReport.GetReportDocument();

            HitRateReport2 hitRateReport2 = new HitRateReport2(dataSet);
            ReportDocument hitRate2 = hitRateReport2.GetReportDocument();
            CrystalReportDecorator crystalReportEntity2 = new CrystalReportDecorator(hitRate2);

            this.crystalReportEntity = new CrystalReportDecorator(this.hitRate);

            this.crystalReportEntity.SavePdf();
            this.crystalReportEntity.SaveXlsx();
            //this.crystalReportEntity.SaveRtf();

            crystalReportEntity2.SavePdf();
            crystalReportEntity2.SaveXlsx();
            //crystalReportEntity2.SaveRtf();
        }

        public Boolean Save()
        {
            crystalReportEntity.SaveFile();

            return true;
        }

        public Boolean SaveAndDownloadAsBase64()
        {
            crystalReportEntity.SaveAndDownloadAsBase64();

            return true;
        }
    }
}
