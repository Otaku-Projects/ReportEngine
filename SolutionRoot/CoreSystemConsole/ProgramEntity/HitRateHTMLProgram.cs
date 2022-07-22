using CoreReport.JasperReport;
using CoreSystemConsole.ReportDataModel;
using JasperReport.ReportEntity;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSystemConsole.ProgramEntity
{
    public class HitRateHTMLProgram
    {
        public HitRateHTMLProgram()
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateHTMLProgram");

            HitRateDataView hitRateDataView1 = new HitRateDataView();
            HitRateDataView hitRateDataView2 = new HitRateDataView();

            hitRateDataView1.CreateDummyData1();
            IDictionary<string, object> dataSetObj1 = hitRateDataView1.GetDataSetObj();

            hitRateDataView2.CreateDummyData2();
            IDictionary<string, object> dataSetObj2 = hitRateDataView2.GetDataSetObj();

            JasperReportDecorator jasperReportDecorator = null;

            HitRateReport1 hitRateReport1 = new HitRateReport1(dataSetObj1);
            jasperReportDecorator = new JasperReportDecorator(hitRateReport1);
            jasperReportDecorator.SaveXlsxByHTML();

            HitRateReport2 hitRateReport2 = new HitRateReport2(dataSetObj2);
            jasperReportDecorator = new JasperReportDecorator(hitRateReport2);
            jasperReportDecorator.SaveXlsxByHTML();
        }
    }
}
