using CoreReport.JasperReport;
using JasperReport.ReportEntity;
using jsreport.Binary;
using jsreport.Local;
using jsreport.Shared;
using jsreport.Types;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSystemConsole.ProgramEntity
{
    public class InvoiceProgram
    {
        public InvoiceProgram()
        {
            Console.WriteLine("Said \"Hello World!\" from InvoiceProgram");

            HitRateDataView hitRateDataView1 = new HitRateDataView();
            HitRateDataView hitRateDataView2 = new HitRateDataView();

            hitRateDataView1.CreateDummyData1();
            IDictionary<string, object> dataSetObj1 = hitRateDataView1.GetDataSetObj();

            hitRateDataView2.CreateDummyData2();
            IDictionary<string, object> dataSetObj2 = hitRateDataView2.GetDataSetObj();

            JasperReportEntity jasperReportEntity = null;

            InvoiceReport1 hitRateReport1 = new InvoiceReport1(dataSetObj1);
            jasperReportEntity = new JasperReportEntity(hitRateReport1);
            jasperReportEntity.SavePdf();

            InvoiceReport2 hitRateReport2 = new InvoiceReport2(dataSetObj2);
            jasperReportEntity = new JasperReportEntity(hitRateReport2);
            jasperReportEntity.SavePdf();
        }
    }
}
