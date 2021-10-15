using CoreReport.JasperReport;
using CoreSystemConsole.ReportDataModel;
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
    public class HitRateXMLProgram
    {
        public HitRateXMLProgram()
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateXMLProgram");

            HitRateDataView hitRateDataView1 = new HitRateDataView();
            HitRateDataView hitRateDataView2 = new HitRateDataView();

            hitRateDataView1.CreateDummyData1();
            IDictionary<string, object> dataSetObj1 = hitRateDataView1.GetDataSetObj();

            hitRateDataView2.CreateDummyData2();
            IDictionary<string, object> dataSetObj2 = hitRateDataView2.GetDataSetObj();

            JasperReportDecorator jasperReportDecorator = null;

            HitRateReport3 hitRateReport3 = new HitRateReport3(dataSetObj2);
            jasperReportDecorator = new JasperReportDecorator(hitRateReport3);
            jasperReportDecorator.SaveXlsx();
        }
    }
}
