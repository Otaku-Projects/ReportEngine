﻿using CoreReport.EPPlus5Report;
using CoreSystemConsole.ReportDataModel;
using EPPlus5Report.ReportEntity;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSystemConsole.ProgramEntity
{
    public class EPPlus5XlsxTemplateProgram
    {
        public EPPlus5XlsxTemplateProgram()
        {
            Console.WriteLine("Said \"Hello World!\" from EPPlus5XlsxTemplateProgram");

            HitRateDataView hitRateDataView1 = new HitRateDataView();
            HitRateDataView hitRateDataView2 = new HitRateDataView();
            HitRateDataView hitRateDataView3 = new HitRateDataView();

            //hitRateDataView1.CreateDummyData1();
            //IDictionary<string, object> dataSetObj1 = hitRateDataView1.GetDataSetObj();
            hitRateDataView2.CreateDummyData2();
            IDictionary<string, object> dataSetObj2 = hitRateDataView2.GetDataSetObj();
            hitRateDataView3.CreateDummyData3();
            IDictionary<string, object> dataSetObj3 = hitRateDataView3.GetDataSetObj();
            DataSet dataSet3 = hitRateDataView3.GetDataSet();

            EPPlus5Decorator epplus5Decorator = null;
            HitRateReportDecorator hitRateRptDecorator = null;

            HitRateReport5 hitRateReport5 = new HitRateReport5(dataSet3);
            hitRateRptDecorator = new HitRateReportDecorator(hitRateReport5);
            hitRateRptDecorator.RenderTemplateAndSaveAsXlsx();

            HitRateReport5 hitRateReport5pdf = new HitRateReport5(dataSet3);
            HitRateReportDecorator hitRateRptDecorator2 = null;
            hitRateRptDecorator2 = new HitRateReportDecorator(hitRateReport5pdf);
            hitRateRptDecorator2.RenderTemplateAndSaveAsPdf();

            //HitRateReport6 hitRateReport6 = new HitRateReport6(dataSet3);
            //HitRateReportDecorator hitRateRptDecorator3 = null;
            //hitRateRptDecorator3 = new HitRateReportDecorator(hitRateReport6);
            //hitRateRptDecorator3.RenderTemplateAndSaveAsXlsx();
        }
    }
}
