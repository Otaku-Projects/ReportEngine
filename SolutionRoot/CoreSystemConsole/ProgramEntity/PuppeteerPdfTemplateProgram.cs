using CoreReport.Puppeteer;
using CoreSystemConsole.ReportDataModel;
using Puppeteer.ReportEntity;
using jsreport.Binary;
using jsreport.Local;
using jsreport.Shared;
using jsreport.Types;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSystemConsole.ProgramEntity
{
    public class PuppeteerPdfTemplateProgram
    {
        public PuppeteerPdfTemplateProgram()
        {
            Console.WriteLine("Said \"Hello World!\" from ITextGroupIText5PdfTemplateProgram");

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

            FileListReportDecorator fileListReportDecorator = null;
            ReportReferenceDecorator reportReferenceDecorator = null;

            //PocFileList pocFileList = new PocFileList(dataSet3);
            PocFileList pocFileList = new PocFileList(dataSetObj2);
            fileListReportDecorator = new FileListReportDecorator(pocFileList);
            fileListReportDecorator.RenderTemplateAndSaveAsPdf();

            ReportReference2 reportReference2 = new ReportReference2(dataSetObj3);
            fileListReportDecorator = new FileListReportDecorator(reportReference2);
            fileListReportDecorator.RenderTemplateAndSaveAsPdf();

            //ReportReference1 reportReference1 = new ReportReference1(dataSetObj3);
            //reportReferenceDecorator = new ReportReferenceDecorator(reportReference1);
            //reportReferenceDecorator.RenderTemplateAndSaveAsPdf();
        }
    }
}
