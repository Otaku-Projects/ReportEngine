using CoreReport.Puppeteer;
using CoreSystemConsole.ReportDataModel;
using IronPDFProject.ReportEntity;
using IronPDFProject.ReportRender;
using Puppeteer.ReportEntity;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSystemConsole.ProgramEntity
{
    public class IronPdfTemplateProgram
    {
        public IronPdfTemplateProgram()
        {
            Console.WriteLine("Said \"Hello World!\" from IronPdfTemplateProgram");

            UrlToPdfReportDecorator urlToPdfReportDecorator = null;
            TestUrlToPdfReport testUrlToPdfReport = new TestUrlToPdfReport();
            urlToPdfReportDecorator = new UrlToPdfReportDecorator(testUrlToPdfReport, "url.pdf");
            urlToPdfReportDecorator.SaveFile();
        }
    }
}
