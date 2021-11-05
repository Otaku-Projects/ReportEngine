using System;
using System.Collections.Generic;
using CoreReport;
using CoreSystemConsole.ProgramEntity;

namespace CoreSystemConsole
{
    class Program
    {
        public static void Main(string[] args)
        {
            Console.WriteLine("Said \"Hello World!\" from CoreSystemConsole");

            // Tick-off the Report Entity Program
            //InvoiceProgram invoiceProgram = new InvoiceProgram();

            //HitRateHTMLProgram hitRateHTMLProgram = new HitRateHTMLProgram();

            //HitRateXMLProgram hitRateXMLProgram = new HitRateXMLProgram();

            //EPPlus5XlsxTemplateProgram ePPlus5XlsxTemplateProgram = new EPPlus5XlsxTemplateProgram();

            //ITextGroupIPdfTemplateProgram iTextGroupIText5PdfTemplateProgram = new ITextGroupIPdfTemplateProgram();

            PuppeteerPdfTemplateProgram puppeteerPdfTemplateProgram = new PuppeteerPdfTemplateProgram();
        }
    }
}
