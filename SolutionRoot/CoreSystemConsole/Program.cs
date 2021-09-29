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
            InvoiceProgram invoiceProgram = new InvoiceProgram();

            //HitRateProgram hitRateProgram = new HitRateProgram();
        }
    }
}
