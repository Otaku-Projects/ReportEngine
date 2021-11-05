
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;

namespace Puppeteer.ReportEntity
{
    public class ReportReference2 : PuppeteerReportEntity
    {
        public ReportReference2(DataSet _dataSet)
        {
            Console.WriteLine("Said \"Hello World!\" from ReportReference2");
            //this.dataSet = _dataSet;
            this.dataSet = _dataSet;
        }

        public ReportReference2(IDictionary<string, object> _dataSetObj)
        {
            Console.WriteLine("Said \"Hello World!\" from ReportReference2");
            //this.dataSet = _dataSet;
            this.dataSetObj = _dataSetObj;
        }
        public override void InitializateMetaData()
        {
        }
        public override void InitializateMainContent()
        {
        }

        public override void InitializateDataGrid()
        {
        }

        public override void InitializateHeaderFooter()
        {
        }
    }
}
