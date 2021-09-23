using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using jsreport.Binary;
using jsreport.Local;
using jsreport.Shared;
using jsreport.Types;

namespace JasperReport.ReportEntity
{
    public class HitRateReport : BaseReportEntity
    {

        public HitRateReport(DataSet _dataSet) { }

        public HitRateReport(IDictionary<string, object> _dataSetObj)
        {
            Console.WriteLine("Said \"Hello World!\" from HitRateReport");
            //this.dataSet = _dataSet;
            this.dataSetObj = _dataSetObj;

            string _templateDirectory = string.Empty;
            string _templateFileLocation = string.Empty;
            _templateDirectory = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"HitRateReport");
            _templateDirectory = Path.Combine(this.templateBaseDirectory, @"HitRateReport");

            if (File.Exists(Path.Combine(_templateDirectory,@"index.html")))
            {
                _templateFileLocation = Path.Combine(_templateDirectory, @"index.html");
            }else if (File.Exists(Path.Combine(_templateDirectory, @"index.htm")))
            {
                _templateFileLocation = Path.Combine(_templateDirectory, @"index.htm");
            }

            this.templateReportFileDirectory = _templateDirectory;
            this.templateReportFileLocation = _templateFileLocation;
        }

    }
}
