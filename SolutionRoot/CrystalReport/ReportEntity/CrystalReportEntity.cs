using CrystalDecisions.CrystalReports.Engine;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CrystalReport.ReportEntity
{
    public class CrystalReportEntity
    {
        protected ReportDocument rptDocument;
        protected string rptFilesFolder;

        public CrystalReportEntity()
        {
            this.rptFilesFolder = @"D:\Documents\ReportEngine\SolutionRoot\CrystalReport\ReportTemplate";
        }

        protected void SetReportDocument(ReportDocument _rptDoc)
        {
            this.rptDocument = _rptDoc;
        }

        public ReportDocument GetReportDocument()
        {
            return this.rptDocument;
        }
    }
}
