using CoreReport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IronPDFProject.ReportEntity
{
    public abstract class IronPdfReportEntity : BaseReportEntity
    {
        public IronPdfReportEntity()
        {
            this.templateBaseDirectory = @"D:\Documents\ReportEngine\SolutionRoot\IronPDFProject\ReportTemplate";
            // this return the start up project directory
            // e.g: "D:\\Documents\\CoreSystem\\WebApi\\" + \ReportTemplate
            this.templateBaseDirectory = Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate");
            // this return the program running directory
            // e.g: "D:\\Documents\\CoreSystem\\WebApi\\bin\\Debug\\net5.0" + \ReportTemplate
            this.templateBaseDirectory = Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "ReportTemplate");
        }
    }
}
