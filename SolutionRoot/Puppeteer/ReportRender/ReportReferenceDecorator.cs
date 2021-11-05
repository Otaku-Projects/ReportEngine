using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreReport;
using System.Reflection;
using System.Drawing;
using System.Diagnostics;
using System.Net;
using Puppeteer.ReportEntity;

namespace CoreReport.Puppeteer
{
    public class ReportReferenceDecorator : PuppeteerDecorator
    {

        public ReportReferenceDecorator() : base()
        {
        }
        public ReportReferenceDecorator(PuppeteerReportEntity _reportEntity, string _filename = "") : base(_reportEntity, _filename = "")
        {
        }

    }
}