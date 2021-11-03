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
using ITextGroupNV.ReportEntity;
using iText.Html2pdf;

namespace CoreReport.ITextGroupNV
{
    public class FileListReportDecorator: IText7Decorator
    {

        public FileListReportDecorator() : base()
        {
        }
        public FileListReportDecorator(ITextReportEntity _reportEntity, string _filename = "") : base(_reportEntity, _filename = "")
        {
        }

    }
}