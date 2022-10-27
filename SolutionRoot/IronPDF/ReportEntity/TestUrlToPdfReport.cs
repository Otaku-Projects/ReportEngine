using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IronPDFProject.ReportEntity
{
    public class TestUrlToPdfReport : UrlToPdfReportEntity
    {
        public TestUrlToPdfReport()
        {
            this.AddUrlPath("https://ironpdf.com/");
            this.AddUrlPath("https://www.google.com/");
            this.AddUrlPath("https://www.hktvmall.com/hktv/zh/search_a/?keyword=%E8%83%BD%E9%87%8F%E6%9E%9C%E5%87%8D&category=AA73600000000&page=0");
        }
    }
}
