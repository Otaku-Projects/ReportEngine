using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyOpenXml
{
    public class ReportFileTuple
    {
        public string fileName { get; set; }
        public Byte[] fileByte { get; set; }
        public ReportFileTuple()
        {
            this.fileName = string.Empty;
            this.fileByte = new Byte[0];
        }
        public ReportFileTuple(string fileName, Byte[] fileByte)
        {
            this.fileName = fileName;
            this.fileByte = fileByte;
        }
    }
}
