using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EasyOpenXml
{
    public class ReportFileTuple : Tuple<string, Byte[]>
    {
        public ReportFileTuple(string one, Byte[] two)
            : base(one, two)
        {

        }

        public string Filename { get { return this.Item1; } }
        public Byte[] FileByte { get { return this.Item2; } }
    }
}
