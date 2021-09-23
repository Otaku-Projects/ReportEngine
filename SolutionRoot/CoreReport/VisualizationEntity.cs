using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreReport
{
    public abstract class VisualizationEntity
    {
        protected string tempRenderFolder = @"D:\\Temp";
        private int numCopies;
        public int NumCopies
        {
            get { return numCopies; }
            set { numCopies = value; }
        }
        public abstract void Display();
        public abstract void SaveFile();
        public abstract void SaveAndDownloadAsBase64();
    }
}
