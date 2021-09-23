using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreReport
{
    public class CrystalReportDecorator : VisualizationDecorator
    {
        /*
        public CrystalReportDecorator(VisualizationEntity _visualizationEntity)
        {

        }
        */

        public CrystalReportDecorator(VisualizationEntity _visualizationEntity) : base(_visualizationEntity)
        {

        }

        public override void Display()
        {
            //base.Display();
        }
        public override void SaveAndDownloadAsBase64()
        {
            //base.SaveAndDownloadAsBase64();
            this.visualizationEntity.SaveAndDownloadAsBase64();
        }
        public override void SaveFile()
        {
            //base.SaveFile();
            this.visualizationEntity.SaveFile();
        }
    }
}
