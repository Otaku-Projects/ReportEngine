using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreReport
{
    public class VisualizationDecorator : VisualizationEntity
    {
        protected VisualizationEntity visualizationEntity;

        public VisualizationDecorator()
        {
        }

        public VisualizationDecorator(VisualizationEntity _visualizationEntity)
        {
            this.visualizationEntity = _visualizationEntity;
        }

        public override void Display()
        {
            this.visualizationEntity.Display();
        }

        public override void SaveAndDownloadAsBase64()
        {
            this.visualizationEntity.SaveAndDownloadAsBase64();
        }

        public override void SaveFile()
        {
            this.visualizationEntity.SaveFile();
        }
    }
}
