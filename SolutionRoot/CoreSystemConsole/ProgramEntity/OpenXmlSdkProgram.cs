using CoreReport.OpenXmlSDK;
using CoreSystemConsole.ReportDataModel;
using OpenXmlSDK.ReportEntity;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CoreSystemConsole.ProgramEntity
{
    public class OpenXmlSdkProgram
    {
        public OpenXmlSdkProgram()
        {
            Console.WriteLine("Said \"Hello World!\" from EPPlus5XlsxTemplateProgram");

            HitRateDataView hitRateDataView2 = new HitRateDataView();
            HitRateDataView hitRateDataView3 = new HitRateDataView();

            hitRateDataView2.CreateDummyData2();
            IDictionary<string, object> dataSetObj2 = hitRateDataView2.GetDataSetObj();
            hitRateDataView3.CreateDummyData3();
            IDictionary<string, object> dataSetObj3 = hitRateDataView3.GetDataSetObj();
            DataSet dataSet3 = hitRateDataView3.GetDataSet();

            OpenXmlSDKDecorator openXmlSDKDecorator = null;

            HitRateReport8 hitRateReport8 = new HitRateReport8(dataSetObj3);
            openXmlSDKDecorator = new OpenXmlSDKDecorator(hitRateReport8);
            openXmlSDKDecorator.RenderTemplateAndSaveAsXlsx();
        }
    }
}
