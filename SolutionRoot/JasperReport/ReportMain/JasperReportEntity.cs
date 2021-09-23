using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreReport;
using JasperReport.ReportEntity;
using jsreport.Binary;
using jsreport.Local;
using jsreport.Shared;
using jsreport.Types;

namespace CoreReport.JasperReport
{
    public class JasperReportEntity : VisualizationEntity
    {
        private string createdBy;
        private DateTime createdDate;
        private DateTime printedDate;
        private string filename;

        private BaseReportEntity reportEntity;
        private IRenderService rs;

        private DataSet dataSet;
        private IDictionary<string, object> dataSetObj;
        private string jasperReportRenderFolder;

        public JasperReportEntity()
        {
            //var _rs = new LocalReporting()
            //    .RunInDirectory(Path.Combine(Directory.GetCurrentDirectory(), "./ReportTemplate"))
            //    .KillRunningJsReportProcesses()
            //    .UseBinary(JsReportBinary.GetBinary())
            //    .Configure(cfg => cfg.AllowedLocalFilesAccess().FileSystemStore().BaseUrlAsWorkingDirectory())
            //    .AsUtility()
            //    .Create();

            //this.rs = _rs;
            this.jasperReportRenderFolder = this.tempRenderFolder;
        }
        public JasperReportEntity(BaseReportEntity _reportEntity, string _filename = "")
        {
            if (string.IsNullOrEmpty(_filename))
            {
                Guid obj = Guid.NewGuid();
                _filename = obj.ToString();
            }

            this.rs = _reportEntity.GetRenderService();
            this.dataSetObj = _reportEntity.GetDataSetObj();
            this.jasperReportRenderFolder = this.tempRenderFolder;

            this.filename = _filename;

            this.reportEntity = _reportEntity;

            this.createdBy = "CoreSystem";
            this.createdDate = new DateTime();
        }

        /*
        public JasperReportEntity(IRenderService _rs, string _filename = "")
        {
            if (string.IsNullOrEmpty(_filename))
            {
                Guid obj = Guid.NewGuid();
                _filename = obj.ToString();
            }

            this.rs = _rs;
            this.jasperReportRenderFolder = this.tempRenderFolder;

            this.filename = _filename;

            this.createdBy = "CoreSystem";
            this.createdDate = new DateTime();
        }
        */

        public void RefreshPrintDate()
        {
            this.printedDate = new DateTime();
        }

        public override void Display()
        {
            throw new NotImplementedException();
        }

        public override void SaveAndDownloadAsBase64()
        {
            this.RefreshPrintDate();
        }

        public override void SaveFile()
        {
            this.RefreshPrintDate();

        }

        public virtual void SaveExcel(string _fileName="")
        {
            this.SaveXlsx(_fileName);
        }

        public virtual void SaveRtf(string _fileName="")
        {

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            try
            {
            }
            catch (Exception ex)
            {
            }
        }

        public virtual void SaveXlsx(string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            try
            {
            }
            catch (Exception ex)
            {
            }
        }

        public virtual void SaveXls(string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            try
            {
            }
            catch (Exception ex)
            {
            }
        }

        public virtual void SavePdf(string _templateFile, string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            try
            {
                DataSet _dataSet = this.dataSet;
                object _reportData = this.ConvertDataSetToObject(_dataSet);
                //var report = rs.RenderAsync(RenderRequest_1_helloWorld).Result;
                //report.Content.CopyTo(File.OpenWrite(_fileName +".pdf"));

                //var invoiceReport = rs.RenderByNameAsync("Invoice", _dataSet).Result;
                //invoiceReport.Content.CopyTo(File.OpenWrite(_fileName + ".pdf"));

                string _fileContent = File.ReadAllText(_templateFile);
                string _scriptFile = File.ReadAllText(Path.Combine(this.reportEntity.GetTemplateFileDirectory(), "helper.js"));
                RenderRequest _renderRequest = this.CreateEmptyRenderRequest(_templateFile);
                _renderRequest.Template.Content = _fileContent;
                _renderRequest.Template.Helpers = _scriptFile;
                _renderRequest.Template.Engine = Engine.Handlebars;
                _renderRequest.Template.Recipe = Recipe.ChromePdf;
                _renderRequest.Data = this.dataSetObj;
                var report = rs.RenderAsync(_renderRequest).Result;
                report.Content.CopyTo(File.OpenWrite(
                    Path.Combine(
                        this.jasperReportRenderFolder
                        , _fileName + ".pdf")
                ));
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        protected RenderRequest CreateEmptyRenderRequest(string _templateFile="")
        {
            RenderRequest _renderRequest = new RenderRequest();
            //string _fileContent = File.ReadAllText(_templateFile);
            _renderRequest.Template = new Template()
            {
                Content = "",
                Engine = Engine.None,
                Recipe = Recipe.ChromePdf
            };
            _renderRequest.Data = new object();

            return _renderRequest;
        }

        protected object ConvertDataSetToObject(DataSet _dataSet)
        {
            /*
                dynamic x = new ExpandoObject();
                x.NewProp = string.Empty;
            //or
                var x = new ExpandoObject() as IDictionary<string, Object>;
                x.Add("NewProp", string.Empty);
             */
        //dynamic _obj = new ExpandoObject();
        var _obj = new ExpandoObject() as IDictionary<string, object>;
            if (_dataSet == null || _dataSet.Tables.Count == 0) return _obj;

            foreach (DataTable _table in _dataSet.Tables)
            {
                //var tableObj = new ExpandoObject() as IDictionary<string, object>;
                List<dynamic> rowList = new List<dynamic>();
                _obj.Add(_table.TableName, rowList);
                //_obj.Add(_table.TableName, new dynamic[_table.Rows.Count]);

                //Array _array = new Array() { };
                //_obj.Add(_table.TableName, _array);

                //dynamic packet = new ExpandoObject();
                foreach (DataRow _row in _table.Rows)
                {
                    //_obj[_table.TableName][_rowCount] = new ExpandoObject();
                    var expandoDict = new ExpandoObject() as IDictionary<String, Object>;
                    foreach (DataColumn col in _table.Columns)
                    {
                        //put every column of this row into the new dictionary
                        expandoDict.Add(col.ColumnName, _row[col.ColumnName]);
                    }
                    rowList.Add(expandoDict);

                    //foreach (object item in _row.ItemArray)
                    //{
                    //    // read item
                    //    expandoDict.Add(item.ToString(), item);
                    //}
                }
            }

            return _obj;

            /*
             

            DataTable _dt = _dataSet.Tables[0];

            var dynamicDt = new List<dynamic>();
            foreach (DataRow row in _dt.Rows)
            {
                dynamic dyn = new ExpandoObject();
                dynamicDt.Add(dyn);
                foreach (DataColumn column in _dt.Columns)
                {
                    var dic = (IDictionary<string, object>)dyn;
                    dic[column.ColumnName] = row[column];
                }
            }
            return dynamicDt;
             */
        }
    }
}