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
using System.Collections;
using Puppeteer.ReportEntity;
using Fluid;
using Newtonsoft.Json.Linq;
using Fluid.Values;
using System.Diagnostics;

namespace CoreReport.Puppeteer
{
    public class PuppeteerDecorator : VisualizationDecorator
    {
        protected string createdBy;
        protected DateTime createdDate;
        protected DateTime printedDate;
        protected string filename;

        protected PuppeteerReportEntity reportEntity;

        protected DataSet dataSet;
        protected IDictionary<string, object> dataSetObj;
        protected string puppeteerRenderFolder;
        protected string puppeteerNodeFolder;

        public List<string> _fonts;

        protected string report_instance_dir;
        protected string report_template_dir;
        protected string fonts_folder;

        public PuppeteerDecorator()
        {
            this.Initialize();
        }
        public PuppeteerDecorator(PuppeteerReportEntity _reportEntity, string _filename = "")
        {
            if (string.IsNullOrEmpty(_filename))
            {
                Guid obj = Guid.NewGuid();
                _filename = obj.ToString();
            }

            this.dataSet = _reportEntity.GetDataSet();
            this.dataSetObj = _reportEntity.GetDataSetObj();

            this.filename = _filename;

            this.reportEntity = _reportEntity;

            this.createdBy = "CoreSystem";
            this.createdDate = new DateTime();

            this.report_instance_dir = string.Empty;
            this.report_template_dir = string.Empty;

            this.Initialize();
        }

        public void Initialize()
        {
            this._fonts = new List<string>();
            this._fonts.Add("NotoSansCJKjp-Regular.otf");
            this._fonts.Add("NotoSansCJKkr-Regular.otf");
            this._fonts.Add("NotoSansCJKsc-Regular.otf");
            this._fonts.Add("NotoSansCJKtc-Regular.otf");

            this.report_instance_dir = this.reportEntity.GetTemplateFileDirectory();
            this.report_template_dir = System.IO.Directory.GetParent(this.report_instance_dir).ToString();
            this.fonts_folder = Path.Combine(report_template_dir, "General", "fonts");

            this.puppeteerRenderFolder = this.tempRenderFolder;
            this.puppeteerNodeFolder = Path.Combine(System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location), "web");
        }

        public void RefreshPrintDate()
        {
            this.printedDate = new DateTime();
        }

        public override void Display()
        {
            throw new NotImplementedException();
        }

        protected void CreatePdfTemplatePropertiesInstance()
        {
            string _pdfTemplateFilePath = this.reportEntity.GetPdfTemplateFilePath();
            if (!File.Exists(_pdfTemplateFilePath))
            {
                throw new FileNotFoundException($"PDF template (HTML file) not found at {_pdfTemplateFilePath}");
            }
        }
        public void GetPdfTemplatePropertiesInstance()
        {
            this.CreatePdfTemplatePropertiesInstance();
        }
        public virtual void RenderTemplateAndSaveAsXlsx(string _fileName = "")
        {
            // you should not call into here, please inherit the decorator and override this function
            throw new NotImplementedException();
        }
        public virtual FileStream RenderTemplateAndSaveAsPdf(string _fileName = "")
        {
            if (string.IsNullOrEmpty(_fileName))
            {
                Guid obj = Guid.NewGuid();
                _fileName = obj.ToString();
            }

            string pdfFilePath = Path.Combine(
                this.puppeteerRenderFolder,
                _fileName + ".pdf");

            string htmlFilePath = Path.Combine(
                this.reportEntity.GetTemplateFileDirectory(),
                _fileName + ".html");

            string nodeFilePath = Path.Combine(
                this.puppeteerNodeFolder,
                "starter.puppeteer-report.js");

            string _report_instance_dir = this.report_instance_dir;
            string _report_template_dir = this.report_template_dir;
            string _fonts_folder = this.fonts_folder;

            try
            {

                IDictionary<string, object> _dataSetObj = this.reportEntity.GetDataSetObj();
                var newDict = new Dictionary<string, object>(_dataSetObj);

                // Convert IDictionary/Dictionary<string, object> To Anonymous Object
                var _expandoObject = new ExpandoObject();
                var eoColl = (ICollection<KeyValuePair<string, object>>)_expandoObject;
                foreach (var kvp in _dataSetObj)
                {
                    eoColl.Add(kvp);
                }
                dynamic eoDynamic = eoColl;

                //eoDynamic.ReportTemplate_Root = Path.Combine("file:///", System.IO.Directory.GetParent(_report_instance_dir).ToString());
                //eoDynamic.ReportInstance_Folder = Path.Combine("file:///", _report_instance_dir);
                eoDynamic.meta_PrintDateTime = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                eoDynamic.meta_DateTime_yyyy_mm_dd = DateTime.Now.ToString("yyyy-MM-dd");
                eoDynamic.meta_DateTime_yyyy_mm_dd_hh_mm_ss = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");

                string htmlRenderResult = string.Empty;
                //dynamic eo = _dataSetObj.Aggregate(new ExpandoObject() as IDictionary<string, Object>,
                //            (a, p) => { a.Add(p.Key, p.Value); return a; });

                # region read report template content
                string _pdfTemplateFilePath = this.reportEntity.GetPdfTemplateFilePath();
                //FileStream htmlTemplateStream = File.Open(_pdfTemplateFilePath, FileMode.Open, FileAccess.Read);

                using var fs = new FileStream(_pdfTemplateFilePath, FileMode.Open, FileAccess.Read);
                using var sr = new StreamReader(fs, Encoding.UTF8);
                string htmlTemplateSource = sr.ReadToEnd();
                #endregion

                TemplateOptions options = new TemplateOptions();
                options.MemberAccessStrategy.Register<IDictionary, object>(
                    (obj, name) => obj[name]
                );
                options.MemberAccessStrategy.MemberNameStrategy = MemberNameStrategies.Default;
                //TemplateContext contextBody = new TemplateContext(_dataSetObj, options);
                //TemplateContext contextBody = new TemplateContext(eoDynamic);

                // When a property of a JObject value is accessed, try to look into its properties
                options.MemberAccessStrategy.Register<JObject, object>((source, name) => source[name]);

                // Convert JToken to FluidValue
                options.ValueConverters.Add(x => x is JObject o ? new ObjectValue(o) : null);
                options.ValueConverters.Add(x => x is JValue v ? v.Value : null);

                string jsonText = Newtonsoft.Json.JsonConvert.SerializeObject(eoDynamic);
                var jsonModel = JObject.Parse(jsonText);

                var _dataSetObj2 = this.reportEntity.GetDataSetObj();
                TemplateContext contextBody = new TemplateContext(jsonModel, options);
                htmlRenderResult = this.RenderDataByFluid(htmlTemplateSource, contextBody);

                // use dictionary
                var model = new Dictionary<string, object>();
                model.Add("meta_CurrentDateTime", DateTime.Now);
                TemplateContext contentMetaData = new TemplateContext(model, options);
                htmlRenderResult = this.RenderFileMetaDataByFluid(htmlRenderResult, contentMetaData);

                // use Anonymous type
                //var model2 = new { Firstname = "Bill", Lastname = "Gates" };
                //TemplateContext contentMetaData2 = new TemplateContext(model2);
                //htmlRenderResult = this.RenderFileMetaDataByFluid(htmlTemplateSource, contentMetaData);

                // save html render result
                File.WriteAllText(@htmlFilePath, htmlRenderResult);

                // Convert html to pdf
                ProcessStartInfo startInfo = new ProcessStartInfo();

                // Sets NODE_PATH variable to installed node_modules directory
                // The new process will have RAYPATH variable created with "test" value
                // All environment variables of the created process are inherited from the
                // current process
                startInfo.EnvironmentVariables["NODE_PATH"] = "%AppData%\npm\node_modules";

                startInfo.FileName = "node";
                //startInfo.Arguments = $"/hidden /readonly /excel_active_sheet {xlsxFilePath} {pdfFilePath}";
                startInfo.Arguments = $"{nodeFilePath} {htmlFilePath} {pdfFilePath}";
                //startInfo.UseShellExecute = false;
                //startInfo.RedirectStandardOutput = true;
                //startInfo.CreateNoWindow = true;
                // convert xlsx to pdf
                Process exeProcess = Process.Start(startInfo);

                //Set a time-out value.
                int timeOut = 15000;

                // wait until it's done or time out.
                exeProcess.WaitForExit(timeOut);

                // Alternatively, if it's an application with a UI that you are waiting to enter into a message loop
                //exeProcess.WaitForInputIdle();

                //Check to see if the process is still running.
                if (exeProcess.HasExited == false)
                    //Process is still running.
                    //Test to see if the process is hung up.
                    if (exeProcess.Responding)
                        //Process was responding; close the main window.
                        exeProcess.CloseMainWindow();
                    else
                        //Process was not responding; force the process to close.
                        exeProcess.Kill();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            if (File.Exists(@htmlFilePath))
            {
                File.Delete(@htmlFilePath);
            }

            FileStream _fileStream = new FileStream(pdfFilePath, FileMode.Open, FileAccess.Read, FileShare.None);
            return _fileStream;
        }
        public virtual string RenderDataByFluid(string htmlTemplateSource, TemplateContext _content)
        {
            string htmlRenderResult = string.Empty;

            FluidParser parser = new FluidParser();

            if (parser.TryParse(htmlTemplateSource, out IFluidTemplate template, out string error))
            {
                htmlRenderResult = template.Render(_content);
            }
            else
            {
                Console.WriteLine($"Error: {error}");
            }

            return htmlRenderResult;
        }
        
        public virtual string RenderFileMetaDataByFluid(string htmlTemplateSource, TemplateContext _content)
        {
            string htmlRenderResult = string.Empty;
            FluidParser parser = new FluidParser();

            if (parser.TryParse(htmlTemplateSource, out IFluidTemplate template, out string error))
            {
                htmlRenderResult = template.Render(_content);
            }
            else
            {
                Console.WriteLine($"Error: {error}");
            }

            return htmlRenderResult;
        }

        public override void SaveAndDownloadAsBase64()
        {
            this.RefreshPrintDate();
        }

        public override void SaveFile()
        {
            this.RefreshPrintDate();
        }

        public virtual void SavePdf(string _templateFile="", string _fileName = "")
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
                Console.WriteLine(ex.ToString());
            }
        }

        protected object ConvertDataSetToObject(DataSet _dataSet)
        {
        var _obj = new ExpandoObject() as IDictionary<string, object>;
            if (_dataSet == null || _dataSet.Tables.Count == 0) return _obj;

            foreach (DataTable _table in _dataSet.Tables)
            {
                List<dynamic> rowList = new List<dynamic>();
                _obj.Add(_table.TableName, rowList);
                foreach (DataRow _row in _table.Rows)
                {
                    var expandoDict = new ExpandoObject() as IDictionary<String, Object>;
                    foreach (DataColumn col in _table.Columns)
                    {
                        //put every column of this row into the new dictionary
                        expandoDict.Add(col.ColumnName, _row[col.ColumnName]);
                    }
                    rowList.Add(expandoDict);
                }
            }

            return _obj;
        }
    }
}