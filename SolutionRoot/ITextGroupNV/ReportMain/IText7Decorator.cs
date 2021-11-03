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
using ITextGroupNV.ReportEntity;
using iText.Html2pdf;
using Fluid;
using iText.Html2pdf.Resolver.Font;
using iText.Layout.Font;
using iText.IO.Font;
using System.Collections;
using Newtonsoft.Json.Linq;
using Fluid.Values;
using iText.StyledXmlParser.Css.Media;
using iText.Layout.Element;
using iText.Layout;
using iText.Kernel.Pdf;
using iText.Layout.Properties;
using iText.Layout.Splitting;
using iText.IO.Font.Otf;
using iText.StyledXmlParser.Jsoup.Select;

namespace CoreReport.ITextGroupNV
{
    public class IText7Decorator : VisualizationDecorator
    {
        protected string createdBy;
        protected DateTime createdDate;
        protected DateTime printedDate;
        protected string filename;

        protected ITextReportEntity reportEntity;

        protected DataSet dataSet;
        protected IDictionary<string, object> dataSetObj;
        protected string iTextReportRenderFolder;
        protected ConverterProperties converterProperties;

        public List<string> _fonts;

        protected string report_instance_dir;
        protected string report_template_dir;
        protected string fonts_folder;

        public IText7Decorator()
        {
            this.iTextReportRenderFolder = this.tempRenderFolder;

            this.Initialize();
        }
        public IText7Decorator(ITextReportEntity _reportEntity, string _filename = "")
        {
            if (string.IsNullOrEmpty(_filename))
            {
                Guid obj = Guid.NewGuid();
                _filename = obj.ToString();
            }

            this.dataSet = _reportEntity.GetDataSet();
            this.dataSetObj = _reportEntity.GetDataSetObj();
            this.iTextReportRenderFolder = this.tempRenderFolder;

            this.filename = _filename;

            this.reportEntity = _reportEntity;

            this.createdBy = "CoreSystem";
            this.createdDate = new DateTime();

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
        }

        public void RefreshPrintDate()
        {
            this.printedDate = new DateTime();
        }

        public override void Display()
        {
            throw new NotImplementedException();
        }

        protected ConverterProperties CreatePdfTemplatePropertiesInstance()
        {
            string _pdfTemplateFilePath = this.reportEntity.GetPdfTemplateFilePath();
            if (!File.Exists(_pdfTemplateFilePath))
            {
                throw new FileNotFoundException($"PDF template (HTML file) not found at {_pdfTemplateFilePath}");
            }

            ConverterProperties _converterProperties = new ConverterProperties();

            string _report_instance_dir = this.report_instance_dir;
            string _report_template_dir = this.report_template_dir;
            string _fonts_folder = this.fonts_folder;
            // adding fonts folder
            FontProvider fontProvider = new DefaultFontProvider(false, false, false);
            foreach (string fontFilename in this._fonts)
            {
                string fontPath = Path.Combine(_fonts_folder, fontFilename);
                FontProgram fontProgram = FontProgramFactory.CreateFont(fontPath);
                fontProvider.AddFont(fontProgram);
            }
            _converterProperties.SetFontProvider(fontProvider);

            this.converterProperties = _converterProperties;

            return _converterProperties;
        }
        public ConverterProperties GetPdfTemplatePropertiesInstance()
        {
            return this.CreatePdfTemplatePropertiesInstance();
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
                this.iTextReportRenderFolder,
                _fileName + ".pdf");

            string _report_instance_dir = this.report_instance_dir;
            string _report_template_dir = this.report_template_dir;
            string _fonts_folder = this.fonts_folder;

            try
            {
                // ConverterProperties not implement IDisposable, cannot use using here
                //using (ConverterProperties _converterProperties = this.StartRenderDataAndMergeToTemplate())
                //{
                ConverterProperties _converterProperties = this.StartRenderDataAndMergeToTemplate();

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

                eoDynamic.ReportTemplate_Root = Path.Combine("file:///", System.IO.Directory.GetParent(_report_instance_dir).ToString());
                eoDynamic.ReportInstance_Folder = Path.Combine("file:///", _report_instance_dir);
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

                // Convert html to pdf
                using (FileStream pdfDest = File.Open(pdfFilePath, FileMode.Create))
                {
                    _converterProperties.SetBaseUri(eoDynamic.ReportInstance_Folder);
                    MediaDeviceDescription mediaDeviceDescription = new MediaDeviceDescription(MediaType.SCREEN);
                    _converterProperties.SetMediaDeviceDescription(mediaDeviceDescription);
                    HtmlConverter.ConvertToPdf(htmlRenderResult, pdfDest, _converterProperties);

                    //IList<IElement> elements = HtmlConverter.ConvertToElements(htmlRenderResult, converterProperties);

                    //FileOutputStream outputStream = new FileOutputStream(pdfDest);
                    //WriterProperties writerProperties = new WriterProperties();
                    //writerProperties.AddXmpMetadata();

                    //PdfWriter writer = new PdfWriter(outputStream);

                    //PdfDocument pdfDoc = new PdfDocument(writer);

                    //Document document = new Document(pdfDoc);
                    //document.SetProperty(Property.SPLIT_CHARACTERS, new DefaultSplitCharacters(){
                    //        @Override
                    //        public boolean isSplitCharacter(GlyphLine text, int glyphPos)
                    //{
                    //    //return super.isSplitCharacter(text, glyphPos);//override this 
                    //    return true;//解决word-break: break-all;不兼容的问题
                    //}
                    //for (IElement element : elements)
                    //{
                    //    document.add((IBlockElement)element);
                    //}
                    //document.close();
                    //});
                    //_converterProperties.SetProperty(Property.SPLIT_CHARACTERS);
                    //});
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
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
        protected virtual ConverterProperties StartRenderDataAndMergeToTemplate()
        {
            ConverterProperties _converterProperties = this.GetPdfTemplatePropertiesInstance();

            this.converterProperties = _converterProperties;
            return _converterProperties;
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