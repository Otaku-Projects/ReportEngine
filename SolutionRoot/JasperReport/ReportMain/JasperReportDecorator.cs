﻿using System;
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
using static JasperReport.ReportEntity.JasperReportEntity;

namespace CoreReport.JasperReport
{
    public class JasperReportDecorator : VisualizationDecorator
    {
        private string createdBy;
        private DateTime createdDate;
        private DateTime printedDate;
        private string filename;

        private JasperReportEntity reportEntity;
        private IRenderService rs;

        private DataSet dataSet;
        private IDictionary<string, object> dataSetObj;
        private string jasperReportRenderFolder;

        public JasperReportDecorator()
        {
            this.jasperReportRenderFolder = this.tempRenderFolder;
        }
        public JasperReportDecorator(JasperReportEntity _reportEntity, string _filename = "")
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

        public virtual void SaveExcel(string _fileName = "")
        {
            this.SaveXlsx(_fileName);
        }

        public virtual void SaveExcelByHTML(string _fileName = "")
        {
            this.SaveXlsxByHTML(_fileName);
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
                RenderRequest _renderRequest = this.CreateXlsxRenderRequest();

                var report = rs.RenderAsync(_renderRequest).Result;
                report.Content.CopyTo(File.OpenWrite(
                    Path.Combine(
                        this.jasperReportRenderFolder
                        , _fileName + ".xlsx")
                ));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        public virtual void SaveXlsxByHTML(string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            try
            {
                RenderRequest _renderRequest = this.CreateXlsxRenderByHTMLRequest();

                var report = rs.RenderAsync(_renderRequest).Result;
                report.Content.CopyTo(File.OpenWrite(
                    Path.Combine(
                        this.jasperReportRenderFolder
                        , _fileName + ".xlsx")
                ));
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
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

        public virtual void SavePdf(string _templateFile="", string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            try
            {
                RenderRequest _renderRequest = this.CreatePdfRenderRequest();

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

        protected RenderRequest CreateEmptyRenderRequest(string _templateFile = "")
        {
            RenderRequest _renderRequest = new RenderRequest();
            _renderRequest.Template = new Template()
            {
                Content = "",
                Engine = Engine.None,
                Recipe = Recipe.ChromePdf
            };
            _renderRequest.Data = new object();

            return _renderRequest;
        }

        protected RenderRequest CreateXlsxRenderRequest()
        {
            RenderRequest _renderRequest = new RenderRequest();

            _renderRequest = this.CreateXlsxRenderByHTMLRequest();
            _renderRequest.Template.Engine = Engine.Handlebars;
            _renderRequest.Template.Recipe = Recipe.Xlsx;

            return _renderRequest;
        }

        protected RenderRequest CreateXlsxRenderByHTMLRequest()
        {
            RenderRequest _renderRequest = new RenderRequest();

            _renderRequest = this.CreateEmptyRenderRequest();

            string _htmlFileContent = string.Empty;
            string _scriptFileContent = string.Empty;

            string _htmlFilePath = string.Empty;
            string _scriptFilePath = string.Empty;
            _htmlFilePath = this.reportEntity.GetPageComponent(PageNature.MainContent).GetHtmlFilePath();
            _scriptFilePath = this.reportEntity.GetPageComponent(PageNature.MainContent).GetScriptFilePath();

            if (!string.IsNullOrEmpty(_htmlFilePath) && File.Exists(_htmlFilePath))
            {
                _htmlFileContent = File.ReadAllText(_htmlFilePath);
            }
            if (!string.IsNullOrEmpty(_scriptFilePath) && File.Exists(_scriptFilePath))
            {
                _scriptFileContent = File.ReadAllText(_scriptFilePath);
            }

            _renderRequest.Template.Content = _htmlFileContent;
            _renderRequest.Template.Helpers = _scriptFileContent;
            _renderRequest.Template.Engine = Engine.Handlebars;
            _renderRequest.Template.Recipe = Recipe.HtmlToXlsx;

            _renderRequest.Data = this.dataSetObj;

            return _renderRequest;
        }

        protected RenderRequest CreatePdfRenderRequest()
        {
            RenderRequest _renderRequest = new RenderRequest();

            _renderRequest = this.CreateEmptyRenderRequest();
            _renderRequest = this.AddMarginToRenderRequest(_renderRequest);

            string _htmlFileContent = string.Empty;
            string _scriptFileContent = string.Empty;

            string _htmlFilePath = string.Empty;
            string _scriptFilePath = string.Empty;
            //_scriptFilePath = Path.Combine(this.reportEntity.GetTemplateFileDirectory(), "helper.js");

            _htmlFilePath = this.reportEntity.GetPageComponent(PageNature.MainContent).GetHtmlFilePath();
            _scriptFilePath = this.reportEntity.GetPageComponent(PageNature.MainContent).GetScriptFilePath();

            if(!string.IsNullOrEmpty(_htmlFilePath) && File.Exists(_htmlFilePath))
            {
                _htmlFileContent = File.ReadAllText(_htmlFilePath);
            }
            if (!string.IsNullOrEmpty(_scriptFilePath) && File.Exists(_scriptFilePath))
            {
                _scriptFileContent = File.ReadAllText(_scriptFilePath);
            }

            _renderRequest.Template.Content = _htmlFileContent;
            _renderRequest.Template.Helpers = _scriptFileContent;
            _renderRequest.Template.Engine = Engine.Handlebars;
            _renderRequest.Template.Recipe = Recipe.ChromePdf;

            _renderRequest = this.AddPdfutilsToRenderRequest(_renderRequest);

            _renderRequest.Data = this.dataSetObj;

            //_renderRequest.Template.Chrome.MarginTop = "1cm";
            //_renderRequest.Template.Chrome.MarginLeft = "1cm";
            //_renderRequest.Template.Chrome.MarginBottom = "1cm";
            //_renderRequest.Template.Chrome.MarginRight = "1cm";

            return _renderRequest;
        }

        protected RenderRequest AddMarginToRenderRequest(RenderRequest _renderRequest)
        {
            _renderRequest.Template.Chrome = new Chrome()
            {
                MarginTop = "2.54cm",
                MarginLeft = "2.54cm",
                MarginBottom = "2.54cm",
                MarginRight = "2.54cm"
            };

            return _renderRequest;
        }

        protected RenderRequest AddPdfutilsToRenderRequest(RenderRequest _renderRequest)
        {
            HeaderFooterOptions hfOption = this.reportEntity.GetHeaderFooterOption();
            Dictionary<PageNature, PageComponent> pageComponents = this.reportEntity.GetPageComponents();

            string htmlFileContent = string.Empty;
            string scriptContent = string.Empty;

            #region Header and Footer
            string _headerFileContent = string.Empty;
            string _footerFileContent = string.Empty;
            string _headerFooterFileContent = string.Empty;

            string _headerFilePath = string.Empty;
            string _footerFilePath = string.Empty;

            _headerFilePath = this.reportEntity.GetPageComponent(PageNature.Header).GetHtmlFilePath();
            _footerFilePath = this.reportEntity.GetPageComponent(PageNature.Footer).GetHtmlFilePath();

            if (!string.IsNullOrEmpty(_headerFilePath) && File.Exists(_headerFilePath))
            {
                _headerFileContent = File.ReadAllText(_headerFilePath);
            }

            if (!string.IsNullOrEmpty(_footerFilePath) && File.Exists(_footerFilePath))
            {
                _footerFileContent = File.ReadAllText(_footerFilePath);
            }

            #endregion
            List<PdfOperation> _pdfOperationList = new List<PdfOperation>();
            foreach (KeyValuePair<PageNature, PageComponent> _pageKV in pageComponents)
            {
                if(_pageKV.Key == PageNature.MainContent) continue;
                htmlFileContent = string.Empty;
                scriptContent = string.Empty;

                htmlFileContent = File.ReadAllText(_pageKV.Value.GetHtmlFilePath());

                // read script if exists, allows render html without script file
                if(File.Exists(_pageKV.Value.GetScriptFilePath()))
                    scriptContent = File.ReadAllText(_pageKV.Value.GetScriptFilePath());

                PdfOperation _pdfOperation = new PdfOperation()
                {
                    Type = PdfOperationType.Merge,
                    RenderForEveryPage = true,
                    Template = new Template
                    {
                        Content = htmlFileContent,
                        Helpers = scriptContent,
                        Engine = Engine.Handlebars,
                        Recipe = Recipe.ChromePdf
                    }
                };
                _pdfOperationList.Add(_pdfOperation);
            }
            _renderRequest.Template.PdfOperations = _pdfOperationList;

            return _renderRequest;
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