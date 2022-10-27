using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CoreReport;
using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;

namespace CoreReport.CrystalReport
{
    public class CrystalReportDecorator : VisualizationDecorator
    {
        private string createdBy;
        private DateTime createdDate;
        private DateTime printedDate;
        private string filename;

        private ReportDocument reportDocument;
        private DataSet dataSet;

        private string crystalReportRenderFolder;

        protected ExportOptions exportOptions;
        protected DiskFileDestinationOptions CrDiskFileDestinationOptions;
        protected PdfRtfWordFormatOptions CrFormatTypeOptions;

        public CrystalReportDecorator(ReportDocument _rptDoc, string _filename = "")
        {
            if (string.IsNullOrEmpty(_filename))
            {
                Guid obj = Guid.NewGuid();
                _filename = obj.ToString();
            }

            this.reportDocument = _rptDoc;
            this.crystalReportRenderFolder = this.tempRenderFolder;

            this.filename = _filename;

            this.createdBy = "CoreSystem";
            this.createdDate = new DateTime();

            this.CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
            this.CrFormatTypeOptions = new PdfRtfWordFormatOptions();
        }

        public CrystalReportDecorator(ReportDocument _rptDoc, DataSet _dataSet, string _filename = "")
        {
            if (string.IsNullOrEmpty(_filename))
            {
                Guid obj = Guid.NewGuid();
                _filename = obj.ToString();
            }

            this.reportDocument = _rptDoc;

            this.dataSet = _dataSet;
            this.filename = _filename;

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

            this.CrDiskFileDestinationOptions.DiskFileName = System.IO.Path.Combine(
                this.crystalReportRenderFolder
                , _fileName + ".rtf");

            try
            {
                ExportOptions _wordExportOptions;

                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions wordFormatTypeOptions = new PdfRtfWordFormatOptions();
                CrDiskFileDestinationOptions.DiskFileName = System.IO.Path.Combine(
                this.crystalReportRenderFolder
                , _fileName + ".rtf");
                _wordExportOptions = this.reportDocument.ExportOptions;
                _wordExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                _wordExportOptions.ExportFormatType = ExportFormatType.WordForWindows;
                _wordExportOptions.ExportDestinationOptions = CrDiskFileDestinationOptions;
                _wordExportOptions.ExportFormatOptions = null;
                //this.reportDocument.Refresh();
                this.reportDocument.Export(_wordExportOptions);
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

            this.CrDiskFileDestinationOptions.DiskFileName = System.IO.Path.Combine(
                this.crystalReportRenderFolder
                , _fileName + ".xlsx");

            try
            {
                ExportOptions CrExportOptions;

                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                ExcelFormatOptions CrFormatTypeOptions = new ExcelFormatOptions();
                CrDiskFileDestinationOptions.DiskFileName = System.IO.Path.Combine(
                this.crystalReportRenderFolder
                , _fileName + ".xlsx");
                CrExportOptions = this.reportDocument.ExportOptions;
                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.ExportFormatType = ExportFormatType.XLSXPagebased;
                //CrExportOptions.ExportFormatType = ExportFormatType.XLSXRecord;
                CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                CrExportOptions.FormatOptions = CrFormatTypeOptions;
                this.reportDocument.Refresh();
                this.reportDocument.Export();
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        public virtual void SaveXls(string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            this.CrDiskFileDestinationOptions.DiskFileName = System.IO.Path.Combine(
                this.crystalReportRenderFolder
                , _fileName + ".xls");

            try
            {
                ExportOptions CrExportOptions;

                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                ExcelFormatOptions CrFormatTypeOptions = new ExcelFormatOptions();
                CrDiskFileDestinationOptions.DiskFileName = System.IO.Path.Combine(
                this.crystalReportRenderFolder
                , _fileName + ".xls");
                CrExportOptions = this.reportDocument.ExportOptions;
                CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                CrExportOptions.ExportFormatType = ExportFormatType.Excel;
                CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                CrExportOptions.FormatOptions = CrFormatTypeOptions;
                this.reportDocument.Refresh();
                this.reportDocument.Export();
            }
            catch (Exception ex)
            {
            }
        }

        public virtual void SavePdf(string _fileName = "")
        {
            this.RefreshPrintDate();

            if (string.IsNullOrEmpty(_fileName))
            {
                _fileName = this.filename;
            }

            this.CrDiskFileDestinationOptions.DiskFileName = System.IO.Path.Combine(
                this.crystalReportRenderFolder
                , _fileName + ".pdf");

            /*
            this.exportOptions = this.reportDocument.exportop;
            {
                this.exportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                this.exportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                this.exportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                this.exportOptions.FormatOptions = CrFormatTypeOptions;
            }
            doc.Export();
            */

            //this.reportDocument.Refresh();
            //this.reportDocument.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, this.CrDiskFileDestinationOptions.DiskFileName);


            try
            {
                ExportOptions CrExportOptions;
                DiskFileDestinationOptions CrDiskFileDestinationOptions = new DiskFileDestinationOptions();
                PdfRtfWordFormatOptions CrFormatTypeOptions = new PdfRtfWordFormatOptions();
                CrDiskFileDestinationOptions.DiskFileName = System.IO.Path.Combine(
                this.crystalReportRenderFolder
                , _fileName + ".pdf");
                CrExportOptions = this.reportDocument.ExportOptions;
                {
                    CrExportOptions.ExportDestinationType = ExportDestinationType.DiskFile;
                    CrExportOptions.ExportFormatType = ExportFormatType.PortableDocFormat;
                    CrExportOptions.DestinationOptions = CrDiskFileDestinationOptions;
                    CrExportOptions.FormatOptions = CrFormatTypeOptions;
                }
                this.reportDocument.Refresh();
                this.reportDocument.Export();
            }
            catch (Exception ex)
            {
            }
        }
    }
}
