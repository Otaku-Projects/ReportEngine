using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace EPPlus5Report.ReportEntity
{
    public class PageComponent {
        private string htmlPath;
        private string scriptPath;
        private string cssPath;
        private string directory;
        public PageComponent()
        {
            this.directory = string.Empty;
            this.htmlPath = string.Empty;
            this.scriptPath = string.Empty;
            this.cssPath = string.Empty;
        }
        public PageComponent(string _directory, string _htmlFileName, string _scriptFileName)
        {
            this.directory = string.Empty;

            this.htmlPath = string.Empty;
            this.scriptPath = string.Empty;
            this.cssPath = string.Empty;

            if (this.SetDirectory(_directory))
            {
                this.SetHtmlFileName(_htmlFileName);
                this.SetScriptFileName(_scriptFileName);
            }

        }
        public string GetHtmlFilePath()
        {
            return this.htmlPath;
        }
        public string GetScriptFilePath()
        {
            return this.scriptPath;
        }
        public string GetHtmlFileName()
        {
            return Path.GetFileName(this.htmlPath);
        }
        public string GetScriptFileName()
        {
            return Path.GetFileName(this.scriptPath);
        }

        public Boolean SetDirectory(string _directory)
        {
            if (Directory.Exists(_directory))
            {
                this.directory = _directory;
                return true;
            }
            else
            {
                return false;
            }
            return false;
        }
        public void SetHtmlPath(string _htmlPath)
        {
            if (Path.IsPathRooted(_htmlPath) && File.Exists(_htmlPath))
            {
                this.htmlPath = _htmlPath;
            }
        }

        public void SetScriptPath(string _scriptPath)
        {
            if (Path.IsPathRooted(_scriptPath) && File.Exists(_scriptPath))
            {
                this.scriptPath = _scriptPath;
            }
        }

        public void SetHtmlFileName(string _htmlFileName)
        {
            string _directory = this.directory;
            this.htmlPath = Path.Combine(_directory, _htmlFileName);
        }

        public void SetScriptFileName(string _scriptFileName)
        {
            string _directory = this.directory;
            this.scriptPath = Path.Combine(_directory, _scriptFileName);
        }
    }

    public abstract class BaseReportEntity
    {
        protected DataSet dataSet;
        protected IDictionary<string, object> dataSetObj;

        protected string rptFilesFolder;

        protected string templateBaseDirectory;
        protected string templateReportFileDirectory;
        //protected string templateReportFileLocation;
        //protected string templateReportHeaderLocation;
        //protected string templateReportFooterLocation;

        protected string xlsxTemplateFileName;

        protected Dictionary<PageNature, PageComponent> pageComponents;

        public abstract void InitializateMetaData();
        public abstract void InitializateMainContent();
        public abstract void InitializateHeaderFooter();
        public abstract void InitializateDataGrid();

        protected HeaderFooterOptions headerFooterOption;

        private List<ExcelDataGrid> dataGridList;
        private List<ExcelDataGrid> dataGridTemplateBackupList;

        public BaseReportEntity()
        {
            this.templateBaseDirectory = @"D:\Documents\ReportEngine\SolutionRoot\JasperReport\ReportTemplate";
            this.templateBaseDirectory = Path.Combine(Directory.GetCurrentDirectory(), "ReportTemplate");

            this.pageComponents = new Dictionary<PageNature, PageComponent>();

            this.xlsxTemplateFileName = string.Empty;

            this.dataGridList = new List<ExcelDataGrid>();
            this.dataGridTemplateBackupList = new List<ExcelDataGrid>();

            this.Initializate();
        }
        protected virtual void Initializate()
        {
            this.InitializateMetaData();
            this.InitializateMainContent();
            this.InitializateHeaderFooter();

            this.InitializateDataGrid();
        }

        protected virtual void AddDataGrid(ExcelDataGrid _dataGrid)
        {
            if (_dataGrid.IsValidAddToDataGridList())
            {
                this.dataGridList.Add(_dataGrid);
            }
        }

        public virtual List<ExcelDataGrid> GetDataGrid()
        {
            return this.dataGridList;
        }

        public virtual void SetDataGrid(List<ExcelDataGrid> _dataGridList)
        {
            this.dataGridList = _dataGridList;
        }

        public virtual void BackupDataGridSetting()
        {
            this.dataGridTemplateBackupList = new List<ExcelDataGrid>();
            this.dataGridList.ForEach((item) =>
            {
                this.dataGridTemplateBackupList.Add(new ExcelDataGrid(item));
            });
        }
        public virtual List<ExcelDataGrid> GetBackupTemplateDataGrid()
        {
            return this.dataGridTemplateBackupList;
        }

        protected virtual void AddPageContent(PageComponent _pageComponent)
        {
            if (this.pageComponents.ContainsKey(PageNature.MainContent))
            {
                this.pageComponents[PageNature.MainContent] = _pageComponent;
            }
            else
            {
                this.pageComponents.Add(PageNature.MainContent, _pageComponent);
            }
        }

        protected virtual void AddPageHeader(PageComponent _pageComponent)
        {
            if (this.pageComponents.ContainsKey(PageNature.Header))
            {
                this.pageComponents[PageNature.Header] = _pageComponent;
            }
            else
            {
                this.pageComponents.Add(PageNature.Header, _pageComponent);
            }

            if (this.pageComponents.ContainsKey(PageNature.Header) && this.pageComponents.ContainsKey(PageNature.Footer))
            {
                this.headerFooterOption = HeaderFooterOptions.HeaderFooterInSeparateFile;
            }
            else if (this.pageComponents.ContainsKey(PageNature.Header))
            {
                this.headerFooterOption = HeaderFooterOptions.Header;
            }
        }
        protected virtual void AddPageFooter(PageComponent _pageComponent)
        {
            this.headerFooterOption = HeaderFooterOptions.Footer;

            if (this.pageComponents.ContainsKey(PageNature.Footer))
            {
                this.pageComponents[PageNature.Footer] = _pageComponent;
            }
            else
            {
                this.pageComponents.Add(PageNature.Footer, _pageComponent);
            }

            if (this.pageComponents.ContainsKey(PageNature.Header) && this.pageComponents.ContainsKey(PageNature.Footer))
            {
                this.headerFooterOption = HeaderFooterOptions.HeaderFooterInSeparateFile;
            }
            else if (this.pageComponents.ContainsKey(PageNature.Footer))
            {
                this.headerFooterOption = HeaderFooterOptions.Footer;
            }
        }
        protected virtual void AddPageHeaderFooter(PageComponent _pageComponent)
        {
            this.headerFooterOption = HeaderFooterOptions.HeaderFooterInSingleFile;

            if (this.pageComponents.ContainsKey(PageNature.Header))
                this.pageComponents.Remove(PageNature.Header);
            if (this.pageComponents.ContainsKey(PageNature.Footer))
                this.pageComponents.Remove(PageNature.Footer);

            if (this.pageComponents.ContainsKey(PageNature.HeaderAndFooter))
            {
                this.pageComponents[PageNature.HeaderAndFooter] = _pageComponent;
            }
            else
            {
                this.pageComponents.Add(PageNature.HeaderAndFooter, _pageComponent);
            }
        }

        public PageComponent GetPageComponent(PageNature _pageNature)
        {
            PageComponent _pc = new PageComponent();
            if (this.pageComponents.ContainsKey(_pageNature))
            {
                _pc = this.pageComponents[_pageNature];
            }
            return _pc;
        }

        public Dictionary<PageNature, PageComponent> GetPageComponents()
        {
            return this.pageComponents;
        }

        protected void SetDataSet(DataSet _dataSet)
        {
            this.dataSet = _dataSet;
        }

        public DataSet GetDataSet()
        {
            return this.dataSet;
        }

        protected void SetDataSetObj(IDictionary<string, object> _dataSetObj)
        {
            this.dataSetObj = _dataSetObj;
        }

        public IDictionary<string, object> GetDataSetObj()
        {
            if (this.dataSetObj == null)
            {
                return this.dataSetObj;
            }
            // renew date, time on each get
            string _tableName = "ReportMetaData";
            this.dataSetObj.Remove(_tableName);
            dynamic _obj = new ExpandoObject();
            _obj = new
            {
                DateTime = DateTime.Now.ToString("dd MMMM yyyy HH:mm")
            };

            this.dataSetObj.Add(_tableName, _obj);

            return this.dataSetObj;
        }

        public string GetTemplateFileDirectory()
        {
            return this.templateReportFileDirectory;
        }
        public string GetXlsxTemplateFilePath()
        {
            return Path.Combine(this.templateReportFileDirectory, this.xlsxTemplateFileName);
        }
        public enum HeaderFooterOptions
        {
            None = 0,
            Header = 11,
            Footer = 12,
            HeaderFooterInSingleFile = 21,
            HeaderFooterInSeparateFile = 22
        }

        public enum PageNature
        {
            None = 0,
            Header = 11,
            Footer = 12,
            HeaderAndFooter = 13,
            MainContent = 14
        }

        public HeaderFooterOptions GetHeaderFooterOption()
        {
            return this.headerFooterOption;
        }

        protected void SetXlsxTemplateFileName(string _xlsxTemplateFileName)
        {
            this.xlsxTemplateFileName = _xlsxTemplateFileName;
        }
        public string GetXlsxTemplateFileName()
        {
            return this.xlsxTemplateFileName;
        }
    }

    public class FileOutputUtil
    {
        static DirectoryInfo _outputDir = null;
        public static DirectoryInfo OutputDir
        {
            get
            {
                return _outputDir;
            }
            set
            {
                _outputDir = value;
                if (!_outputDir.Exists)
                {
                    _outputDir.Create();
                }
            }
        }
        public static FileInfo GetFileInfo(string file, bool deleteIfExists = true)
        {
            var fi = new FileInfo(OutputDir.FullName + Path.DirectorySeparatorChar + file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }
        public static FileInfo GetFileInfo(DirectoryInfo altOutputDir, string file, bool deleteIfExists = true)
        {
            var fi = new FileInfo(altOutputDir.FullName + Path.DirectorySeparatorChar + file);
            if (deleteIfExists && fi.Exists)
            {
                fi.Delete();  // ensures we create a new workbook
            }
            return fi;
        }


        internal static DirectoryInfo GetDirectoryInfo(string directory)
        {
            var di = new DirectoryInfo(_outputDir.FullName + Path.DirectorySeparatorChar + directory);
            if (!di.Exists)
            {
                di.Create();
            }
            return di;
        }
    }
}
