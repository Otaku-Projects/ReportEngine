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
    public class PageComponent
    {
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

}
