using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IronPDFProject.ReportEntity
{
    public class UrlToPdfReportEntity : IronPdfReportEntity
    {
        protected List<string> UrlPathList;
        public UrlToPdfReportEntity()
        {
        }

        public void ClearUrlPath()
        {
            this.UrlPathList = new List<string>();
        }
        public void AddUrlPath(string urlPath)
        {
            if (!this.UrlPathList.Contains(urlPath))
                this.UrlPathList.Add(urlPath);
        }
        public bool RemoveUrlPath(string urlPath)
        {
            return this.UrlPathList.Remove(urlPath);
        }
        public List<string> GetUrlPath()
        {
            return this.UrlPathList;
        }

        public override void InitializateMetaData()
        {
            this.ClearUrlPath();
        }

        public override void InitializateMainContent()
        {
            //throw new NotImplementedException();
            // check is URL valid
            // validate the URL pattern

            // try to visit the URL

        }

        public override void InitializateDataGrid()
        {
            //throw new NotImplementedException();
        }

        public override void InitializateHeaderFooter()
        {
            //throw new NotImplementedException();
        }
    }
}
