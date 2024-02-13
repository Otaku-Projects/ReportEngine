// See https://aka.ms/new-console-template for more information
using EasyOpenXml;
using System.Reflection;

Console.WriteLine("Hello, World!");

GeneratedClass sampleExcel = new GeneratedClass();
List<ReportFileTuple> filesList1 = sampleExcel.DownloadExcel();

string filePath = $"{System.IO.Path.GetDirectoryName(Assembly.GetEntryAssembly().Location)}{System.IO.Path.DirectorySeparatorChar}";

if(filesList1 != null && filesList1.Count > 0)
{
    try
    {
        foreach (ReportFileTuple file1 in filesList1)
        {
            using (var fs = new FileStream(filePath + file1.Filename, FileMode.Create))
            {
                fs.Write(file1.FileByte, 0, file1.FileByte.Length);
            }
        }
    }
    catch (Exception e)
    {
        Console.WriteLine(e.ToString());
    }
}