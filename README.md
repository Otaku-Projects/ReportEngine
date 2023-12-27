# Brief of background
In charge of developing the whole report generation module for a new system, going to explore, try, and test several report (excel/pdf) generation engines.

## User/Functional Requirements
The preliminary user/functional requirements were rough, so I expanded the required functionality to provide flexibility.

1.1 Common Content and Features
* Table, border
* Cell value, cell data type (date/time/numeric/text/image/URL)
* Font size, font family(font style), alignment
* Color, background-color
* Formula, formula function
* File meta, created by, create datetime, last saved by, last saved datetime

1.2. Advanced features
* Templating (how to implement something like mail merge/repeating table body, how easy to use) 
* Print area (page break for printing)
* Repeat Table header for printing
* Conditional Style (special color, style on condition)
* For PDF, embed custom font family file (TTF/OTF/WOFF/WOFF2/SVG) for styling the PDF text content

2. Performance
* Generate 100000 rows with plain text data
* Generate 100000 rows with formula(sub total, total) data

4. Security
* Generate file in memory stream, create file in computer physical storage (disk) is not preferred, alternative, delete the file on daily end
* Encrypt document (Require open password), if can't, alternative, compress in zip file, encrypt in zip layer
* Watermark
* For PDF, access right restriction (deny copy/print...)

# Scope of Reviewed Report Engines/Library
Try and error for study how to implmenet report(xlsx, pdf) generation in C# .net Core5

|                         | Excel              | PDF                | Status                     |
|-------------------------|--------------------|--------------------|----------------------------|
| Back End - Project Name |                    |                    |                            |
| CrystalReport           | :heavy_check_mark: | :heavy_check_mark: | Implemented                |
| JasperReport            | :heavy_check_mark: | :heavy_check_mark: | Implemented                |
| EPPlus5                 | :heavy_check_mark: |                    | Implemented                |
| ITextGroupNV            |                    | :heavy_check_mark: | Implemented                |
| Puppeteer               |                    | :heavy_check_mark: | Implemented                |
| IronPDFProject          |                    | :heavy_check_mark: | Implemented                |
| OpenXmlSDK              | :heavy_check_mark: |                    | explored and prepared wiki |
| Front End - JS Library  |                    |                    |                            |
| parallax/jsPDF          |                    | write pdf          |                            |
| mozilla/pdf.js          |                    | read pdf           |                            |
| SheetJS                 | :heavy_check_mark: |                    |                            |

P.S

In observation, the front-end JS libraries feature is limited and can't satisfy my needs compared to the library which is based on c#.
If you want a lite and the simplest report read/write function, just check out the below JavaScript/TypeScript listed library, those were not included in this project because this project is not focusing on the report read/write in front-end.

- parallax/jsPDF (write pdf)
- mozilla/pdf.js (read pdf)
- SheetJS (read/write excel)

# Demo
***Before you run (F5) on Visual Studio, make sure you read the "Installation" and complete the "Configuration"***
<br><br>
**If you want to test the Crystal Report**

change the startup project to "CoreSystemConsoleInNet"

**If you want to test others report engines**

change the startup project to "CoreSystemConsole"
<br><br>
Then, run (press F5) and wait its finish, the report should be generated and placed at the directory `tempRenderFolder`

# Configuration and setup

1. Control which reports you would like to test, comment and uncomment the lines in

> `SolutionRoot\CoreSystemConsole\Program.cs`

```
// Tick-off the Report Entity Program
//InvoiceProgram invoiceProgram = new InvoiceProgram();

//HitRateHTMLProgram hitRateHTMLProgram = new HitRateHTMLProgram();

//HitRateXMLProgram hitRateXMLProgram = new HitRateXMLProgram();

//EPPlus5XlsxTemplateProgram ePPlus5XlsxTemplateProgram = new EPPlus5XlsxTemplateProgram();

ITextGroupIPdfTemplateProgram iTextGroupIText5PdfTemplateProgram = new ITextGroupIPdfTemplateProgram();
```

2. Control the report generate folder, open and edit

> `SolutionRoot\CoreReport\VisualizationEntity.cs`

```
protected string tempRenderFolder = @"D:\\Temp"; // report will be generated in this directory
```


# Installation

Some engines installation is not required, because those library develop under pure C# and already installed by NuGet Package Manager

Some report installation is required, because those library rely on the external executable program like JAVA…

for examples, JasperReport and Crystal Report need to install

The installation steps details are described below

## Crystal Report
### Pre-installation
Before run, please install Crystal Reports, Developer for Visual Studio Downloads
https://wiki.scn.sap.com/wiki/display/BOBJ/Crystal+Reports%2C+Developer+for+Visual+Studio+Downloads

### Documentation
Connecting to Object Collections
https://help.sap.com/viewer/0d6684e153174710b8b2eb114bb7f843/SP21/en-US/45afd8f46e041014910aba7db0e91070.html

### Example and Tutorial
Tutorial: Connecting to Object Collections
https://help.sap.com/viewer/0d6684e153174710b8b2eb114bb7f843/SP21/en-US/45c50fec6e041014910aba7db0e91070.html

## Jasper Reports
### Pre-installation
Before run, please install .NET jsreport sdk(jsreport binary, jsreport local) by nuget
https://jsreport.net/learn/dotnet

### Documentation
jsreport documentation
https://jsreport.net/learn

Recipes
https://jsreport.net/learn/recipes

Templating engines
https://jsreport.net/learn/templating-engines

.Net local reporting
https://jsreport.net/learn/dotnet-local

.Net Client
https://jsreport.net/learn/dotnet-client

### Example and Tutorial
GitHub jsreport/jsreport-dotnet
https://github.com/jsreport/jsreport-dotnet

#### Page header, footer, page number
Merge dynamic header with items
 
https://playground.jsreport.net/w/admin/ihh7laK2

Merge header and footer with page numbers

https://playground.jsreport.net/w/admin/kMI4FBmw

Merge with render for every page enabled

[https://playground.jsreport.net/w/admin/1A7l_UG_](https://playground.jsreport.net/w/admin/1A7l_UG_)

## Open XML SDK
### Open XML SDK 2.5 Productivity Tool
Import the Excel file to the productivity tool, and click generate to generate a C# code. The c# code could be generate a same excel file.
if you want a table in it, you should create a simplest Excel with 1 table header and 1 table body, then modify the code for printing table body in a loop.

https://learn.microsoft.com/en-us/answers/questions/466445/where-can-i-download-open-xml-sdk-2-5-productivity

https://web.archive.org/web/20190116000204/https://www.microsoft.com/en-us/download/details.aspx?id=30425

1. install Open XML SDK 2.5 for Microsoft Office

2. install Open XML SDK 2.5 Productivity Tool for Microsoft Office


# Conclusion
In general speaking, declear you need, excel or pdf or both, read or write or both

the generation approach listed below:

For excel

- front end, use javascript to generate xlsx (in xml format), less implement time, hard to do comprehensive layout
- back end, framework/engine providing a design tool to design and save the layout as a template, allowed to feed the data set(s) to template, to 
- back end, having a template excel in back end, read and copy the template then fill your data in the cell by row/column

For pdf

- front end/back end, use javascript to call pdf api, create pdf components with coordinate (width, height, x, y), hard to handle comprehensive layout
- front end, use canvas HTML element to capture the a specific area of screen in browser and print as a pdf
- front end, use javascript to call browser print function to print your page as a pdf
- back end, convert a excel to pdf
- back end, convert html (maybe with limited css) to pdf

## For excel manipulation read/write (xlsx, xls)
After the test, I rank the tool from 1 to bigger number, 1 is the most perferable.
1. EPPlus5, api is straight forward, easy to understand and use, implemented excel like behaviors, most advcanced features (chart, pivot table, header, footer, print number, cell validation..etc), support xlsx
2. Jasper Report, Java based program, officia provide a c# wrapped for call support xlsx, xls
3. OpenXmlSDK, microsoft provides basic API, required to read dehumanized, complex, extremely long documentation
4. Crystal Report, support xlsx, xls, bad excel generation because of the design, for details please read below 4 Urls
> https://archive.sap.com/documents/docs/DOC-39608
> 
> https://userapps.support.sap.com/sap/support/knowledge/en/1198296?fbclid=IwAR0_KR9veTxUJG_LituJlLSBYrvG6BZN3_OUm-JEZSiFa9enoZp-Jysa54Q
> 
> https://answers.sap.com/questions/424754/how-to-merge-columns-when-exporting-crystal-report.html?fbclid=IwAR0WjV8zsw_6Fd5OG3s-BNCyzbVuYToHD1xCMIgh0O1mNRFIqbEXSCrlcUA
> 
> https://stackoverflow.com/questions/28045209/can-grow-proprity-of-a-crystal-report-field-doesnt-push-down-lines-correctly?fbclid=IwAR2KEHM-rtmA-FHfun3NrsS_rDZLdVotuiy-14u_u7ih7vbgcjLsUoGQejA

## For pdf manipulation write
After the test, I rank the tool from 1 to bigger number, 1 is the most perferable.
1. puppeteer/puppeteer, PejmanNik/puppeteer-report
2. Crystal Report
3. iText7
4. Jasper Report
