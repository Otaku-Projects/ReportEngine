# Brief of background
In charge of developing the whole report generation module for a new system, going to explore, try, and test several report (excel/pdf) generation engines.

## User/Functional Requirements
The preliminary user/functional requirements were unclear, so I included many functionalities from basic to advanced that going for testing.

Then try to find the best suitable tools for generate Excel and/or PDF, it is not necessary to have a tool that can generate both Excel and pdf 
(I think that would be expensive if one contains almost all listed features and able to generate in Excel and PDF). 

A tool for Excel, a tool for PDF would also be accepted.

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
<code style="color : red">If you have time before all you start</code>

$${\color{red}Read \space the \space documentation \space in \space "Develop \space Guide" \space Folder}$$

Some engines develop under pure C# and/or libraries, components already installed by NuGet Package Manager

except some of those rely on the external executable program like JAVA…

for examples, JasperReport, Crystal Report, EPPlus, Open XML need to install 3rd party SDK / library to make it work

to save the time, just right click on the project on VS Solution Explorer, select unload the project<br>
For CoreSystemConsole, remove the project under Dependencies > Projects to ignore in build<br>
For CoreSystemConsoleInNet, remove the project under Reference to ignore in build<br>

The installation beief are described in Wiki, please read "/Developer Guide/Developer Document.docx" for the details

# Conclusion
Keep this page clear and short, this moved to Wiki

https://github.com/Otaku-Projects/ReportEngine/wiki
