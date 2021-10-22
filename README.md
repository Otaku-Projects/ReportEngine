# ReportEngine
Try and error for study how to implmenet report(xlsx, pdf) generation in C# .net framework and/or .net Core5

- Crystal Report (Tested)
- Jasper Report (Tested)
- EPPlus (Testing)

## Configuation
open SolutionRoot\CoreReport\VisualizationEntity.cs

update the value
```
protected string tempRenderFolder = @"D:\\Temp"; // report will be generated in this directory
```

## Demo
- Crystal Report
Change the startup project to "CoreSystemConsoleInNet"

- Jasper Report
Change the startup project to "CoreSystemConsole"

## Program Structure
I use "Design Pattern - Decorator" to separate the coding files by reporting enginer.
>Let's said a system contains many functions, a report function represented by a menu item in navigation menu.
>
>In general, a report function provides the selection criteria, user select the criteria 
>
>Then, click "Export Xlsx" or "Export Pdf" button to generate report file in xlsx, pdf as they want.

"Decorator" Design Pattern gives a report program easy to switch the report enginer, also allows different reports use various engines in a single system

https://www.dofactory.com/net/decorator-design-pattern#realworld


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

https://playground.jsreport.net/w/admin/1A7l_UG_

## EPPlus5
EPPlus5 is open source, but you are required to purchase license for commercial use

The library(ies) were installed under the project through Package Manager Console

[EPPlus]:https://www.nuget.org/packages/EPPlus

```
Install-Package EPPlus -Version 5.8.0
```

## iText Group
Some products of iText 7 Suite is open source, but you are required to purchase license for commercial use

The library(ies) were installed under the project through Package Manager Console

[itext7]:https://github.com/itext/itext7-dotnet
[itext7.pdfhtml]:https://github.com/itext/i7n-pdfhtml

```
Install-Package itext7 -Version 7.1.16
Install-Package itext7.pdfhtml -Version 3.0.5
```
