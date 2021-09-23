# ReportEngine
Try and error for study how to implmenet report(xlsx, pdf) generation in C# .net framework and/or .net Core5

- Crystal Report (Tested)
- Jasper Report (Testing)
- EPPlus (Planed to test)

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

## Crystal Report
### Pre-installation
Crystal Reports, Developer for Visual Studio Downloads
https://wiki.scn.sap.com/wiki/display/BOBJ/Crystal+Reports%2C+Developer+for+Visual+Studio+Downloads

### Documentation
Connecting to Object Collections
https://help.sap.com/viewer/0d6684e153174710b8b2eb114bb7f843/SP21/en-US/45afd8f46e041014910aba7db0e91070.html

### Example and Tutorial
Tutorial: Connecting to Object Collections
https://help.sap.com/viewer/0d6684e153174710b8b2eb114bb7f843/SP21/en-US/45c50fec6e041014910aba7db0e91070.html

## Jasper Reports
### Pre-installation
install .NET jsreport sdk(jsreport binary, jsreport local) by nuget
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