"use strict";

const path = require("path");
const puppeteer = require("puppeteer");

// print process.argv
var args = process.argv.slice(2);
console.dir(args.length);
console.dir(args);
args.forEach(function (val, index, array) {
  console.log(index + ': ' + val);
});
	  
var htmlFile;
var pdf_dest_file;

if(args.length >=2){
	htmlFile = args[0];
	pdf_dest_file = args[1];
}else{
	htmlFile = path.resolve("D:/Documents/ReportEngine/SolutionRoot/ITextGroupNV/ReportTemplate/PocFileList/index.html");
	htmlFile = path.resolve("D:/Documents/ReportEngine/SolutionRoot/Puppeteer/ReportTemplate/ReportReference1/index.html");

	pdf_dest_file = path.resolve("./puppeteer.report.pdf");
}

(async () => {
  const browser = await puppeteer.launch();
  const page = await browser.newPage();
  await page.goto("file://" + htmlFile);
  await page.pdf({
		path: pdf_dest_file,
		format: "a4",
		printBackground: true,
		displayHeaderFooter: true,
		  headerTemplate: "<div/>",
		  footerTemplate: `<div style="width: 210mm;text-align: right; font-size: 10px;"><span style="margin-right: 1rem"><span class="pageNumber"></span> of <span class="totalPages"></span> Pages</span></div>`
	});
  await browser.close();
})();