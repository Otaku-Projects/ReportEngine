"use strict";

const path = require("path");
const report = require("puppeteer-report");
const puppeteer = require("puppeteer");

// print process.argv
var args = process.argv.slice(2);

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
	htmlFile = path.resolve("D:/Documents/ReportEngine/SolutionRoot/Puppeteer/ReportTemplate/ReportReference2/index.html");

	pdf_dest_file = path.resolve("./puppeteer.report.pdf");
}

(async () => {
	const browser = await puppeteer.launch({
	  args: ["--no-sandbox", "--disable-setuid-sandbox", "--disable-dev-shm-usage"],
	});

	try {
	  
	  // Generates a PDF with 'screen' media type.
	  await report.pdf(browser, htmlFile, {
		path: pdf_dest_file,
		format: "a4",
		printBackground: true,
	  });
	} finally {
	  await browser.close();
	}
})();