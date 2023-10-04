using System;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using ReportBuilder;
using System.Xml.Linq;


using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.IO;
using Cell = DocumentFormat.OpenXml.Spreadsheet.Cell;
using System.Linq;
using DocumentFormat.OpenXml.Vml.Office;
using System.Reflection.Emit;

Console.WriteLine("Hello, World!");



//string filePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "output.xlsx");
//string fileName = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "PieChartExample.xlsx");

// Create a new Excel document.
// Specify the file name and path to save in the Downloads folder.
//string fileName = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), "Downloads", "PieChartExample.xlsx");

// Create a new Excel document.
// Specify the file name and path to save in the Downloads folder.
//string downloadsFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);

string filePath = System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "output.xlsx");



//  https://www.dritsoftware.com/docs/netspreadsheet/openxmlsdk/Charts/ChartTypes/InsertBarChart.html
DoFile.InsertBarChart(filePath);  // do works
//DoFile.InsertArea3DChart2(filePath);

