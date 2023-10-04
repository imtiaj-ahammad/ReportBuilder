using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing;

namespace ReportBuilder
{
    public static class DoFile
    {
        public static void AddDataToSheet(WorksheetPart worksheetPart, int[] data)
        {
            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

            for (int i = 0; i < data.Length; i++)
            {
                Row row = new Row();
                Cell cell = new Cell() { CellValue = new CellValue(data[i].ToString()), DataType = CellValues.Number };
                row.Append(cell);
                sheetData.Append(row);
            }
        }

        public static void AddBarChartToSheet(WorksheetPart worksheetPart, int[] data)
        {
            DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();

            BarChart barChart = new BarChart(
                new BarDirection() { Val = BarDirectionValues.Column },
                new BarGrouping() { Val = BarGroupingValues.Clustered },
                new VaryColors() { Val = false }
            );

            CategoryAxisData categoryAxisData = new CategoryAxisData();
            StringReference stringReference = new StringReference();
            DocumentFormat.OpenXml.Drawing.Charts.Formula formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = "Sheet1!$A$2:$A$" + (data.Length + 1) };
            stringReference.Append(formula);
            categoryAxisData.Append(stringReference);

            BarChartSeries barChartSeries = new BarChartSeries();
            barChartSeries.Append(categoryAxisData);

            DocumentFormat.OpenXml.Drawing.Charts.Values values = new DocumentFormat.OpenXml.Drawing.Charts.Values();
            NumberReference numberReference = new NumberReference();
            DocumentFormat.OpenXml.Drawing.Charts.Formula formulaValues = new DocumentFormat.OpenXml.Drawing.Charts.Formula { Text = "Sheet1!$B$2:$B$" + (data.Length + 1) };
            numberReference.Append(formulaValues);
            values.Append(numberReference);
            barChartSeries.Append(values);

            barChart.Append(barChartSeries);

            ChartSpace chartSpace = new ChartSpace();
            chartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            chartSpace.Append(barChart);

            chartPart.ChartSpace = chartSpace;

            drawingsPart.ChartParts.Append(chartPart);
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();

            worksheetPart.Worksheet.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
        }

        public static void GeneratePieChart(string filePath)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                Sheets sheets = workbookPart.Workbook.AppendChild(new Sheets());
                Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Sheet 1o" };
                sheets.Append(sheet);

                WorkbookStylesPart workbookStylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
                workbookStylesPart.Stylesheet = new Stylesheet();
                workbookStylesPart.Stylesheet.Save();

                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                drawingsPart.WorksheetDrawing = new WorksheetDrawing();
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                GenerateChartPart(chartPart);

                drawingsPart.WorksheetDrawing.Save();

                document.Close();
            }
        }

        private static void GenerateChartPart(ChartPart chartPart)
        {
            C.ChartSpace chartSpace = new C.ChartSpace();
            chartSpace.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");

            C.Chart chart = new C.Chart();
            C.PlotArea plotArea = new C.PlotArea();
            C.PieChart pieChart = new C.PieChart();

            C.PieChartSeries pieChartSeries = new C.PieChartSeries();
            C.Index index = new C.Index() { Val = new UInt32Value(0u) };
            C.Order order = new C.Order() { Val = new UInt32Value(0u) };
            C.SeriesText seriesText = new C.SeriesText();
            C.StringReference stringReference = new C.StringReference();
            C.Formula formula = new C.Formula { Text = "'Sheet 1'!$A$2:$A$4" };
            C.StringCache stringCache = new C.StringCache();
            C.PointCount pointCount = new C.PointCount() { Val = new UInt32Value(3U) };

            C.StringPoint stringPoint1 = new C.StringPoint() { Index = (UInt32Value)0U };
            C.NumericValue numericValue1 = new C.NumericValue();
            numericValue1.Text = "Apples";
            stringPoint1.Append(numericValue1);

            C.StringPoint stringPoint2 = new C.StringPoint() { Index = (UInt32Value)1U };
            C.NumericValue numericValue2 = new C.NumericValue();
            numericValue2.Text = "Bananas";
            stringPoint2.Append(numericValue2);

            C.StringPoint stringPoint3 = new C.StringPoint() { Index = (UInt32Value)2U };
            C.NumericValue numericValue3 = new C.NumericValue();
            numericValue3.Text = "Cherries";
            stringPoint3.Append(numericValue3);

            stringCache.Append(pointCount);
            stringCache.Append(stringPoint1);
            stringCache.Append(stringPoint2);
            stringCache.Append(stringPoint3);
            stringReference.Append(formula);
            stringReference.Append(stringCache);
            seriesText.Append(stringReference);

            pieChartSeries.Append(index);
            pieChartSeries.Append(order);
            pieChartSeries.Append(seriesText);
            pieChart.Append(pieChartSeries);
            plotArea.Append(pieChart);
            chart.Append(plotArea);

            C.Legend legend = new C.Legend();
            C.LegendPosition legendPosition = new C.LegendPosition() { Val = C.LegendPositionValues.Right };
            legend.Append(legendPosition);
            chart.Append(legend);

            chartSpace.Append(chart);
            chartPart.ChartSpace = chartSpace;
        }

        #region barChart
        public static void InsertBarChart(string filePath)
        {
            using (var spreadsheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                Workbook workbook = new Workbook();
                workbookPart.Workbook = workbook;
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet();
                FileVersion fv = new FileVersion();
                fv.ApplicationName = "Microsoft Office Excel";

                Sheets sheets = new Sheets();
                Sheet sheet = new Sheet();
                sheet.Name = "Sheet1";
                sheet.SheetId = 1;
                sheet.Id = workbookPart.GetIdOfPart(worksheetPart);
                sheets.Append(sheet);
                workbook.Append(fv);
                workbook.Append(sheets);

                AddData(workbookPart, worksheetPart);

                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
                worksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
                { Id = worksheetPart.GetIdOfPart(drawingsPart) });


                // Add a new chart and set the chart language to English-US.
                ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
                chartPart.ChartSpace = new ChartSpace();
                chartPart.ChartSpace.RoundedCorners = new RoundedCorners() { Val = false };

                //Set default style
                var alternateContent = chartPart.ChartSpace.AppendChild(new AlternateContent());
                var choice = alternateContent.AppendNewAlternateContentChoice();
                var c14Style = new DocumentFormat.OpenXml.Office2010.Drawing.Charts.Style() { Val = 102 };
                choice.Requires = "c14";
                choice.AppendChild(c14Style);
                var fallback = alternateContent.AppendNewAlternateContentFallback();
                var cStyle = new DocumentFormat.OpenXml.Drawing.Charts.Style() { Val = 2 };
                fallback.AppendChild(cStyle);


                //chartPart.ChartSpace.EditingLanguage = new EditingLanguage{Val = new StringValue("en-US")};

                var chart = chartPart.ChartSpace.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart());
                chart.AutoTitleDeleted = new AutoTitleDeleted() { Val = true };
                PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
                //Layout layout = plotArea.AppendChild<Layout>(new Layout());
                var barChart = plotArea.AppendChild(new BarChart());
                barChart.BarDirection = new BarDirection() { Val = BarDirectionValues.Bar };
                barChart.AppendChild(new Grouping() { Val = new EnumValue<GroupingValues>(GroupingValues.Standard) });
                barChart.VaryColors = new VaryColors() { Val = true };


                var ser0 = barChart.AppendChild(new BarChartSeries());
                ser0.Index = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = new UInt32Value(0u) };
                ser0.Order = new Order() { Val = new UInt32Value(0u) };
                ser0.SeriesText = new SeriesText();
                ser0.SeriesText.StringReference = AddStringReference("Sheet1!$J$1", "Series 1");
                ser0.InvertIfNegative = new InvertIfNegative() { Val = false };
                var cat0 = new CategoryAxisData();
                cat0.StringReference =
                    AddStringReference("Sheet1!$I$2:$I$5", "Point 1", "Point 2", "Point 3", "Point 4");
                ser0.AppendChild(cat0);
                var val0 = new DocumentFormat.OpenXml.Drawing.Charts.Values();
                val0.NumberReference = AddNumberReference("Sheet1!$J$2:$J$5", "2", "4", "6", "8");
                ser0.AppendChild(val0);

                var ser1 = barChart.AppendChild(new BarChartSeries());
                ser1.Index = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = new UInt32Value(1u) };
                ser1.Order = new Order() { Val = new UInt32Value(1u) };
                ser1.SeriesText = new SeriesText();
                ser1.SeriesText.StringReference = AddStringReference("Sheet1!$K$1", "Series 2");
                ser1.InvertIfNegative = new InvertIfNegative() { Val = false };
                var cat1 = new CategoryAxisData();
                cat1.StringReference =
                    AddStringReference("Sheet1!$I$2:$I$5", "Point 1", "Point 2", "Point 3", "Point 4");
                ser1.AppendChild(cat1);
                var val1 = new DocumentFormat.OpenXml.Drawing.Charts.Values();
                val1.NumberReference = AddNumberReference("Sheet1!$K$2:$K$5", "4", "3", "2", "1");
                ser1.AppendChild(val1);

                var ser2 = barChart.AppendChild(new BarChartSeries());
                ser2.Index = new DocumentFormat.OpenXml.Drawing.Charts.Index() { Val = new UInt32Value(2u) };
                ser2.Order = new Order() { Val = new UInt32Value(2u) };
                ser2.SeriesText = new SeriesText();
                ser2.SeriesText.StringReference = AddStringReference("Sheet1!$L$1", "Series 3");
                ser2.InvertIfNegative = new InvertIfNegative() { Val = false };
                var cat2 = new CategoryAxisData();
                cat2.StringReference =
                    AddStringReference("Sheet1!$I$2:$I$5", "Point 1", "Point 2", "Point 3", "Point 4");
                ser2.AppendChild(cat2);
                var val2 = new DocumentFormat.OpenXml.Drawing.Charts.Values();
                val2.NumberReference = AddNumberReference("Sheet1!$L$2:$L$5", "2", "1", "-1", "-2");
                ser2.AppendChild(val2);

                barChart.Append(new GapWidth() { Val = 150 });
                barChart.Append(new Overlap() { Val = 0 });
                barChart.Append(new AxisId() { Val = 1250099810u });
                barChart.Append(new AxisId() { Val = 1687146081u });

                var catAx = plotArea.AppendChild<CategoryAxis>(new CategoryAxis());
                catAx.Append(new AxisId() { Val = 1250099810u });
                catAx.Append(new Scaling()
                { Orientation = new Orientation() { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax } });
                catAx.Append(new Delete() { Val = false });
                catAx.Append(new AxisPosition() { Val = AxisPositionValues.Bottom });
                catAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = false });
                catAx.Append(new TickLabelPosition() { Val = TickLabelPositionValues.Low });

                #region CategoryAxisFont
                catAx.TextProperties = new DocumentFormat.OpenXml.Drawing.Charts.TextProperties();
                catAx.TextProperties.BodyProperties = new DocumentFormat.OpenXml.Drawing.BodyProperties();
                catAx.TextProperties.BodyProperties.RightToLeftColumns = new BooleanValue(false);
                catAx.TextProperties.BodyProperties.Anchor = new EnumValue<TextAnchoringTypeValues>(TextAnchoringTypeValues.Top);
                catAx.TextProperties.ListStyle = new ListStyle();
                //The actual formatting is set on the default run properties
                var catAxP = new Paragraph();
                catAxP.ParagraphProperties = new ParagraphProperties();
                var catAxDefRPr = new DefaultRunProperties();
                //Set the Bold attribute
                catAxDefRPr.Bold = true;
                var catAxDefRPrSolidFill = catAxDefRPr.AppendChild(new SolidFill());
                catAxDefRPrSolidFill.PresetColor = new PresetColor() { Val = PresetColorValues.Red };
                catAxP.ParagraphProperties.AppendChild(catAxDefRPr);
                #endregion
                catAx.Append(new CrossingAxis() { Val = new UInt32Value(1687146081u) });
                catAx.Append(new LabelOffset() { Val = new UInt16Value((ushort)100) });

                ValueAxis valAx = plotArea.AppendChild<ValueAxis>(new ValueAxis());
                valAx.Append(new AxisId() { Val = 1687146081u });
                valAx.Append(new Scaling() { Orientation = new Orientation() { Val = DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax } });
                valAx.Append(new Delete() { Val = new BooleanValue(false) });
                valAx.Append(new AxisPosition() { Val = AxisPositionValues.Bottom });
                valAx.Append(new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat() { FormatCode = "General", SourceLinked = false });
                valAx.Append(new TickLabelPosition() { Val = TickLabelPositionValues.NextTo });
                valAx.Append(new CrossingAxis() { Val = 1250099810u });
                valAx.Append(new Crosses() { Val = CrossesValues.AutoZero });
                valAx.Append(new CrossBetween() { Val = CrossBetweenValues.Between });

                chart.Append(new PlotVisibleOnly() { Val = true });
                chart.Append(new DisplayBlanksAs() { Val = DisplayBlanksAsValues.Zero });
                chart.Append(new ShowDataLabelsOverMaximum() { Val = false });

                // Save the chart part.
                chartPart.ChartSpace.Save();

                // Position the chart on the worksheet using a TwoCellAnchor object.
                drawingsPart.WorksheetDrawing = new WorksheetDrawing();
                var twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("0"),
                    new ColumnOffset("5"),
                    new RowId("0"),
                    new RowOffset("5")));
                twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("5"),
                    new ColumnOffset("152405"),
                    new RowId("12"),
                    new RowOffset("5")));

                // Append a GraphicFrame to the TwoCellAnchor object.
                var graphicFrame =
                    twoCellAnchor.AppendChild(new DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame());
                graphicFrame.Macro = "";
                graphicFrame.NonVisualGraphicFrameProperties = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties();
                graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties();
                graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Id = new UInt32Value(1u);
                graphicFrame.NonVisualGraphicFrameProperties.NonVisualDrawingProperties.Name = "Chart 1";
                graphicFrame.NonVisualGraphicFrameProperties.NonVisualGraphicFrameDrawingProperties = new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties();

                graphicFrame.Transform = new Transform();
                graphicFrame.Transform.Offset = new Offset();
                graphicFrame.Transform.Offset.X = 0;
                graphicFrame.Transform.Offset.Y = 0;
                graphicFrame.Transform.Extents = new Extents();
                graphicFrame.Transform.Extents.Cx = 0;
                graphicFrame.Transform.Extents.Cy = 0;

                graphicFrame.Graphic = new Graphic();
                graphicFrame.Graphic.GraphicData = new GraphicData();
                graphicFrame.Graphic.GraphicData.Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart";
                graphicFrame.Graphic.GraphicData.AppendChild(new ChartReference()
                { Id = drawingsPart.GetIdOfPart(chartPart) });

                twoCellAnchor.Append(new ClientData());

                drawingsPart.WorksheetDrawing.Save();
                workbookPart.Workbook.Save();
                spreadsheetDocument.Close();
            }
        }

        static NumberReference AddNumberReference(string formula, params string[] values)
        {
            var numRef = new NumberReference();

            numRef.Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formula };
            numRef.NumberingCache = new NumberingCache();
            numRef.NumberingCache.FormatCode = new FormatCode("General");
            numRef.NumberingCache.PointCount = new PointCount() { Val = new UInt32Value(Convert.ToUInt32(values.Length)) };

            uint index = 0;
            foreach (var value in values)
            {
                numRef.NumberingCache.AppendChild(new NumericPoint() { Index = new UInt32Value(index) })
                    .Append(new NumericValue(value));
                index++;
            }

            return numRef;
        }

        static StringReference AddStringReference(string formula, params string[] values)
        {
            var strRef = new StringReference();

            strRef.Formula = new DocumentFormat.OpenXml.Drawing.Charts.Formula() { Text = formula };
            strRef.StringCache = new StringCache();
            strRef.StringCache.PointCount = new PointCount() { Val = new UInt32Value(Convert.ToUInt32(values.Length)) };

            uint index = 0;
            foreach (var value in values)
            {
                strRef.StringCache.AppendChild(new StringPoint { Index = new UInt32Value(index) })
                    .Append(new NumericValue(value));
                index++;
            }

            return strRef;
        }

        static void AddData(WorkbookPart workbookPart, WorksheetPart worksheetPart)
        {
            SheetData sheetData = new SheetData();

            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = new SharedStringTable();
            sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("Series 1")));
            sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("Series 2")));
            sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("Series 3")));
            sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("Point 1")));
            sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("Point 2")));
            sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("Point 3")));
            sharedStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text("Point 4")));

            var row1 = new Row() { RowIndex = 1 };
            var row2 = new Row() { RowIndex = 2 };
            var row3 = new Row() { RowIndex = 3 };
            var row4 = new Row() { RowIndex = 4 };
            var row5 = new Row() { RowIndex = 5 };
            sheetData.Append(row1);
            sheetData.Append(row2);
            sheetData.Append(row3);
            sheetData.Append(row4);
            sheetData.Append(row5);

            var cellJ1 = new Cell() { CellReference = "J1", DataType = new EnumValue<CellValues>(CellValues.SharedString) };
            cellJ1.CellValue = new CellValue("0"); //Index to the shared string "Series 1"
            var cellK1 = new Cell() { CellReference = "K1", DataType = new EnumValue<CellValues>(CellValues.SharedString) };
            cellK1.CellValue = new CellValue("1");
            var cellL1 = new Cell() { CellReference = "L1", DataType = new EnumValue<CellValues>(CellValues.SharedString) };
            cellL1.CellValue = new CellValue("2");

            row1.Append(cellJ1);
            row1.Append(cellK1);
            row1.Append(cellL1);

            row2.Append(new Cell() { CellReference = "I2", CellValue = new CellValue("3"), DataType = new EnumValue<CellValues>(CellValues.SharedString) });
            row2.Append(new Cell() { CellReference = "J2", CellValue = new CellValue("2") });
            row2.Append(new Cell() { CellReference = "K2", CellValue = new CellValue("4") });
            row2.Append(new Cell() { CellReference = "L2", CellValue = new CellValue("2") });

            row3.Append(new Cell() { CellReference = "I3", CellValue = new CellValue("4"), DataType = new EnumValue<CellValues>(CellValues.SharedString) });
            row3.Append(new Cell() { CellReference = "J3", CellValue = new CellValue("4") });
            row3.Append(new Cell() { CellReference = "K3", CellValue = new CellValue("3") });
            row3.Append(new Cell() { CellReference = "L3", CellValue = new CellValue("1") });

            row4.Append(new Cell() { CellReference = "I4", CellValue = new CellValue("5"), DataType = new EnumValue<CellValues>(CellValues.SharedString) });
            row4.Append(new Cell() { CellReference = "J4", CellValue = new CellValue("6") });
            row4.Append(new Cell() { CellReference = "K4", CellValue = new CellValue("2") });
            row4.Append(new Cell() { CellReference = "L4", CellValue = new CellValue("-1") });

            row5.Append(new Cell() { CellReference = "I5", CellValue = new CellValue("6"), DataType = new EnumValue<CellValues>(CellValues.SharedString) });
            row5.Append(new Cell() { CellReference = "J5", CellValue = new CellValue("8") });
            row5.Append(new Cell() { CellReference = "K5", CellValue = new CellValue("1") });
            row5.Append(new Cell() { CellReference = "L5", CellValue = new CellValue("-2") });

            worksheetPart.Worksheet.Append(sheetData);
        }
        #endregion barChart


    }
}
