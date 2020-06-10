/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Data.SQLite;
using System.Threading.Tasks;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using OfficeOpenXml.Drawing;
using System.Collections.Generic;
using OfficeOpenXml.Table;
using System.Data;
using System.IO;
using OfficeOpenXml.Drawing.Chart.ChartEx;

namespace EPPlusSamples
{
    class ChartsAndThemesSample
    {
        private class RegionData
        {
            public string Region { get; set; }
            public int SoldUnits { get; set; }
            public double TotalSales { get; set; }
            public double Margin { get; set; }
        }
        /// <summary>
        /// Sample 15 - Creates various charts and apply a theme if supplied in the parameter themeFile.
        /// </summary>
        public static async Task<string> RunAsync(string connectionString, FileInfo xlFile, FileInfo themeFile)
        {
            using (var package = new ExcelPackage())
            {
                //Load a theme file if set. Thmx files can be exported from Excel. This will change the appearance for the workbook.                
                if (themeFile != null)
                {
                    package.Workbook.ThemeManager.Load(themeFile);
                    /*** Themes can also be altered. For example, uncomment this code to set the Accent1 to a blue color ***/
                    //package.Workbook.ThemeManager.CurrentTheme.ColorScheme.Accent1.SetRgbColor(Color.FromArgb(32, 78, 224));
                }

                /*********************************************************************************************************
                 * About chart styles: 
                 * 
                 * Chart styles can be applied to charts using the Chart.StyleManager.SetChartMethod method.
                 * The chart styles can either be set by the two enums ePresetChartStyle and ePresetChartStyleMultiSeries or by setting the Chart Style Number.
                 * 
                 * Note: Chart styles in Excel changes depending on many parameters (like number of series, axis types and more), so the enums will not always reflect the style index in Excel. 
                 * The enums are for the most common scenarios.
                 * If you want to reflect a specific style please use the Chart Style Number for the chart in Excel. 
                 * The chart style number can be fetched by recording a macro in Excel and click the style you want to apply.
                 * 
                 * Chart style do not alter visibility of chart objects like datalabels or chart titles like Excel do. That must be set in code before setting the style.
                 *********************************************************************************************************/

                //The first method adds a worksheet with four 3D charts with different styles. The last chart applies an exported chart template file (*.crtx) to the chart.
                await Add3DCharts(connectionString, package);
                
                //This method adds four line charts with different chart elements like up-down bars, error bars, drop lines and high-low lines.
                await AddLineCharts(connectionString, package);
                
                //Adds a scatter chart with a moving average trendline.
                AddScatterChart(package);
                
                //Adds a bubble-chartsheet
                AddBubbleChartsWorksheet(package);
                
                //Adds a radar chart
                AddRadarChart(package);

                //Adds a sunburst and a treemap chart 
                await AddSunburstAndTreemapChart(connectionString, package);

                //Add a box & whisker and a histogram chart 
                AddBoxWhiskerAndParetoHistogramChart(package);

                // Add a waterfall chart
                AddWaterfallChart(package);

                // Add a funnel chart
                AddFunnelChart(package);

                await AddRegionMapChart(connectionString, package);

                //Add an area chart using a chart template (chrx file)
                await AddAreaFromChartTemplate(connectionString, package);

                //Save our new workbook in the output directory and we are done!
                package.SaveAs(xlFile);
                return xlFile.FullName;
            }
        }

        private static async Task AddSunburstAndTreemapChart(string connectionString, ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Sunburst & Treemap Chart");
            var range = await LoadSalesFromDatabase(connectionString, ws);

            var sunburstChart = ws.Drawings.AddSunburstChart("SunburstChart1");
            var sbSerie = sunburstChart.Series.Add(ws.Cells[2, 4, range.Rows, 4], ws.Cells[2, 1, range.Rows, 3]);
            sbSerie.HeaderAddress = ws.Cells["D1"];
            sunburstChart.SetPosition(1, 0, 6, 0);
            sunburstChart.SetSize(800, 800);
            sunburstChart.Title.Text = "Sales";            
            sunburstChart.Legend.Add();
            sunburstChart.Legend.Position = eLegendPosition.Bottom;
            sbSerie.DataLabel.Add(true, true);
            sunburstChart.StyleManager.SetChartStyle(ePresetChartStyle.SunburstChartStyle4);


            var treemapChart = ws.Drawings.AddTreemapChart("TreemapChart1");
            var tmSerie = treemapChart.Series.Add(ws.Cells[2, 4, range.Rows, 4], ws.Cells[2, 1, range.Rows, 3]);
            treemapChart.Title.Font.Fill.Style = eFillStyle.SolidFill;
            treemapChart.Title.Font.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Background2);
            tmSerie.HeaderAddress = ws.Cells["D1"];
            treemapChart.SetPosition(1, 0, 19, 0);
            treemapChart.SetSize(1000, 800);
            treemapChart.Title.Text = "Sales";
            treemapChart.Legend.Add();
            treemapChart.Legend.Position = eLegendPosition.Right;
            tmSerie.DataLabel.Add(true, true);
            tmSerie.ParentLabelLayout = eParentLabelLayout.Banner;
            treemapChart.StyleManager.SetChartStyle(ePresetChartStyle.TreemapChartStyle3);
            
        }
        private static void AddBoxWhiskerAndParetoHistogramChart(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("BoxAndWhiskerChart");
            AddBoxWhiskerData(ws);

            var boxWhiskerChart = ws.Drawings.AddBoxWhiskerChart("BoxAndWhisker1");
            var bwSerie1 = boxWhiskerChart.Series.Add(ws.Cells[2, 1, 11, 1], null);
            bwSerie1.HeaderAddress = ws.Cells["A1"];
            var bwSerie2 = boxWhiskerChart.Series.Add(ws.Cells[2, 2, 11, 2], null);
            bwSerie2.HeaderAddress = ws.Cells["B1"];
            var bwSerie3 = boxWhiskerChart.Series.Add(ws.Cells[2, 3, 11, 3], null);
            bwSerie3.HeaderAddress = ws.Cells["C1"];
            boxWhiskerChart.SetPosition(1, 0, 6, 0);
            boxWhiskerChart.SetSize(800, 800);
            boxWhiskerChart.Title.Text = "Number series";
            boxWhiskerChart.StyleManager.SetChartStyle(ePresetChartStyle.BoxWhiskerChartStyle4);

            var histogramChart = ws.Drawings.AddHistogramChart("Pareto", true);
            histogramChart.SetPosition(1, 0, 19, 0);
            histogramChart.SetSize(800, 800);
            var hgSerie = histogramChart.Series.Add(ws.Cells[2, 3, 15, 3], null);
            hgSerie.HeaderAddress = ws.Cells["C1"];
            hgSerie.Binning.Size = 4;
        }

        private static void AddWaterfallChart(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("WaterfallChart");

            ws.SetValue("A1", "Description");
            ws.SetValue("A2", "Initial Saldo");
            ws.SetValue("A3", "Food");
            ws.SetValue("A4", "Beer");
            ws.SetValue("A5", "Transfer");
            ws.SetValue("A6", "Electrical Bill");
            ws.SetValue("A7", "Cell Phone");
            ws.SetValue("A8", "Car Repair");

            ws.SetValue("B1", "Saldo/transaction");
            ws.SetValue("B2", 1000);
            ws.SetValue("B3", -237.5);
            ws.SetValue("B4", -33.75);
            ws.SetValue("B5", 200);
            ws.SetValue("B6", -153.4);
            ws.SetValue("B7", -49);
            ws.SetValue("B8", -258.47);
            ws.Cells["B9"].Formula="SUM(B2:B8)";
            ws.Calculate();
            ws.Cells.AutoFitColumns();
            var waterfallChart = ws.Drawings.AddWaterfallChart("Waterfall1");
            waterfallChart.Title.Text = "Saldo and Transaction";
            waterfallChart.SetPosition(1, 0, 6, 0);
            waterfallChart.SetSize(800, 400);
            var wfSerie = waterfallChart.Series.Add(ws.Cells[2, 2, 9, 2], ws.Cells[2, 1, 8, 1]);

            var dp=wfSerie.DataPoints.Add(0);
            dp.SubTotal = true;
            dp = wfSerie.DataPoints.Add(7);
            dp.SubTotal = true;
        }
        private static void AddFunnelChart(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("FunnelChart");

            ws.SetValue("A1", "Stage");
            ws.SetValue("A2", "Leads");
            ws.SetValue("A3", "Prospects");
            ws.SetValue("A4", "Meeting");
            ws.SetValue("A5", "Negotiation");
            ws.SetValue("A6", "Project");
            ws.SetValue("A7", "Close");

            ws.SetValue("B1", "Number");
            ws.SetValue("B2", 3500);
            ws.SetValue("B3", 1000);
            ws.SetValue("B4", 200);
            ws.SetValue("B5", 100);
            ws.SetValue("B6", 95);
            ws.SetValue("B7", 92);
            ws.Tables.Add(ws.Cells["A1:B7"], "SalesTable");
            ws.Cells.AutoFitColumns();

            var funnelChart = ws.Drawings.AddFunnelChart("FunnelChart");
            funnelChart.Title.Text = "Sales process";
            funnelChart.SetPosition(1, 0, 6, 0);
            funnelChart.SetSize(800, 400);            
            var fSerie = funnelChart.Series.Add(ws.Cells[2, 2, 7, 2], ws.Cells[2, 1, 7, 1]);
            fSerie.DataLabel.Add(false, true);
        }
        private static async Task AddRegionMapChart(string connectionString, ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("RegionMapChart");

            var range = await LoadSalesFromDatabase(connectionString, ws);

            //Region map charts 
            var regionChart = ws.Drawings.AddRegionMapChart("RegionMapChart");
            regionChart.Title.Text = "Sales";
            regionChart.SetPosition(1, 0, 6, 0);
            regionChart.SetSize(1200, 600);
            
            var rmSerie = regionChart.Series.Add(ws.Cells[2, 4, range.End.Row, 4], ws.Cells[2, 1, range.End.Row, 3]);

            rmSerie.ColorBy = eColorBy.Value;
            
            //Color settings only apply when color by is Value
            rmSerie.Colors.NumberOfColors = eNumberOfColors.ThreeColor;
            rmSerie.Colors.MinColor.Color.SetSchemeColor(eSchemeColor.Accent3);
            rmSerie.Colors.MidColor.Color.SetHslColor(180, 50, 50);
            rmSerie.Colors.MidColor.ValueType = eColorValuePositionType.Number;
            rmSerie.Colors.MidColor.PositionValue = 200;
            rmSerie.Colors.MaxColor.Color.SetRgbPercentageColor(75, 25, 25);
            rmSerie.Colors.MaxColor.ValueType = eColorValuePositionType.Percent;
            rmSerie.Colors.MaxColor.PositionValue = 85;

            rmSerie.ProjectionType = eProjectionType.Mercator;
        }
        private static void AddBoxWhiskerData(ExcelWorksheet ws)
        {
            ws.Cells["A1"].Value = "Primes";
            ws.Cells["A2"].Value = 2;
            ws.Cells["A3"].Value = 3;
            ws.Cells["A4"].Value = 5;
            ws.Cells["A5"].Value = 7;
            ws.Cells["A6"].Value = 11;
            ws.Cells["A7"].Value = 13;
            ws.Cells["A8"].Value = 17;
            ws.Cells["A9"].Value = 19;
            ws.Cells["A10"].Value = 23;
            ws.Cells["A11"].Value = 29;
            ws.Cells["A12"].Value = 31;
            ws.Cells["A13"].Value = 37;
            ws.Cells["A14"].Value = 41;
            ws.Cells["A15"].Value = 43;

            ws.Cells["B1"].Value = "Even";
            ws.Cells["B2"].Value = 2;
            ws.Cells["B3"].Value = 4;
            ws.Cells["B4"].Value = 6;
            ws.Cells["B5"].Value = 8;
            ws.Cells["B6"].Value = 10;
            ws.Cells["B7"].Value = 12;
            ws.Cells["B8"].Value = 14;
            ws.Cells["B9"].Value = 16;
            ws.Cells["B10"].Value = 18;
            ws.Cells["B11"].Value = 20;
            ws.Cells["B12"].Value = 22;
            ws.Cells["B13"].Value = 24;
            ws.Cells["B14"].Value = 26;
            ws.Cells["B15"].Value = 28;

            ws.Cells["C1"].Value = "Random";
            ws.Cells["C2"].Value = 2;
            ws.Cells["C3"].Value = 3;
            ws.Cells["C4"].Value = 7;
            ws.Cells["C5"].Value = 12;
            ws.Cells["C6"].Value = 15;
            ws.Cells["C7"].Value = 18;
            ws.Cells["C8"].Value = 19;
            ws.Cells["C9"].Value = 23;
            ws.Cells["C10"].Value = 25;
            ws.Cells["C11"].Value = 30;
            ws.Cells["C12"].Value = 35;
            ws.Cells["C13"].Value = 37;
            ws.Cells["C14"].Value = 40;
            ws.Cells["C15"].Value = 42;
        }

        private static async Task AddAreaFromChartTemplate(string connectionString, ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Area chart from template");
            var range = await LoadFromDatabase(connectionString, ws);

            //Add an Area chart from a template file. The crtx file has it's own theme, so it does not change with the theme.
            var areaChart = (ExcelAreaChart)ws.Drawings.AddChartFromTemplate(FileInputUtil.GetFileInfo("15-ChartsAndThemes", "AreaChartStyle3.crtx"), "areaChart");
            var areaSerie = areaChart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 16, 1]);
            areaSerie.Header = "Order Value";
            areaChart.SetPosition(1, 0, 6, 0);
            areaChart.SetSize(1200, 400);
            areaChart.Title.Text = "Area Chart";

            range.AutoFitColumns(0);
        }

        private static void AddScatterChart(ExcelPackage package)
        {
            //Add a scatter chart on the data with one serie per row. 
            var ws = package.Workbook.Worksheets.Add("Scatter Chart");
            
            CreateIceCreamData(ws);

            var chart = ws.Drawings.AddScatterChart("ScatterChart1", eScatterChartType.XYScatter);
            chart.SetPosition(1, 0, 3, 0);
            chart.To.Column = 18;
            chart.To.Row = 20;
            chart.XAxis.Format = "yyyy-mm";
            chart.XAxis.Title.Text = "Period";
            chart.XAxis.MajorGridlines.Width = 1;
            chart.YAxis.Format = "$#,##0";
            chart.YAxis.Title.Text = "Sales";

            chart.Legend.Position = eLegendPosition.Bottom;

            var serie = chart.Series.Add(ws.Cells[3, 2, 14, 2], ws.Cells[3, 1, 14, 1]);
            serie.HeaderAddress = ws.Cells["A1"];
            var tr=serie.TrendLines.Add(eTrendLine.MovingAvgerage);
            tr.Name = "Icecream Sales-Monthly Average";
            chart.StyleManager.SetChartStyle(ePresetChartStyle.ScatterChartStyle12);
        }

        private static void AddBubbleChartsWorksheet(ExcelPackage package)
        {
            ExcelWorksheet wsData = LoadBubbleChartData(package);

            //Add a bubble chart on the data with one serie per row. 
            var wsChart = package.Workbook.Worksheets.AddChart("Bubble Chart", eChartType.Bubble);
            var chart = ((ExcelBubbleChart)wsChart.Chart);
            for (int row = 2; row <= 7; row++)
            {
                var serie = chart.Series.Add(wsData.Cells[row, 2], wsData.Cells[row, 3], wsData.Cells[row, 4]);
                serie.HeaderAddress = wsData.Cells[row, 1];
            }

            chart.DataLabel.Position = eLabelPosition.Center;
            chart.DataLabel.ShowSeriesName = true;
            chart.DataLabel.ShowBubbleSize = true;
            chart.Title.Text = "Sales per Region";
            chart.XAxis.Title.Text = "Total Sales";
            chart.XAxis.Title.Font.Size = 12;
            chart.XAxis.MajorGridlines.Width = 1;
            chart.YAxis.Title.Text = "Sold Units";
            chart.YAxis.Title.Font.Size = 12;
            chart.Legend.Position = eLegendPosition.Bottom;

            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.BubbleChartStyle10);
        }

        private static async Task Add3DCharts(string connectionString, ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("3D Charts");

            var range = await LoadFromDatabase(connectionString, ws);

            //Add a column chart
            var chart = ws.Drawings.AddBarChart("column3dChart", eBarChartType.ColumnClustered3D);
            var serie = chart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 26, 1]);
            serie.Header = "Order Value";
            chart.SetPosition(0, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Column Chart 3D";

            //Set style 9 and Colorful Palette 3. 
            chart.StyleManager.SetChartStyle(ePresetChartStyle.Column3dChartStyle9, ePresetChartColors.ColorfulPalette3);

            //Add a line chart
            var lineChart = ws.Drawings.AddLineChart("line3dChart", eLineChartType.Line3D);
            var lineSerie = lineChart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 16, 1]);
            lineSerie.Header = "Order Value";
            lineChart.SetPosition(21, 0, 6, 0);
            lineChart.SetSize(1200, 400);
            lineChart.Title.Text = "Line 3D";
            //Set Line3D Style 1
            lineChart.StyleManager.SetChartStyle(ePresetChartStyle.Line3dChartStyle1);

            //Add a bar chart
            chart = ws.Drawings.AddBarChart("bar3dChart", eBarChartType.BarStacked3D);
            serie = chart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 16, 1]);
            serie.Header = "Order Value";
            serie = chart.Series.Add(ws.Cells[2, 3, 16, 3], ws.Cells[2, 1, 16, 1]);
            serie.Header = "Tax";
            serie = chart.Series.Add(ws.Cells[2, 4, 16, 4], ws.Cells[2, 1, 16, 1]);
            serie.Header = "Freight";

            chart.SetPosition(42, 0, 6, 0);
            chart.SetSize(1200, 600);
            chart.Title.Text = "Bar Chart 3D";
            //Set the color
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.StackedBar3dChartStyle7, ePresetChartColors.ColorfulPalette1);

            range.AutoFitColumns(0);
        }
        private static async Task AddLineCharts(string connectionString, ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("LineCharts");

            var range = await LoadFromDatabase(connectionString, ws);

            //Add a line chart
            var chart = ws.Drawings.AddLineChart("LineChartWithDroplines", eLineChartType.Line);
            var serie = chart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 16, 1]);
            serie.Header = "Order Value";
            chart.SetPosition(0, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Line Chart With Droplines";
            chart.AddDropLines();
            chart.DropLine.Border.Width = 2;
            //Set style 12
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle12);

            //Add a line chart with Error Bars
            chart = ws.Drawings.AddLineChart("LineChartWithErrorBars", eLineChartType.Line);
            serie = chart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 16, 1]);
            serie.Header = "Order Value";
            chart.SetPosition(21, 0, 6, 0);
            chart.SetSize(1200, 400);   //Make this chart wider to make room for the datatable.
            chart.Title.Text = "Line Chart With Error Bars";
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Percentage);
            serie.ErrorBars.Value = 5;
            chart.PlotArea.CreateDataTable();

            //Set style 2
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle2);

            //Add a line chart with Error Bars
            chart = ws.Drawings.AddLineChart("LineChartWithUpDownBars", eLineChartType.Line);
            var serie1 = chart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 16, 1]);
            serie1.Header = "Order Value";
            var serie2 = chart.Series.Add(ws.Cells[2, 3, 16, 3], ws.Cells[2, 1, 16, 1]);
            serie2.Header = "Tax";
            var serie3 = chart.Series.Add(ws.Cells[2, 4, 16, 4], ws.Cells[2, 1, 16, 1]);
            serie3.Header = "Freight";
            chart.SetPosition(42, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Line Chart With Up/Down Bars";
            chart.AddUpDownBars(true, true);

            //Set style 10, Note: As this is a line chart with multiple series, we use the enum for multiple series. Charts with multiple series usually has a subset of of the chart styles in Excel.
            //Another option to set the style is to use the Excel Style number, in this case 236: chart.StyleManager.SetChartStyle(236)
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.LineChartStyle9);
            range.AutoFitColumns(0);


            //Add a line chart with high/low Bars
            chart = ws.Drawings.AddLineChart("LineChartWithHighLowLines", eLineChartType.Line);
            serie1 = chart.Series.Add(ws.Cells[2, 2, 26, 2], ws.Cells[2, 1, 26, 1]);
            serie1.Header = "Order Value";
            serie2 = chart.Series.Add(ws.Cells[2, 3, 26, 3], ws.Cells[2, 1, 26, 1]);
            serie2.Header = "Tax";
            serie3 = chart.Series.Add(ws.Cells[2, 4, 26, 4], ws.Cells[2, 1, 26, 1]);
            serie3.Header = "Freight";
            chart.SetPosition(63, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Line Chart With High/Low Lines";
            chart.AddHighLowLines();

            //Set the style using the Excel ChartStyle number. The chart style must exist in the ExcelChartStyleManager.StyleLibrary[]. 
            //Styles can be added and removed from this library. By default it is loaded with the styles for EPPlus supported chart types.
            chart.StyleManager.SetChartStyle(237);
            range.AutoFitColumns(0);
        }
        private static void AddRadarChart(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("RadarChart");

            var dt=GetCarDataTable();
            ws.Cells["A1"].LoadFromDataTable(dt, true);
            ws.Cells.AutoFitColumns();

            var chart = ws.Drawings.AddRadarChart("RadarChart1", eRadarChartType.RadarFilled);
            var serie = chart.Series.Add(ws.Cells["B2:B5"], ws.Cells["A2:A5"]);
            serie.HeaderAddress = ws.Cells["B1"];
            serie = chart.Series.Add(ws.Cells["C2:C5"], ws.Cells["A2:A5"]);
            serie.HeaderAddress = ws.Cells["C1"];
            serie = chart.Series.Add(ws.Cells["D2:D5"], ws.Cells["A2:A5"]);
            serie.HeaderAddress = ws.Cells["D1"];
            serie = chart.Series.Add(ws.Cells["E2:E5"], ws.Cells["A2:A5"]);
            serie.HeaderAddress = ws.Cells["E1"];

            chart.Legend.Position = eLegendPosition.Top;
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.RadarChartStyle4);

            //If you want to apply custom styling do that after setting the chart style so its not overwritten.
            chart.Legend.Effect.SetPresetShadow(ePresetExcelShadowType.OuterTopLeft);

            chart.SetPosition(0, 0, 6, 0);
            chart.To.Column = 17;
            chart.To.Row = 30;
        }

        private static DataTable GetCarDataTable()
        {
            var dt = new DataTable();
            dt.Columns.Add("Car", typeof(string));
            dt.Columns.Add("Acceleration Index", typeof(int));
            dt.Columns.Add("Size Index", typeof(int));
            dt.Columns.Add("Polution Index", typeof(int));
            dt.Columns.Add("Retro Index", typeof(int));
            dt.Rows.Add("Volvo 242", 1, 3, 4, 4);
            dt.Rows.Add("Lamborghini Countach", 5, 1, 5, 4);
            dt.Rows.Add("Tesla Model S", 5, 2, 1, 1);
            dt.Rows.Add("Hummer H1", 2, 5, 5, 2);

            return dt;
        }

        private static async Task<ExcelRangeBase> LoadFromDatabase(string connectionString, ExcelWorksheet ws)
        {
            ExcelRangeBase range;
            using (var sqlConn = new SQLiteConnection(connectionString))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select orderdate as OrderDate, SUM(ordervalue) as OrderValue, SUM(tax) As Tax,SUM(freight) As Freight from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId Where Currency='USD' group by OrderDate ORDER BY OrderDate desc limit 15", sqlConn))
                { 
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        range = await ws.Cells["A1"].LoadFromDataReaderAsync(sqlReader, true);
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = true;
                        range.Offset(0, 0, range.Rows, 1).Style.Numberformat.Format = "yyyy-MM-dd";
                    }
                    //Set the numberformat
                }
            }
            return range;
        }
        private static async Task<ExcelRangeBase> LoadSalesFromDatabase(string connectionString, ExcelWorksheet ws)
        {
            ExcelRangeBase range;
            using (var sqlConn = new SQLiteConnection(connectionString))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select s.continent, s.country, s.city, SUM(OrderValue) As Sales from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId Where Currency='USD' group by s.continent, s.country, s.city ORDER BY s.continent, s.country, s.city", sqlConn))
                {
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        range = await ws.Cells["A1"].LoadFromDataReaderAsync(sqlReader, true);
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = true;
                        range.Offset(0, 3, range.Rows, 3).Style.Numberformat.Format = "#,##0";
                    }
                    //Set the numberformat
                }
            }
            return range;
        }

        private static void CreateIceCreamData(ExcelWorksheet ws)
        {
            ws.SetValue("A1", "Icecream Sales-2019");
            ws.SetValue("A2", "Date");
            ws.SetValue("B2", "Sales");
            ws.SetValue("A3", new DateTime(2019, 1, 1));
            ws.SetValue("B3", 2500);
            ws.SetValue("A4", new DateTime(2019, 2, 1));
            ws.SetValue("B4", 3000);
            ws.SetValue("A5", new DateTime(2019, 3, 1));
            ws.SetValue("B5", 2700);
            ws.SetValue("A6", new DateTime(2019, 4, 1));
            ws.SetValue("B6", 4400);
            ws.SetValue("A7", new DateTime(2019, 5, 1));
            ws.SetValue("B7", 6900);
            ws.SetValue("A8", new DateTime(2019, 6, 1));
            ws.SetValue("B8", 11200);
            ws.SetValue("A9", new DateTime(2019, 7, 1));
            ws.SetValue("B9", 13200);
            ws.SetValue("A10", new DateTime(2019, 8, 1));
            ws.SetValue("B10", 12400);
            ws.SetValue("A11", new DateTime(2019, 9, 1));
            ws.SetValue("B11", 8700);
            ws.SetValue("A12", new DateTime(2019, 10, 1));
            ws.SetValue("B12", 4800);
            ws.SetValue("A13", new DateTime(2019, 11, 1));
            ws.SetValue("B13", 2000);
            ws.SetValue("A14", new DateTime(2019, 12, 1));
            ws.SetValue("B14", 2400);
            ws.Cells["A3:A14"].Style.Numberformat.Format = "yyyy-MM";
            ws.Cells["B3:B14"].Style.Numberformat.Format = "#,##0kr";
        }
        private static ExcelWorksheet LoadBubbleChartData(ExcelPackage package)
        {
            var data = new List<RegionData>()
            {
                    new RegionData(){ Region = "North", SoldUnits=500, TotalSales=4800, Margin=0.200 },
                    new RegionData(){ Region = "Central", SoldUnits=900, TotalSales=7330, Margin=0.333 },
                    new RegionData(){ Region = "South", SoldUnits=400, TotalSales=3700, Margin=0.150 },
                    new RegionData(){ Region = "East", SoldUnits=350, TotalSales=4400, Margin=0.102 },
                    new RegionData(){ Region = "West", SoldUnits=700, TotalSales=6900, Margin=0.218 },
                    new RegionData(){ Region = "Stockholm", SoldUnits=1200, TotalSales=8250, Margin=0.350 }
            };
            var wsData = package.Workbook.Worksheets.Add("ChartData");
            wsData.Cells["A1"].LoadFromCollection(data, true, TableStyles.Medium15);
            wsData.Cells["B2:C7"].Style.Numberformat.Format = "#,##0";
            wsData.Cells["D2:D7"].Style.Numberformat.Format = "#,##0.00%";

            var shape = wsData.Drawings.AddShape("Shape1", eShapeStyle.Rect);
            shape.Text = "This worksheet contains the data for the bubble-chartsheet";
            shape.SetPosition(1, 0, 6, 0);
            shape.Effect.SetPresetShadow(ePresetExcelShadowType.OuterBottomLeft);
            return wsData;
        }

    }
}
