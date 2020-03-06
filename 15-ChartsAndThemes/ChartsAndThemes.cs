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
        /// Sample 15 - Load a theme and create a column chart
        /// </summary>
        public static async Task<string> RunAsync(string connectionString)
        {
            using (var package = new ExcelPackage())
            {
                //Load a theme file (Can be exported from Excel). This will change the appearance for the workbook. Uncomment to get the default office theme.
                package.Workbook.ThemeManager.Load(FileInputUtil.GetFileInfo("15-ChartsAndThemes", "integral.thmx"));

                /*** Themes can also be altered. For example, uncomment this code to set the Accent1 to a blue color ***/
                //package.Workbook.ThemeManager.CurrentTheme.ColorScheme.Accent1.SetRgbColor(Color.FromArgb(32, 78, 224));

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

                // save our new workbook in the output directory and we are done!
                var xlFile = FileOutputUtil.GetFileInfo("15-ChartsAndThemes.xlsx");
                package.SaveAs(xlFile);
                return xlFile.FullName;
            }
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
            var ws = package.Workbook.Worksheets.Add("3D Charts With Integral Theme");

            var range = await LoadFromDatabase(connectionString, ws);

            //Add a column chart
            var chart = ws.Drawings.AddBarChart("column3dChart", eBarChartType.ColumnClustered3D);
            var serie = chart.Series.Add(ws.Cells[2, 7, 26, 7], ws.Cells[2, 6, 26, 6]);
            serie.Header = "Order Value";
            chart.SetPosition(0, 0, 10, 0);
            chart.SetSize(1000, 300);
            chart.Title.Text = "Column Chart 3D";

            //Set style 9 and Colorful Palette 3
            chart.StyleManager.SetChartStyle(ePresetChartStyle.Bar3dChartStyle9, ePresetChartColors.ColorfulPalette3);

            //Add a bar chart
            chart = ws.Drawings.AddBarChart("bar3dChart", eBarChartType.ColumnClustered3D);
            serie = chart.Series.Add(ws.Cells[2, 7, 26, 7], ws.Cells[2, 6, 26, 6]);
            serie.Header = "Order Value";
            chart.SetPosition(17, 0, 10, 0);
            chart.SetSize(1000, 300);
            chart.Title.Text = "Bar Chart 3D";
            //Set the color
            chart.StyleManager.SetChartStyle(ePresetChartStyle.Column3dChartStyle7, ePresetChartColors.MonochromaticPalette1);

            //Add a line chart
            var lineChart = ws.Drawings.AddLineChart("line3dChart", eLineChartType.Line3D);
            var lineSerie = lineChart.Series.Add(ws.Cells[2, 7, 26, 7], ws.Cells[2, 6, 26, 6]);
            lineSerie.Header = "Order Value";
            lineChart.SetPosition(34, 0, 10, 0);
            lineChart.SetSize(1000, 300);
            lineChart.Title.Text = "Line 3D";
            //Set Line3D Style 1
            lineChart.StyleManager.SetChartStyle(ePresetChartStyle.Line3dChartStyle1);

            //Add an Area chart from a template file.
            var areaChart = (ExcelAreaChart)ws.Drawings.AddChartFromTemplate(FileInputUtil.GetFileInfo("15-ChartsAndThemes", "AreaChartStyle3.crtx"), "areaChart");
            var areaSerie = areaChart.Series.Add(ws.Cells[2, 7, 26, 7], ws.Cells[2, 6, 26, 6]);
            areaSerie.Header = "Order Value";
            areaChart.SetPosition(51, 0, 10, 0);
            areaChart.SetSize(1000, 300);
            areaChart.Title.Text = "Area Chart";


            range.AutoFitColumns(0);
        }
        private static async Task AddLineCharts(string connectionString, ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("LineCharts");

            var range = await LoadFromDatabase(connectionString, ws);

            //Add a line chart
            var chart = ws.Drawings.AddLineChart("LineChartWithDroplines", eLineChartType.Line);
            var serie = chart.Series.Add(ws.Cells[2, 7, 26, 7], ws.Cells[2, 6, 26, 6]);
            serie.Header = "Order Value";
            chart.SetPosition(0, 0, 10, 0);
            chart.SetSize(1000, 300);
            chart.Title.Text = "Line Chart With Droplines";
            chart.AddDropLines();
            chart.DropLine.Border.Width = 2;
            //Set style 12
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle12);

            //Add a line chart with Error Bars
            chart = ws.Drawings.AddLineChart("LineChartWithErrorBars", eLineChartType.Line);
            serie = chart.Series.Add(ws.Cells[2, 7, 26, 7], ws.Cells[2, 6, 26, 6]);
            serie.Header = "Order Value";
            chart.SetPosition(17, 0, 10, 0);
            chart.SetSize(1000, 300);
            chart.Title.Text = "Line Chart With Error Bars";
            serie.AddErrorBars(eErrorBarType.Both, eErrorValueType.Percentage);
            serie.ErrorBars.Value = 5;
            chart.PlotArea.CreateDataTable();

            //Set style 2
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle2);

            //Add a line chart with Error Bars
            chart = ws.Drawings.AddLineChart("LineChartWithUpDownBars", eLineChartType.Line);
            var serie1 = chart.Series.Add(ws.Cells[2, 7, 26, 7], ws.Cells[2, 6, 26, 6]);
            serie1.Header = "Order Value";
            var serie2 = chart.Series.Add(ws.Cells[2, 8, 26, 8], ws.Cells[2, 6, 26, 6]);
            serie2.Header = "Tax";
            var serie3 = chart.Series.Add(ws.Cells[2, 9, 26, 9], ws.Cells[2, 6, 26, 6]);
            serie3.Header = "Freight";
            chart.SetPosition(34, 0, 10, 0);
            chart.SetSize(1000, 300);
            chart.Title.Text = "Line Chart With Up/Down Bars";
            chart.AddUpDownBars(true, true);

            //Set style 10, Note: As this is a line chart with multiple series, we use the enum for multiple series. Charts with multiple series usually has a subset of of the chart styles in Excel.
            //Another option to set the style is to use the Excel Style number, in this case 236: chart.StyleManager.SetChartStyle(236);
            chart.StyleManager.SetChartStyle(ePresetChartStyleMultiSeries.LineChartStyle9);
            range.AutoFitColumns(0);


            //Add a line chart with high/low Bars
            chart = ws.Drawings.AddLineChart("LineChartWithHighLowLines", eLineChartType.Line);
            serie1 = chart.Series.Add(ws.Cells[2, 7, 26, 7], ws.Cells[2, 6, 26, 6]);
            serie1.Header = "Order Value";
            serie2 = chart.Series.Add(ws.Cells[2, 8, 26, 8], ws.Cells[2, 6, 26, 6]);
            serie2.Header = "Tax";
            serie3 = chart.Series.Add(ws.Cells[2, 9, 26, 9], ws.Cells[2, 6, 26, 6]);
            serie3.Header = "Freight";
            chart.SetPosition(51, 0, 10, 0);
            chart.SetSize(1000, 300);
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
                using (var sqlCmd = new SQLiteCommand("select companyName as CompanyName, [name] as Name, email as Email, country as Country, o.OrderId as OrderId, orderdate as OrderDate, ordervalue as OrderValue, tax As Tax,freight As Freight, currency Currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY OrderDate, OrderValue desc limit 25", sqlConn))
                { 
                    using (var sqlReader = sqlCmd.ExecuteReader())
                    {
                        range = await ws.Cells["A1"].LoadFromDataReaderAsync(sqlReader, true);
                        range.Offset(0, 0, 1, range.Columns).Style.Font.Bold = true;
                        range.Offset(0, 5, range.Rows, 1).Style.Numberformat.Format = "yyyy-MM-dd";
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
