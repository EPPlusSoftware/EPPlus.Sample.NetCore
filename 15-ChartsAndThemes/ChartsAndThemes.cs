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

namespace EPPlusSamples
{
    class ChartsAndThemesSample
    {
        /// <summary>
        /// Sample 15 - Load a theme and create a column chart
        /// </summary>
        public static async Task<string> RunAsync(string connectionString)
        {
            using (var package = new ExcelPackage())
            {
                //Load a theme file (Can be exported from Excel)
                package.Workbook.ThemeManager.Load(FileInputUtil.GetFileInfo("15-ChartsAndThemes", "integral.thmx"));
                //package.Workbook.ThemeManager.Load(FileInputUtil.GetFileInfo("15-ChartsAndThemes", "WoodType.thmx"));

                /*** Themes can also be altered, for example uncomment this code to set the Accent1 to a blue color ***/
                //package.Workbook.ThemeManager.CurrentTheme.ColorScheme.Accent1.SetRgbColor(Color.FromArgb(32, 78, 224));

                await Add3DCharts(connectionString, package);
                await AddLineCharts(connectionString, package);

                // save our new workbook in the output directory and we are done!
                var xlFile = FileOutputUtil.GetFileInfo("15-ChartsAndThemes.xlsx");
                package.SaveAs(xlFile);
                return xlFile.FullName;
            }
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

            //Set style 10, Note: As this is a line chart with multiple series, this turnes up as Style 9 in Excel. Charts with multiple series has a subset of of the chart styles.
            //Another option to set the style is to use the Excel Style number, in this case 236: chart.StyleManager.SetChartStyle(236);
            chart.StyleManager.SetChartStyle(ePresetChartStyle.LineChartStyle10);
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

            //Set the style using the Excel ChartStyle number. The chart style must exist in the ExcelChartStyleManager.StyleLibrary[]. Styles can be added and removed from this library.
            chart.StyleManager.SetChartStyle(237);
            range.AutoFitColumns(0);
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
    }
}
