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
                 * Chart style do not alter visibility of chart objects like data labels or chart titles like Excel do. That must be set in code before setting the style.
                 *********************************************************************************************************/

                //The first method adds a worksheet with four 3D charts with different styles. The last chart applies an exported chart template file (*.crtx) to the chart.
                await ThreeDimensionalCharts.Add3DCharts(connectionString, package);
                
                //This method adds four line charts with different chart elements like up-down bars, error bars, drop lines and high-low lines.
                await LineChartsSample.Add(connectionString, package);
                
                //Adds a scatter chart with a moving average trendline.
                ScatterChartSample.Add(package);

                //Adds a column chart with a legend where we style and remove individual legend items.
                await ColumnChartWithLegendSample.Add(connectionString, package);

                //Adds a bubble-chartsheet
                ChartWorksheetSample.Add(package);
                
                //Adds a radar chart
                RadarChartSample.Add(package);

                //Adds a Volume-High-Low-Close stock chart
                StockChartSample.Add(package);

                //Adds a sunburst and a treemap chart 
                await SunburstAndTreemapChartSample.Add(connectionString, package);

                //Adds a box & whisker and a histogram chart 
                BoxWhiskerHistogramChartSample.Add(package);

                // Adds a waterfall chart
                WaterfallChartSample.Add(package);

                // Adds a funnel chart
                FunnelChartSample.Add(package);

                await RegionMapChartSample.Add(connectionString, package);

                //Add an area chart using a chart template (chrx file)
                await ChartTemplateSample.AddAreaChart(connectionString, package);

                //Save our new workbook in the output directory and we are done!
                package.SaveAs(xlFile);
                return xlFile.FullName;
            }
        }
    }
}
