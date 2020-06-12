using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public class ThreeDimensionalCharts : ChartSampleBase
    {
        public static async Task Add3DCharts(string connectionString, ExcelPackage package)
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
    }
}
