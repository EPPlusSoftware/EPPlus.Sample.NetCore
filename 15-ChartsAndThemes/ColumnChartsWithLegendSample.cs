using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Drawing;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public class ColumnChartWithLegendSample : ChartSampleBase
    {
        public static async Task Add(string connectionString, ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("ColumnCharts");

            var range = await LoadFromDatabase(connectionString, ws);

            //Add a line chart
            var chart = ws.Drawings.AddBarChart("ColumnChartWithLegend", eBarChartType.ColumnStacked);
            var serie1 = chart.Series.Add(ws.Cells[2, 2, 16, 2], ws.Cells[2, 1, 16, 1]);
            serie1.Header = "Order Value";
            var serie2 = chart.Series.Add(ws.Cells[2, 3, 16, 3], ws.Cells[2, 1, 16, 1]);
            serie2.Header = "Tax";
            var serie3 = chart.Series.Add(ws.Cells[2, 4, 16, 4], ws.Cells[2, 1, 16, 1]);
            serie3.Header = "Freight";
            chart.SetPosition(0, 0, 6, 0);
            chart.SetSize(1200, 400);
            chart.Title.Text = "Column chart";

            //Set style 10
            chart.StyleManager.SetChartStyle(ePresetChartStyle.ColumnChartStyle10);

            chart.Legend.Entries[0].Font.Fill.Color = Color.Red;
            chart.Legend.Entries[1].Font.Fill.Color = Color.Green;
            chart.Legend.Entries[2].Deleted = true;

            range.AutoFitColumns(0);
        }
    }
}
