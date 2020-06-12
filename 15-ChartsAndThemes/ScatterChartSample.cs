using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
namespace EPPlusSamples
{
    public class ScatterChartSample : ChartSampleBase
    {
        public static void Add(ExcelPackage package)
        {
            //Adda a scatter chart on the data with one serie per row. 
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
            var tr = serie.TrendLines.Add(eTrendLine.MovingAvgerage);
            tr.Name = "Icecream Sales-Monthly Average";
            chart.StyleManager.SetChartStyle(ePresetChartStyle.ScatterChartStyle12);
        }
    }
}
