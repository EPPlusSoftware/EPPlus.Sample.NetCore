using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public class LineChartsSample : ChartSampleBase
    {
        public static async Task Add(string connectionString, ExcelPackage package)
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
    }
}
