using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.Style;
namespace EPPlusSamples
{
    public class RadarChartSample : ChartSampleBase
    {
        public static void AddRadarChart(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("RadarChart");

            var dt = GetCarDataTable();
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

    }
}
