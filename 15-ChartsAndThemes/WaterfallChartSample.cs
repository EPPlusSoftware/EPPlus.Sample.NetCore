using OfficeOpenXml;
namespace EPPlusSamples
{
    public class WaterfallChartSample
    {
        public static void Add(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("WaterfallChart");

            LoadWaterfallChartData(ws);
            var waterfallChart = ws.Drawings.AddWaterfallChart("Waterfall1");
            waterfallChart.Title.Text = "Saldo and Transaction";
            waterfallChart.SetPosition(1, 0, 6, 0);
            waterfallChart.SetSize(800, 400);
            var wfSerie = waterfallChart.Series.Add(ws.Cells[2, 2, 9, 2], ws.Cells[2, 1, 8, 1]);

            var dp = wfSerie.DataPoints.Add(0);
            dp.SubTotal = true;
            dp = wfSerie.DataPoints.Add(7);
            dp.SubTotal = true;
        }

        private static void LoadWaterfallChartData(ExcelWorksheet ws)
        {
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
            ws.Cells["B9"].Formula = "SUM(B2:B8)";
            ws.Calculate();
            ws.Cells.AutoFitColumns();
        }
    }
}
