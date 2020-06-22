using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusSamples
{
    public class FunnelChartSample
    {
        public static void Add(ExcelPackage package)
        {
            ExcelWorksheet ws = LoadFunnelChartData(package);

            var funnelChart = ws.Drawings.AddFunnelChart("FunnelChart");
            funnelChart.Title.Text = "Sales process";
            funnelChart.SetPosition(1, 0, 6, 0);
            funnelChart.SetSize(800, 400);
            var fSerie = funnelChart.Series.Add(ws.Cells[2, 2, 7, 2], ws.Cells[2, 1, 7, 1]);
            fSerie.DataLabel.Add(false, true);
            funnelChart.StyleManager.SetChartStyle(OfficeOpenXml.Drawing.Chart.Style.ePresetChartStyle.FunnelChartStyle9);
        }

        private static ExcelWorksheet LoadFunnelChartData(ExcelPackage package)
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
            return ws;
        }
    }
}
