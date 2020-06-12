using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using OfficeOpenXml.Drawing.Chart.Style;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public class SunburstAndTreemapChartSample : ChartSampleBase
    {
        public static async Task Add(string connectionString, ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Sunburst & Treemap Chart");
            var range = await LoadSalesFromDatabase(connectionString, ws);

            var sunburstChart = ws.Drawings.AddSunburstChart("SunburstChart1");
            var sbSerie = sunburstChart.Series.Add(ws.Cells[2, 4, range.Rows, 4], ws.Cells[2, 1, range.Rows, 3]);
            sbSerie.HeaderAddress = ws.Cells["D1"];
            sunburstChart.SetPosition(1, 0, 6, 0);
            sunburstChart.SetSize(800, 800);
            sunburstChart.Title.Text = "Sales";            
            sunburstChart.Legend.Add();
            sunburstChart.Legend.Position = eLegendPosition.Bottom;
            sbSerie.DataLabel.Add(true, true);
            sunburstChart.StyleManager.SetChartStyle(ePresetChartStyle.SunburstChartStyle4);


            var treemapChart = ws.Drawings.AddTreemapChart("TreemapChart1");
            var tmSerie = treemapChart.Series.Add(ws.Cells[2, 4, range.Rows, 4], ws.Cells[2, 1, range.Rows, 3]);
            treemapChart.Title.Font.Fill.Style = eFillStyle.SolidFill;
            treemapChart.Title.Font.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Background2);
            tmSerie.HeaderAddress = ws.Cells["D1"];
            treemapChart.SetPosition(1, 0, 19, 0);
            treemapChart.SetSize(1000, 800);
            treemapChart.Title.Text = "Sales";
            treemapChart.Legend.Add();
            treemapChart.Legend.Position = eLegendPosition.Right;
            tmSerie.DataLabel.Add(true, true);
            tmSerie.ParentLabelLayout = eParentLabelLayout.Banner;
            treemapChart.StyleManager.SetChartStyle(ePresetChartStyle.TreemapChartStyle3);            
        }
    }
}
