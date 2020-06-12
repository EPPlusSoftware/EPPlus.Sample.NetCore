using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Drawing.Chart.ChartEx;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public class RegionMapChartSample : ChartSampleBase
    {
        public static async Task Add(string connectionString, ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("RegionMapChart");

            var range = await LoadSalesFromDatabase(connectionString, ws);

            //Region map charts 
            var regionChart = ws.Drawings.AddRegionMapChart("RegionMapChart");
            regionChart.Title.Text = "Sales";
            regionChart.SetPosition(1, 0, 6, 0);
            regionChart.SetSize(1200, 600);

            var rmSerie = regionChart.Series.Add(ws.Cells[2, 4, range.End.Row, 4], ws.Cells[2, 1, range.End.Row, 3]);

            rmSerie.ColorBy = eColorBy.Value;

            //Color settings only apply when ColorBy is set to Value
            rmSerie.Colors.NumberOfColors = eNumberOfColors.ThreeColor;
            rmSerie.Colors.MinColor.Color.SetSchemeColor(eSchemeColor.Accent3);
            rmSerie.Colors.MidColor.Color.SetHslColor(180, 50, 50);
            rmSerie.Colors.MidColor.ValueType = eColorValuePositionType.Number;
            rmSerie.Colors.MidColor.PositionValue = 200;
            rmSerie.Colors.MaxColor.Color.SetRgbPercentageColor(75, 25, 25);
            rmSerie.Colors.MaxColor.ValueType = eColorValuePositionType.Percent;
            rmSerie.Colors.MaxColor.PositionValue = 85;

            rmSerie.ProjectionType = eProjectionType.Mercator;
        }
    }
}
