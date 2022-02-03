using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EPPlusSamples
{
    public static class HtmlRangeExportSample
    {
        const string scriptInclude = "<script type=\"text/javascript\" src=\"https://www.gstatic.com/charts/loader.js\" /><script type=\"text/javascript\" src=\"GoogleChartSetup.js\" />";

        const string chartPlaceholder = "<div id=\"bar-chart\" style=\"height:300px\"></div>";
        //This sample demonstrates how to copy entire worksheet, ranges and how to exclude different cell properties.
        public static void Run()
        {
            var outputFolder = FileUtil.GetDirectoryInfo("HtmlOutput");
            File.Copy("GoogleChartSetup.js", outputFolder.FullName, true);
            //Start by using the excel file generated in sample 28
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("17-FxReportFromDatabase.xlsx")))
            {
                var wsRateSample = p.Workbook.Worksheets["Rates Sample"];

                var range = wsRateSample.Cells[20, 1, wsRateSample.Dimension.End.Row, wsRateSample.Dimension.End.Column];
                var exporter = range.CreateHtmlExporter();
                exporter.Settings.HeaderRows = 2;
                exporter.Settings.TableId = "my-table";
                var htmlTable = exporter.GetHtmlString();
                var cssTable = exporter.GetCssString();

                var html = $"<html><head>{scriptInclude}<style type=\"css/text\">{cssTable}</style></head><body>{chartPlaceholder}</body></html>";


            }
        }
    }
}