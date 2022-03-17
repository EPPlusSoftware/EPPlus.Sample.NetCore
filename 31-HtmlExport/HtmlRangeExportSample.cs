using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public static class HtmlRangeExportSample
    {
        //This sample demonstrates how to copy entire worksheet, ranges and how to exclude different cell properties.
        //More advanced samples using charts and json exports are available in our samples web site available 
        //here: https://samples.epplussoftware.com/HtmlExport, https://samples.epplussoftware.com/JsonExport
        public async static Task RunAsync()
        {
            var outputFolder = FileUtil.GetDirectoryInfo("HtmlOutput");

            await ExportGettingStartedAsync(outputFolder);
            
            ExportSalesReport(outputFolder);

            await ExcludeCssAsync(outputFolder);

            ExportMultipleRanges(outputFolder);
        }

        private static async Task ExportGettingStartedAsync(DirectoryInfo outputFolder)
        {
            //Start by using the excel file generated in sample 8
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("01-GettingStarted.xlsx")))
            {
                var ws = p.Workbook.Worksheets["Inventory"];
                //Will create the html exporter for min and max bounds of the worksheet (ws.Dimensions)
                var exporter = ws.Cells.CreateHtmlExporter();   
                
                //Get the html and styles in one call. 
                var html=await exporter.GetSinglePageAsync();
                await File.WriteAllTextAsync(FileUtil.GetFileInfo(outputFolder, "Range-01-GettingStarted.html", true).FullName,
                    html);
            }
        }

        private static void ExportSalesReport(DirectoryInfo outputFolder)
        {
            //Start by using the excel file generated in sample 8
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("08-Salesreport.xlsx")))
            {
                var ws = p.Workbook.Worksheets["Sales"];
                var exporter = ws.Cells.CreateHtmlExporter();   //Will create the html exporter for min and max bounds of the worksheet (ws.Dimensions)
                exporter.Settings.HeaderRows = 4;               //We have three header rows.
                exporter.Settings.TableId = "my-table";         //We can set an id of the worksheet if we want to use it in css or javascript.

                //By default EPPlus include the normal font in the css for the table. This can be tuned off and replaces by your own settings.
                exporter.Settings.Css.IncludeNormalFont = false;
                //AdditionalCssElements is a collection where you can add your own styles for the table. You can also clear default styles set by EPPlus.
                exporter.Settings.Css.AdditionalCssElements.Add("font-family", "verdana");

                //EPPlus will not set column width and row heights by default, as this doesn't go well with todays responsive designs.
                //If you want fixed widths/heights set the proprties below to true...
                //Note that individual width and height are set direcly on the colspan-elements and tr-elements.
                //Default width and heights are set via the classes epp-dcw and epp-drh (with default StyleClassPrefix.).
                exporter.Settings.SetColumnWidth = true;
                exporter.Settings.SetRowHeight = true;

                //Get the html...
                var htmlTable = exporter.GetHtmlString();
                //...and the styles
                var cssTable = exporter.GetCssString();

                //EPPlus will not add the Excel grid lines, but you can easily add your own in the css...
                cssTable += "#my-table th,td {border:solid thin lightgray}";

                var html = $"<html><head><style type=\"text/css\">{cssTable}</style></head>{htmlTable}</html>";

                File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Range-02-Salesreport.html", true).FullName,
                    html);
            }
        }
        private static async Task ExcludeCssAsync(DirectoryInfo outputFolder)
        {
            //Start by using the excel file generated in sample 20
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("20-CreateAFileSystemReport.xlsx")))
            {
                var ws = p.Workbook.Worksheets[0];
                var range = ws.Cells[1, 1, 5, ws.Dimension.End.Column];

                var exporter = range.CreateHtmlExporter();
                //Css can be excluded on style level, if you don't want some style or you want to add your own.
                exporter.Settings.Css.CssExclude.Font = eFontExclude.Bold | eFontExclude.Italic | eFontExclude.Underline;

                var html = await exporter.GetSinglePageAsync();
                await File.WriteAllTextAsync(FileUtil.GetFileInfo(outputFolder, "Range-03-ExcludeCss.html", true).FullName,
                    html);
            }
        }
        private static void ExportMultipleRanges(DirectoryInfo outputFolder)
        {
            //Start by using the excel file generated in sample 15
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("15-ChartsAndThemes.xlsx")))
            {
                //Now we will use the sample 15 and read two ranges from two different worksheets and combine them to use the same CSS.
                //To do so we create an HTML exporter on the workbook level and adds the ranges we want to use.
                var ws3D = p.Workbook.Worksheets["3D Charts"];
                var wsStock = p.Workbook.Worksheets["Stock Chart"];

                //We mark the top and bottom two values with red and green.
                ws3D.Cells["B13,B7"].Style.Fill.SetBackground(Color.Green);
                ws3D.Cells["B14,B5"].Style.Fill.SetBackground(Color.Red);

                //We mark the top and bottom two rows with red and green.
                wsStock.Cells["A3:E4"].Style.Fill.SetBackground(Color.Green);
                wsStock.Cells["A7:E8"].Style.Fill.SetBackground(Color.Red);

                //Create the exporter. The workbook exporter exports ranges only. If you want to export tables, use the exporter available on the table object.
                var rngExporter = p.Workbook.CreateHtmlExporter(
                    ws3D.Cells["A1:D16"],
                    wsStock.Cells["A1:E11"]);

                //Get the html for the ranges in the HTML. The argument index referece to the ranges supplied when creating the exporter. 
                var html3D = rngExporter.GetHtmlString(0);
                var htmlStock = rngExporter.GetHtmlString(1);
                var css = rngExporter.GetCssString();

                //We also exports a table and merge the css the range css.
                var tblChartData = p.Workbook.Worksheets["ChartData"].Tables[0];
                var tblExporter = tblChartData.CreateHtmlExporter();
                var tblHtml = tblExporter.GetHtmlString();
                var tblCss = tblExporter.GetCssString();

                var htmlTemplate = "<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{0}</style></head>\r\n<body>\r\n{1}<hr>{2}<hr>{3}</body>\r\n</html>";
                File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Range-04-MultipleRanges.html", true).FullName,
                    string.Format(htmlTemplate, css+tblCss, html3D, htmlStock, tblHtml));
            }
        }
    }
}