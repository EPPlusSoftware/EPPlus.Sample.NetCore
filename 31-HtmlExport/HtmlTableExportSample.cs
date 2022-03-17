using OfficeOpenXml;
using OfficeOpenXml.Export.HtmlExport;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace EPPlusSamples
{
    public static class HtmlTableExportSample
    {
        //This sample demonstrates how to export html from a table.
        //More advanced samples using charts and json exports are available in our samples web site available 
        //here: https://samples.epplussoftware.com/HtmlExport, https://samples.epplussoftware.com/JsonExport
        public static void Run()
        {
            var outputFolder = FileUtil.GetDirectoryInfo("HtmlOutput");
            //Start by using the excel file generated in sample 28
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("28-Tables.xlsx")))
            {
                var wsSimpleTable = p.Workbook.Worksheets["SimpleTable"];

                ExportSimpleTable1(outputFolder, wsSimpleTable);
                ExportSimpleTable2(outputFolder, wsSimpleTable);

                var wsStyleTables = p.Workbook.Worksheets["StyleTables"];
                ExportStyleTables(outputFolder, wsStyleTables);

                //This samples exports the filtered table from the slicer sample.
                var wsSlicer = p.Workbook.Worksheets["Slicer"];
                ExportSlicerTables1(outputFolder, wsSlicer);

                //Exports three tables and combine the html and css 
                ExportMultipleTables(outputFolder);
            }
        }
        private static void ExportSimpleTable1(DirectoryInfo outputFolder, ExcelWorksheet wsSimpleTable)
        {
            var table1 = wsSimpleTable.Tables[0];
            //Create the exporter for the table.
            var htmlExporter = table1.CreateHtmlExporter();

            //EPPlus will minify the css and html by default, but for this sample we want it easier to read.
            htmlExporter.Settings.Minify = false;

            // The GetSinglePage method generates en single page. You can also add a string parameter with your own HTML where where the styles and table html is inserted.
            var fullHtml = htmlExporter.GetSinglePage();
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-01-Table1_SinglePage.html", true).FullName,
                fullHtml);

            //In most cases you want to keep the html and the styles separated, so you will retrive the html and the css in separate calls...
            var tableHtml = htmlExporter.GetHtmlString();
            var tableCss = htmlExporter.GetCssString();

            //First create the html file and reference the the css.
            var html = $"<html><head><link rel=\"stylesheet\" href=\"Table-01-Table1.css\"</head>{tableHtml}</html>";
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-01-Table1.html", true).FullName,
                html);

            //The css is written to a separate file.
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-01-Table1.css", true).FullName, tableCss);
        }
        private static void ExportSimpleTable2(DirectoryInfo outputFolder, ExcelWorksheet wsSimpleTable)
        {
            var table2 = wsSimpleTable.Tables[1];

            //Create the exporter for the table.
            var htmlExporter = table2.CreateHtmlExporter();
            //EPPlus will generate Accessibility and data attributes by default, but you can turn it of in the settings.
            htmlExporter.Settings.Accessibility.TableSettings.AddAccessibilityAttributes = false;
            htmlExporter.Settings.RenderDataAttributes = false;

            var html = htmlExporter.GetSinglePage();
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-Table2.html", true).FullName,
                html);

            //We can also change the table style to get a different styling.
            //Here we change to Medium15...
            table2.TableStyle = OfficeOpenXml.Table.TableStyles.Medium15;
            html = htmlExporter.GetSinglePage();
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-table2_Medium15.html", true).FullName,
                html);

            //...Here we use Dark2...
            table2.TableStyle = OfficeOpenXml.Table.TableStyles.Dark2;
            html = htmlExporter.GetSinglePage();
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-table2_Dark2.html", true).FullName,
                html);
        }

        private static void ExportStyleTables(DirectoryInfo outputFolder, ExcelWorksheet wsStyleTables)
        {
            //The last row of the cell contains uncalculated cell (they calculate when opened in Excel),
            //but in EPPlus we need to calculate them first to get a result in cell A254 in the totals row.
            wsStyleTables.Calculate();

            var table1 = wsStyleTables.Tables[0];
            var htmlExporter = table1.CreateHtmlExporter();

            //This sample exports the table as well as some individually cell styles. The headers have font italic and the totals row has a custom formatted text.
            //Also note that Column 2 has hyper links create for the mail addresses.
            var html = htmlExporter.GetSinglePage();
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-Styling_table1_with_hyperlinks.html", true).FullName,
                html);

            var table2 = wsStyleTables.Tables[1];
            htmlExporter = table2.CreateHtmlExporter();

            //Table 2 contains a custom table style.
            html = htmlExporter.GetSinglePage();
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-02-Styling_table2.html", true).FullName,
                html);
        }
        private static void ExportSlicerTables1(DirectoryInfo outputFolder, ExcelWorksheet wsSlicer)
        {
            var table1 = wsSlicer.Tables[0];
            var htmlExporter = table1.CreateHtmlExporter();

            //This sample exports the table filtered by the selection in the slicer (that applies the filter on the table).
            //By default EPPlus will remove hidden rows.
            var html = htmlExporter.GetSinglePage();
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-03-Slicer.html", true).FullName,
                html);

            //You can change this option by setting eHiddenState.Include in the settings.
            //You can also set the it to eHiddenState.IncludeButHide if you want to apply your own filtering.
            htmlExporter.Settings.HiddenRows = eHiddenState.Include;
            html = htmlExporter.GetSinglePage();
            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-03-Slicer_table_all_rows.html", true).FullName,
                html);
        }
        private static void ExportMultipleTables(DirectoryInfo outputFolder)
        {
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("04-LoadingData.xlsx")))
            {
                //Now we will use the third worksheet from sample 4, that contains three tables with different styles.
                var wsList = p.Workbook.Worksheets["FromList"];
              
                var tbl1 = wsList.Tables[0];
                var exporter1 = tbl1.CreateHtmlExporter();
                var tbl1Html = exporter1.GetHtmlString();
                var tbl1Css = exporter1.GetCssString();

                var tbl2 = wsList.Tables[1];
                var exporter2 = tbl2.CreateHtmlExporter();
                var tbl2Html = exporter2.GetHtmlString();                                
                //We have already exported the css once, so we don't want shared css classes to be added again.
                exporter2.Settings.Css.IncludeSharedClasses = false;
                var tbl2Css = exporter2.GetCssString();

                var tbl3 = wsList.Tables[2];
                var exporter3 = tbl3.CreateHtmlExporter();
                var tbl3Html = exporter3.GetHtmlString();
                exporter3.Settings.Css.IncludeSharedClasses = false;
                
                var tbl3Css = exporter3.GetCssString();

                //As the tables have different table styles we add all of the css's.
                //If multiple tables have the same table style, you should only add one of them.
                var css = tbl1Css + tbl2Css + tbl3Css;

                var htmlTemplate = "<html>\r\n<head>\r\n<style type=\"text/css\">\r\n{0}</style></head>\r\n<body>\r\n{1}<hr>{2}<hr>{3}</body>\r\n</html>";
                File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "Table-04-MultipleTables.html", true).FullName,
                    string.Format(htmlTemplate, css, tbl1Html, tbl2Html, tbl3Html));
            }
        }

    }
}