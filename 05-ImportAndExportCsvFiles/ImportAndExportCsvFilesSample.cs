/*************************************************************************************************
  Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/24/2020         Jan Källman & Mats Alm       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using System.IO;
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing.Chart;
using System.Globalization;
using System.Threading.Tasks;
using OfficeOpenXml.Drawing.Chart.Style;

namespace EPPlusSamples.LoadDataFromCsvFilesIntoTables
{
    /// <summary>
    /// This sample shows how to load/save CSV files using the LoadFromText and SaveToText methods, how to use tables and
    /// how to use charts with more than one chart type and secondary axis
    /// </summary>
    public static class ImportAndExportCsvFilesSample
    {
        /// <summary>
        /// Loads two CSV files into tables and adds a chart to each sheet.
        /// </summary>
        /// <param name="outputDir"></param>
        /// <returns></returns>
        public static async Task<string> Run()
        {
            FileInfo newFile = FileOutputUtil.GetFileInfo(@"05-LoadDataFromCsvFilesIntoTables.xlsx");
            
            using (ExcelPackage package = new ExcelPackage())
            {
                LoadFile1(package);                 //Load the text file without async
                await LoadFile2Async(package);      //Load the second text file with async
                await ExportTableAsync(package);
                await package.SaveAsAsync(newFile);
            }
            return newFile.FullName;
        }

        private static async Task ExportTableAsync(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets[1];
            var tbl = ws.Tables[0];
            var format = new ExcelOutputTextFormat
            {
                Delimiter = ';',
                Culture = new CultureInfo("en-GB"),
                Encoding = new UTF8Encoding(),   
                SkipLinesEnd=1  //Skip the totals row                
            };
            await ws.Cells[tbl.Address.Address].SaveToTextAsync(FileOutputUtil.GetFileInfo("05-ExportedFromEPPlus.csv"), format);

            Console.WriteLine($"Writing the text file 'ExportedTable.csv'...");
        }

        private static void LoadFile1(ExcelPackage package)
        {
            //Create the Worksheet
            var sheet = package.Workbook.Worksheets.Add("Csv1");

            //Create the format object to describe the text file
            var format = new ExcelTextFormat
            {
                TextQualifier = '"',
                SkipLinesBeginning = 2,
                SkipLinesEnd = 1
            };

            var file1 = FileInputUtil.GetFileInfo("05-ImportAndExportCsvFiles", "Sample5-1.txt");

            //Now read the file into the sheet. Start from cell A1. Create a table with style 27. First row contains the header.
            Console.WriteLine("Load the text file...");
            var range = sheet.Cells["A1"].LoadFromText(file1, format, TableStyles.Medium27, true);

            Console.WriteLine("Format the table...");
            //Tables don't support custom styling at this stage(you can of course format the cells), but we can create a Namedstyle for a column...
            var dateStyle = package.Workbook.Styles.CreateNamedStyle("TableDate");
            dateStyle.Style.Numberformat.Format = "YYYY-MM";

            var numStyle = package.Workbook.Styles.CreateNamedStyle("TableNumber");
            numStyle.Style.Numberformat.Format = "#,##0.0";

            //Now format the table...
            var tbl = sheet.Tables[0];
            tbl.ShowTotal = true;
            tbl.Columns[0].TotalsRowLabel = "Total";
            tbl.Columns[0].DataCellStyleName = "TableDate";
            tbl.Columns[1].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[1].DataCellStyleName = "TableNumber";
            tbl.Columns[2].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[2].DataCellStyleName = "TableNumber";
            tbl.Columns[3].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[3].DataCellStyleName = "TableNumber";
            tbl.Columns[4].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[4].DataCellStyleName = "TableNumber";
            tbl.Columns[5].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[5].DataCellStyleName = "TableNumber";
            tbl.Columns[6].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[6].DataCellStyleName = "TableNumber";
            
            Console.WriteLine("Create the chart...");
            //Now add a stacked areachart...
            var chart = sheet.Drawings.AddChart("chart1", eChartType.AreaStacked);
            chart.SetPosition(0, 630);
            chart.SetSize(800, 600);

            //Create one series for each column...
            for (int col = 1; col < 7; col++)
            {
                var ser = chart.Series.Add(range.Offset(1, col, range.End.Row - 1, 1), range.Offset(1, 0, range.End.Row - 1, 1));
                ser.HeaderAddress = range.Offset(0, col, 1, 1);
            }
            
            //Set the style to 27.
            chart.Style = eChartStyle.Style27;

            sheet.View.ShowGridLines = false;
            sheet.Calculate();
            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();
        }

        private static async Task LoadFile2Async(ExcelPackage package)
        {
            //Create the Worksheet
            var sheet = package.Workbook.Worksheets.Add("Csv2");

            //Create the format object to describe the text file
            var format = new ExcelTextFormat
            {
                Delimiter = '\t',       //Tab
                SkipLinesBeginning = 1
            };
            CultureInfo ci = new CultureInfo("sv-SE");          //Use your choice of Culture
            ci.NumberFormat.NumberDecimalSeparator = ",";       //Decimal is comma
            format.Culture = ci;

            //Now read the file into the sheet.
            Console.WriteLine("Load the text file...");
           var file2 = FileInputUtil.GetFileInfo("05-ImportAndExportCsvFiles", "Sample5-2.txt");

            var range = await sheet.Cells["A1"].LoadFromTextAsync(file2, format);

            //Add a formula
            range.Offset(1, range.End.Column, range.End.Row - range.Start.Row, 1).FormulaR1C1 = "RC[-1]-RC[-2]";

            //Add a table...
            var tbl = sheet.Tables.Add(range.Offset(0,0,range.End.Row-range.Start.Row+1, range.End.Column-range.Start.Column+2),"Table");
            tbl.ShowTotal = true;
            tbl.Columns[0].TotalsRowLabel = "Total";
            tbl.Columns[1].TotalsRowFormula = "COUNT(3,Table[Product])";    //Add a custom formula
            tbl.Columns[2].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[3].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[4].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[5].TotalsRowFunction = RowFunctions.Sum;
            tbl.Columns[5].Name = "Profit";
            tbl.TableStyle = TableStyles.Medium10;

            sheet.Cells[sheet.Dimension.Address].AutoFitColumns();

            //Add a chart with two charttypes (Column and Line) and a secondary axis...
            var chart = sheet.Drawings.AddChart("chart2", eChartType.ColumnStacked);
            chart.SetPosition(0, 540);
            chart.SetSize(800, 600);

            var serie1= chart.Series.Add(range.Offset(1, 3, range.End.Row - 1, 1), range.Offset(1, 1, range.End.Row - 1, 1));
            serie1.Header = "Purchase Price";
            var serie2 = chart.Series.Add(range.Offset(1, 5, range.End.Row - 1, 1), range.Offset(1, 1, range.End.Row - 1, 1));
            serie2.Header = "Profit";

            //Add a Line series
            var chartType2 = chart.PlotArea.ChartTypes.Add(eChartType.LineStacked);
            chartType2.UseSecondaryAxis = true;
            var serie3 = chartType2.Series.Add(range.Offset(1, 2, range.End.Row - 1, 1), range.Offset(1, 0, range.End.Row - 1, 1));
            serie3.Header = "Items in stock";

            //By default the secondary XAxis is not visible, but we want to show it...
            chartType2.XAxis.Deleted = false;
            chartType2.XAxis.TickLabelPosition = eTickLabelPosition.High;
            
            //Set the max value for the Y axis...
            chartType2.YAxis.MaxValue = 50;

            //chart.Style = eChartStyle.Style26;
            chart.StyleManager.SetChartStyle(ePresetChartStyle.ComboChartStyle2);

            sheet.View.ShowGridLines = false;
            sheet.Calculate();
        }
    }
}
