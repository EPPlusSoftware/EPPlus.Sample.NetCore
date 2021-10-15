/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/13/2021         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Drawing;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public static class CopyRangeSample
    {
        //This sample demonstrates how to copy entire worksheet, ranges and how to exclude different cell properties.
        public static void Run()
        {
            using (var p = new ExcelPackage())
            {
                var sourceFile = FileUtil.GetFileInfo("08-Salesreport.xlsx");
                var sourcePackage = new ExcelPackage(sourceFile);                
                var sourceWs = sourcePackage.Workbook.Worksheets[0];

                //Copy the entire source worksheet to a new worksheet.
                CopyEntireWorksheet(p, sourceWs);
                //Copy a range from the source worksheet to the new worksheet.
                //This samples demonstrates how to exclude different options to exclude different parts of the cell properties
                CopyRange(p, sourceWs);
                //Copy a range with values only, removing any formula.
                CopyValues(p);
                //Copy styles 
                CopyStyles(p, sourceWs);
                
                p.SaveAs(FileUtil.GetCleanFileInfo("30-CopyRangeSamples.xlsx"));
            }
        }

        private static void CopyValues(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("CopyValues");
            //Add some numbers and formulas and calculate the worksheet
            ws.Cells["A1:A10"].FillNumber(1);
            ws.Cells["B1:B9"].Formula = "A1+A2";
            ws.Cells["B10"].Formula = "Sum(B1:B9)";
            ws.Calculate();

            //Now, copy the values starting at cell D1 without the formulas.
            ws.Cells["A1:B10"].Copy(ws.Cells["D1"], ExcelRangeCopyOptionFlags.ExcludeFormulas);
        }

        private static void CopyEntireWorksheet(ExcelPackage p, ExcelWorksheet sourceWs)
        {
            //To copy the entire worksheet just add the source worksheet as parameter 2 when adding the new worksheet.
            p.Workbook.Worksheets.Add("CopySalesReport", sourceWs);
        }

        private static void CopyRange(ExcelPackage p, ExcelWorksheet sourceWs)
        {
            var ws = p.Workbook.Worksheets.Add("CopyRangeOfReport");

            //Use the first 10 rows of the sales report in sample 8 as the source.
            var sourceRange = sourceWs.Cells["A1:G10"]; 

            //Copy the source full range starting from C1.
            //Copy always start from the top left cell of the destination range and copies the full range.
            sourceRange.Copy(ws.Cells["C1"]);
            
            //Copy the same source range to C15 and exclude the hyperlinks.
            //We also remove the Hyperlink style from the range containing the hyperlinks.
            sourceRange.Copy(ws.Cells["C15"], ExcelRangeCopyOptionFlags.ExcludeHyperLinks);
            ws.Cells["D19:D24"].StyleName = "Normal";

            //Copy the values only, excluding merged cells, styles and hyperlinks.
            sourceRange.Copy(ws.Cells["C30"], ExcelRangeCopyOptionFlags.ExcludeMergedCells, ExcelRangeCopyOptionFlags.ExcludeStyles , ExcelRangeCopyOptionFlags.ExcludeHyperLinks);

            //Copy styles and merged cells, excluding values and hyperlinks.
            sourceRange.Copy(ws.Cells["C45"], ExcelRangeCopyOptionFlags.ExcludeValues, ExcelRangeCopyOptionFlags.ExcludeHyperLinks);
        }
        private static void CopyStyles(ExcelPackage p, ExcelWorksheet sourceWs)
        {
            var ws = p.Workbook.Worksheets.Add("CopyStyles");
            
            //Create a new random report 
            FillRangeWithRandomData(ws);
            
            //Copy the styles from the sales report.
            //If the destination range is larger that the source range styles are filled down and right using the last column/row.
            sourceWs.Cells["A1:G5"].CopyStyles(ws.Cells["A1:G50"]);
            
            ws.Cells.AutoFitColumns();
        }

        private static void FillRangeWithRandomData(ExcelWorksheet ws)
        {
            ws.Cells["A1"].Value = "EPPlus";
            ws.Cells["A2"].Value = "New Random Report";
            ws.Cells["A4"].Value = "Color";
            ws.Cells["B4"].Value = "Category";
            ws.Cells["C4"].Value = "Country";
            ws.Cells["D4"].Value = "Id";
            ws.Cells["E4"].Value = "Date";
            ws.Cells["F4"].Value = "Amount";
            ws.Cells["G4"].Value = "Currency";

            ws.Cells["A5:A50"].FillList(new string[] { "Red", "Green", "Blue", "Pink", "Black" });
            ws.Cells["B5:B50"].FillList(new string[] { "New", "Old" });
            ws.Cells["C5:C50"].FillList(new string[] { "Usa", "France", "India" });

            ws.Cells["D5:D50"].FillNumber(1, 10);
            ws.Cells["E5:E50"].FillDateTime(DateTime.Today);
            ws.Cells["F5:F50"].FillNumber(1000, 50);
            ws.Cells["G5:G50"].FillList(new string[] { "USD", "EUR", "INR" });
        }
    }
}
