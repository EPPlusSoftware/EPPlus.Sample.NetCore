using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    /// <summary>
    /// This sample demonstrates how to sort ranges with EPPlus.
    /// </summary>
    public static class SortingRangesSample
    {
        private static string[] _letters = new string[] { "A", "B", "C", "D" };
        private static string[] _tShirtSizes = new string[] { "S", "M", "L", "XL", "XXL" };
        private const int StartRow = 3;
        public static void Run()
        {
            using (var p = new ExcelPackage())
            {
                CreateWorksheetsAndLoadData(p);
                var sheet1 = p.Workbook.Worksheets[0];
                var sheet2 = p.Workbook.Worksheets[1];
                // Sort the range by column 0, then by column 1 descending
                sheet1.Cells["A3:D17"].Sort(x => x.SortBy.Column(0).ThenSortBy.Column(1, eSortOrder.Descending));

                // Sort the range left to right by row 0 (using a custom list), then by row 1
                sheet2.Cells["A3:K5"].Sort(x => x.SortLeftToRightBy.Row(0).UsingCustomList("S", "M", "L", "XL", "XXL").ThenSortBy.Row(1));

                p.SaveAs(FileUtil.GetCleanFileInfo("30-SortingRanges.xlsx"));
            }
        }

        private static void CreateWorksheetsAndLoadData(ExcelPackage p)
        {
            var rnd = new Random((int)DateTime.UtcNow.ToOADate());

            var sheet1 = p.Workbook.Worksheets.Add("Sort top down");
            sheet1.Cells["A1"].Value = "To view the sort state in Excel 2019 with english localization, select the range A3:D17, right click and chose 'Sort' followed by 'Custom sort'";
            // create random data for this sheet
            for(var row = StartRow; row < (StartRow + 15); row++)
            {
                for(var col = 1; col < 5; col++)
                {
                    if(col == 1)
                    {
                        var ix = rnd.Next(0, _letters.Length - 1);
                        sheet1.Cells[row, 1].Value = _letters[ix];
                    }
                    else if(col == 4)
                    {
                        // Add a formula in the right most column to demonstrate that the formulas will be shifted when sorted.
                        sheet1.Cells[row, 4].Formula = $"SUM(B{row}:C{row})";
                    }
                    else
                    {
                        sheet1.Cells[row, col].Value = rnd.Next(14, 555);
                    }
                }
            }

            var sheet2 = p.Workbook.Worksheets.Add("Sort left to right");
            sheet2.Cells["A1"].Value = "To view the sort state in Excel 2019 with english localization, select the range A3:K5, right click and chose 'Sort' followed by 'Custom sort'";
            // create random data for this sheet
            for (var col = 1; col < 12; col++)
            {
                for (var row = 3; row < 6; row++)
                {
                    if (row == 3)
                    {
                        var ix = rnd.Next(0, _tShirtSizes.Length - 1);
                        sheet2.Cells[row, col].Value = _tShirtSizes[ix];
                        sheet2.Cells[row, col].Style.HorizontalAlignment = OfficeOpenXml.Style.ExcelHorizontalAlignment.Right;
                    }
                    else
                    {
                        sheet2.Cells[row, col].Value = rnd.Next(14, 555);
                    }
                }
            }
        }
    }
}
