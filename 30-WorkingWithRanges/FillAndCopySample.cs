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

namespace EPPlusSamples
{
    public static class FillAndCopySample
    {
        public static void Run()
        {
            using (var p = new ExcelPackage())
            {
                FillDate(p);
                p.SaveAs(FileOutputUtil.GetFileInfo("30-FillAndCopySamples.xlsx"));
            }
        }

        private static void FillDate(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("FillDateTime Samples");

            //Fill dates starting from the value in the first cell. By default a 1 day increse is assumed and fill is performed per column downwards.
            ws.SetValue("A2", new DateTime(2021, 1, 1));
            ws.Cells["A2:A60"].FillDateTime();

            //Fill dates using the starting value from a fixed start value instead of using the first cell.
            ws.Cells["B2:B60"].FillDateTime(new DateTime(2021, 6, 30));

            ws.Cells[2, 1, 60, 2].Style.Numberformat.Format = "yyyy-mm-dd";

            //Fill dates per last day of the quater. If the start value is the last day of the month, this is used for all dates in the fill. 
            //This sample excludes weekends and adds some holiday dates. 
            ws.Cells["C2:C60"].FillDateTime(x =>
            {
                x.StartValue = new DateTime(2015, 6, 30);
                x.DateUnit = eDateTimeUnit.Month;
                x.StepValue = 3;
                x.NumberFormat = "yyyy-mm-dd";
                x.SetExcludedWeekdays(DayOfWeek.Saturday, DayOfWeek.Sunday);
                x.SetHolidayDates(
                    new DateTime(2015, 12, 31),
                    new DateTime(2018, 12, 31),
                    new DateTime(2019, 12, 31),
                    new DateTime(2020, 12, 31),
                    new DateTime(2021, 12, 31),
                    new DateTime(2024, 12, 31),
                    new DateTime(2025, 12, 31),
                    new DateTime(2026, 12, 31),
                    new DateTime(2027, 12, 31),
                    new DateTime(2029, 12, 31)
                );
            });


            ws.Columns[1, 5].AutoFit();
        }
    }
}