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

namespace EPPlusSamples
{
    public static class FillRangeSample
    {
        public static void Run()
        {
            using (var p = new ExcelPackage())
            {
                FillNumber(p);
                FillDate(p);
                FillList(p);
                p.SaveAs(FileUtil.GetCleanFileInfo("30-FillRangeSamples.xlsx"));
            }
        }
        /// <summary>
        /// Several samples how to use the FillNumber method
        /// </summary>
        /// <param name="p">The package</param>
        private static void FillNumber(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("FillNumber Samples");

            ws.SetValue(1, 1, 50);
            //Fill the range by adding 1 for each cell starting from the value in the top-left cell.
            ws.Cells["A1:A20"].FillNumber();

            //Fill the range by subtracting 2 from the initial value 30.
            ws.Cells["B1:B20"].FillNumber(30, -2);

            //Fill by starting with 100 and increase 5% for each cell. Fill by left by row
            ws.Cells["D2:AA2"].FillNumber(x =>
            {
                x.CalculationMethod = eCalculationMethod.Multiply;
                x.StartValue = 100;
                x.StepValue = 1.05;
                x.Direction = eFillDirection.Row;
            });
        }
        /// <summary>
        /// Several samples how to use the FillDate method
        /// </summary>
        /// <param name="p">The package</param>
        private static void FillDate(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("FillDateTime Samples");

            //Fill dates starting from the value in the first cell. By default a 1 day increase is assumed and fill is performed per column downwards.
            ws.SetValue("A2", new DateTime(2021, 1, 1));
            ws.Cells["A2:A60"].FillDateTime();

            //Fill dates using the starting value from a fixed start value instead of using the first cell.
            ws.Cells["B2:B60"].FillDateTime(new DateTime(2021, 6, 30));

            ws.Cells[2, 1, 60, 2].Style.Numberformat.Format = "yyyy-mm-dd";

            //Fill dates per last day of the quarter. If the start value is the last day of the month, this is used for all dates in the fill. 
            //This sample exclude weekends and adds some holiday dates. 
            //Column C2 and D2 are used as start values.
            ws.Cells["C2"].Value = new DateTime(2015, 6, 30);
            ws.Cells["D2"].Value = new DateTime(2009, 2, 28);
            ws.Cells["C2:D60"].FillDateTime(x =>
            {                
                x.DateTimeUnit = eDateTimeUnit.Month;
                x.StepValue = 3;
                x.NumberFormat = "yyyy-mm-dd";
                x.SetExcludedWeekdays(DayOfWeek.Saturday, DayOfWeek.Sunday); //We exclude weekends. The day before is used instead.
                x.SetExcludedDates(                                          //These dates are also excluded. The day before is used instead.
                    new DateTime(2010, 12, 31),
                    new DateTime(2012, 12, 31),
                    new DateTime(2013, 12, 31),
                    new DateTime(2014, 12, 31),
                    new DateTime(2015, 12, 31),
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

            //You can also fill row-wise and fill right-to-left. 
            //Note that as row 6 don't have a start value it's left blank in the fill.
            //We also set an end value where the fill will stop.
            ws.Cells["AA4"].Value = new DateTime(2030, 1, 1);
            ws.Cells["AA5"].Value = new DateTime(2030, 1, 5);
            ws.Cells["AA7"].Value = new DateTime(2030, 1, 10);
            ws.Cells["F4:AA7"].FillDateTime(x =>
            {
                x.Direction = eFillDirection.Row;
                x.StartPosition = eFillStartPosition.BottomRight;
                x.EndValue = new DateTime(2030, 1, 20);
                x.NumberFormat = "yyyy-mm-dd";
            });

            ws.Columns[1, 27].AutoFit();
        }
        /// <summary>
        /// Several samples how to use the FillList method
        /// </summary>
        /// <param name="p">The package</param>
        private static void FillList(ExcelPackage p)
        {
            var ws = p.Workbook.Worksheets.Add("Fill List Samples");

            var list = new[] { "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday" };
            ws.Cells["A1:A20"].FillList(list);

            ws.Cells["B1:B20"].FillList(list,  x=> { x.StartIndex=6; });
            ws.Cells["E1:AA1"].FillList(list, x => 
            { 
                x.Direction = eFillDirection.Row;
            });

            //Fill the list of primes starting from the bottom-up.
            //We set the range to the size of the list so it's not repeated.
            var primes = new List<int>{ 2,5,7,11,13,17,19,23,29,997,1009 };
            ws.Cells[1,3,primes.Count, 3].FillList(primes, x =>
            {
                x.NumberFormat = "#,##0";
                x.StartPosition = eFillStartPosition.BottomRight;
            });

        }
    }
}