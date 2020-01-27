/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/27/2020         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace EPPlusSamples.FormulaCalculation
{
    /// <summary>
    /// This sample demonstrates the formula calculation engine of EPPlus by opening an existing
    /// workbook and calculate the formulas in it.
    /// </summary>
    public class CalculateExistingWorkbook
    {
        //private static Stream GetResource(string name)
        //{
        //    var assembly = Assembly.GetExecutingAssembly();
        //    return assembly.GetManifestResourceStream(name);

        //}

        private static void RemoveCalculatedFormulaValues(ExcelWorkbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                foreach (var cell in worksheet.Cells)
                {
                    // if there is a formula in the cell, the following code keeps the formula but clears the calculated value.
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        var formula = cell.Formula;
                        cell.Value = null;
                        cell.Formula = formula;
                    }
                }
            }
        }

        public void Run()
        {
            //var resourceStream = GetResource("EPPlusSampleApp.Core.FormulaCalculation.FormulaCalcSample.xlsx");
            var filePath = FileInputUtil.GetFileInfo("06-FormulaCalculation", "FormulaCalcSample.xlsx").FullName;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                // Read the value from the workbook. This is calculated by Excel.
                double? totalSales = package.Workbook.Worksheets["Sales"].Cells["E10"].GetValue<double?>();
                Console.WriteLine("Total sales read from Cell E10: {0}", totalSales.Value);

                // This code removes all calculated values
                RemoveCalculatedFormulaValues(package.Workbook);

                // totalSales from cell C10 should now be empty
                totalSales = package.Workbook.Worksheets["Sales"].Cells["E10"].GetValue<double?>();
                Console.WriteLine("Total sales read from Cell E10: {0}", totalSales.HasValue ? totalSales.Value.ToString() : "null");


                // ************** 1. Calculate the entire workbook **************
                package.Workbook.Calculate();

                // totalSales should now be recalculated
                totalSales = package.Workbook.Worksheets["Sales"].Cells["E10"].GetValue<double?>();
                Console.WriteLine("Total sales read from Cell E10: {0}", totalSales.HasValue ? totalSales.Value.ToString() : "null");

                // ************** 2. Calculate a worksheet **************

                // This code removes all calculated values
                RemoveCalculatedFormulaValues(package.Workbook);

                package.Workbook.Worksheets["Sales"].Calculate();

                // totalSales should now be recalculated
                totalSales = package.Workbook.Worksheets["Sales"].Cells["E10"].GetValue<double?>();
                Console.WriteLine("Total sales read from Cell E10: {0}", totalSales.HasValue ? totalSales.Value.ToString() : "null");

                // ************** 3. Calculate a range **************

                // This code removes all calculated values
                RemoveCalculatedFormulaValues(package.Workbook);

                package.Workbook.Worksheets["Sales"].Cells["E10"].Calculate();

                // totalSales should now be recalculated
                totalSales = package.Workbook.Worksheets["Sales"].Cells["E10"].GetValue<double?>();
                Console.WriteLine("Total sales read from Cell E10: {0}", totalSales.HasValue ? totalSales.Value.ToString() : "null");
            }
            
        }
    }
}
