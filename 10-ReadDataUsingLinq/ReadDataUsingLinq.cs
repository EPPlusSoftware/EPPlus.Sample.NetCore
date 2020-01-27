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
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Xml;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Linq;
namespace EPPlusSamples
{
    public static class ReadDataUsingLinq
    {
        /// <summary>
        /// This sample shows how to use Linq with the Cells collection
        /// </summary>
        /// <param name="outputDir">The path where sample7.xlsx is</param>
        public static void Run()
        {
	        Console.WriteLine("Now open sample 9 again and perform some Linq queries...");
		    Console.WriteLine();

            FileInfo existingFile = FileOutputUtil.GetFileInfo("09-PerformanceAndProtection.xlsx", false);
            using (ExcelPackage package = new ExcelPackage(existingFile))
            {
                ExcelWorksheet sheet = package.Workbook.Worksheets[0];
                
                //Select all cells in column d between 9990 and 10000
                var query1= (from cell in sheet.Cells["d:d"] where cell.Value is double && (double)cell.Value >= 9990 && (double)cell.Value <= 10000 select cell);

                Console.WriteLine("Print all cells with value between 9990 and 10000 in column D ...");
                Console.WriteLine();

                int count = 0;
                foreach (var cell in query1)
                {
                    Console.WriteLine("Cell {0} has value {1:N0}", cell.Address, cell.Value);
                    count++;
                }

                Console.WriteLine("{0} cells found ...",count);
                Console.WriteLine();

                //Select all bold cells
                Console.WriteLine("Now get all bold cells from the entire sheet...");
                var query2 = (from cell in sheet.Cells[sheet.Dimension.Address] where cell.Style.Font.Bold select cell);
                //If you have a clue where the data is, specify a smaller range in the cells indexer to get better performance (for example "1:1,65536:65536" here)
                count = 0;
                foreach (var cell in query2)
                {
                    if (!string.IsNullOrEmpty(cell.Formula))
                    {
                        Console.WriteLine("Cell {0} is bold and has a formula of {1:N0}", cell.Address, cell.Formula);
                    }
                    else
                    {
                        Console.WriteLine("Cell {0} is bold and has a value of {1:N0}", cell.Address, cell.Value);
                    }
                    count++;
                }

                //Here we use more than one column in the where clause. We start by searching column D, then use the Offset method to check the value of column C.
                var query3 = (from cell in sheet.Cells["d:d"]
                              where cell.Value is double && 
                                    (double)cell.Value >= 9500 && (double)cell.Value <= 10000 && 
                                    cell.Offset(0, -1).GetValue<DateTime>().Year == DateTime.Today.Year+1 
                              select cell);

                Console.WriteLine();
                Console.WriteLine("Print all cells with a value between 9500 and 10000 in column D and the year of Column C is {0} ...", DateTime.Today.Year + 1);
                Console.WriteLine();

                count = 0;
                foreach (var cell in query3)    //The cells returned here will all be in column D, since that is the address in the indexer. Use the Offset method to print any other cells from the same row.
                {
                    Console.WriteLine("Cell {0} has value {1:N0} Date is {2:d}", cell.Address, cell.Value, cell.Offset(0, -1).GetValue<DateTime>());
                    count++;
                }
            }
        }
    }
}
