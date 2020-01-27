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
using OfficeOpenXml;
using System.Data.SqlClient;
using System.Drawing;
using OfficeOpenXml.Style;
using System.Data.SQLite;
using System.Threading.Tasks;
using System.Threading;
using OfficeOpenXml.Table;

namespace EPPlusSamples.SalesReport
{
    class UsingAsyncAwaitSample
    {
        /// <summary>
        /// Shows a few different ways to load / save asynchronous
        /// </summary>
        /// <param name="connectionString">The connection string to the SQLite database</param>
        public static async Task RunAsync(string connectionString)
        {
            var file = FileOutputUtil.GetFileInfo("03-AsyncAwait.xlsx");
            using (ExcelPackage package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets.Add("Sheet1");

                using (var sqlConn = new SQLiteConnection(connectionString))
                {
                    sqlConn.Open();
                    using (var sqlCmd = new SQLiteCommand("select CompanyName, [Name], Email, Country, o.OrderId, orderdate, ordervalue, currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                    {
                        var range = await ws.Cells["B2"].LoadFromDataReaderAsync(sqlCmd.ExecuteReader(), true, "Table1", TableStyles.Medium10);
                        range.AutoFitColumns();
                    }
                }

                await package.SaveAsync();
            }

            //Load the package async again.
            using (var package = new ExcelPackage())
            {
                await package.LoadAsync(file);

                var newWs = package.Workbook.Worksheets.Add("AddedSheet2");
                var range = await newWs.Cells["A1"].LoadFromTextAsync(FileInputUtil.GetFileInfo("03-UsingAsyncAwait", "Importfile.txt"), new ExcelTextFormat { Delimiter='\t' });
                range.AutoFitColumns();

                await package.SaveAsAsync(FileOutputUtil.GetFileInfo("03-AsyncAwait-LoadedAndModified.xlsx"));
            }
        }
    }
}
