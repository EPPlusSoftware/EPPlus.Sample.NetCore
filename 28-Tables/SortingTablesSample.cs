using OfficeOpenXml;
using OfficeOpenXml.Sorting;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Text;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    /// <summary>
    /// This sample demonstrates how to sort Excel tables in EPPlus.
    /// </summary>
    public static class SortingTablesSample
    {
        public static async Task RunAsync(string connectionString)
        {
            var file = FileOutputUtil.GetFileInfo("28-SortingTables.xlsx");
            using (ExcelPackage package = new ExcelPackage(file))
            {
                // Sheet 1
                var sheet1 = package.Workbook.Worksheets.Add("Sheet1");
                sheet1.Cells["B1"].Value = "This table is sorted by country DESC, then name ASC, then orderValue ASC";
                using (var sqlConn = new SQLiteConnection(connectionString))
                {
                    sqlConn.Open();
                    using (var sqlCmd = new SQLiteCommand("select CompanyName, [Name], Email, c.Country, o.OrderId, orderdate, ordervalue, currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                    {
                        var range = await sheet1.Cells["B3"].LoadFromDataReaderAsync(sqlCmd.ExecuteReader(), true, "Table1", TableStyles.Medium10);
                        range.AutoFitColumns();
                    }
                }
                // sort this table by country DESC, then by sales persons name ASC, then by Order value ASC
                var table1 = sheet1.Tables[0];
                table1.Sort(x => x.SortBy.ColumnNamed("country", eSortOrder.Descending)
                                    .ThenSortBy.ColumnNamed("name")
                                    .ThenSortBy.ColumnNamed("orderValue"));


                // Sheet 2
                var sheet2 = package.Workbook.Worksheets.Add("Using custom list");
                sheet2.Cells["B1"].Value = "This table is sorted by country with a custom list, then name ASC, then orderValue ASC. The custom lists ensures that Greenland and Costa Rica comes first in the sort";
                using (var sqlConn = new SQLiteConnection(connectionString))
                {
                    sqlConn.Open();
                    using (var sqlCmd = new SQLiteCommand("select CompanyName, [Name], Email, c.Country, o.OrderId, orderdate, ordervalue, currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                    {
                        var range = await sheet2.Cells["B3"].LoadFromDataReaderAsync(sqlCmd.ExecuteReader(), true, "Table2", TableStyles.Medium10);
                        range.AutoFitColumns();
                    }
                }
                // sort this table by country ASC, then by sales persons name ASC, then by Order value ASC
                var table2 = sheet2.Tables["Table2"];
                table2.Sort(x => x.SortBy.ColumnNamed("country", eSortOrder.Descending).UsingCustomList("Greenland", "Costa Rica")
                                    .ThenSortBy.ColumnNamed("name")
                                    .ThenSortBy.ColumnNamed("orderValue"));

                await package.SaveAsync();
            }
        }
    }
}
