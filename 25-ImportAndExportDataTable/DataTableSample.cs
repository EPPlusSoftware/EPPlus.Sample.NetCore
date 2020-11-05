using OfficeOpenXml;
using OfficeOpenXml.Export.ToDataTable;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SQLite;
using System.Text;

namespace EPPlusSamples._25_ImportAndExportDataTable
{
    public class DataTableSample
    {
        public static void Run(string connectionString)
        {
            using (var sqlConn = new SQLiteConnection(connectionString))
            {
                sqlConn.Open();
                using (var sqlCmd = new SQLiteCommand("select CompanyName as 'Company Name', [Name] as Name, Email as 'E-Mail', c.Country as Country, orderdate as 'Order Date', (ordervalue) as 'Order Value',tax as Tax, freight As Freight, currency As Currency from Customer c inner join Orders o on c.CustomerId=o.CustomerId inner join SalesPerson s on o.salesPersonId = s.salesPersonId ORDER BY 1,2 desc", sqlConn))
                {
                    var reader = sqlCmd.ExecuteReader();
                    var dataTable = new DataTable();
                    dataTable.Load(reader);

                    // Create a workbook 
                    using(var package = new ExcelPackage())
                    {
                        var sheet = package.Workbook.Worksheets.Add("DataTable Samples");

                        /***** Load from DataTable ***/

                        // Import the DataTable using LoadFromDataTable
                        sheet.Cells["A1"].LoadFromDataTable(dataTable, true, TableStyles.Dark11);

                        // Now let's export this data back to a DataTable. We know that the data is in a 
                        // table, so we are using the ExcelTables interface to get the range
                        var dt1 = sheet.Tables[0].ToDataTable();
                        PrintDataTable(dt1);


                        /***** Export to DataTable ***/

                        // Export a specific range instead of the entire table
                        // and use the config action to set the table name
                        var dt2 = sheet.Cells["A1:F11"].ToDataTable(o => o.DataTableName = "dt2");
                        PrintDataTable(dt2);

                        // Configure some properties on how the table is generated
                        var dt3 = sheet.Cells["A1:F11"].ToDataTable(c =>
                        {
                            // set name and namespace
                            c.DataTableName = "MyDataTable";
                            c.DataTableNamespace = "MyNamespace";
                            // Removes spaces in column names when read from the first row
                            c.ColumnNameParsingStrategy = NameParsingStrategy.RemoveSpace;
                            // Rename the third column from E-Mail to EmailAddress
                            c.Mappings.Add(2, "EmailAddress");
                            // Ensure that the OrderDate column is casted to DateTime (in Excel it can sometimes be stored as a double/OADate)
                            c.Mappings.Add(4, "OrderDate", typeof(DateTime));
                            // Change the OrderValue to a string
                            c.Mappings.Add(5, "OrderValue", typeof(string), false, cellVal => "Val: " + cellVal.ToString());
                            // Skip the first 2 rows
                            c.SkipNumberOfRowsStart = 2;
                            // Skip the last 100 rows
                            c.SkipNumberOfRowsEnd = 4;

                        });
                        PrintDataTable(dt3);

                        // Export to existing DataTable

                        // Create the DataTable
                        var dataTable2 = new DataTable("myDataTable", "myNamespace");
                        dataTable2.Columns.Add("Company Name", typeof(string));
                        dataTable2.Columns.Add("E-Mail");
                        sheet.Cells["A1:F11"].ToDataTable(o => o.FirstRowIsColumnNames = true, dataTable2);
                        PrintDataTable(dataTable2);

                        // Create the DataTable, use mappings if names of columns/range headers differ
                        var dataTable3 = new DataTable("myDataTableWithMappings", "myNamespace");
                        var col1 = dataTable3.Columns.Add("CompanyName");
                        var col2 = dataTable3.Columns.Add("Email");
                        sheet.Cells["A1:F11"].ToDataTable(o =>
                        {
                            o.FirstRowIsColumnNames = true;
                            o.Mappings.Add(0, col1);
                            o.Mappings.Add(1, col2);
                        }
                        , dataTable3);
                        PrintDataTable(dataTable3);

                    }

                }
            }  
        }

        private static void PrintDataTable(DataTable table)
        {
            Console.WriteLine();
            Console.WriteLine("DATATABLE name=" + table.TableName);
            var cols = new StringBuilder();
            foreach (var col in table.Columns)
            {
                cols.AppendFormat("'{0}' ", ((DataColumn)col).ColumnName);
            }
            Console.WriteLine("Columns:");
            Console.WriteLine(cols.ToString());
            Console.WriteLine();

            Console.WriteLine("First 10 rows:");
            for(var r = 0; r < table.Rows.Count && r < 10; r++)
            {
                for(var c = 0; c < table.Columns.Count; c++)
                {
                    var col = table.Columns[c] as DataColumn;
                    var row = table.Rows[r] as DataRow;
                    var val = col.DataType == typeof(string) ? "'" + row[col.ColumnName] + "'" : row[col.ColumnName];


                    Console.Write(c == 0 ? val : ", " + val);
                }
                Console.WriteLine();
            }
        }
    }
}
