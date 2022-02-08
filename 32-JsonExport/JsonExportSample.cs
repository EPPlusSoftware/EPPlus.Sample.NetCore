using OfficeOpenXml;
using System.IO;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public static class JsonExportSample
    {
        //This sample demonstrates how to export html from a table.
        //More advanced samples using charts and json exports are available in our samples web site available 
        //here: https://samples.epplussoftware.com/JsonExport
        public static async Task Run()
        {
            var outputFolder = FileUtil.GetDirectoryInfo("JsonOutput");

            //Start by using the excel file generated in sample 28
            using (var p = new ExcelPackage(FileUtil.GetFileInfo("28-Tables.xlsx")))
            {
                var wsSimpleTable = p.Workbook.Worksheets["SimpleTable"];

                ExportTable1(outputFolder, wsSimpleTable);

                var wsStyleTables = p.Workbook.Worksheets["StyleTables"];
                await ExportTableWithHyperlink(outputFolder, wsStyleTables);
            }
        }

        private static void ExportTable1(DirectoryInfo outputFolder, ExcelWorksheet wsSimpleTable)
        {
            var table1 = wsSimpleTable.Tables[0];
            
            //First export the table directly from the table object.
            //When exporting a table the data type is set on the column.
            var json = table1.ToJson(x =>
            {
                x.Minify = false;
            });

            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "TableSample1_As_Table.json", true).FullName, json);

            //When exporting the range data types are set on the cell level.
            //You can alter this by AddDataTypesOn, --> x.AddDataTypesOn=eDataTypeOn.OnColumn
            json = table1.Range.ToJson(x =>
            {
                x.Minify = false;                
            });

            File.WriteAllText(FileUtil.GetFileInfo(outputFolder, "TableSample1_As_Range.json", true).FullName, json);
        }
        private static async Task ExportTableWithHyperlink(DirectoryInfo outputFolder, ExcelWorksheet wsStyleTables)
        {
            var table1 = wsStyleTables.Tables[0];


            using (var fs = new FileStream(FileUtil.GetFileInfo(outputFolder, "TableSample2_hyperlinks.json", true).FullName, FileMode.Create, FileAccess.Write))
            {
                await table1.SaveToJsonAsync(fs, x =>
                {
                    x.AddDataTypesOn = eDataTypeOn.NoDataTypes; //Skip data types.
                    x.Minify = false;
                });
                fs.Close();
            }
        }
    }
}