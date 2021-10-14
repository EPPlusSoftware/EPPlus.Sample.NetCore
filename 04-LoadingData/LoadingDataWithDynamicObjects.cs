/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  07/22/2020         EPPlus Software AB           EPPlus 5.2.1
 *************************************************************************************************/
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.LoadFunctions.Params;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Text;

namespace EPPlusSamples.LoadingData
{
    public static class LoadingDataWithDynamicObjects
    {
        public static void Run()
        {
            // Create a list of dynamic objects
            dynamic p1 = new ExpandoObject();
            p1.Id = 1;
            p1.FirstName = "Ivan";
            p1.LastName = "Horvat";
            p1.Age = 21;
            dynamic p2 = new ExpandoObject();
            p2.Id = 2;
            p2.FirstName = "John";
            p2.LastName = "Doe";
            p2.Age = 45;
            dynamic p3 = new ExpandoObject();
            p3.Id = 3;
            p3.FirstName = "Sven";
            p3.LastName = "Svensson";
            p3.Age = 68;

            List<ExpandoObject> items = new List<ExpandoObject>()
            {
                p1,
                p2,
                p3
            };

            // Create a workbook with a worksheet and load the data into a table
            using(var package = new ExcelPackage(FileUtil.GetCleanFileInfo("04-LoadDynamicObjects.xlsx")))
            {
                var sheet = package.Workbook.Worksheets.Add("Dynamic");
                sheet.Cells["A1"].LoadFromDictionaries(items, c =>
                {
                    // Print headers using the property names
                    c.PrintHeaders = true;
                    // insert a space before each capital letter in the header
                    c.HeaderParsingType = HeaderParsingTypes.CamelCaseToSpace;
                    // when TableStyle is not TableStyles.None the data will be loaded into a table with the 
                    // selected style.
                    c.TableStyle = TableStyles.Medium1;
                });
                package.Save();
            }

            // Load data from json (in this case a file)
            var jsonItems = JsonConvert.DeserializeObject<IEnumerable<ExpandoObject>>(File.ReadAllText(FileUtil.GetFileInfo("04-LoadingData", "testdata.json").FullName));
            using (var package = new ExcelPackage(FileUtil.GetCleanFileInfo("04-LoadJsonFromFile.xlsx")))
            {
                var sheet = package.Workbook.Worksheets.Add("Dynamic");
                sheet.Cells["A1"].LoadFromDictionaries(jsonItems, c =>
                {
                    // Print headers using the property names
                    c.PrintHeaders = true;
                    // insert a space before each capital letter in the header
                    c.HeaderParsingType = HeaderParsingTypes.CamelCaseToSpace;
                    // when TableStyle is not TableStyles.None the data will be loaded into a table with the 
                    // selected style.
                    c.TableStyle = TableStyles.Medium1;
                });
                sheet.Cells["D:D"].Style.Numberformat.Format = "yyyy-mm-dd";
                sheet.Cells[1, 1, sheet.Dimension.End.Row, sheet.Dimension.End.Column].AutoFitColumns();
                package.Save();
            }
        }
    }
}
