using OfficeOpenXml;
using System;
using System.IO;
namespace EPPlusSamples
{
    /// <summary>
    /// This sample demonstrates how work with External references in EPPlus.
    /// EPPlus supports adding, updating and removing external workbooks of type xlsx, xlsm and xlst. EPPlus also use the external reference cache for External workbooks. 
    /// EPPlus will also preserve DDE and OLE links.
    /// </summary>
    public static class ExternalReferencesSample
    {
        public static void Run()
        {
            var file = FileOutputUtil.GetFileInfo("29-ExternalLinks.xlsx");
            using (var p = new ExcelPackage(file))
            {
                //Add a reference to the file created by sample 28.
                var externalLinkFile1 = FileOutputUtil.GetFileInfo("28-Tables.xlsx", false);
                var externalWorkbook1 = p.Workbook.ExternalReferences.AddExternalWorkbook(externalLinkFile1);

                AddWorksheetWithExternalReferences(p, externalWorkbook1);

                var externalLinkFile2 = FileOutputUtil.GetFileInfo("01-GettingStarted.xlsx", false);
                var externalWorkbook2 = p.Workbook.ExternalReferences.AddExternalWorkbook(externalLinkFile2);

                AddWorksheetWithExternalReferencesInFormulas(p, externalWorkbook2);


                p.Save();
            }
        }

        private static void AddWorksheetWithExternalReferences(ExcelPackage p, OfficeOpenXml.ExternalReferences.ExcelExternalWorkbook externalWorkbook)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet1");
            //You can access individual cells like this using the index of the external reference in brackets...
            //[1] reference to the the first item in the ExternalReferences collection. This is the externalWorkbook.Index property
            ws.Cells["A1:C3"].Formula = "[1]SimpleTable!A1";

            //You can also reference a value and set a format. Here we use the index property instead of hardcoding it in the formula.
            ws.Cells["F1"].Formula = $"[{externalWorkbook.Index}]Slicer!F38";
            ws.Cells["F1"].Style.Numberformat.Format = "yyyy-MM-dd";

            //Now, Calculate. This will update the cache and get the values.
            //If you only want to update the cache you can use externalWorkbook.UpdateCache();            
            ws.Calculate();
            ws.Cells.AutoFitColumns();
            Console.WriteLine($"Cell F1 with an external link has value: {ws.Cells["F1"].Value} - formatted: {ws.Cells["F1"].Text}");
        }
        private static void AddWorksheetWithExternalReferencesInFormulas(ExcelPackage p, OfficeOpenXml.ExternalReferences.ExcelExternalWorkbook externalWorkbook)
        {
            var ws = p.Workbook.Worksheets.Add("Sheet2");

            ws.Cells["A1"].Value = "Quantity*Price:";
            ws.Cells["B1:B3"].Formula = "[2]Inventory!C2*[2]Inventory!D2";  //Here we reference the second external reference, so index is [2]

            ws.Cells["B4"].Formula = "Sum(B1:B3)";
            ws.Cells["C4"].Formula = "[2]Inventory!E5";

            ws.Cells["A4"].Value = "SUM:";
            ws.Cells["A4"].AddComment("Sum of external cells matches the sum from cell E5 in the original workbook", "EPPlus Software");
            ws.Cells["B4:C4"].Style.Font.Bold = true;
            ws.Cells["B4:C4"].Style.Numberformat.Format = "#,##0";

            ws.Calculate();

            ws.Cells.AutoFitColumns();
        }
    }
}
