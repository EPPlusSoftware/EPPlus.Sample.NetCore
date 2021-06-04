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
            ReadFileWithExternalLink();

            var file = FileOutputUtil.GetFileInfo("29-ExternalLinks.xlsx");
            using (var p = new ExcelPackage(file))
            {

                AddWorksheetWithExternalReferences(p);

                AddWorksheetWithExternalReferencesInFormulas(p);


                p.Save();
            }
        }

        private static void ReadFileWithExternalLink()
        {
            var externalFile = FileInputUtil.GetFileInfo("29-External References", "WorkbookWithExternalLinks.xlsx");
            using (var p = new ExcelPackage(externalFile))
            {
                var ws = p.Workbook.Worksheets[0];
                //This worksheet contains references to an external workbook that is not available. 
                Console.WriteLine("Values from Excel:");
                Console.WriteLine($"Cell C1 formula : {ws.Cells["C1"].Formula} with value {ws.Cells["C1"].Value}");
                Console.WriteLine($"Cell C2 formula : {ws.Cells["C2"].Formula} with value {ws.Cells["C2"].Value}");
                Console.WriteLine($"Cell C3 formula : {ws.Cells["C3"].Formula} with value {ws.Cells["C3"].Value}");
                Console.WriteLine($"Cell C5 formula : {ws.Cells["C5"].Formula} with value {ws.Cells["C5"].Text}");
                Console.WriteLine($"Cell C6 formula : {ws.Cells["C6"].Formula} with value {ws.Cells["C6"].Value}");
                Console.WriteLine();

                ws.ClearFormulaValues();
                ws.Calculate();

                Console.WriteLine("Values after calculation in EPPLus from external reference cache:");
                Console.WriteLine($"Cell C1 formula : {ws.Cells["C1"].Formula} with value {ws.Cells["C1"].Value}");
                Console.WriteLine($"Cell C2 formula : {ws.Cells["C2"].Formula} with value {ws.Cells["C2"].Value}");
                Console.WriteLine($"Cell C3 formula : {ws.Cells["C3"].Formula} with value {ws.Cells["C3"].Value}");
                Console.WriteLine($"Cell C5 formula : {ws.Cells["C5"].Formula} with value {ws.Cells["C5"].Text}");
                Console.WriteLine($"Cell C6 formula : {ws.Cells["C6"].Formula} with value {ws.Cells["C6"].Value}");
                Console.WriteLine();

                var externalReference = p.Workbook.ExternalReferences[0].As.ExternalWorkbook;
                p.Workbook.ExternalReferences.Directories.Add(FileInputUtil.GetSubDirectory("29-External References","Data"));
                externalReference.Load();

                ws.ClearFormulaValues();
                ws.Calculate();

                Console.WriteLine("Values after calculation in EPPLus when the external package has been loaded:");
                Console.WriteLine($"Cell C1 formula : {ws.Cells["C1"].Formula} with value {ws.Cells["C1"].Value}");
                Console.WriteLine($"Cell C2 formula : {ws.Cells["C2"].Formula} with value {ws.Cells["C2"].Value}");
                Console.WriteLine($"Cell C3 formula : {ws.Cells["C3"].Formula} with value {ws.Cells["C3"].Value}");
                Console.WriteLine($"Cell C5 formula : {ws.Cells["C5"].Formula} with value {ws.Cells["C5"].Text}");
                Console.WriteLine($"Cell C6 formula : {ws.Cells["C6"].Formula} with value {ws.Cells["C6"].Value}");
                Console.WriteLine();
            }
        }

        private static void AddWorksheetWithExternalReferences(ExcelPackage p)
        {
            //Add a reference to the file created by sample 28.
            var externalLinkFile = FileOutputUtil.GetFileInfo("28-Tables.xlsx", false);
            var externalWorkbook = p.Workbook.ExternalReferences.AddExternalWorkbook(externalLinkFile);

            var ws = p.Workbook.Worksheets.Add("Sheet1");
            //You can access individual cells like this using the index of the external reference in brackets...
            //[1] reference to the the first item in the ExternalReferences collection. This is the externalWorkbook. Index property
            ws.Cells["A1:C3"].Formula = "[1]SimpleTable!A1";

            //You can also reference a value and set a format. Here we use the index property instead of hardcoding it in the formula.
            ws.Cells["F1"].Formula = $"[{externalWorkbook.Index}]Slicer!F38";
            ws.Cells["F1"].Style.Numberformat.Format = "yyyy-MM-dd";

            //Now, Calculate. As the workbook is loaded EPPlus will use the actual values in the package.
            ws.Calculate();

            //The cache stores cell values that are referenced when calculating formulas in the workbook, so formulas can be calculated without access to the external workbook.
            externalWorkbook.UpdateCache();
            var f38Value = externalWorkbook.CachedWorksheets["Slicer"].CellValues["F38"].Value;
            Console.WriteLine($"Value of cached value in {externalWorkbook.File.Name} worksheet Slicer cell F38 is : {f38Value}");
            
            ws.Cells.AutoFitColumns();
            Console.WriteLine($"Cell F1 with an external link has value: {ws.Cells["F1"].Value} - formatted: {ws.Cells["F1"].Text}");
        }
        private static void AddWorksheetWithExternalReferencesInFormulas(ExcelPackage p)
        {
            var externalLinkFile = FileOutputUtil.GetFileInfo("01-GettingStarted.xlsx", false);
            var externalWorkbook = p.Workbook.ExternalReferences.AddExternalWorkbook(externalLinkFile);

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

            //If you only want to update the cache you can use externalWorkbook.UpdateCache();
            //Note: The cache is updated when the package is saved if externalWorkbook.CacheStatue is eExternalWorkbookCacheStatus.NotUpdated.
            //      If the cache status is LoadedFromPackage or Updated, you must make sure to update the cache, if necessary, before saving the packge
            externalWorkbook.UpdateCache();


            ws.Cells.AutoFitColumns();
        }
    }
}
