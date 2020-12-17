using OfficeOpenXml;
using OfficeOpenXml.Attributes;
using OfficeOpenXml.Table;
using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusSamples.LoadingData
{
    [EpplusTable(TableStyle = TableStyles.Dark1, PrintHeaders = true, AutofitColumns = true, AutoCalculate = false, ShowTotal = true, ShowFirstColumn = true)]
    [
        EpplusFormulaTableColumn(Order = 6, NumberFormat = "€#,##0.00", Header = "Tax amount", FormulaR1C1 = "RC[-2] * RC[-1]", TotalsRowFunction = RowFunctions.Sum, TotalsRowNumberFormat = "€#,##0.00"),
        EpplusFormulaTableColumn(Order = 7, NumberFormat = "€#,##0.00", Header = "Net salary", Formula = "E2-G2", TotalsRowFunction = RowFunctions.Sum, TotalsRowNumberFormat = "€#,##0.00")
    ]
    internal class Actor
    {
        [EpplusIgnore]
        public int Id { get; set; }

        [EpplusTableColumn(Order = 3)]
        public string LastName { get; set; }
        [EpplusTableColumn(Order = 1, Header = "First name")]
        public string FirstName { get; set; }
        [EpplusTableColumn(Order = 2)]
        public string MiddleName { get; set; }

        [EpplusTableColumn(Order = 0, NumberFormat = "yyyy-MM-dd", TotalsRowLabel = "Total")]
        public DateTime Birthdate { get; set; }

        [EpplusTableColumn(Order = 4, NumberFormat = "€#,##0.00", TotalsRowFunction = RowFunctions.Sum, TotalsRowNumberFormat = "€#,##0.00")]
        public double Salary { get; set; }

        [EpplusTableColumn(Order = 5, NumberFormat = "0%", TotalsRowFormula = "Table1[[#Totals],[Tax amount]]/Table1[[#Totals],[Salary]]", TotalsRowNumberFormat = "0 %")]
        public double Tax { get; set; }
    }

    [EpplusTable(TableStyle = TableStyles.Medium1, PrintHeaders = true, AutofitColumns = true, AutoCalculate = true, ShowLastColumn = true)]
    internal class Actor2 : Actor
    {

    }

    public static class LoadingDataFromCollectionWithAttributes
    {
        public static void Run()
        {
            // sample data
            var actors = new List<Actor>
            {
                new Actor{ Salary = 256.24, Tax = 0.21, FirstName = "John", MiddleName = "Bernhard", LastName = "Doe", Birthdate = new DateTime(1950, 3, 15) },
                new Actor{ Salary = 278.55, Tax = 0.23, FirstName = "Sven", MiddleName = "Bertil", LastName = "Svensson", Birthdate = new DateTime(1962, 6, 10)},
                new Actor{ Salary = 315.34, Tax = 0.28, FirstName = "Lisa", MiddleName = "Maria", LastName = "Gonzales", Birthdate = new DateTime(1971, 10, 2)}
            };

            var subclassActors = new List<Actor2>
            {
                new Actor2{ Salary = 256.24, Tax = 0.21, FirstName = "John", MiddleName = "Bernhard", LastName = "Doe", Birthdate = new DateTime(1950, 3, 15) },
                new Actor2{ Salary = 278.55, Tax = 0.23, FirstName = "Sven", MiddleName = "Bertil", LastName = "Svensson", Birthdate = new DateTime(1962, 6, 10)},
                new Actor2{ Salary = 315.34, Tax = 0.28, FirstName = "Lisa", MiddleName = "Maria", LastName = "Gonzales", Birthdate = new DateTime(1971, 10, 2)}
            };

            using (var package = new ExcelPackage(FileOutputUtil.GetFileInfo("04-LoadFromCollectionAttributes.xlsx")))
            {
                // using the Actor class above
                var sheet = package.Workbook.Worksheets.Add("Actors");
                sheet.Cells["A1"].LoadFromCollection(actors);

                // using a subclass where we have overridden the EpplusTableAttribute (different TableStyle and highlight last column instead of the first).
                var subclassSheet = package.Workbook.Worksheets.Add("Using subclass with attributes");
                subclassSheet.Cells["A1"].LoadFromCollection(subclassActors);
                
                package.Save();
            }
        }
    }
}
