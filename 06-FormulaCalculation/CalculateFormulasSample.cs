using System;
using System.Collections.Generic;
using System.Text;

namespace EPPlusSamples.FormulaCalculation
{
    /// <summary>
    /// Sample 17 demonstrates the formula calculation engine of EPPlus.
    /// </summary>
    static class CalculateFormulasSample
    {
        private static CalculateExistingWorkbook CalculateExistingWorkbook = new CalculateExistingWorkbook();

        private static BuildAndCalculateWorkbook BuildAndCalculateWorkbook = new BuildAndCalculateWorkbook();

        private static AddFormulaFunction AddFormulaFunction = new AddFormulaFunction();

        public static void Run()
        {
            CalculateExistingWorkbook.Run();
            BuildAndCalculateWorkbook.Run();
            AddFormulaFunction.Run();
        }

    }
}
