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
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;
using System.Drawing;
using OfficeOpenXml.Drawing.Chart.Style;

namespace EPPlusSamples.OpenWorkbookAddDataAndChart
{
    public class OpenWorkbookAndAddDataAndChartSample
    {
        /// <summary>
        /// Sample 7 - open Sample 1 and add 2 new rows and a Piechart
        /// </summary>
        public static string Run()
        {
            FileInfo newFile = FileUtil.GetCleanFileInfo("07-OpenWorkbookAndAddDataAndChartSample.xlsx");
            FileInfo templateFile = FileUtil.GetFileInfo("07-OpenWorkbookAddDataAndChart", "ExistingWorkbook.xlsx");

            using (ExcelPackage package = new ExcelPackage(newFile, templateFile))
            {
                //Open the first worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                worksheet.InsertRow(5, 2);

                worksheet.Cells["A5"].Value = "12010";
                worksheet.Cells["B5"].Value = "Drill";
                worksheet.Cells["C5"].Value = 20;
                worksheet.Cells["D5"].Value = 8;

                worksheet.Cells["A6"].Value = "12011";
                worksheet.Cells["B6"].Value = "Crowbar";
                worksheet.Cells["C6"].Value = 7;
                worksheet.Cells["D6"].Value = 23.48;

                worksheet.Cells["E2:E6"].FormulaR1C1 = "RC[-2]*RC[-1]";                

                var name = worksheet.Names.Add("SubTotalName", worksheet.Cells["C7:E7"]);
                name.Style.Font.Italic = true;
                name.Formula = "SUBTOTAL(9,C2:C6)";

                //Format the new rows
                worksheet.Cells["C5:C6"].Style.Numberformat.Format = "#,##0";
                worksheet.Cells["D5:E6"].Style.Numberformat.Format = "#,##0.00";

                var chart = worksheet.Drawings.AddPieChart("PieChart", ePieChartType.Pie3D);
                
                chart.Title.Text = "Total";
                //From row 1 colum 5 with five pixels offset
                chart.SetPosition(0, 0, 5, 5);
                chart.SetSize(600, 300);

                ExcelAddress valueAddress = new ExcelAddress(2, 5, 6, 5);
                var ser = (chart.Series.Add(valueAddress.Address, "B2:B6") as ExcelPieChartSerie);
                chart.DataLabel.ShowCategory = true;
                chart.DataLabel.ShowPercent = true;

                chart.Legend.Border.LineStyle = eLineStyle.Solid;
                chart.Legend.Border.Fill.Style = eFillStyle.SolidFill;
                chart.Legend.Border.Fill.Color = Color.DarkBlue;

                //Set the chart style to match the preset style for 3D pie charts.
                chart.StyleManager.SetChartStyle(ePresetChartStyle.Pie3dChartStyle3);
                                                 
                //Switch the PageLayoutView back to normal
                worksheet.View.PageLayoutView = false;
                // save our new workbook and we are done!
                package.Save();
            }

            return newFile.FullName;
        }
    }
}
