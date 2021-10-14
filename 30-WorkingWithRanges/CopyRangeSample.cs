/*************************************************************************************************
  Required Notice: Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  10/13/2021         EPPlus Software AB           Initial release EPPlus 5
 *************************************************************************************************/

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Drawing;
using System.Threading.Tasks;

namespace EPPlusSamples
{
    public static class CopyRangeSample
    {
        public static void Run(string connectionStr)
        {
            using (var p = new ExcelPackage())
            {
                var sourceFile = FileUtil.GetFileInfo("08-Salesreport.xlsx");
                var sourcePackage = new ExcelPackage(sourceFile);                
                var sourceWs = sourcePackage.Workbook.Worksheets[0];
                RangeCopy(p, sourceWs);
                p.SaveAs(FileUtil.GetCleanFileInfo("30-CopyRangeSamples.xlsx"));
            }
        }

        private static void RangeCopy(ExcelPackage p, ExcelWorksheet sourceWs)
        {
            var ws = p.Workbook.Worksheets.Add("CopyFullTable");

            var sourceRange = sourceWs.Cells["A1:G10"]; //Copy the first 10 rows
            sourceRange.Copy(ws.Cells["C1"]);

            sourceRange.Copy(ws.Cells["C15"], ExcelRangeCopyOptionFlags.ExcludeHyperLinks);

            sourceRange.Copy(ws.Cells["C30"], ExcelRangeCopyOptionFlags.ExcludeMergedCells , ExcelRangeCopyOptionFlags.ExcludeStyles , ExcelRangeCopyOptionFlags.ExcludeHyperLinks);
        }
    }
}
