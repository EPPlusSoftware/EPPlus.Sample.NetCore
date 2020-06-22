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
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Drawing;

namespace EPPlusSamples
{
    class WorkingWithVbaSample
    {
        public static void Run()
        {
            //Create a macro-enabled workbook from scratch.
            SimpleVba();
            
            //Open Sample 1 and add code to change the chart to a bubble chart.
            AddABubbleChart();

            //Simple battleships game from scratch.
            CreateABattleShipsGame();
        }
        private static void SimpleVba()
        {
            ExcelPackage pck = new ExcelPackage();

            //Add a worksheet.
            var ws=pck.Workbook.Worksheets.Add("VBA Sample");
            ws.Drawings.AddShape("VBASampleRect", eShapeStyle.RoundRect);
            
            //Create a vba project             
            pck.Workbook.CreateVBAProject();

            //Now add some code to update the text of the shape...
            var sb = new StringBuilder();

            sb.AppendLine("Private Sub Workbook_Open()");
            sb.AppendLine("    [VBA Sample].Shapes(\"VBASampleRect\").TextEffect.Text = \"This text is set from VBA!\"");
            sb.AppendLine("End Sub");
            pck.Workbook.CodeModule.Code = sb.ToString();

            //And Save as xlsm
            FileInfo fi = FileOutputUtil.GetFileInfo("21.1-SimpleVba.xlsm");
            pck.SaveAs(fi);
        }
        private static void AddABubbleChart()
        {
            FileInfo sample1File = FileOutputUtil.GetFileInfo("01-GettingStarted.xlsx", false);
            //Open Sample 1 again
            ExcelPackage pck = new ExcelPackage(sample1File);
            var p = new ExcelPackage();
            //Create a vba project             
            pck.Workbook.CreateVBAProject();

            //Now add some code that creates a bubble chart...
            var sb = new StringBuilder();

            sb.AppendLine("Public Sub CreateBubbleChart()");
            sb.AppendLine("Dim co As ChartObject");
            sb.AppendLine("Set co = Inventory.ChartObjects.Add(10, 100, 400, 200)");
            sb.AppendLine("co.Chart.SetSourceData Source:=Range(\"'Inventory'!$B$1:$E$5\")");
            sb.AppendLine("co.Chart.ChartType = xlBubble3DEffect         'Add a bubblechart");
            sb.AppendLine("End Sub");

            //Create a new module and set the code
            var module = pck.Workbook.VbaProject.Modules.AddModule("EPPlusGeneratedCode");
            module.Code = sb.ToString();

            //Call the newly created sub from the workbook open event
            pck.Workbook.CodeModule.Code = "Private Sub Workbook_Open()\r\nCreateBubbleChart\r\nEnd Sub";

            //Optionally, Sign the code with your company certificate.
            //X509Store store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);
            //pck.Workbook.VbaProject.Signature.Certificate = store.Certificates[0];

            //And Save as xlsm
            FileInfo fi =FileOutputUtil.GetFileInfo("21.2-AddABubbleChartVba.xlsm");
            pck.SaveAs(fi);
        }
        private static void CreateABattleShipsGame()
        {
            //Now, lets do something a little bit more fun.
            //We are going to create a simple battleships game from scratch.

            ExcelPackage pck = new ExcelPackage();

            //Add a worksheet.
            var ws = pck.Workbook.Worksheets.Add("Battleship");

            ws.View.ShowGridLines = false;
            ws.View.ShowHeaders = false;

            ws.DefaultColWidth = 3;
            ws.DefaultRowHeight = 15;

            int gridSize=10;

            //Create the boards
            var board1 = ws.Cells[2, 2, 2 + gridSize - 1, 2 + gridSize - 1];
            var board2 = ws.Cells[2, 4+gridSize-1, 2 + gridSize-1, 4 + (gridSize-1)*2];
            CreateBoard(board1);
            CreateBoard(board2);
            ws.Select("B2");
            ws.Protection.IsProtected = true;
            ws.Protection.AllowSelectLockedCells = true;

            //Create the VBA Project
            pck.Workbook.CreateVBAProject();
            //Password protect your code
            pck.Workbook.VbaProject.Protection.SetPassword("EPPlus");

             var codeDir = FileInputUtil.GetSubDirectory("21-VBA", "VBA-Code");

            //Add all the code from the textfiles in the Vba-Code sub-folder.
            pck.Workbook.CodeModule.Code = GetCodeModule(codeDir, "ThisWorkbook.txt");

            //Add the sheet code
            ws.CodeModule.Code = GetCodeModule(codeDir, "BattleshipSheet.txt");
            var m1=pck.Workbook.VbaProject.Modules.AddModule("Code");
            string code = GetCodeModule(codeDir, "CodeModule.txt");

            //Insert your ships on the right board. you can changes these, but don't cheat ;)
            var ships = new string[]{
                "N3:N7",
                "P2:S2",
                "V9:V11",
                "O10:Q10",
                "R11:S11"};
            
            //Note: For security reasons you should never mix external data and code(to avoid code injections!), especially not on a webserver. 
            //If you deside to do that anyway, be very careful with the validation of the data.
            //Be extra careful if you sign the code.
            //Read more here http://en.wikipedia.org/wiki/Code_injection

            code = string.Format(code, ships[0],ships[1],ships[2],ships[3],ships[4], board1.Address, board2.Address);  //Ships are injected into the constants in the module
            m1.Code = code;

            //Ships are displayed with a black background
            string shipsaddress = string.Join(",", ships);
            ws.Cells[shipsaddress].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
            ws.Cells[shipsaddress].Style.Fill.BackgroundColor.SetColor(Color.Black);

            var m2 = pck.Workbook.VbaProject.Modules.AddModule("ComputerPlay");
            m2.Code = GetCodeModule(codeDir, "ComputerPlayModule.txt"); 
            var c1 = pck.Workbook.VbaProject.Modules.AddClass("Ship",false);
            c1.Code = GetCodeModule(codeDir, "ShipClass.txt"); 

            //Add the info text shape.
            var tb = ws.Drawings.AddShape("txtInfo", eShapeStyle.Rect);
            tb.SetPosition(1, 0, 27, 0);
            tb.Fill.Color = Color.LightSlateGray;
            var rt1 = tb.RichText.Add("Battleships");
            rt1.Bold = true;
            tb.RichText.Add("\r\nDouble-click on the left board to make your move. Find and sink all ships to win!");

            //Set the headers.
            ws.SetValue("B1", "Computer Grid");
            ws.SetValue("M1", "Your Grid");
            ws.Row(1).Style.Font.Size = 18;

            AddChart(ws.Cells["B13"], "chtHitPercent", "Player");
            AddChart(ws.Cells["M13"], "chtComputerHitPercent", "Computer");

            ws.Names.Add("LogStart", ws.Cells["B24"]);
            ws.Cells["B24:X224"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            ws.Cells["B25:X224"].Style.Font.Name = "Consolas";
            ws.SetValue("B24", "Log");
            ws.Cells["B24"].Style.Font.Bold = true;
            ws.Cells["B24:X24"].Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black);
            var cf=ws.Cells["B25:B224"].ConditionalFormatting.AddContainsText();
            cf.Text = "hit";
            cf.Style.Font.Color.Color = Color.Red;

            //If you have a valid certificate for code signing you can use this code to set it.
            ///*** Try to find a cert valid for signing... ***/
            //X509Store store = new X509Store(StoreLocation.CurrentUser);
            //store.Open(OpenFlags.ReadOnly);   
            //foreach (var cert in store.Certificates)
            //{
            //    if (cert.HasPrivateKey && cert.NotBefore <= DateTime.Today && cert.NotAfter >= DateTime.Today)
            //    {
            //        pck.Workbook.VbaProject.Signature.Certificate = cert;
            //        break;
            //    }
            //}

            var fi = FileOutputUtil.GetFileInfo(@"21.3-CreateABattleShipsGameVba.xlsm");
            pck.SaveAs(fi);
        }

        private static string GetCodeModule(DirectoryInfo codeDir, string fileName)
        {
            return File.ReadAllText(FileOutputUtil.GetFileInfo(codeDir, fileName, false).FullName);
        }

        private static void AddChart(ExcelRange rng,string name, string prefix)
        {
            var chrt = rng.Worksheet.Drawings.AddPieChart(name, ePieChartType.Pie);
            chrt.SetPosition(rng.Start.Row-1, 0, rng.Start.Column-1, 0);
            chrt.To.Row = rng.Start.Row+9;
            chrt.To.Column = rng.Start.Column + 9;
            chrt.Style = eChartStyle.Style18;
            chrt.DataLabel.ShowPercent = true;

            var serie = chrt.Series.Add(rng.Offset(2, 2, 1, 2), rng.Offset(1, 2, 1, 2));
            serie.Header = "Hits";
            
            chrt.Title.Text = "Hit ratio";
            
            var n1 = rng.Worksheet.Names.Add(prefix + "Misses", rng.Offset(2, 2));
            n1.Value = 0;
            var n2 = rng.Worksheet.Names.Add(prefix + "Hits", rng.Offset(2, 3));
            n2.Value = 0;
            rng.Offset(1, 2).Value = "Misses";
            rng.Offset(1, 3).Value = "Hits";            
        }

        private static void CreateBoard(ExcelRange rng)
        {
            //Create a gradiant background with one dark and one light blue color
            rng.Style.Fill.Gradient.Color1.SetColor(Color.FromArgb(0x80, 0x80, 0XFF));
            rng.Style.Fill.Gradient.Color2.SetColor(Color.FromArgb(0x20, 0x20, 0XFF));
            rng.Style.Fill.Gradient.Type = ExcelFillGradientType.None;
            for (int col = 0; col <= rng.End.Column - rng.Start.Column; col++)
            {
                for (int row = 0; row <= rng.End.Row - rng.Start.Row; row++)
                {
                    if (col % 4 == 0)
                    {
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 45;
                    }
                    if (col % 4 == 1)
                    {
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 70;
                    }
                    if (col % 4 == 2)
                    {
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 110;
                    }
                    else
                    {
                        rng.Offset(row, col, 1, 1).Style.Fill.Gradient.Degree = 135;
                    }
                }
            }
            //Set the inner cell border to thin, light gray
            rng.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            rng.Style.Border.Top.Color.SetColor(Color.Gray);
            rng.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            rng.Style.Border.Right.Color.SetColor(Color.Gray);
            rng.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            rng.Style.Border.Left.Color.SetColor(Color.Gray);
            rng.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            rng.Style.Border.Bottom.Color.SetColor(Color.Gray);

            //Solid black border around the board.
            rng.Style.Border.BorderAround(ExcelBorderStyle.Medium, Color.Black);
        }
    }
}
