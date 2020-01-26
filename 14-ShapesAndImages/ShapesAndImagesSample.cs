/*************************************************************************************************
  Copyright (C) EPPlus Software AB. 
  This software is licensed under PolyForm Noncommercial License 1.0.0 
  and may only be used for noncommercial purposes 
  https://polyformproject.org/licenses/noncommercial/1.0.0/

  A commercial license to use this software can be purchased at https://epplussoftware.com
 *************************************************************************************************
  Date               Author                       Change
 *************************************************************************************************
  01/24/2020         Jan Källman & Mats Alm       Initial release EPPlus 5
 *************************************************************************************************/
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using OfficeOpenXml;
using System.Drawing;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing;

namespace EPPlusSamples.EncryptionAndProtection
{
    public static class ShapesAndImagesSample
    {
        public static void Run()
        {
            //The output package
            var outputFile = FileOutputUtil.GetFileInfo("14-ShapesAndImages.xlsx");

            //Create the template...
            using (ExcelPackage package = new ExcelPackage(outputFile))
            {
                FillAndColorSamples(package);
                EffectSamples(package);
                ThreeDSamples(package);
                PictureSample(package);
                package.Save();
            }
        }
        private static void PictureSample(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Picture");
            
            //Add an jpg image and apply some effects.
            var pic = ws.Drawings.AddPicture("Landscape", FileInputUtil.GetFileInfo("14-ShapesAndImages", "LandscapeView.jpg"));
            pic.SetPosition(2, 0, 1, 0);
            pic.Effect.SetPresetShadow(ePresetExcelShadowType.OuterBottomRight);
            pic.Effect.OuterShadow.Distance = 10;
            pic.Effect.SetPresetSoftEdges(ePresetExcelSoftEdgesType.SoftEdge5Pt);

            //Add the same image, but with 25 percent of the size. Let the position be absolute.
            pic = ws.Drawings.AddPicture("LandscapeSmall", FileInputUtil.GetFileInfo("14-ShapesAndImages", "LandscapeView.jpg"));
            pic.SetPosition(2, 0, 16, 0);
            pic.SetSize(25);
            pic.ChangeCellAnchor(eEditAs.Absolute);

            //Add the same image again, but let the picure move and resize when rows and colums are resized.
            pic = ws.Drawings.AddPicture("LandscapeRotated", FileInputUtil.GetFileInfo("14-ShapesAndImages", "LandscapeView.jpg"));
            pic.SetPosition(30, 0, 16, 0);
            pic.ChangeCellAnchor(eEditAs.TwoCell);
        }

        private static void FillAndColorSamples(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("Fills And Colors");

            //Drawing with a Solid fill
            var drawing = ws.Drawings.AddShape("SolidFill", eShapeStyle.RoundRect);
            drawing.SetPosition(0, 5, 0, 5);
            drawing.SetSize(250, 250);
            drawing.Fill.Style = eFillStyle.SolidFill;
            drawing.Fill.SolidFill.Color.SetSchemeColor(eSchemeColor.Accent6);
            drawing.Text = "RoundRect With Solid Fill";

            //Drawing with a pattern fill
            drawing = ws.Drawings.AddShape("PatternFill", eShapeStyle.SmileyFace);
            drawing.SetPosition(0, 5, 4, 5);
            drawing.SetSize(250, 250);
            drawing.Fill.Style = eFillStyle.PatternFill;
            drawing.Fill.PatternFill.PatternType = eFillPatternStyle.DiagBrick;
            drawing.Fill.PatternFill.BackgroundColor.SetPresetColor(ePresetColor.Yellow);
            drawing.Fill.PatternFill.ForegroundColor.SetSystemColor(eSystemColor.GrayText);
            drawing.Border.Width = 2;
            drawing.Border.Fill.Style = eFillStyle.SolidFill;
            drawing.Border.Fill.SolidFill.Color.SetHslColor(90, 50, 25);
            drawing.Font.Fill.Color = Color.Black;
            drawing.Font.Bold = true;
            drawing.Text = "Smiley With Pattern Fill";

            //Drawing with a Gradient fill
            drawing = ws.Drawings.AddShape("GradientFill", eShapeStyle.Heart);
            drawing.SetPosition(0, 5, 8, 5);
            drawing.SetSize(250, 250);
            drawing.Fill.Style = eFillStyle.GradientFill;
            drawing.Fill.GradientFill.Colors.AddRgb(0, Color.DarkRed);
            drawing.Fill.GradientFill.Colors.AddRgb(30, Color.Red);
            drawing.Fill.GradientFill.Colors.AddRgbPercentage(65, 100, 0, 0);
            drawing.Fill.GradientFill.Colors[2].Color.Transforms.AddAlpha(75);
            drawing.Text = "Heart with Gradient";

            //Drawing with a blip fill
            drawing = ws.Drawings.AddShape("BlipFill", eShapeStyle.Bevel);
            drawing.SetPosition(0, 5, 12, 5);
            drawing.SetSize(250, 250);
            drawing.Fill.Style = eFillStyle.BlipFill;

            var image = new Bitmap(FileInputUtil.GetFileInfo("14-ShapesAndImages", "EPPlusLogo.jpg").FullName);
            drawing.Fill.BlipFill.Image = image;
            drawing.Fill.BlipFill.Stretch = true;
            drawing.Text = "Blip Fill";
        }
        private static void EffectSamples(ExcelPackage package)
        {
            var ws=package.Workbook.Worksheets.Add("Effects");

            /**** Shadow effects ****/
            var drawing = ws.Drawings.AddShape("OuterShadow", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetShadow(ePresetExcelShadowType.OuterBottomRight);
            drawing.SetPosition(0, 5, 0, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Outer Shadow - Bottom Right";

            drawing = ws.Drawings.AddShape("InnerShadow", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetShadow(ePresetExcelShadowType.InnerTopLeft);
            drawing.SetPosition(0, 5, 4, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Inner Shadow - Top Left";

            drawing = ws.Drawings.AddShape("PerspectiveBelowShadow", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetShadow(ePresetExcelShadowType.PerspectiveBelow);
            drawing.SetPosition(0, 5, 8, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Perspective Shadow - Below";

            /**** Glow effects ****/
            drawing = ws.Drawings.AddShape("Glow Accent1 5pt", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetGlow(ePresetExcelGlowType.Accent1_5Pt);
            drawing.SetPosition(20, 5, 0, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Glow - Accent1 - 5pt";

            drawing = ws.Drawings.AddShape("Glow Accent2 8pt", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetGlow(ePresetExcelGlowType.Accent2_8Pt);
            drawing.SetPosition(20, 5, 4, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Glow - Accent2 - 8pt";

            drawing = ws.Drawings.AddShape("Glow Accent4 11pt", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetGlow(ePresetExcelGlowType.Accent4_11Pt);
            drawing.SetPosition(20, 5, 8, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Glow - Accent4 - 11pt";

            drawing = ws.Drawings.AddShape("Glow Accent5 18pt", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetGlow(ePresetExcelGlowType.Accent5_18Pt);
            drawing.SetPosition(20, 5, 12, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Glow - Accent5 - 18pt";

            /**** Reflection effects ****/
            drawing = ws.Drawings.AddShape("Full Reflection", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetReflection(ePresetExcelReflectionType.Full4Pt);
            drawing.SetPosition(40, 5, 0, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Reflection - Full 4Pt";

            drawing = ws.Drawings.AddShape("Half Reflection", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetReflection(ePresetExcelReflectionType.Half8Pt);
            drawing.SetPosition(40, 5, 4, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Reflection - Half 8Pt";

            drawing = ws.Drawings.AddShape("Tight Touching Reflection", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetReflection(ePresetExcelReflectionType.TightTouching);
            drawing.SetPosition(40, 5, 8, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "Reflection - Tight Touching Reflection";

            drawing = ws.Drawings.AddShape("Soft Edges 10Pt", eShapeStyle.RoundRect);
            drawing.Effect.SetPresetSoftEdges(ePresetExcelSoftEdgesType.SoftEdge10Pt);
            drawing.SetPosition(70, 0, 0, 0);
            drawing.SetSize(250, 250);
            drawing.Text = "Soft Edges - 10Pt";
        }
        private static void ThreeDSamples(ExcelPackage package)
        {
            var ws = package.Workbook.Worksheets.Add("3D");

            //Create a shape with 3D - TopBevel - Round
            var drawing = ws.Drawings.AddShape("3D - TopBevel - Round", eShapeStyle.RoundRect);
            drawing.ThreeD.TopBevel.BevelType = eBevelPresetType.Circle;
            //Default height and width is 6, but we can alter it.
            drawing.ThreeD.TopBevel.Height = 5;
            drawing.ThreeD.TopBevel.Width = 5;
            drawing.SetPosition(0, 5, 0, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "3D - TopBevel - Angle";

            //Create a shape with 3D - TopBevel - ArtDeco and change the 3D camera
            drawing = ws.Drawings.AddShape("3D - TopBevel - ArtDeco", eShapeStyle.RoundRect);
            drawing.ThreeD.TopBevel.BevelType = eBevelPresetType.ArtDeco;
            drawing.ThreeD.Scene.Camera.CameraType = ePresetCameraType.PerspectiveLeft;
            drawing.ThreeD.MaterialType = ePresetMaterialType.TranslucentPowder;
            drawing.SetPosition(0, 5, 5, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "3D - TopBevel - ArtDeco";

            //Create a shape with 3D Camera PerspectiveRelaxedModerately, alter the beveltype and Lightrig 
            drawing = ws.Drawings.AddShape("3D Camera PerspectiveRelaxedModerately", eShapeStyle.RoundRect);
            drawing.ThreeD.MaterialType=ePresetMaterialType.Metal;
            drawing.ThreeD.Scene.Camera.CameraType=ePresetCameraType.PerspectiveRelaxedModerately;
            drawing.ThreeD.TopBevel.BevelType = eBevelPresetType.HardEdge;
            drawing.ThreeD.Scene.LightRig.RigType=eRigPresetType.Sunrise;
            drawing.SetPosition(0, 5, 10, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "3D - Camera - PerspectiveRelaxedModerately - Metal";

            //Create a shape with 3D Camera With Extrusion Color & Contour Color
            drawing = ws.Drawings.AddShape("3D Camera With Extr & Contour", eShapeStyle.RoundRect);
            drawing.ThreeD.MaterialType = ePresetMaterialType.Plastic;
            drawing.ThreeD.Scene.Camera.CameraType = ePresetCameraType.IsometricOffAxis2Right;
            drawing.ThreeD.TopBevel.BevelType = eBevelPresetType.Convex;
            drawing.ThreeD.TopBevel.BevelType = eBevelPresetType.Circle;
            drawing.ThreeD.Scene.LightRig.RigType = eRigPresetType.BrightRoom;

            drawing.ThreeD.ExtrusionColor.SetRgbColor(Color.Red);
            drawing.ThreeD.ExtrusionHeight = 6;

            drawing.ThreeD.ContourColor.SetHslColor(240, 32, 27); //Dark blue
            drawing.ThreeD.ContourWidth = 3;

            drawing.ThreeD.Scene.LightRig.RigType = eRigPresetType.BrightRoom;

            drawing.SetPosition(0, 5, 15, 5);
            drawing.SetSize(250, 250);
            drawing.Text = "3D Camera With Extrusion & Contour";
        }
    }
}
