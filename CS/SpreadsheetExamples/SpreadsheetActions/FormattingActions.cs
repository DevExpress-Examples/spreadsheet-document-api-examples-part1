using DevExpress.Spreadsheet;
using System;
using System.Drawing;

namespace SpreadsheetExamples
{
    public static class FormattingActions
    {
        #region Actions
        public static Action<IWorkbook> CreateModifyApplyStyleAction = CreateModifyApplyStyle;
        public static Action<IWorkbook> FormatCellAction = FormatCell;
        public static Action<IWorkbook> SetDateFormatsAction = SetDateFormats;
        public static Action<IWorkbook> SetNumberFormatsAction = SetNumberFormats;
        public static Action<IWorkbook> CustomNumberFormatAction = CustomNumberFormat;
        public static Action<IWorkbook> ChangeCellColorsAction = ChangeCellColors;
        public static Action<IWorkbook> ChangeCellGradientFillAction = ChangeCellGradientFill;
        public static Action<IWorkbook> SpecifyCellFontAction = SpecifyCellFont;
        public static Action<IWorkbook> AlignCellContentsAction = AlignCellContents;
        public static Action<IWorkbook> AddCellBordersAction = AddCellBorders;
        #endregion

        static void CreateModifyApplyStyle(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];

                #region #CreateNewStyle
                // Add a new style under the "My Style" name to the Styles collection of the workbook.
                Style myStyle = workbook.Styles.Add("My Style");

                // Specify formatting characteristics for the style.
                myStyle.BeginUpdate();
                try
                {
                    // Set the font color to Blue.
                    myStyle.Font.Color = Color.Blue;

                    // Set the font size to 12.
                    myStyle.Font.Size = 12;

                    // Set the horizontal alignment to Center.
                    myStyle.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                    // Set the background.
                    myStyle.Fill.BackgroundColor = Color.LightBlue;
                }
                finally
                {
                    myStyle.EndUpdate();
                }
                #endregion #CreateNewStyle

                #region #DuplicateExistingStyle
                // Add a new style under the "My Good Style" name to the Styles collection.
                Style myGoodStyle = workbook.Styles.Add("My Good Style");

                // Copy all format settings from the built-in Good style.
                myGoodStyle.CopyFrom(BuiltInStyleId.Good);
                #endregion #DuplicateExistingStyle

                #region #ModifyExistingStyle
                // Change the required formatting characteristics of the style.
                myGoodStyle.BeginUpdate();
                try
                {
                    myGoodStyle.Fill.BackgroundColor = Color.LightYellow;
                    // ...
                }
                finally
                {
                    myGoodStyle.EndUpdate();
                }
                #endregion #ModifyExistingStyle

                #region #ApplyStyles
                // Access the built-in "Good" MS Excel style from the Styles collection of the workbook.
                Style styleGood = workbook.Styles[BuiltInStyleId.Good];

                // Apply the "Good" style to a range of cells.
                worksheet.Range["A1:C4"].Style = styleGood;

                // Access a custom style that has been previously created in the loaded document by its name.
                Style customStyle = workbook.Styles["Custom Style"];

                // Apply the custom style to the cell.
                worksheet.Cells["D6"].Style = customStyle;

                // Apply the "Good" style to the eighth row.
                worksheet.Rows[7].Style = styleGood;

                // Apply the custom style to the "H" column.
                worksheet.Columns["H"].Style = customStyle;
                #endregion #ApplyStyles
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void FormatCell(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {


                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Cells["B2"].Value = "Test";
                worksheet.Range["C3:E6"].Value = "Test";

                #region #CellFormatting
                // Access the cell to be formatted.
                Cell cell = worksheet.Cells["B2"];

                // Specify font settings (font name, color, size and style).
                cell.Font.Name = "MV Boli";
                cell.Font.Color = Color.Blue;
                cell.Font.Size = 14;
                cell.Font.FontStyle = SpreadsheetFontStyle.Bold;

                // Specify cell background color.
                cell.Fill.BackgroundColor = Color.LightSkyBlue;

                // Specify text alignment in the cell. 
                cell.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                #endregion #CellFormatting

                #region #RangeFormatting
                // Access the range of cells to be formatted.
                CellRange range = worksheet.Range["C3:E6"];

                // Begin updating of the range formatting. 
                Formatting rangeFormatting = range.BeginUpdateFormatting();

                // Specify font settings (font name, color, size and style).
                rangeFormatting.Font.Name = "MV Boli";
                rangeFormatting.Font.Color = Color.Blue;
                rangeFormatting.Font.Size = 14;
                rangeFormatting.Font.FontStyle = SpreadsheetFontStyle.Bold;

                // Specify cell background color.
                rangeFormatting.Fill.BackgroundColor = Color.LightSkyBlue;

                // Specify text alignment in cells.
                rangeFormatting.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                rangeFormatting.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                // End updating of the range formatting.
                range.EndUpdateFormatting(rangeFormatting);
                #endregion #RangeFormatting
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void SetDateFormats(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Range["A1:F1"].ColumnWidthInCharacters = 15;
                worksheet.Range["A1:F1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                #region #DateTimeFormats
                worksheet.Range["A1:F1"].Formula = "= Now()";

                // Apply different date display formats.
                worksheet.Cells["A1"].NumberFormat = "m/d/yy";

                worksheet.Cells["B1"].NumberFormat = "d-mmm-yy";

                worksheet.Cells["C1"].NumberFormat = "dddd";

                // Apply different time display formats.
                worksheet.Cells["D1"].NumberFormat = "m/d/yy h:mm";

                worksheet.Cells["E1"].NumberFormat = "h:mm AM/PM";

                worksheet.Cells["F1"].NumberFormat = "h:mm:ss";
                #endregion #DateTimeFormats
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void SetNumberFormats(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {

                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Range["A1:H1"].ColumnWidthInCharacters = 12;
                worksheet.Range["A1:H1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                #region #NumberFormats
                // Display 111 as 111.
                worksheet.Cells["A1"].Value = 111;
                worksheet.Cells["A1"].NumberFormat = "#####";

                // Display 222 as 00222.
                worksheet.Cells["B1"].Value = 222;
                worksheet.Cells["B1"].NumberFormat = "00000";

                // Display 12345678 as 12,345,678.
                worksheet.Cells["C1"].Value = 12345678;
                worksheet.Cells["C1"].NumberFormat = "#,#";

                // Display .126 as 0.13.
                worksheet.Cells["D1"].Value = .126;
                worksheet.Cells["D1"].NumberFormat = "0.##";

                // Display 74.4 as 74.400.
                worksheet.Cells["E1"].Value = 74.4;
                worksheet.Cells["E1"].NumberFormat = "##.000";

                // Display 1.6 as 160.0%.
                worksheet.Cells["F1"].Value = 1.6;
                worksheet.Cells["F1"].NumberFormat = "0.0%";

                // Display 4321 as $4,321.00.
                worksheet.Cells["G1"].Value = 4321;
                worksheet.Cells["G1"].NumberFormat = "$#,##0.00";

                // Display 8.75 as 8 3/4.
                worksheet.Cells["H1"].Value = 8.75;
                worksheet.Cells["H1"].NumberFormat = "# ?/?";
                #endregion #NumberFormats
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void CustomNumberFormat(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {

                Worksheet worksheet = workbook.Worksheets[0];

                worksheet["A1"].Value = "Number Format:";
                worksheet["A2"].Value = "Values:";
                worksheet["A3"].Value = "Formatted Values:";
                worksheet["A1:A3"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet["A1:E3"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left;
                worksheet["A1:E3"].ColumnWidthInCharacters = 17;

                worksheet.MergeCells(worksheet["B1:E1"]);
                worksheet["B1:E1"].Value = "[Green]#.00;[Red]#.00;[Blue]0.00;[Magenta]\"product: \"@";
                worksheet["B1:E1"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                worksheet["B1:E3"].Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin);

                #region #CustomNumberFormat
                // Set cell values.
                worksheet["B2:B3"].Value = -15.50;
                worksheet["C2:C3"].Value = 555;
                worksheet["D2:D3"].Value = 0;
                worksheet["E2:E3"].Value = "Name";

                //Apply custom number format.
                worksheet["B3:E3"].NumberFormat = "[Green]#.00;[Red]#.00;[Blue]0.00;[Magenta]\"product: \"@";
                #endregion #CustomNumberFormat
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void ChangeCellColors(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Range["C3:H10"].Merge();
                worksheet.Range["C3:H10"].Value = "Test";
                worksheet.Range["C3:H10"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Range["C3:H10"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
                worksheet.Cells["A1"].Value = "Test";

                #region #ColorCells
                // Format an individual cell.
                worksheet.Cells["A1"].Font.Color = Color.Red;
                worksheet.Cells["A1"].FillColor = Color.Yellow;

                // Format a range of cells.
                CellRange range = worksheet.Range["C3:H10"];
                Formatting rangeFormatting = range.BeginUpdateFormatting();
                rangeFormatting.Font.Color = Color.Blue;
                rangeFormatting.Fill.BackgroundColor = Color.LightBlue;
                rangeFormatting.Fill.PatternType = PatternType.LightHorizontal;
                rangeFormatting.Fill.PatternColor = Color.Violet;
                range.EndUpdateFormatting(rangeFormatting);
                #endregion #ColorCells
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void ChangeCellGradientFill(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Range["C3:F8"].Merge();
                worksheet.Range["C3:F8"].Value = "Test";
                worksheet.Range["C3:F8"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Range["C3:F8"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

                worksheet.Range["C13:F18"].Merge();
                worksheet.Range["C13:F18"].Value = "Test";
                worksheet.Range["C13:F18"].Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                worksheet.Range["C13:F18"].Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

                #region #GradientLinear
                // Specify a linear gradient fill for a cell.
                Fill fillA1 = worksheet.Cells["A1"].Fill;
                fillA1.FillType = FillType.Gradient;
                fillA1.Gradient.Type = GradientFillType.Linear;
                // Set the tilt for the gradient line in degrees.
                // 90 degree angle defines a color transition from the top to the bottom.
                fillA1.Gradient.Degree = 90;
                // Specify two gradient colors. 
                // The position of a color stop should be either 0 (start) or 1 (end).
                fillA1.Gradient.Stops.Add(0, Color.Yellow);
                fillA1.Gradient.Stops.Add(1, Color.SkyBlue);

                // Specify a linear gradient fill for a range.
                worksheet.Range["C3:F8"].Fill.FillType = FillType.Gradient;
                GradientFill rangeGradient1 = worksheet.Range["C3:F8"].Fill.Gradient;
                rangeGradient1.Type = GradientFillType.Linear;
                // Set the tilt for the gradient line in degrees.
                // 45 degree angle defines a color transition from top left to bottom right.
                rangeGradient1.Degree = 45;
                // Specify two gradient colors. 
                // The position of a color stop should be either 0 (start) or 1 (end).
                GradientStopCollection rangeStops1 = rangeGradient1.Stops;
                rangeStops1.Add(0, Color.BlanchedAlmond);
                rangeStops1.Add(1, Color.Blue);
                #endregion #GradientLinear

                #region #GradientPath
                // Specify a path gradient fill for a cell.
                Fill fillB1 = worksheet.Cells["A3"].Fill;
                fillB1.FillType = FillType.Gradient;
                GradientFill cellGradient = fillB1.Gradient;
                cellGradient.Type = GradientFillType.Path;
                // Set the relative position of a convergence point.
                cellGradient.RectangleLeft = 0f;
                cellGradient.RectangleRight = 0f;
                cellGradient.RectangleTop = 0.5f;
                cellGradient.RectangleBottom = 0.5f;
                // Specify two gradient colors. 
                // The position of a color stop should be either 0 (start) or 1 (end).
                GradientStopCollection cellStops2 = cellGradient.Stops;
                cellStops2.Add(0, Color.Yellow);
                cellStops2.Add(1, Color.Red);

                // Specify a path gradient fill for a range.
                worksheet.Range["C13:F18"].Fill.FillType = FillType.Gradient;
                GradientFill rangeGradient2 = worksheet.Range["C13:f18"].Fill.Gradient;
                rangeGradient2.Type = GradientFillType.Path;
                // Set the relative position of a convergence point.
                rangeGradient2.RectangleLeft = 0.5f;
                rangeGradient2.RectangleRight = 0.5f;
                rangeGradient2.RectangleTop = 0.5f;
                rangeGradient2.RectangleBottom = 0.5f;
                // Specify two gradient colors. 
                // The position of a color stop should be either 0 (start) or 1 (end).
                GradientStopCollection rangeStops2 = rangeGradient2.Stops;
                rangeStops2.Add(0, Color.Orange);
                rangeStops2.Add(1, Color.Green);
                #endregion #GradientPath
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void SpecifyCellFont(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];

                worksheet.Cells["A1"].Value = "Font Attributes";
                worksheet.Cells["A1"].ColumnWidthInCharacters = 20;

                #region #FontSettings
                // Access the Font object.
                SpreadsheetFont cellFont = worksheet.Cells["A1"].Font;
                // Set the font name.
                cellFont.Name = "Times New Roman";
                // Set the font size.
                cellFont.Size = 14;
                // Set the font color.
                cellFont.Color = Color.Blue;
                // Format text as bold.
                cellFont.Bold = true;
                // Set font to be underlined.
                cellFont.UnderlineType = UnderlineType.Single;
                #endregion #FontSettings
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void AlignCellContents(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                Worksheet worksheet = workbook.Worksheets[0];

                CellRange range = worksheet.Range["A1:A4"];
                range.ColumnWidthInCharacters = 30;
                workbook.Unit = DevExpress.Office.DocumentUnit.Inch;
                range.RowHeight = 1;

                #region #AlignCellContents
                Cell cellA1 = worksheet.Cells["A1"];
                cellA1.Value = "Right and top";
                cellA1.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right;
                cellA1.Alignment.Vertical = SpreadsheetVerticalAlignment.Top;

                Cell cellA2 = worksheet.Cells["A2"];
                cellA2.Value = "Center";
                cellA2.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
                cellA2.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;

                Cell cellA3 = worksheet.Cells["A3"];
                cellA3.Value = "Left and bottom, indent";
                cellA3.Alignment.Indent = 2;

                Cell cellA4 = worksheet.Cells["A4"];
                cellA4.Value = "The Alignment.WrapText property is applied to wrap the text within a cell";
                cellA4.Alignment.WrapText = true;

                Cell cellA5 = worksheet.Cells["A5"];
                cellA5.Value = "Rotation by 45 degrees";
                cellA5.Alignment.RotationAngle = 45;
                #endregion #AlignCellContents
            }
            finally
            {
                workbook.EndUpdate();
            }
        }

        static void AddCellBorders(IWorkbook workbook)
        {
            workbook.BeginUpdate();
            try
            {
                #region #CellBorders
                Worksheet worksheet = workbook.Worksheets[0];

                // Set each particular border for the cell.
                Cell cellB2 = worksheet.Cells["B2"];
                Borders cellB2Borders = cellB2.Borders;
                cellB2Borders.LeftBorder.LineStyle = BorderLineStyle.MediumDashDot;
                cellB2Borders.LeftBorder.Color = Color.Pink;
                cellB2Borders.TopBorder.LineStyle = BorderLineStyle.MediumDashDotDot;
                cellB2Borders.TopBorder.Color = Color.HotPink;
                cellB2Borders.RightBorder.LineStyle = BorderLineStyle.MediumDashed;
                cellB2Borders.RightBorder.Color = Color.DeepPink;
                cellB2Borders.BottomBorder.LineStyle = BorderLineStyle.Medium;
                cellB2Borders.BottomBorder.Color = Color.Red;

                // Set all borders for cells.
                Cell cellC4 = worksheet.Cells["C4"];
                cellC4.Borders.SetAllBorders(Color.Orange, BorderLineStyle.MediumDashed);
                Cell cellD6 = worksheet.Cells["D6"];
                cellD6.Borders.SetOutsideBorders(Color.Gold, BorderLineStyle.Double);
                #endregion #CellBorders

                #region #CellRangeBorders
                // Set all borders for the range of cells in one step.
                CellRange range1 = worksheet.Range["B8:F13"];
                range1.Borders.SetAllBorders(Color.Green, BorderLineStyle.Double);

                // Set all inside and outside borders separately for the range of cells.
                CellRange range2 = worksheet.Range["C15:F18"];
                range2.SetInsideBorders(Color.SkyBlue, BorderLineStyle.MediumDashed);
                range2.Borders.SetOutsideBorders(Color.DeepSkyBlue, BorderLineStyle.Medium);

                // Set all horizontal and vertical borders separately for the range of cells.
                CellRange range3 = worksheet.Range["D21:F23"];
                Formatting range3Formatting = range3.BeginUpdateFormatting();
                Borders range3Borders = range3Formatting.Borders;
                range3Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.MediumDashDot;
                range3Borders.InsideHorizontalBorders.Color = Color.DarkBlue;
                range3Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.MediumDashDotDot;
                range3Borders.InsideVerticalBorders.Color = Color.Blue;
                range3.EndUpdateFormatting(range3Formatting);

                // Set each particular border for the range of cell. 
                CellRange range4 = worksheet.Range["E25:F26"];
                Formatting range4Formatting = range4.BeginUpdateFormatting();
                Borders range4Borders = range4Formatting.Borders;
                range4Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thick);
                range4Borders.LeftBorder.Color = Color.Violet;
                range4Borders.TopBorder.Color = Color.Violet;
                range4Borders.RightBorder.Color = Color.DarkViolet;
                range4Borders.BottomBorder.Color = Color.DarkViolet;
                range4.EndUpdateFormatting(range4Formatting);
                #endregion #CellRangeBorders
            }
            finally
            {
                workbook.EndUpdate();
            }
        }
    }
}
