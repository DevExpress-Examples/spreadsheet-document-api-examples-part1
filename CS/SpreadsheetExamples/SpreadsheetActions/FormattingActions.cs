using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DevExpress.Spreadsheet;
using System.Drawing;

namespace SpreadsheetExamples {
    public static class FormattingActions {
        #region Actions
        public static Action<Workbook> ApplayStyleAction = ApplayStyle;
        public static Action<Workbook> CreateModifyStyleAction = CreateModifyStyle;
        public static Action<Workbook> FormatCellAction = FormatCell;
        public static Action<Workbook> SetDateFormatsAction = SetDateFormats;
        public static Action<Workbook> SetNumberFormatsAction = SetNumberFormats;
        public static Action<Workbook> ChangeCellColorsAction = ChangeCellColors;
        public static Action<Workbook> SpecifyCellFontAction = SpecifyCellFont;
        public static Action<Workbook> AlignCellContentsAction = AlignCellContents;
        public static Action<Workbook> AddCellBordersAction = AddCellBorders;
        #endregion

        static void ApplayStyle(Workbook workbook) {
            #region #ApplyCellStyle
            Worksheet worksheet = workbook.Worksheets[0];

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
            #endregion #ApplyCellStyle
        }

        static void CreateModifyStyle(Workbook workbook) {
            #region #CreateNewStyle
            // Add a new style under the "My Style" name to the Styles collection of the workbook.
            Style myStyle = workbook.Styles.Add("My Style");

            // Specify formatting characteristics for the style.
            myStyle.BeginUpdate();
            try {
                // Set the font color to Blue.
                myStyle.Font.Color = Color.Blue;

                // Set the font size to 12.
                myStyle.Font.Size = 12;

                // Set the horizontal alignment to Center.
                myStyle.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;

                // Set the background.
                myStyle.Fill.BackgroundColor = Color.LightBlue;
                myStyle.Fill.PatternType = PatternType.LightGray;
                myStyle.Fill.PatternColor = Color.Yellow;
            } 
            finally {
                myStyle.EndUpdate();
            }
            #endregion #CreateNewStyle

            #region #DuplicateExistingStyle
            // Add a new style under the "My Good Style" name to the Styles collection.
            Style myGoodStyle = workbook.Styles.Add("My Good Style");

            // Copy all format settings from the built-in Good style.
            myGoodStyle.CopyFrom(BuiltInStyleId.Good);

            // Modify the required formatting characteristics if needed.
            // ...
            #endregion #DuplicateExistingStyle

            #region #ModifyExistingStyle
            // Access the style to be modified.
            Style customStyle = workbook.Styles["Custom Style"];

            // Change the required formatting characteristics of the style.
            customStyle.BeginUpdate();
            try {
                customStyle.Fill.BackgroundColor = Color.Gold;
                // ...
            } finally {
                customStyle.EndUpdate();
            }
            #endregion #ModifyExistingStyle
        }

        static void FormatCell(Workbook workbook) {

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
            Range range = worksheet.Range["C3:E6"];

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


        static void SetDateFormats(Workbook workbook) {

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

        static void SetNumberFormats(Workbook workbook) {

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

        static void ChangeCellColors(Workbook workbook) {

            Worksheet worksheet = workbook.Worksheets[0];

            worksheet.Range["C3:D4"].Merge();
            worksheet.Range["C3:D4"].Value = "Test";
            worksheet.Cells["A1"].Value = "Test";

            #region #ColorCells
            // Format an individual cell.
            worksheet.Cells["A1"].Font.Color = Color.Red;
            worksheet.Cells["A1"].FillColor = Color.Yellow;

            // Format a range of cells.
            Range range = worksheet.Range["C3:D4"];
            Formatting rangeFormatting = range.BeginUpdateFormatting();
            rangeFormatting.Font.Color = Color.Blue;
            rangeFormatting.Fill.BackgroundColor = Color.LightBlue;
            rangeFormatting.Fill.PatternType = PatternType.LightHorizontal;
            rangeFormatting.Fill.PatternColor = Color.Violet;
            range.EndUpdateFormatting(rangeFormatting);
            #endregion #ColorCells
        }

        static void SpecifyCellFont(Workbook workbook) {

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
            cellFont.UnderlineType = UnderlineType.Double;
            #endregion #FontSettings
        }

        static void AlignCellContents(Workbook workbook) {

            Worksheet worksheet = workbook.Worksheets[0];

            Range range = worksheet.Range["A1:B3"];
            range.ColumnWidthInCharacters = 30;
            range.RowHeight = 200;

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
            cellA3.Alignment.Indent = 1;

            Cell cellB1 = worksheet.Cells["B1"];
            cellB1.Value = "The Alignment.ShrinkToFit property is applied";
            cellB1.Alignment.ShrinkToFit = true;

            Cell cellB2 = worksheet.Cells["B2"];
            cellB2.Value = "Rotated Cell Contents";
            cellB2.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center;
            cellB2.Alignment.Vertical = SpreadsheetVerticalAlignment.Center;
            cellB2.Alignment.RotationAngle = 15;

            Cell cellB3 = worksheet.Cells["B3"];
            cellB3.Value = "The Alignment.WrapText property is applied to wrap the text within a cell";
            cellB3.Alignment.WrapText = true;
            #endregion #AlignCellContents
        }

        static void AddCellBorders(Workbook workbook) {
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
            cellB2Borders.DiagonalBorderType = DiagonalBorderType.Up;
            cellB2Borders.DiagonalBorderLineStyle = BorderLineStyle.Thick;
            cellB2Borders.DiagonalBorderColor = Color.Red;

            // Set diagonal borders for the cell.
            Cell cellC4 = worksheet.Cells["C4"];
            Borders cellC4Borders = cellC4.Borders;
            cellC4Borders.SetDiagonalBorders(Color.Orange, BorderLineStyle.Double, DiagonalBorderType.UpAndDown);

            // Set all outside borders for the cell in one step. 
            Cell cellD6 = worksheet.Cells["D6"];
            cellD6.Borders.SetOutsideBorders(Color.Gold, BorderLineStyle.Double);
            #endregion #CellBorders

            #region #CellRangeBorders
            // Set all borders for the range of cells in one step.
            Range range1 = worksheet.Range["B8:F13"];
            range1.Borders.SetAllBorders(Color.Green, BorderLineStyle.Double);

            // Set all inside and outside borders separately for the range of cells.
            Range range2 = worksheet.Range["C15:F18"];
            range2.SetInsideBorders(Color.SkyBlue, BorderLineStyle.MediumDashed);
            range2.Borders.SetOutsideBorders(Color.DeepSkyBlue, BorderLineStyle.Medium);

            // Set all horizontal and vertical borders separately for the range of cells.
            Range range3 = worksheet.Range["D21:F23"];
            Formatting range3Formatting = range3.BeginUpdateFormatting();
            Borders range3Borders = range3Formatting.Borders;
            range3Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.MediumDashDot;
            range3Borders.InsideHorizontalBorders.Color = Color.DarkBlue;
            range3Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.MediumDashDotDot;
            range3Borders.InsideVerticalBorders.Color = Color.Blue;
            range3.EndUpdateFormatting(range3Formatting);

            // Set each particular border for the range of cell. 
            Range range4 = worksheet.Range["E25:F26"];
            Formatting range4Formatting = range4.BeginUpdateFormatting();
            Borders range4Borders = range4Formatting.Borders;
            range4Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thick);
            range4Borders.LeftBorder.Color = Color.Violet;
            range4Borders.TopBorder.Color = Color.Violet;
            range4Borders.RightBorder.Color = Color.DarkViolet;
            range4Borders.BottomBorder.Color = Color.DarkViolet;
            range4Borders.DiagonalBorderType = DiagonalBorderType.UpAndDown;
            range4Borders.DiagonalBorderLineStyle = BorderLineStyle.MediumDashed;
            range4Borders.DiagonalBorderColor = Color.BlueViolet;
            range4.EndUpdateFormatting(range4Formatting);
            #endregion #CellRangeBorders
        }
    }
}
