Imports System
Imports DevExpress.Spreadsheet
Imports System.Drawing

Namespace SpreadsheetExamples

    Public Module FormattingActions

#Region "Actions"
        Public ApplyStyleAction As Action(Of Workbook) = AddressOf ApplyStyle

        Public CreateModifyStyleAction As Action(Of Workbook) = AddressOf CreateModifyStyle

        Public FormatCellAction As Action(Of Workbook) = AddressOf FormatCell

        Public SetDateFormatsAction As Action(Of Workbook) = AddressOf SetDateFormats

        Public SetNumberFormatsAction As Action(Of Workbook) = AddressOf SetNumberFormats

        Public ChangeCellColorsAction As Action(Of Workbook) = AddressOf ChangeCellColors

        Public SpecifyCellFontAction As Action(Of Workbook) = AddressOf SpecifyCellFont

        Public AlignCellContentsAction As Action(Of Workbook) = AddressOf AlignCellContents

        Public AddCellBordersAction As Action(Of Workbook) = AddressOf AddCellBorders

#End Region
        Private Sub ApplyStyle(ByVal workbook As Workbook)
#Region "#ApplyCellStyle"
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Access the built-in "Good" MS Excel style from the Styles collection of the workbook.
            Dim styleGood As Style = workbook.Styles(BuiltInStyleId.Good)
            ' Apply the "Good" style to a cell range.
            worksheet.Range("A1:C4").Style = styleGood
            ' Access the previously created custom style by its name.
            Dim customStyle As Style = workbook.Styles("Custom Style")
            ' Apply the custom style to the cell.
            worksheet.Cells("D6").Style = customStyle
            ' Apply the "Good" style to the eighth row.
            worksheet.Rows(7).Style = styleGood
            ' Apply the custom style to the "H" column.
            worksheet.Columns("H").Style = customStyle
#End Region  ' #ApplyCellStyle
        End Sub

        Private Sub CreateModifyStyle(ByVal workbook As Workbook)
#Region "#CreateNewStyle"
            ' Add a new style under the "My Style" name to the workbook's style collection.
            Dim myStyle As Style = workbook.Styles.Add("My Style")
            ' Specify formatting characteristics for the style.
            myStyle.BeginUpdate()
            Try
                ' Specify the font color.
                myStyle.Font.Color = Color.Blue
                ' Specify the font size.
                myStyle.Font.Size = 12
                ' Specify the horizontal alignment.
                myStyle.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                ' Set the background.
                myStyle.Fill.BackgroundColor = Color.LightBlue
                myStyle.Fill.PatternType = PatternType.LightGray
                myStyle.Fill.PatternColor = Color.Yellow
            Finally
                myStyle.EndUpdate()
            End Try

#End Region  ' #CreateNewStyle
#Region "#DuplicateExistingStyle"
            ' Add a new style under the "My Good Style" name to the workbook's style collection.
            Dim myGoodStyle As Style = workbook.Styles.Add("My Good Style")
            ' Copy all format settings from the built-in Good style.
            myGoodStyle.CopyFrom(BuiltInStyleId.Good)
            ' Modify the required formatting characteristics if needed.
            ' ...
#End Region  ' #DuplicateExistingStyle
#Region "#ModifyExistingStyle"
            ' Access the style to be modified.
            Dim customStyle As Style = workbook.Styles("Custom Style")
            ' Change the required formatting characteristics of the style.
            customStyle.BeginUpdate()
            Try
                ' ...
                customStyle.Fill.BackgroundColor = Color.Gold
            Finally
                customStyle.EndUpdate()
            End Try
#End Region  ' #ModifyExistingStyle
        End Sub

        Private Sub FormatCell(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            worksheet.Cells("B2").Value = "Test"
            worksheet.Range("C3:E6").Value = "Test"
#Region "#CellFormatting"
            ' Access the "B2" cell.
            Dim cell As Cell = worksheet.Cells("B2")
            ' Specify font settings (font name, color, size and style).
            cell.Font.Name = "MV Boli"
            cell.Font.Color = Color.Blue
            cell.Font.Size = 14
            cell.Font.FontStyle = SpreadsheetFontStyle.Bold
            ' Specify the cell background color.
            cell.Fill.BackgroundColor = Color.LightSkyBlue
            ' Specify text alignment. 
            cell.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
            cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
#End Region  ' #CellFormatting
#Region "#RangeFormatting"
            ' Access the "C3:E6" cell range.
            Dim range As CellRange = worksheet.Range("C3:E6")
            ' Start to update the cell range formatting. 
            Dim rangeFormatting As Formatting = range.BeginUpdateFormatting()
            ' Specify font settings (font name, color, size and style).
            rangeFormatting.Font.Name = "MV Boli"
            rangeFormatting.Font.Color = Color.Blue
            rangeFormatting.Font.Size = 14
            rangeFormatting.Font.FontStyle = SpreadsheetFontStyle.Bold
            ' Specify the cell background color.
            rangeFormatting.Fill.BackgroundColor = Color.LightSkyBlue
            ' Specify text alignment.
            rangeFormatting.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
            rangeFormatting.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            ' Finalize to update the cell range formatting.
            range.EndUpdateFormatting(rangeFormatting)
#End Region  ' #RangeFormatting
        End Sub

        Private Sub SetDateFormats(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            worksheet.Range("A1:F1").ColumnWidthInCharacters = 15
            worksheet.Range("A1:F1").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
#Region "#DateTimeFormats"
            worksheet.Range("A1:F1").Formula = "= Now()"
            ' Apply different date display formats.
            worksheet.Cells("A1").NumberFormat = "m/d/yy"
            worksheet.Cells("B1").NumberFormat = "d-mmm-yy"
            worksheet.Cells("C1").NumberFormat = "dddd"
            ' Apply different time display formats.
            worksheet.Cells("D1").NumberFormat = "m/d/yy h:mm"
            worksheet.Cells("E1").NumberFormat = "h:mm AM/PM"
            worksheet.Cells("F1").NumberFormat = "h:mm:ss"
#End Region  ' #DateTimeFormats
        End Sub

        Private Sub SetNumberFormats(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            worksheet.Range("A1:H1").ColumnWidthInCharacters = 12
            worksheet.Range("A1:H1").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
#Region "#NumberFormats"
            ' Display 111 as 111.
            worksheet.Cells("A1").Value = 111
            worksheet.Cells("A1").NumberFormat = "#####"
            ' Display 222 as 00222.
            worksheet.Cells("B1").Value = 222
            worksheet.Cells("B1").NumberFormat = "00000"
            ' Display 12345678 as 12,345,678.
            worksheet.Cells("C1").Value = 12345678
            worksheet.Cells("C1").NumberFormat = "#,#"
            ' Display .126 as 0.13.
            worksheet.Cells("D1").Value = .126
            worksheet.Cells("D1").NumberFormat = "0.##"
            ' Display 74.4 as 74.400.
            worksheet.Cells("E1").Value = 74.4
            worksheet.Cells("E1").NumberFormat = "##.000"
            ' Display 1.6 as 160.0%.
            worksheet.Cells("F1").Value = 1.6
            worksheet.Cells("F1").NumberFormat = "0.0%"
            ' Display 4321 as $4,321.00.
            worksheet.Cells("G1").Value = 4321
            worksheet.Cells("G1").NumberFormat = "$#,##0.00"
            ' Display 8.75 as 8 3/4.
            worksheet.Cells("H1").Value = 8.75
            worksheet.Cells("H1").NumberFormat = "# ?/?"
#End Region  ' #NumberFormats
        End Sub

        Private Sub ChangeCellColors(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            worksheet.Range("C3:D4").Merge()
            worksheet.Range("C3:D4").Value = "Test"
            worksheet.Cells("A1").Value = "Test"
#Region "#ColorCells"
            ' Format an individual cell.
            worksheet.Cells("A1").Font.Color = Color.Red
            worksheet.Cells("A1").FillColor = Color.Yellow
            ' Format a cell range.
            Dim range As CellRange = worksheet.Range("C3:D4")
            Dim rangeFormatting As Formatting = range.BeginUpdateFormatting()
            rangeFormatting.Font.Color = Color.Blue
            rangeFormatting.Fill.BackgroundColor = Color.LightBlue
            rangeFormatting.Fill.PatternType = PatternType.LightHorizontal
            rangeFormatting.Fill.PatternColor = Color.Violet
            range.EndUpdateFormatting(rangeFormatting)
#End Region  ' #ColorCells
        End Sub

        Private Sub SpecifyCellFont(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            worksheet.Cells("A1").Value = "Font Attributes"
            worksheet.Cells("A1").ColumnWidthInCharacters = 20
#Region "#FontSettings"
            ' Access the Font object.
            Dim cellFont As SpreadsheetFont = worksheet.Cells("A1").Font
            ' Specify the font name.
            cellFont.Name = "Times New Roman"
            ' Specify the font size.
            cellFont.Size = 14
            ' Specify the font color.
            cellFont.Color = Color.Blue
            ' Format text as bold.
            cellFont.Bold = True
            ' Specify the font underline type.
            cellFont.UnderlineType = UnderlineType.Double
#End Region  ' #FontSettings
        End Sub

        Private Sub AlignCellContents(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            Dim range As CellRange = worksheet.Range("A1:B3")
            range.ColumnWidthInCharacters = 30
            range.RowHeight = 200
#Region "#AlignCellContents"
            ' Align the "A1" cell content.
            Dim cellA1 As Cell = worksheet.Cells("A1")
            cellA1.Value = "Right and top"
            cellA1.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right
            cellA1.Alignment.Vertical = SpreadsheetVerticalAlignment.Top
            ' Align the "A2" cell content.
            Dim cellA2 As Cell = worksheet.Cells("A2")
            cellA2.Value = "Center"
            cellA2.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            cellA2.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
            ' Align the "A3" cell content.
            Dim cellA3 As Cell = worksheet.Cells("A3")
            cellA3.Value = "Left and bottom, indent"
            cellA3.Alignment.Indent = 1
            ' Align the "B1" cell content.
            Dim cellB1 As Cell = worksheet.Cells("B1")
            cellB1.Value = "The Alignment.ShrinkToFit property is applied"
            cellB1.Alignment.ShrinkToFit = True
            ' Align the "B2" cell content.
            Dim cellB2 As Cell = worksheet.Cells("B2")
            cellB2.Value = "Rotated Cell Contents"
            cellB2.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            cellB2.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
            cellB2.Alignment.RotationAngle = 15
            ' Align the "B3" cell content.
            Dim cellB3 As Cell = worksheet.Cells("B3")
            cellB3.Value = "The Alignment.WrapText property is applied to wrap the text within a cell"
            cellB3.Alignment.WrapText = True
#End Region  ' #AlignCellContents
        End Sub

        Private Sub AddCellBorders(ByVal workbook As Workbook)
#Region "#CellBorders"
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Specify borders for the "B2" cell.
            Dim cellB2 As Cell = worksheet.Cells("B2")
            Dim cellB2Borders As Borders = cellB2.Borders
            cellB2Borders.LeftBorder.LineStyle = BorderLineStyle.MediumDashDot
            cellB2Borders.LeftBorder.Color = Color.Pink
            cellB2Borders.TopBorder.LineStyle = BorderLineStyle.MediumDashDotDot
            cellB2Borders.TopBorder.Color = Color.HotPink
            cellB2Borders.RightBorder.LineStyle = BorderLineStyle.MediumDashed
            cellB2Borders.RightBorder.Color = Color.DeepPink
            cellB2Borders.BottomBorder.LineStyle = BorderLineStyle.Medium
            cellB2Borders.BottomBorder.Color = Color.Red
            cellB2Borders.DiagonalBorderType = DiagonalBorderType.Up
            cellB2Borders.DiagonalBorderLineStyle = BorderLineStyle.Thick
            cellB2Borders.DiagonalBorderColor = Color.Red
            ' Specify diagonal borders for the "C4" cell.
            Dim cellC4 As Cell = worksheet.Cells("C4")
            Dim cellC4Borders As Borders = cellC4.Borders
            cellC4Borders.SetDiagonalBorders(Color.Orange, BorderLineStyle.Double, DiagonalBorderType.UpAndDown)
            ' Specify outside borders for the "D6" cell. 
            Dim cellD6 As Cell = worksheet.Cells("D6")
            cellD6.Borders.SetOutsideBorders(Color.Gold, BorderLineStyle.Double)
#End Region  ' #CellBorders
#Region "#CellRangeBorders"
            ' Specify all borders for the "B8:F13" cell range.
            Dim range1 As CellRange = worksheet.Range("B8:F13")
            range1.Borders.SetAllBorders(Color.Green, BorderLineStyle.Double)
            ' Specify inside and outside borders for the "C15:F18" cell range.
            Dim range2 As CellRange = worksheet.Range("C15:F18")
            range2.SetInsideBorders(Color.SkyBlue, BorderLineStyle.MediumDashed)
            range2.Borders.SetOutsideBorders(Color.DeepSkyBlue, BorderLineStyle.Medium)
            ' Specify horizontal and vertical borders for the "D21:F23" cell range.
            Dim range3 As CellRange = worksheet.Range("D21:F23")
            Dim range3Formatting As Formatting = range3.BeginUpdateFormatting()
            Dim range3Borders As Borders = range3Formatting.Borders
            range3Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.MediumDashDot
            range3Borders.InsideHorizontalBorders.Color = Color.DarkBlue
            range3Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.MediumDashDotDot
            range3Borders.InsideVerticalBorders.Color = Color.Blue
            range3.EndUpdateFormatting(range3Formatting)
            ' Specify borders for the "E25:F26" cell range. 
            Dim range4 As CellRange = worksheet.Range("E25:F26")
            Dim range4Formatting As Formatting = range4.BeginUpdateFormatting()
            Dim range4Borders As Borders = range4Formatting.Borders
            range4Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thick)
            range4Borders.LeftBorder.Color = Color.Violet
            range4Borders.TopBorder.Color = Color.Violet
            range4Borders.RightBorder.Color = Color.DarkViolet
            range4Borders.BottomBorder.Color = Color.DarkViolet
            range4Borders.DiagonalBorderType = DiagonalBorderType.UpAndDown
            range4Borders.DiagonalBorderLineStyle = BorderLineStyle.MediumDashed
            range4Borders.DiagonalBorderColor = Color.BlueViolet
            range4.EndUpdateFormatting(range4Formatting)
#End Region  ' #CellRangeBorders
        End Sub
    End Module
End Namespace
