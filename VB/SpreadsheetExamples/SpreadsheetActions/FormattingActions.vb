Imports DevExpress.Spreadsheet
Imports System
Imports System.Drawing

Namespace SpreadsheetExamples

    Public Module FormattingActions

#Region "Actions"
        Public CreateModifyApplyStyleAction As Action(Of IWorkbook) = AddressOf CreateModifyApplyStyle

        Public FormatCellAction As Action(Of IWorkbook) = AddressOf FormatCell

        Public SetDateFormatsAction As Action(Of IWorkbook) = AddressOf SetDateFormats

        Public SetNumberFormatsAction As Action(Of IWorkbook) = AddressOf SetNumberFormats

        Public CustomNumberFormatAction As Action(Of IWorkbook) = AddressOf CustomNumberFormat

        Public ChangeCellColorsAction As Action(Of IWorkbook) = AddressOf ChangeCellColors

        Public ChangeCellGradientFillAction As Action(Of IWorkbook) = AddressOf ChangeCellGradientFill

        Public SpecifyCellFontAction As Action(Of IWorkbook) = AddressOf SpecifyCellFont

        Public AlignCellContentsAction As Action(Of IWorkbook) = AddressOf AlignCellContents

        Public AddCellBordersAction As Action(Of IWorkbook) = AddressOf AddCellBorders

#End Region
        Private Sub CreateModifyApplyStyle(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
#Region "#CreateNewStyle"
                ' Add a new style under the "My Style" name to the Styles collection of the workbook.
                Dim myStyle As Style = workbook.Styles.Add("My Style")
                ' Specify formatting characteristics for the style.
                myStyle.BeginUpdate()
                Try
                    ' Set the font color to Blue.
                    myStyle.Font.Color = Color.Blue
                    ' Set the font size to 12.
                    myStyle.Font.Size = 12
                    ' Set the horizontal alignment to Center.
                    myStyle.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                    ' Set the background.
                    myStyle.Fill.BackgroundColor = Color.LightBlue
                Finally
                    myStyle.EndUpdate()
                End Try

#End Region  ' #CreateNewStyle
#Region "#DuplicateExistingStyle"
                ' Add a new style under the "My Good Style" name to the Styles collection.
                Dim myGoodStyle As Style = workbook.Styles.Add("My Good Style")
                ' Copy all format settings from the built-in Good style.
                myGoodStyle.CopyFrom(BuiltInStyleId.Good)
#End Region  ' #DuplicateExistingStyle
#Region "#ModifyExistingStyle"
                ' Change the required formatting characteristics of the style.
                myGoodStyle.BeginUpdate()
                Try
                    ' ...
                    myGoodStyle.Fill.BackgroundColor = Color.LightYellow
                Finally
                    myGoodStyle.EndUpdate()
                End Try

#End Region  ' #ModifyExistingStyle
#Region "#ApplyStyles"
                ' Access the built-in "Good" MS Excel style from the Styles collection of the workbook.
                Dim styleGood As Style = workbook.Styles(BuiltInStyleId.Good)
                ' Apply the "Good" style to a range of cells.
                worksheet.Range("A1:C4").Style = styleGood
                ' Access a custom style that has been previously created in the loaded document by its name.
                Dim customStyle As Style = workbook.Styles("Custom Style")
                ' Apply the custom style to the cell.
                worksheet.Cells("D6").Style = customStyle
                ' Apply the "Good" style to the eighth row.
                worksheet.Rows(7).Style = styleGood
                ' Apply the custom style to the "H" column.
#End Region  ' #ApplyStyles
                worksheet.Columns("H").Style = customStyle
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Sub FormatCell(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                worksheet.Cells("B2").Value = "Test"
                worksheet.Range("C3:E6").Value = "Test"
#Region "#CellFormatting"
                ' Access the cell to be formatted.
                Dim cell As Cell = worksheet.Cells("B2")
                ' Specify font settings (font name, color, size and style).
                cell.Font.Name = "MV Boli"
                cell.Font.Color = Color.Blue
                cell.Font.Size = 14
                cell.Font.FontStyle = SpreadsheetFontStyle.Bold
                ' Specify cell background color.
                cell.Fill.BackgroundColor = Color.LightSkyBlue
                ' Specify text alignment in the cell. 
                cell.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                cell.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
#End Region  ' #CellFormatting
#Region "#RangeFormatting"
                ' Access the range of cells to be formatted.
                Dim range As CellRange = worksheet.Range("C3:E6")
                ' Begin updating of the range formatting. 
                Dim rangeFormatting As Formatting = range.BeginUpdateFormatting()
                ' Specify font settings (font name, color, size and style).
                rangeFormatting.Font.Name = "MV Boli"
                rangeFormatting.Font.Color = Color.Blue
                rangeFormatting.Font.Size = 14
                rangeFormatting.Font.FontStyle = SpreadsheetFontStyle.Bold
                ' Specify cell background color.
                rangeFormatting.Fill.BackgroundColor = Color.LightSkyBlue
                ' Specify text alignment in cells.
                rangeFormatting.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                rangeFormatting.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                ' End updating of the range formatting.
#End Region  ' #RangeFormatting
                range.EndUpdateFormatting(rangeFormatting)
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Sub SetDateFormats(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
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
#End Region  ' #DateTimeFormats
                worksheet.Cells("F1").NumberFormat = "h:mm:ss"
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Sub SetNumberFormats(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
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
#End Region  ' #NumberFormats
                worksheet.Cells("H1").NumberFormat = "# ?/?"
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Sub CustomNumberFormat(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                worksheet("A1").Value = "Number Format:"
                worksheet("A2").Value = "Values:"
                worksheet("A3").Value = "Formatted Values:"
                worksheet("A1:A3").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                worksheet("A1:E3").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Left
                worksheet("A1:E3").ColumnWidthInCharacters = 17
                worksheet.MergeCells(worksheet("B1:E1"))
                worksheet("B1:E1").Value = "[Green]#.00;[Red]#.00;[Blue]0.00;[Magenta]""product: ""@"
                worksheet("B1:E1").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                worksheet("B1:E3").Borders.SetAllBorders(Color.Black, BorderLineStyle.Thin)
#Region "#CustomNumberFormat"
                ' Set cell values.
                worksheet("B2:B3").Value = -15.50
                worksheet("C2:C3").Value = 555
                worksheet("D2:D3").Value = 0
                worksheet("E2:E3").Value = "Name"
                'Apply custom number format.
#End Region  ' #CustomNumberFormat
                worksheet("B3:E3").NumberFormat = "[Green]#.00;[Red]#.00;[Blue]0.00;[Magenta]""product: ""@"
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Sub ChangeCellColors(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                worksheet.Range("C3:H10").Merge()
                worksheet.Range("C3:H10").Value = "Test"
                worksheet.Range("C3:H10").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                worksheet.Range("C3:H10").Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                worksheet.Cells("A1").Value = "Test"
#Region "#ColorCells"
                ' Format an individual cell.
                worksheet.Cells("A1").Font.Color = Color.Red
                worksheet.Cells("A1").FillColor = Color.Yellow
                ' Format a range of cells.
                Dim range As CellRange = worksheet.Range("C3:H10")
                Dim rangeFormatting As Formatting = range.BeginUpdateFormatting()
                rangeFormatting.Font.Color = Color.Blue
                rangeFormatting.Fill.BackgroundColor = Color.LightBlue
                rangeFormatting.Fill.PatternType = PatternType.LightHorizontal
                rangeFormatting.Fill.PatternColor = Color.Violet
#End Region  ' #ColorCells
                range.EndUpdateFormatting(rangeFormatting)
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Sub ChangeCellGradientFill(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                worksheet.Range("C3:F8").Merge()
                worksheet.Range("C3:F8").Value = "Test"
                worksheet.Range("C3:F8").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                worksheet.Range("C3:F8").Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                worksheet.Range("C13:F18").Merge()
                worksheet.Range("C13:F18").Value = "Test"
                worksheet.Range("C13:F18").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                worksheet.Range("C13:F18").Alignment.Vertical = SpreadsheetVerticalAlignment.Center
#Region "#GradientLinear"
                ' Specify a linear gradient fill for a cell.
                Dim fillA1 As Fill = worksheet.Cells("A1").Fill
                fillA1.FillType = FillType.Gradient
                fillA1.Gradient.Type = GradientFillType.Linear
                ' Set the tilt for the gradient line in degrees.
                ' 90 degree angle defines a color transition from the top to the bottom.
                fillA1.Gradient.Degree = 90
                ' Specify two gradient colors. 
                ' The position of a color stop should be either 0 (start) or 1 (end).
                fillA1.Gradient.Stops.Add(0, Color.Yellow)
                fillA1.Gradient.Stops.Add(1, Color.SkyBlue)
                ' Specify a linear gradient fill for a range.
                worksheet.Range("C3:F8").Fill.FillType = FillType.Gradient
                Dim rangeGradient1 As GradientFill = worksheet.Range("C3:F8").Fill.Gradient
                rangeGradient1.Type = GradientFillType.Linear
                ' Set the tilt for the gradient line in degrees.
                ' 45 degree angle defines a color transition from top left to bottom right.
                rangeGradient1.Degree = 45
                ' Specify two gradient colors. 
                ' The position of a color stop should be either 0 (start) or 1 (end).
                Dim rangeStops1 As GradientStopCollection = rangeGradient1.Stops
                rangeStops1.Add(0, Color.BlanchedAlmond)
                rangeStops1.Add(1, Color.Blue)
#End Region  ' #GradientLinear
#Region "#GradientPath"
                ' Specify a path gradient fill for a cell.
                Dim fillB1 As Fill = worksheet.Cells("A3").Fill
                fillB1.FillType = FillType.Gradient
                Dim cellGradient As GradientFill = fillB1.Gradient
                cellGradient.Type = GradientFillType.Path
                ' Set the relative position of a convergence point.
                cellGradient.RectangleLeft = 0F
                cellGradient.RectangleRight = 0F
                cellGradient.RectangleTop = 0.5F
                cellGradient.RectangleBottom = 0.5F
                ' Specify two gradient colors. 
                ' The position of a color stop should be either 0 (start) or 1 (end).
                Dim cellStops2 As GradientStopCollection = cellGradient.Stops
                cellStops2.Add(0, Color.Yellow)
                cellStops2.Add(1, Color.Red)
                ' Specify a path gradient fill for a range.
                worksheet.Range("C13:F18").Fill.FillType = FillType.Gradient
                Dim rangeGradient2 As GradientFill = worksheet.Range("C13:f18").Fill.Gradient
                rangeGradient2.Type = GradientFillType.Path
                ' Set the relative position of a convergence point.
                rangeGradient2.RectangleLeft = 0.5F
                rangeGradient2.RectangleRight = 0.5F
                rangeGradient2.RectangleTop = 0.5F
                rangeGradient2.RectangleBottom = 0.5F
                ' Specify two gradient colors. 
                ' The position of a color stop should be either 0 (start) or 1 (end).
                Dim rangeStops2 As GradientStopCollection = rangeGradient2.Stops
                rangeStops2.Add(0, Color.Orange)
#End Region  ' #GradientPath
                rangeStops2.Add(1, Color.Green)
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Sub SpecifyCellFont(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                worksheet.Cells("A1").Value = "Font Attributes"
                worksheet.Cells("A1").ColumnWidthInCharacters = 20
#Region "#FontSettings"
                ' Access the Font object.
                Dim cellFont As SpreadsheetFont = worksheet.Cells("A1").Font
                ' Set the font name.
                cellFont.Name = "Times New Roman"
                ' Set the font size.
                cellFont.Size = 14
                ' Set the font color.
                cellFont.Color = Color.Blue
                ' Format text as bold.
                cellFont.Bold = True
                ' Set font to be underlined.
#End Region  ' #FontSettings
                cellFont.UnderlineType = UnderlineType.Single
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Sub AlignCellContents(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                Dim range As CellRange = worksheet.Range("A1:A4")
                range.ColumnWidthInCharacters = 30
                workbook.Unit = DevExpress.Office.DocumentUnit.Inch
                range.RowHeight = 1
#Region "#AlignCellContents"
                Dim cellA1 As Cell = worksheet.Cells("A1")
                cellA1.Value = "Right and top"
                cellA1.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Right
                cellA1.Alignment.Vertical = SpreadsheetVerticalAlignment.Top
                Dim cellA2 As Cell = worksheet.Cells("A2")
                cellA2.Value = "Center"
                cellA2.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
                cellA2.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
                Dim cellA3 As Cell = worksheet.Cells("A3")
                cellA3.Value = "Left and bottom, indent"
                cellA3.Alignment.Indent = 2
                Dim cellA4 As Cell = worksheet.Cells("A4")
                cellA4.Value = "The Alignment.WrapText property is applied to wrap the text within a cell"
                cellA4.Alignment.WrapText = True
                Dim cellA5 As Cell = worksheet.Cells("A5")
                cellA5.Value = "Rotation by 45 degrees"
#End Region  ' #AlignCellContents
                cellA5.Alignment.RotationAngle = 45
            Finally
                workbook.EndUpdate()
            End Try
        End Sub

        Private Sub AddCellBorders(ByVal workbook As IWorkbook)
            workbook.BeginUpdate()
            Try
#Region "#CellBorders"
                Dim worksheet As Worksheet = workbook.Worksheets(0)
                ' Set each particular border for the cell.
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
                ' Set all borders for cells.
                Dim cellC4 As Cell = worksheet.Cells("C4")
                cellC4.Borders.SetAllBorders(Color.Orange, BorderLineStyle.MediumDashed)
                Dim cellD6 As Cell = worksheet.Cells("D6")
                cellD6.Borders.SetOutsideBorders(Color.Gold, BorderLineStyle.Double)
#End Region  ' #CellBorders
#Region "#CellRangeBorders"
                ' Set all borders for the range of cells in one step.
                Dim range1 As CellRange = worksheet.Range("B8:F13")
                range1.Borders.SetAllBorders(Color.Green, BorderLineStyle.Double)
                ' Set all inside and outside borders separately for the range of cells.
                Dim range2 As CellRange = worksheet.Range("C15:F18")
                range2.SetInsideBorders(Color.SkyBlue, BorderLineStyle.MediumDashed)
                range2.Borders.SetOutsideBorders(Color.DeepSkyBlue, BorderLineStyle.Medium)
                ' Set all horizontal and vertical borders separately for the range of cells.
                Dim range3 As CellRange = worksheet.Range("D21:F23")
                Dim range3Formatting As Formatting = range3.BeginUpdateFormatting()
                Dim range3Borders As Borders = range3Formatting.Borders
                range3Borders.InsideHorizontalBorders.LineStyle = BorderLineStyle.MediumDashDot
                range3Borders.InsideHorizontalBorders.Color = Color.DarkBlue
                range3Borders.InsideVerticalBorders.LineStyle = BorderLineStyle.MediumDashDotDot
                range3Borders.InsideVerticalBorders.Color = Color.Blue
                range3.EndUpdateFormatting(range3Formatting)
                ' Set each particular border for the range of cell. 
                Dim range4 As CellRange = worksheet.Range("E25:F26")
                Dim range4Formatting As Formatting = range4.BeginUpdateFormatting()
                Dim range4Borders As Borders = range4Formatting.Borders
                range4Borders.SetOutsideBorders(Color.Black, BorderLineStyle.Thick)
                range4Borders.LeftBorder.Color = Color.Violet
                range4Borders.TopBorder.Color = Color.Violet
                range4Borders.RightBorder.Color = Color.DarkViolet
                range4Borders.BottomBorder.Color = Color.DarkViolet
#End Region  ' #CellRangeBorders
                range4.EndUpdateFormatting(range4Formatting)
            Finally
                workbook.EndUpdate()
            End Try
        End Sub
    End Module
End Namespace
