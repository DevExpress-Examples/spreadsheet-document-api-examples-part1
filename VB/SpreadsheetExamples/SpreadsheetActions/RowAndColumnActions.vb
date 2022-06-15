Imports System
Imports DevExpress.Spreadsheet
Imports System.Drawing

Namespace SpreadsheetExamples

    Public Module RowAndColumnActions

#Region "Actions"
        Public InsertRowsColumnsAction As Action(Of Workbook) = AddressOf InsertRowsColumns

        Public DeleteRowsColumnsAction As Action(Of Workbook) = AddressOf DeleteRowsColumns

        Public CopyRowsColumnsAction As Action(Of Workbook) = AddressOf CopyRowsColumns

        Public ShowHideRowsColumnsAction As Action(Of Workbook) = AddressOf ShowHideRowsColumns

        Public SpecifyRowsHeightColumnsWidthAction As Action(Of Workbook) = AddressOf SpecifyRowsHeightColumnsWidth

        Public GroupRowsColumnsAction As Action(Of Workbook) = AddressOf GroupRowsColumns

#End Region
        Private Sub InsertRowsColumns(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Populate cells with data.
            For i As Integer = 0 To 10 - 1
                worksheet.Cells(i, 0).Value = i + 1
                worksheet.Cells(0, i).Value = i + 1
            Next

#Region "#InsertRows"
            ' Insert the third row.
            worksheet.Rows("3").Insert()
            ' Insert the fifth row.
            worksheet.Rows.Insert(4)
            ' Insert five rows (from row 9 to row 13).
            worksheet.Rows.Insert(8, 5)
            ' Insert two rows above the "L15:M16" cell range.
            worksheet.InsertCells(worksheet.Range("L15:M16"), InsertCellsMode.EntireRow)
#End Region  ' #InsertRows
#Region "#InsertColumns"
            ' Insert column "C".
            worksheet.Columns("C").Insert()
            ' Insert column "E".
            worksheet.Columns.Insert(4)
            ' Insert three columns (from column "G" to column "I").
            worksheet.Columns.Insert(6, 3)
            ' Insert two columns to the left of the "L15:M16" cell range.
            worksheet.InsertCells(worksheet.Range("L15:M16"), InsertCellsMode.EntireColumn)
#End Region  ' #InsertColumns
        End Sub

        Private Sub DeleteRowsColumns(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets("Sheet1")
            ' Fill cells with data.
            For i As Integer = 0 To 15 - 1
                worksheet.Cells(i, 0).Value = i + 1
                worksheet.Cells(0, i).Value = i + 1
            Next

#Region "#DeleteRows"
            ' Delete the second row.
            worksheet.Rows(1).Delete()
            ' Delete the third row.
            worksheet.Rows.Remove(2)
            ' Delete three rows (from row 10 to row 12).
            worksheet.Rows.Remove(9, 3)
            ' Delete a row that contains the "B2" cell.
            worksheet.DeleteCells(worksheet.Cells("B2"), DeleteMode.EntireRow)
#End Region  ' #DeleteRows
#Region "#DeleteColumns"
            ' Delete the second column.
            worksheet.Columns(1).Delete()
            ' Delete the third column.
            worksheet.Columns.Remove(2)
            ' Delete three columns (from column "J" to column "L").
            worksheet.Columns.Remove(9, 3)
            ' Delete a column that contains the "B2" cell.
            worksheet.DeleteCells(worksheet.Cells("B2"), DeleteMode.EntireColumn)
#End Region  ' #DeleteColumns
        End Sub

        Private Sub CopyRowsColumns(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Modify the second row. 
            worksheet.Cells("A2").Value = "Row 2"
            worksheet.Rows("2").Height = 150
            worksheet.Rows("2").Alignment.Vertical = SpreadsheetVerticalAlignment.Center
            worksheet.Rows("2").FillColor = Color.LightCyan
            ' Modify the "B" column.
            worksheet.Cells("B1").Value = "ColumnB"
            worksheet.Columns("B").Borders.SetOutsideBorders(Color.CadetBlue, BorderLineStyle.Thick)
#Region "#CopyRowsColumns"
            ' Copy all data from the second row to the fifth row.
            worksheet.Rows("5").CopyFrom(worksheet.Rows("2"))
            ' Copy only borders from the "B" column to the "E" column.
            worksheet.Columns("E").CopyFrom(worksheet.Columns("B"), PasteSpecial.Borders)
#End Region  ' #CopyRowsColumns
        End Sub

        Private Sub ShowHideRowsColumns(ByVal workbook As Workbook)
#Region "#ShowHideRowsColumns"
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Hide the eighth row.
            worksheet.Rows(7).Visible = False
            ' Hide the fourth column.
            worksheet.Columns(3).Visible = False
            ' Hide columns from 5 to 7.
            worksheet.Columns.Hide(5, 7)
            ' Hide rows from 6 to 8.
            worksheet.Rows.Hide(5, 7)
            ' Hide the tenth row.
            worksheet.Rows(9).Height = 0
            ' Hide the tenth column.
            worksheet.Columns(9).Width = 0
#End Region  ' #ShowHideRowsColumns
        End Sub

        Private Sub SpecifyRowsHeightColumnsWidth(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            worksheet.Range("B1:J1").Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            worksheet.Cells("B1").Value = "30 characters"
            worksheet.Cells("C1").Value = "15 mm"
            worksheet.Cells("E1").Value = "100 points"
            worksheet.Cells("F1").Value = "70 points"
            worksheet.Cells("G1").Value = "70 points"
            worksheet.Cells("H1").Value = "70 points"
            worksheet.Cells("J1").Value = "30 characters"
            worksheet.Cells("K1").Value = "15 mm"
            worksheet.Cells("A3").Value = "50 points"
            worksheet.Cells("A5").Value = "2 inches"
            worksheet.Cells("A7").Value = "50 points"
            Dim rowHeightValues As Formatting = worksheet.Range("A3:A7").BeginUpdateFormatting()
            rowHeightValues.Alignment.RotationAngle = 90
            rowHeightValues.Alignment.Vertical = SpreadsheetVerticalAlignment.Center
            rowHeightValues.Alignment.Horizontal = SpreadsheetHorizontalAlignment.Center
            worksheet.Range("A3:A7").EndUpdateFormatting(rowHeightValues)
#Region "#RowHeight"
            ' Set the height of the third row to 50 points.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point
            worksheet.Rows(2).Height = 50
            ' Set the height of the row that contains the "C5" cell to 2 inches.
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch
            worksheet.Cells("C5").RowHeight = 2
            ' Set the height of the seventh row to the height of the third row.
            worksheet.Rows("7").Height = worksheet.Rows("3").Height
            ' Set the default row height to 30 points.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point
            worksheet.DefaultRowHeight = 30
#End Region  ' #RowHeight
#Region "#ColumnWidth"
            ' Set the "B" column width to 30 characters of the default font that is specified by the Normal style.
            worksheet.Columns("B").WidthInCharacters = 30
            ' Set the "C" column width to 15 millimeters.
            workbook.Unit = DevExpress.Office.DocumentUnit.Millimeter
            worksheet.Columns("C").Width = 15
            ' Set the width of the column that contains the "E15" cell to 100 points.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point
            worksheet.Cells("E15").ColumnWidth = 100
            ' Set the width of all columns that contain the "F4:H7" cell range (the "F", "G" and "H" columns) to 70 points.
            worksheet.Range("F4:H7").ColumnWidth = 70
            ' Set the "J" column width to the "B" column width value.
            worksheet.Columns("J").Width = worksheet.Columns("B").Width
            ' Copy the "C" column width value and assign it to the "K" column width.
            worksheet.Columns("K").CopyFrom(worksheet.Columns("C"), PasteSpecial.ColumnWidths)
            ' Set the default column width to 40 pixels.
            worksheet.DefaultColumnWidthInPixels = 40
#End Region  ' #ColumnWidth
        End Sub

        Private Sub GroupRowsColumns(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
#Region "#GroupRows"
            ' Group ten rows (from the second row to the eleventh row).
            worksheet.Rows.Group(1, 10, False)
#End Region  ' #GroupRows
#Region "#GroupColumns"
            ' Group eight columns (from the third column to the tenth column).
            worksheet.Columns.Group(2, 9, True)
#End Region  ' #GroupColumns
        End Sub
    End Module
End Namespace
