Imports System
Imports DevExpress.Spreadsheet
Imports System.Drawing

Namespace SpreadsheetExamples
    Public NotInheritable Class RowAndColumnActions

        Private Sub New()
        End Sub

        #Region "Actions"
        Public Shared InsertRowsColumnsAction As Action(Of Workbook) = AddressOf InsertRowsColumns
        Public Shared DeleteRowsColumnsAction As Action(Of Workbook) = AddressOf DeleteRowsColumns
        Public Shared CopyRowsColumnsAction As Action(Of Workbook) = AddressOf CopyRowsColumns
        Public Shared ShowHideRowsColumnsAction As Action(Of Workbook) = AddressOf ShowHideRowsColumns
        Public Shared SpecifyRowsHeightColumnsWidthAction As Action(Of Workbook) = AddressOf SpecifyRowsHeightColumnsWidth
        Public Shared GroupRowsColumnsAction As Action(Of Workbook) = AddressOf GroupRowsColumns
        #End Region

        Private Shared Sub InsertRowsColumns(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Populate cells with data.
            For i As Integer = 0 To 9
                worksheet.Cells(i, 0).Value = i + 1
                worksheet.Cells(0, i).Value = i + 1
            Next i

'            #Region "#InsertRows"
            ' Insert a new row 3.
            worksheet.Rows("3").Insert()

            ' Insert a new row into the worksheet at the 5the position.
            worksheet.Rows.Insert(4)

            ' Insert five rows into the worksheet at the 9th position.
            worksheet.Rows.Insert(8, 5)

            ' Insert two rows above the "L15:M16" cell range.
            worksheet.InsertCells(worksheet.Range("L15:M16"), InsertCellsMode.EntireRow)
'            #End Region ' #InsertRows

'            #Region "#InsertColumns"
            ' Insert a new column C.
            worksheet.Columns("C").Insert()

            ' Insert a new column into the worksheet at the 5th position.
            worksheet.Columns.Insert(4)

            ' Insert three columns into the worksheet at the 7th position.
            worksheet.Columns.Insert(6, 3)

            ' Insert two columns to the left of the "L15:M16" cell range.
            worksheet.InsertCells(worksheet.Range("L15:M16"), InsertCellsMode.EntireColumn)
'            #End Region ' #InsertColumns
        End Sub

        Private Shared Sub DeleteRowsColumns(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets("Sheet1")

            ' Fill cells with data.
            For i As Integer = 0 To 14
                worksheet.Cells(i, 0).Value = i + 1
                worksheet.Cells(0, i).Value = i + 1
            Next i

'            #Region "#DeleteRows"
            ' Delete the 2nd row from the worksheet.
            worksheet.Rows(1).Delete()

            ' Delete the 3rd row from the worksheet.
            worksheet.Rows.Remove(2)

            ' Delete three rows from the worksheet starting from the 10th row.
            worksheet.Rows.Remove(9, 3)

            ' Delete a row that contains the "B2"cell.
            worksheet.DeleteCells(worksheet.Cells("B2"), DeleteMode.EntireRow)
'            #End Region ' #DeleteRows

'            #Region "#DeleteColumns"
            ' Delete the 2nd column from the worksheet.
            worksheet.Columns(1).Delete()

            ' Delete the 3rd column from the worksheet.
            worksheet.Columns.Remove(2)

            ' Delete three columns from the worksheet starting from the 10th column.
            worksheet.Columns.Remove(9, 3)

            ' Delete a column that contains the "B2"cell.
            worksheet.DeleteCells(worksheet.Cells("B2"), DeleteMode.EntireColumn)
'            #End Region ' #DeleteColumns
        End Sub

        Private Shared Sub CopyRowsColumns(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Modify the 2nd row. 
            worksheet.Cells("A2").Value = "Row 2"
            worksheet.Rows("2").Height = 150
            worksheet.Rows("2").Alignment.Vertical = SpreadsheetVerticalAlignment.Center
            worksheet.Rows("2").FillColor = Color.LightCyan

            ' Modify the "B" column.
            worksheet.Cells("B1").Value = "ColumnB"
            worksheet.Columns("B").Borders.SetOutsideBorders(Color.CadetBlue, BorderLineStyle.Thick)

'            #Region "#CopyRowsColumns"
            ' Copy all data from the 2nd row to the 5th row.
            worksheet.Rows("5").CopyFrom(worksheet.Rows("2"))

            ' Copy only borders from the "B" column to the "E" column.
            worksheet.Columns("E").CopyFrom(worksheet.Columns("B"), PasteSpecial.Borders)
'            #End Region ' #CopyRowsColumns
        End Sub

        Private Shared Sub ShowHideRowsColumns(ByVal workbook As Workbook)
'            #Region "#ShowHideRowsColumns"
            Dim worksheet As Worksheet = workbook.Worksheets(0)

            ' Hide the 8th row of the worksheet.
            worksheet.Rows(7).Visible = False

            ' Hide the 4th column of the worksheet.
            worksheet.Columns(3).Visible = False
'            #End Region ' #ShowHideRowsColumns
        End Sub

        Private Shared Sub SpecifyRowsHeightColumnsWidth(ByVal workbook As Workbook)

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

'            #Region "#RowHeight"
            ' Set the height of the 3rd row to 50 points.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point
            worksheet.Rows(2).Height = 50

            ' Set the height of the row that contains the "C5" cell to 2 inches.
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch
            worksheet.Cells("C5").RowHeight = 2

            ' Set the height of the 7th row to the height of the 3rd row.
            worksheet.Rows("7").Height = worksheet.Rows("3").Height

            ' Set the default row height to 30 points.
            workbook.Unit = DevExpress.Office.DocumentUnit.Point
            worksheet.DefaultRowHeight = 30
'            #End Region ' #RowHeight

'            #Region "#ColumnWidth"
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
'            #End Region ' #ColumnWidth

        End Sub

        Private Shared Sub GroupRowsColumns(ByVal workbook As Workbook)

            Dim worksheet As Worksheet = workbook.Worksheets(0)

'            #Region "#GroupRows"
            ' Group ten rows (from the second row to the eleventh row).
            worksheet.Rows.Group(1, 10, False)
'            #End Region ' #GroupRows

'            #Region "#GroupColumns"
            ' Group eight columns (from the third column to the tenth column).
            worksheet.Columns.Group(2, 9, True)
'            #End Region ' #GroupColumns
        End Sub
    End Class
End Namespace
