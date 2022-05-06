Imports System
Imports System.Drawing
Imports DevExpress.Spreadsheet

Namespace SpreadsheetExamples

    Public Module WorksheetActions

#Region "Actions"
        Public AssignActiveWorksheetAction As Action(Of Workbook) = AddressOf AssignActiveWorksheet

        Public AddWorksheetAction As Action(Of Workbook) = AddressOf AddWorksheet

        Public RemoveWorksheetAction As Action(Of Workbook) = AddressOf RemoveWorksheet

        Public RenameWorksheetAction As Action(Of Workbook) = AddressOf RenameWorksheet

        Public CopyWorksheetWithinWorkbookAction As Action(Of Workbook) = AddressOf CopyWorksheetWithinWorkbook

        Public CopyWorksheetBetweenWorkbooksAction As Action(Of Workbook) = AddressOf CopyWorksheetBetweenWorkbooks

        Public MoveWorksheetAction As Action(Of Workbook) = AddressOf MoveWorksheet

        Public ShowHideWorksheetAction As Action(Of Workbook) = AddressOf ShowHideWorksheet

        Public ShowHideGridlinesAction As Action(Of Workbook) = AddressOf ShowHideGridlines

        Public ShowHideHeadingsAction As Action(Of Workbook) = AddressOf ShowHideHeadings

        Public PageSetupAction As Action(Of Workbook) = AddressOf PageSetup

        Public ZoomWorksheetAction As Action(Of Workbook) = AddressOf ZoomWorksheet

#End Region
        Private Sub AssignActiveWorksheet(ByVal workbook As Workbook)
#Region "#ActiveWorksheet"
            ' Set the second worksheet under the "Sheet2" name as active.
            workbook.Worksheets.ActiveWorksheet = workbook.Worksheets("Sheet2")
#End Region  ' #ActiveWorksheet
        End Sub

        Private Sub AddWorksheet(ByVal workbook As Workbook)
#Region "#AddWorksheet"
            ' Add a new worksheet to the workbook. The worksheet is appended to the end of the worksheet collection
            ' under the name "SheetN", where N is a number that is greater by 1
            ' than the maximum number used in worksheet names of the same type.
            workbook.Worksheets.Add()
            ' Add a new worksheet under the specified name.
            workbook.Worksheets.Add().Name = "TestSheet1"
            workbook.Worksheets.Add("TestSheet2")
            ' Add a new worksheet to the specified position in the worksheet collection.
            workbook.Worksheets.Insert(1, "TestSheet3")
            workbook.Worksheets.Insert(3)
#End Region  ' #AddWorksheet
        End Sub

        Private Sub RemoveWorksheet(ByVal workbook As Workbook)
#Region "#DeleteWorksheet"
            ' Delete the "Sheet2" worksheet.
            workbook.Worksheets.Remove(workbook.Worksheets("Sheet2"))
            ' Delete the first worksheet.
            workbook.Worksheets.RemoveAt(0)
#End Region  ' #DeleteWorksheet
        End Sub

        Private Sub RenameWorksheet(ByVal workbook As Workbook)
#Region "#RenameWorksheet"
            ' Rename the second worksheet.
            workbook.Worksheets(1).Name = "Renamed Sheet"
#End Region  ' #RenameWorksheet
        End Sub

        Private Sub CopyWorksheetWithinWorkbook(ByVal workbook As Workbook)
            workbook.Worksheets("Sheet1").Cells.FillColor = Color.LightSteelBlue
            workbook.Worksheets("Sheet1").Cells("A1").ColumnWidthInCharacters = 50
            workbook.Worksheets("Sheet1").Cells("A1").Value = "Sheet1's Content"
#Region "#CopyWorksheet"
            ' Add a new worksheet to a workbook.
            workbook.Worksheets.Add("Sheet1_Copy")
            ' Copy all information (content and formatting) to the newly created worksheet 
            ' from the "Sheet1" worksheet.
            workbook.Worksheets("Sheet1_Copy").CopyFrom(workbook.Worksheets("Sheet1"))
#End Region  ' #CopyWorksheet
        End Sub

        Private Sub CopyWorksheetBetweenWorkbooks(ByVal workbook As Workbook)
#Region "#CopyWorksheetsBetweenWorkbooks"
            ' Create a source workbook.
            Dim sourceWorkbook As Workbook = New Workbook()
            ' Add a new worksheet.
            sourceWorkbook.Worksheets.Add()
            ' Modify the second worksheet of the source workbook.
            sourceWorkbook.Worksheets(1).Cells("A1").Value = "A worksheet to be copied"
            sourceWorkbook.Worksheets(1).Cells("A1").Font.Color = Color.ForestGreen
            ' Copy the second worksheet of the source workbook into the first worksheet of another workbook.
            workbook.Worksheets(0).CopyFrom(sourceWorkbook.Worksheets(1))
#End Region  ' #CopyWorksheetsBetweenWorkbooks
        End Sub

        Private Sub MoveWorksheet(ByVal workbook As Workbook)
#Region "#MoveWorksheet"
            ' Move the first worksheet to the position of the last worksheet within the workbook.
            workbook.Worksheets(0).Move(workbook.Worksheets.Count - 1)
#End Region  ' #MoveWorksheet
        End Sub

        Private Sub ShowHideWorksheet(ByVal workbook As Workbook)
#Region "#ShowHideWorksheet"
            ' Hide the "Sheet2" worksheet and disable access to this worksheet in the user interface.
            ' Use the Worksheet.Visible property to unhide this worksheet.
            workbook.Worksheets("Sheet2").VisibilityType = WorksheetVisibilityType.VeryHidden
            ' Hide the "Sheet3" worksheet. 
            ' You can unhide this worksheet from the user interface.
            workbook.Worksheets("Sheet3").Visible = False
#End Region  ' #ShowHideWorksheet
        End Sub

        Private Sub ShowHideGridlines(ByVal workbook As Workbook)
#Region "#ShowHideGridlines"
            ' Hide gridlines on the first worksheet.
            workbook.Worksheets(0).ActiveView.ShowGridlines = False
#End Region  ' #ShowHideGridlines
        End Sub

        Private Sub ShowHideHeadings(ByVal workbook As Workbook)
#Region "#ShowHideHeadings"
            ' Hide row and column headings in the first worksheet.
            workbook.Worksheets(0).ActiveView.ShowHeadings = False
#End Region  ' #ShowHideHeadings
        End Sub

        Private Sub PageSetup(ByVal workbook As Workbook)
#Region "#ViewType"
            ' Select the worksheet view type.
            workbook.Worksheets(0).ActiveView.ViewType = WorksheetViewType.PageLayout
#End Region  ' #ViewType
#Region "#PageOrientation"
            ' Set the page orientation to Landscape.
            workbook.Worksheets(0).ActiveView.Orientation = PageOrientation.Landscape
#End Region  ' #PageOrientation
#Region "#PageMargins"
            ' Specifies inches as the workbook's measurement units.
            workbook.Unit = DevExpress.Office.DocumentUnit.Inch
            ' Access page margins.
            Dim pageMargins As Margins = workbook.Worksheets(0).ActiveView.Margins
            ' Specify page margins.
            pageMargins.Left = 1
            pageMargins.Top = 1.5F
            pageMargins.Right = 1
            pageMargins.Bottom = 0.8F
            ' Specify header and footer margins.
            pageMargins.Header = 1
            pageMargins.Footer = 0.4F
#End Region  ' #PageMargins
#Region "#PaperSize"
            ' Select the page's paper size.
            workbook.Worksheets(0).ActiveView.PaperKind = System.Drawing.Printing.PaperKind.A4
#End Region  ' #PaperSize
        End Sub

        Private Sub ZoomWorksheet(ByVal workbook As Workbook)
#Region "#WorksheetZoom"
            ' Zoom in the worksheet view. 
            workbook.Worksheets(0).ActiveView.Zoom = 150
#End Region  ' #WorksheetZoom
        End Sub
    End Module
End Namespace
