Imports System
Imports System.Drawing
#Region "#printingUsings"
Imports DevExpress.Spreadsheet
Imports DevExpress.XtraPrinting

#End Region  ' #printingUsings
Namespace SpreadsheetExamples

    Public Module PrintingActions

        Public PrintAction As Action(Of Workbook) = AddressOf Print

        Private Sub Print(ByVal workbook As Workbook)
            Dim worksheet As Worksheet = workbook.Worksheets(0)
            ' Generate a simple multiplication table.
            Dim topHeader As CellRange = worksheet.Range.FromLTRB(1, 0, 20, 0)
            topHeader.Formula = "=COLUMN() - 1"
            Dim leftCaption As CellRange = worksheet.Range.FromLTRB(0, 1, 0, 20)
            leftCaption.Formula = "=ROW() - 1"
            Dim tableRange As CellRange = worksheet.Range.FromLTRB(1, 1, 20, 20)
            tableRange.Formula = "=(ROW()-1)*(COLUMN()-1)"
            ' Format headers of the multiplication table.
            Dim rangeFormatting As Formatting = topHeader.BeginUpdateFormatting()
            rangeFormatting.Borders.BottomBorder.LineStyle = BorderLineStyle.Thin
            rangeFormatting.Borders.BottomBorder.Color = Color.Black
            topHeader.EndUpdateFormatting(rangeFormatting)
            rangeFormatting = leftCaption.BeginUpdateFormatting()
            rangeFormatting.Borders.RightBorder.LineStyle = BorderLineStyle.Thin
            rangeFormatting.Borders.RightBorder.Color = Color.Black
            leftCaption.EndUpdateFormatting(rangeFormatting)
            rangeFormatting = tableRange.BeginUpdateFormatting()
            rangeFormatting.Fill.BackgroundColor = Color.LightBlue
            tableRange.EndUpdateFormatting(rangeFormatting)
#Region "#WorksheetPrintOptions"
            worksheet.ActiveView.Orientation = PageOrientation.Landscape
            '  Display row and column headings.
            worksheet.ActiveView.ShowHeadings = True
            worksheet.ActiveView.PaperKind = System.Drawing.Printing.PaperKind.A4
            ' Access an object that contains print options.
            Dim printOptions As WorksheetPrintOptions = worksheet.PrintOptions
            '  Print in black and white.
            printOptions.BlackAndWhite = True
            '  Do not print gridlines.
            printOptions.PrintGridlines = False
            '  Scale the print area to fit to a page.
            printOptions.FitToPage = True
            '  Print a dash instead of a cell error message.
            printOptions.ErrorsPrintMode = ErrorsPrintMode.Dash
#End Region  ' #WorksheetPrintOptions
#Region "#PrintWorkbook"
            ' Invoke the Print Preview dialog for the workbook.
            Using printingSystem As PrintingSystem = New PrintingSystem()
                Using link As PrintableComponentLink = New PrintableComponentLink(printingSystem)
                    link.Component = workbook
                    link.CreateDocument()
                    link.ShowPreviewDialog()
                End Using
            End Using
#End Region  ' #PrintWorkbook
        End Sub
    End Module
End Namespace
