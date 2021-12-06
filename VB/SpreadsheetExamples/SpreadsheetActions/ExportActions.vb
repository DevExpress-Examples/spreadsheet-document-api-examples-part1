Imports System
Imports System.IO
Imports System.Drawing
Imports System.Diagnostics
Imports DevExpress.Spreadsheet

Namespace SpreadsheetExamples

    Public Module ExportActions

#Region "Actions"
        Public ExportToPdfAction As Action(Of Workbook) = AddressOf ExportToPdf

#End Region
        Private Sub ExportToPdf(ByVal workbook As Workbook)
            workbook.Worksheets(0).Cells("D8").Value = "This document is exported to the PDF format."
#Region "#ExportToPdf"
            ' Export the workbook to PDF.
            Using pdfFileStream As FileStream = New FileStream("Documents\Document_PDF.pdf", FileMode.Create)
                workbook.ExportToPdf(pdfFileStream)
            End Using

#End Region  ' #ExportToPdf
            Call Process.Start("Documents\Document_PDF.pdf")
        End Sub
    End Module
End Namespace
