Imports System
Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Diagnostics
Imports System.Collections.Generic
Imports DevExpress.Spreadsheet

Namespace SpreadsheetExamples
    Public NotInheritable Class ExportActions

        Private Sub New()
        End Sub

        #Region "Actions"
        Public Shared ExportToPdfAction As Action(Of Workbook) = AddressOf ExportToPdf
        #End Region

        Private Shared Sub ExportToPdf(ByVal workbook As Workbook)
            workbook.Worksheets(0).Cells("D8").Value = "This document is exported to the PDF format."

'            #Region "#ExportToPdf"
            Using pdfFileStream As New FileStream("Documents\Document_PDF.pdf", FileMode.Create)
                workbook.ExportToPdf(pdfFileStream)
            End Using
'            #End Region ' #ExportToPdf
            Process.Start("Documents\Document_PDF.pdf")
        End Sub
    End Class
End Namespace
