Attribute VB_Name = "modExportador"
Option Explicit

Public Sub ExportarDocumentoPDF(ByVal documento As Document, ByVal caminhoDocx As String)
    On Error GoTo Falhou
    Dim pdf As String: pdf = Left$(caminhoDocx, InStrRev(caminhoDocx, ".")) & "pdf"
    documento.ExportAsFixedFormat pdf, wdExportFormatPDF, False, wdExportOptimizeForPrint, wdExportAllDocument, , , wdExportDocumentContent, True, True, wdExportCreateHeadingBookmarks, True, True, False
    RegistrarLog "PDF exportado: " & pdf
    Exit Sub
Falhou: MostrarErro "Não foi possível exportar o PDF", Err.Number, Err.Description: Err.Raise Err.Number
End Sub
