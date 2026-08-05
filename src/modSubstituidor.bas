Attribute VB_Name = "modSubstituidor"
Option Explicit

Public Sub GerarDocumentosPreenchidos(ByVal caminhos As Collection, ByVal valores As Object, ByVal exportarPDF As Boolean)
    On Error GoTo Falhou
    Dim caminho As Variant, documento As Document, destino As String, fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    For Each caminho In caminhos
        Set documento = Documents.Open(CStr(caminho), ReadOnly:=False, AddToRecentFiles:=False, Visible:=False)
        SubstituirNoDocumento documento, valores
        destino = CaminhoBase() & "\" & PASTA_GERADOS & "\" & fso.GetFileName(CStr(caminho))
        documento.SaveAs2 destino, wdFormatXMLDocument
        If exportarPDF Then ExportarDocumentoPDF documento, destino
        documento.Close wdDoNotSaveChanges
        RegistrarLog "Documento gerado: " & destino
    Next
    Exit Sub
Falhou:
    If Not documento Is Nothing Then documento.Close wdDoNotSaveChanges
    MostrarErro "Falha ao substituir marcadores", Err.Number, Err.Description: Err.Raise Err.Number
End Sub

Private Sub SubstituirNoDocumento(ByVal documento As Document, ByVal valores As Object)
    Dim historia As Range, intervalo As Range, forma As Shape, comentario As Comment
    For Each historia In documento.StoryRanges
        Set intervalo = historia
        Do While Not intervalo Is Nothing
            SubstituirNoRange intervalo, valores
            Set intervalo = intervalo.NextStoryRange
        Loop
    Next
    For Each forma In documento.Shapes: SubstituirNoShape forma, valores: Next
    For Each comentario In documento.Comments: SubstituirNoRange comentario.Range, valores: Next
End Sub

Private Sub SubstituirNoRange(ByVal intervalo As Range, ByVal valores As Object)
    Dim chave As Variant
    For Each chave In valores.Keys
        With intervalo.Duplicate.Find
            .ClearFormatting: .Replacement.ClearFormatting
            .Text = "{{" & CStr(chave) & "}}": .Replacement.Text = CStr(valores(chave))
            .Forward = True: .Wrap = wdFindStop: .Format = False: .MatchCase = False: .MatchWildcards = False
            .Execute Replace:=wdReplaceAll
        End With
        RegistrarLog "Campo substituído: " & CStr(chave)
    Next
End Sub

Private Sub SubstituirNoShape(ByVal forma As Shape, ByVal valores As Object)
    On Error Resume Next
    If forma.TextFrame.HasText Then SubstituirNoRange forma.TextFrame.TextRange, valores
End Sub
