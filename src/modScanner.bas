Attribute VB_Name = "modScanner"
Option Explicit

Public Function LocalizarMarcadores(ByVal caminhos As Collection) As Object
    On Error GoTo Falhou
    Dim campos As Object, caminho As Variant, documento As Document
    Set campos = CriarDicionario()
    For Each caminho In caminhos
        Set documento = Documents.Open(CStr(caminho), ReadOnly:=True, AddToRecentFiles:=False, Visible:=False)
        EscanearDocumento documento, campos
        documento.Close wdDoNotSaveChanges
        RegistrarLog "Documento escaneado: " & CStr(caminho)
    Next
    Set LocalizarMarcadores = campos
    Exit Function
Falhou:
    If Not documento Is Nothing Then documento.Close wdDoNotSaveChanges
    MostrarErro "Falha durante o escaneamento", Err.Number, Err.Description: Err.Raise Err.Number
End Function

Private Sub EscanearDocumento(ByVal documento As Document, ByVal campos As Object)
    Dim historia As Range, intervalo As Range, forma As Shape, comentario As Comment
    For Each historia In documento.StoryRanges
        Set intervalo = historia
        Do While Not intervalo Is Nothing
            EscanearTexto intervalo.Text, campos, documento.Name
            Set intervalo = intervalo.NextStoryRange
        Loop
    Next
    For Each forma In documento.Shapes: EscanearShape forma, campos, documento.Name: Next
    For Each comentario In documento.Comments: EscanearTexto comentario.Range.Text, campos, documento.Name: Next
End Sub

Private Sub EscanearShape(ByVal forma As Shape, ByVal campos As Object, ByVal nomeDocumento As String)
    On Error Resume Next
    If forma.TextFrame.HasText Then EscanearTexto forma.TextFrame.TextRange.Text, campos, nomeDocumento
    If forma.TextFrame2.HasText Then EscanearTexto forma.TextFrame2.TextRange.Text, campos, nomeDocumento
End Sub

Private Sub EscanearTexto(ByVal texto As String, ByVal campos As Object, ByVal nomeDocumento As String)
    Dim ocorrencia As Object, regex As Object
    Set regex = CriarRegExpMarcadores()
    For Each ocorrencia In regex.Execute(texto)
        If Not campos.Exists(Trim$(CStr(ocorrencia.SubMatches(0)))) Then
            campos.Add Trim$(CStr(ocorrencia.SubMatches(0))), vbNullString
            RegistrarLog "Campo encontrado em " & nomeDocumento & ": " & CStr(ocorrencia.SubMatches(0))
        End If
    Next
End Sub
