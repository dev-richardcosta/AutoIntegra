Attribute VB_Name = "modUtil"
Option Explicit

Public Function CriarDicionario() As Object
    Set CriarDicionario = CreateObject("Scripting.Dictionary")
    CriarDicionario.CompareMode = 1
End Function

Public Function CriarRegExpMarcadores() As Object
    Dim regex As Object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True: regex.IgnoreCase = True: regex.Pattern = PADRAO_MARCADOR
    Set CriarRegExpMarcadores = regex
End Function

Public Function CaminhoBase() As String
    If Len(ThisDocument.Path) = 0 Then Err.Raise vbObjectError + 100, NOME_PROJETO, "Salve o documento que contém a macro antes de executar o gerador."
    CaminhoBase = ThisDocument.Path
End Function

Public Function FormatarRotulo(ByVal nomeCampo As String) As String
    FormatarRotulo = Replace$(StrConv(LCase$(nomeCampo), vbProperCase), "_", " ")
End Function

Public Sub RegistrarLog(ByVal mensagem As String)
    On Error Resume Next
    Dim numero As Integer: numero = FreeFile
    Open CaminhoBase() & "\" & PASTA_LOGS & "\processamento.log" For Append As #numero
    Print #numero, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " | " & mensagem
    Close #numero
End Sub

Public Sub MostrarErro(ByVal contexto As String, ByVal numero As Long, ByVal descricao As String)
    RegistrarLog "ERRO [" & contexto & "] " & CStr(numero) & " - " & descricao
    MsgBox contexto & ":" & vbCrLf & descricao, vbExclamation, NOME_PROJETO
End Sub
